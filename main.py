"""
========================================
動画の無音時間をカットする編集ツール
========================================

【このファイルでできること】
- 動画ファイルを選んで、無音部分を自動でカット
- カットした動画を別名で保存（元ファイルは上書きしない）
- 処理の様子をログで確認できる

【動作確認済み環境】
- Python 3.9 〜 3.14
- MoviePy 2.2.1
- pydub 0.25.1

【必要なライブラリ】
pip install moviepy pydub

（ドラッグ＆ドロップは追加インストール不要）

【FFmpegについて】
moviepy と pydub は内部で FFmpeg を使います。
FFmpeg を別途インストールしてください。
  → https://ffmpeg.org/download.html
  Windows の場合は PATH に追加するか、
  このファイルと同じフォルダに ffmpeg.exe を置いてください。

【ショートカットキー】
  Space / Enter  : 実行（ファイル選択済みのとき）
  O              : ファイルを開く
  R              : 出力動画を再生（完了後）
  E              : 出力フォルダをエクスプローラーで開く（完了後）
  N              : 次のファイルへ（完了後・クリア）
  Ctrl + L       : ログをクリア

【変更履歴】
v1.3.0 - 作業効率改善
  - ドラッグ＆ドロップを追加インストールなしで実現
      Windows: ctypes で IDropTarget を実装
      Mac/Linux: Tk の組み込み DND にフォールバック
  - 完了後に「▶ 再生」ボタン・「フォルダを開く」ボタンが出現
  - ショートカットキーを全面追加
  - 完了後に「次のファイルへ」でワンクリックでリセット
  - 最後に開いたフォルダを記憶（次回のファイルダイアログがそこから始まる）
  - メッセージボックスを廃止→ステータスバーとログで全通知
  - 完了後アクションパネルをスライドイン表示

v1.2.0 - UIデザイン全面リニューアル
  - ダークテーマ（チャコール）の完全カスタムUIに変更
  - ttk テーマ依存をなくしWindows/Mac/Linux で同じ見た目に
  - ドラッグ＆ドロップゾーン・スライダー・カスタム進捗バー

v1.1.0 - MoviePy v2 / Python 3.14 対応・バグ修正

【作成者メモ】
初心者でも読みやすいよう、1関数1役割を意識しています。
将来の拡張ポイントは「# [拡張ポイント]」のコメントで示しています。
"""

# =============================================
# ライブラリの読み込み
# =============================================
import os
import sys
import uuid
import tempfile
import subprocess  # 再生・フォルダを開く・ドラッグ＆ドロップで使う
import tkinter as tk
from tkinter import filedialog
import threading

# EXE環境で imageio / moviepy の metadata 読み取りに失敗する対策
try:
    import importlib.metadata as importlib_metadata

    _original_version = importlib_metadata.version

    def _safe_version(package_name: str) -> str:
        try:
            return _original_version(package_name)
        except importlib_metadata.PackageNotFoundError:
            if getattr(sys, "frozen", False):
                if package_name in ("imageio", "imageio-ffmpeg", "imageio_ffmpeg", "moviepy", "pydub"):
                    return "0"
            raise

    importlib_metadata.version = _safe_version

except Exception:
    pass


# ---- 動画処理・無音検出ライブラリ ----
LIBRARIES_OK = False
IMPORT_ERROR_MESSAGE = ""

try:
    from moviepy import VideoFileClip, concatenate_videoclips
    from pydub import AudioSegment
    from pydub.silence import detect_nonsilent
    LIBRARIES_OK = True
except ImportError as e:
    IMPORT_ERROR_MESSAGE = str(e)
except Exception as e:
    IMPORT_ERROR_MESSAGE = str(e)

# =============================================
# FFmpeg パス解決（EXE化対応）
# =============================================

def get_ffmpeg_dir() -> str:
    """
    FFmpeg の実行ファイルがあるフォルダパスを返す関数

    探す順番:
      1. EXE と同じフォルダ内の ffmpeg フォルダ
         例: SilenceCut/ffmpeg/ffmpeg.exe
      2. EXE / スクリプトと同じフォルダ（直置き）
         例: SilenceCut/ffmpeg.exe
      3. PATH が通った場所（システムインストール済み）

    PyInstaller でビルドした EXE の場合:
      sys._MEIPASS ... 一時展開フォルダ（onefile の場合）
      sys.executable の dirname ... EXE 自身のフォルダ（onedir の場合）
      どちらも試す。

    戻り値:
      ffmpeg.exe が見つかったフォルダパス、見つからなければ空文字列
    """
    import shutil

    # ── 候補パスを順番に作る ──────────────────────────
    candidates = []

    # ① PyInstaller onefile: 一時展開先の ffmpeg サブフォルダ / 直下
    if hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, "ffmpeg"))
        candidates.append(sys._MEIPASS)

    # ② EXE / スクリプトと同じフォルダの ffmpeg サブフォルダ / 直下
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    candidates.append(os.path.join(exe_dir, "ffmpeg"))
    candidates.append(exe_dir)

    # ③ スクリプト実行時: __file__ のフォルダ（開発時用）
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidates.append(os.path.join(script_dir, "ffmpeg"))
        candidates.append(script_dir)
    except NameError:
        pass  # __file__ がない環境（PyInstaller frozen）では無視

    # ── 候補を順番に確認 ──────────────────────────────
    for folder in candidates:
        ffmpeg_path = os.path.join(folder, "ffmpeg.exe")
        if os.path.isfile(ffmpeg_path):
            return folder  # フォルダパスを返す

    # ④ PATH からも探す（システムインストール済みの場合）
    if shutil.which("ffmpeg"):
        return ""  # PATH が通っているので空文字でよい（追加設定不要）

    return ""  # 見つからなかった


def setup_ffmpeg_path():
    """
    FFmpeg のパスを moviepy / pydub / imageio に設定する関数

    この関数を main() の最初に呼ぶことで、
    EXE と同じフォルダに ffmpeg.exe を置くだけで動くようになる。

    設定方法:
      - 環境変数 PATH に追加
      - imageio_ffmpeg の FFMPEG_BINARY を上書き
      - pydub の AudioSegment.converter を上書き
    どれか一つでも通れば moviepy / pydub 両方が動く。
    """
    ffmpeg_dir = get_ffmpeg_dir()
    if not ffmpeg_dir:
        return  # PATH が通っているか見つからない → 何もしない

    ffmpeg_exe = os.path.join(ffmpeg_dir, "ffmpeg.exe")

    # ── 環境変数 PATH の先頭に追加 ────────────────────
    # moviepy も pydub も内部で PATH を探すのでこれだけで大体動く
    current_path = os.environ.get("PATH", "")
    if ffmpeg_dir not in current_path:
        os.environ["PATH"] = ffmpeg_dir + os.pathsep + current_path

    # ── imageio_ffmpeg に直接パスを教える（moviepy が使う） ────
    try:
        import imageio_ffmpeg
        # imageio_ffmpeg はこの変数でパスを上書きできる
        imageio_ffmpeg._utils.get_ffmpeg_exe = lambda: ffmpeg_exe
    except Exception:
        pass

    # ── pydub に直接パスを教える ───────────────────────
    try:
        from pydub import AudioSegment
        AudioSegment.converter = ffmpeg_exe
        AudioSegment.ffmpeg    = ffmpeg_exe
        AudioSegment.ffprobe   = os.path.join(ffmpeg_dir, "ffprobe.exe")
    except Exception:
        pass


# =============================================
# デザイントークン（色・フォント・サイズを一元管理）
# =============================================
BG_BASE      = "#16181D"
BG_SURFACE   = "#1E2028"
BG_ELEVATED  = "#252830"
BORDER_COLOR = "#2E3140"

ACCENT       = "#4AE3A0"
ACCENT_DARK  = "#2DBF80"

TEXT_PRIMARY   = "#E8EAF0"
TEXT_SECONDARY = "#8B90A0"
TEXT_MUTED     = "#4E5260"

LOG_BG      = "#12141A"
LOG_INFO    = "#7A8090"
LOG_SUCCESS = "#4AE3A0"
LOG_WARNING = "#F0C060"
LOG_ERROR   = "#F07080"

FONT_LABEL  = ("Yu Gothic UI", 9)
FONT_HINT   = ("Yu Gothic UI", 8)
FONT_VALUE  = ("Yu Gothic UI Semibold", 11, "bold")
FONT_BUTTON = ("Yu Gothic UI", 12, "bold")
FONT_LOG    = ("Consolas", 9)

WINDOW_TITLE     = "SilenceCut"
WINDOW_WIDTH     = 720
WINDOW_HEIGHT    = 700
WINDOW_MIN_WIDTH = 600
WINDOW_MIN_HEIGHT = 580

# =============================================
# 処理パラメーターのデフォルト値
# =============================================
# [拡張ポイント] 将来は config.json から読み込む形にできます
DEFAULT_SILENCE_THRESH_DBFS     = -40
DEFAULT_MIN_SILENCE_DURATION_MS = 500
DEFAULT_PADDING_MS              = 100
DEFAULT_SEEK_STEP_MS            = 100

SUPPORTED_VIDEO_EXTENSIONS = (
    ("動画ファイル", "*.mp4 *.avi *.mov *.mkv *.wmv *.flv"),
    ("すべてのファイル", "*.*"),
)
OUTPUT_SUFFIX = "_silence_cut"




# =============================================
# カスタムウィジェット
# =============================================

class FlatButton(tk.Canvas):
    """
    フルカスタムの角丸ボタン（ホバー・無効状態対応）

    tkinter 標準の Button は見た目の自由度が低いため
    Canvas で描いて click イベントを自前で処理します。
    """

    def __init__(self, parent, text="", command=None,
                 bg=ACCENT, fg="#000000",
                 hover_bg=ACCENT_DARK,
                 disabled_bg="#2A2D38", disabled_fg=TEXT_MUTED,
                 font=FONT_BUTTON, radius=8, height=48, **kwargs):
        super().__init__(
            parent,
            bg=parent.cget("bg"),
            highlightthickness=0,
            height=height,
            cursor="hand2",
            **kwargs,
        )
        self._text       = text
        self._command    = command
        self._bg_normal  = bg
        self._bg_hover   = hover_bg
        self._bg_dis     = disabled_bg
        self._fg_normal  = fg
        self._fg_dis     = disabled_fg
        self._font       = font
        self._radius     = radius
        self._enabled    = True
        self._hovered    = False

        self.bind("<Configure>", self._redraw)
        self.bind("<Enter>",     self._on_enter)
        self.bind("<Leave>",     self._on_leave)
        self.bind("<Button-1>",  self._on_click)

    def _redraw(self, _=None):
        self.delete("all")
        w, h, r = self.winfo_width(), self.winfo_height(), self._radius
        bg = (self._bg_dis if not self._enabled
              else self._bg_hover if self._hovered
              else self._bg_normal)
        fg = self._fg_dis if not self._enabled else self._fg_normal

        pts = [r,0, w-r,0, w,0, w,r, w,h-r, w,h,
               w-r,h, r,h, 0,h, 0,h-r, 0,r, 0,0, r,0]
        self.create_polygon(pts, smooth=True, fill=bg, outline="")
        self.create_text(w//2, h//2, text=self._text,
                         fill=fg, font=self._font, anchor="center")

    def _on_enter(self, _):
        self._hovered = True;  self._redraw()

    def _on_leave(self, _):
        self._hovered = False; self._redraw()

    def _on_click(self, _):
        if self._enabled and self._command:
            self._command()

    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        self.config(cursor="hand2" if enabled else "arrow")
        self._redraw()

    def set_text(self, text: str):
        self._text = text
        self._redraw()


class SlimProgressBar(tk.Canvas):
    """細くシンプルなカスタム進捗バー（0.0〜1.0）"""

    def __init__(self, parent, height=3, **kwargs):
        super().__init__(parent, bg=parent.cget("bg"),
                         highlightthickness=0, height=height, **kwargs)
        self._value = 0.0
        self.bind("<Configure>", self._redraw)

    def set_value(self, ratio: float):
        self._value = max(0.0, min(1.0, ratio))
        self._redraw()

    def _redraw(self, _=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        self.create_rectangle(0, 0, w, h, fill=BG_ELEVATED, outline="")
        if self._value > 0:
            self.create_rectangle(0, 0, int(w * self._value), h,
                                  fill=ACCENT, outline="")


# =============================================
# 処理ロジック
# =============================================

def detect_nonsilent_segments(audio_file_path, silence_thresh_dbfs,
                               min_silence_duration_ms, seek_step_ms):
    """
    音声ファイルから「有音区間」を検出して返す関数

    戻り値: [[開始ms, 終了ms], ...]
    """
    try:
        audio = AudioSegment.from_file(audio_file_path)
        return detect_nonsilent(
            audio,
            min_silence_len=min_silence_duration_ms,
            silence_thresh=silence_thresh_dbfs,
            seek_step=seek_step_ms,
        )
    except Exception as e:
        raise RuntimeError(f"音声解析中にエラーが発生しました: {e}") from e


def apply_padding_to_segments(segments, padding_ms, total_duration_ms):
    """有音区間の前後に余白を追加する関数"""
    return [
        [max(0, s - padding_ms), min(total_duration_ms, e + padding_ms)]
        for s, e in segments
    ]


def merge_overlapping_segments(segments):
    """重なり合っている区間をひとつにまとめる関数"""
    if not segments:
        return []
    segs = sorted(segments, key=lambda x: x[0])
    merged = [list(segs[0])]
    for cs, ce in segs[1:]:
        if cs <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], ce)
        else:
            merged.append([cs, ce])
    return merged


def generate_output_file_path(input_file_path):
    """出力ファイルのパスを自動生成する関数（元ファイルは上書きしない）"""
    d = os.path.dirname(input_file_path)
    name, ext = os.path.splitext(os.path.basename(input_file_path))
    return os.path.join(d, f"{name}{OUTPUT_SUFFIX}{ext}")


def make_unique_temp_audio_path():
    """重複しない一時音声ファイルのパスを生成する関数"""
    return os.path.join(
        tempfile.gettempdir(),
        f"silence_cutter_temp_{uuid.uuid4().hex[:8]}.wav",
    )


def open_path_in_explorer(path: str):
    """
    ファイルまたはフォルダを OS のファイルマネージャーで開く関数

    Windows: エクスプローラー
    Mac    : Finder
    Linux  : xdg-open
    """
    try:
        if sys.platform == "win32":
            if os.path.isfile(path):
                # ファイルの場合は親フォルダを開いて該当ファイルを選択状態にする
                subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
            else:
                subprocess.Popen(["explorer", os.path.normpath(path)])
        elif sys.platform == "darwin":
            subprocess.Popen(["open", "-R", path] if os.path.isfile(path)
                             else ["open", path])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(path)])
    except Exception:
        pass


def play_video_in_default_player(path: str):
    """
    OS のデフォルト動画プレイヤーで動画を再生する関数

    Windows: os.startfile  /  Mac: open  /  Linux: xdg-open
    """
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


def cut_silence_from_video(input_path, silence_thresh_dbfs, min_silence_duration_ms,
                            padding_ms, seek_step_ms, log_callback, progress_callback):
    """
    動画の無音部分をカットして保存するメイン処理関数

    戻り値:
        成功: (出力パス, 元の長さ秒, 出力の長さ秒)
        失敗: None
    """
    video_clip = None
    subclips   = []
    final_clip = None
    temp_audio = make_unique_temp_audio_path()

    try:
        log_callback("動画ファイルを読み込んでいます...", "info")
        progress_callback(10)

        video_clip = VideoFileClip(input_path)
        total_sec  = video_clip.duration
        total_ms   = int(total_sec * 1000)
        log_callback(f"動画の長さ: {total_sec:.1f} 秒", "info")

        if video_clip.audio is None:
            log_callback("音声トラックがありません。音声付きの動画を選んでください。", "error")
            return None

        log_callback("音声を抽出しています...", "info")
        progress_callback(20)
        video_clip.audio.write_audiofile(temp_audio, logger=None)

        log_callback("無音区間を検出しています...", "info")
        progress_callback(40)
        nonsilent = detect_nonsilent_segments(
            audio_file_path=temp_audio,
            silence_thresh_dbfs=silence_thresh_dbfs,
            min_silence_duration_ms=min_silence_duration_ms,
            seek_step_ms=seek_step_ms,
        )

        if not nonsilent:
            log_callback("有音区間が見つかりませんでした。", "warning")
            log_callback("しきい値を小さくするか、最小無音時間を短くしてください。", "warning")
            return None

        log_callback(f"有音区間: {len(nonsilent)} 個検出", "info")
        progress_callback(55)

        padded  = apply_padding_to_segments(nonsilent, padding_ms, total_ms)
        final_segs = merge_overlapping_segments(padded)
        log_callback(f"統合後: {len(final_segs)} 区間", "info")

        log_callback("動画を組み立てています...", "info")
        progress_callback(65)

        for s_ms, e_ms in final_segs:
            s = s_ms / 1000.0
            e = min(e_ms / 1000.0, total_sec)
            if e - s < 0.01:
                continue
            subclips.append(video_clip.subclipped(s, e))  # MoviePy v2

        if not subclips:
            log_callback("有効なクリップが作れませんでした。", "error")
            return None

        log_callback(f"{len(subclips)} クリップを結合しています...", "info")
        progress_callback(75)
        final_clip = concatenate_videoclips(subclips, method="chain")

        output_path = generate_output_file_path(input_path)
        if os.path.exists(output_path):
            log_callback(f"上書き: {os.path.basename(output_path)}", "warning")

        log_callback(f"保存先: {output_path}", "info")
        progress_callback(85)

        final_clip.write_videofile(output_path, logger=None,
                                   audio_codec="aac", threads=4)
        progress_callback(100)

        out_sec  = final_clip.duration
        cut_sec  = total_sec - out_sec
        cut_pct  = (cut_sec / total_sec * 100) if total_sec > 0 else 0

        log_callback("─" * 36, "info")
        log_callback(
            f"完了  {total_sec:.1f}s → {out_sec:.1f}s  "
            f"({cut_pct:.1f}% 短縮 / {cut_sec:.1f}s 削減)",
            "success",
        )

        return (output_path, total_sec, out_sec)

    except Exception as e:
        log_callback(f"エラー: {e}", "error")
        return None

    finally:
        for c in subclips:
            try: c.close()
            except Exception: pass
        if final_clip:
            try: final_clip.close()
            except Exception: pass
        if video_clip:
            try: video_clip.close()
            except Exception: pass
        if os.path.exists(temp_audio):
            try: os.remove(temp_audio)
            except Exception: pass


def _parse_first_dnd_path(data: str) -> str:
    """
    tkinterdnd2 の event.data から最初のファイルパスを取り出す関数

    tkinterdnd2 が返すパスの形式（3パターン）:
      ① スペースなし 1ファイル : C:/Users/user/video.mp4
      ② スペースあり 1ファイル : {C:/Users/my name/my video.mp4}
      ③ 複数ファイル           : {C:/path/a.mp4} {C:/path/b.mp4}
                                 C:/path/a.mp4 C:/path/b.mp4（スペースなし複数）

    問題のある処理:
      raw.split()[0] → スペース含みパスを "C:/Users/my" で切ってしまう

    正しい処理:
      波括弧が含まれているかどうかで分岐する。
      波括弧あり → 最初の {} ブロックを取り出す
      波括弧なし → 全体を1つのパスとして使う（スペース区切りなら先頭だけ）

    引数:
        data: event.data の文字列

    戻り値:
        パス文字列（見つからなければ空文字列）
    """
    raw = data.strip()
    if not raw:
        return ""

    if "{" in raw:
        # 波括弧あり形式: {パス} か {パス1} {パス2} ...
        # 最初の { } ブロックだけ取り出す
        start = raw.find("{")
        end   = raw.find("}", start)
        if start != -1 and end != -1:
            return raw[start + 1 : end]
        # } がない異常ケース: 波括弧だけ除去して返す
        return raw.replace("{", "").strip()
    else:
        # 波括弧なし形式: そのままか、スペース区切りの複数パス
        # 先頭の1つだけを使う（スペース区切り複数ファイルは先頭のみ採用）
        return raw.split()[0] if raw else ""


def validate_inputs(file_path, silence_thresh_str, min_silence_str, padding_str):
    """
    ユーザーの入力値が正しいかチェックする関数

    戻り値: エラーがなければ None / エラーがあればエラーメッセージ文字列
    """
    if not file_path or not file_path.strip():
        return "動画ファイルを選択してください。"
    if not os.path.isfile(file_path):
        return "選択したファイルが見つかりません。もう一度選択してください。"
    try:
        if float(silence_thresh_str) > 0:
            return "無音しきい値は 0 以下の値にしてください（例: -40）"
    except ValueError:
        return "無音しきい値には数値を入力してください（例: -40）"
    try:
        v = float(min_silence_str)
        if v <= 0: return "最小無音時間は 0 より大きい値にしてください（例: 0.5）"
        if v > 60: return "最小無音時間が長すぎます（60秒以内）"
    except ValueError:
        return "最小無音時間には数値を入力してください（例: 0.5）"
    try:
        v = float(padding_str)
        if v < 0: return "余白時間は 0 以上の値にしてください（例: 0.1）"
        if v > 5: return "余白時間が長すぎます（5秒以内）"
    except ValueError:
        return "余白時間には数値を入力してください（例: 0.1）"
    return None


# =============================================
# GUIアプリケーション本体
# =============================================

class SilenceCutterApp:
    """
    動画無音カットツールのGUIアプリクラス

    ===== 主なメソッドの役割 =====
    __init__                : 初期化・変数準備
    _build_ui               : 全体レイアウト組み立て
    _build_header           : タイトル行 + ショートカットヒント
    _build_drop_zone        : ファイル選択ゾーン
    _build_params_section   : スライダー3本
    _build_run_button       : 実行ボタン
    _build_progress_area    : 進捗バー + ステータス
    _build_result_panel     : 完了後アクションパネル（再生・フォルダ・次へ）
    _build_log_section      : ログ
    _setup_drag_and_drop    : ドラッグ＆ドロップの初期化
    _setup_shortcuts        : ショートカットキーの登録
    _set_file               : ファイルセット共通処理
    _on_select_file         : ファイル選択ダイアログ
    _on_start_process       : 実行
    _run_process            : バックグラウンド処理スレッド
    _on_process_complete    : 完了後の表示
    _on_process_failed      : 失敗後の表示
    _show_result_panel      : 完了後パネルを表示
    _hide_result_panel      : 完了後パネルを非表示
    _on_play_output         : 出力動画を再生
    _on_open_folder         : フォルダを開く
    _on_next_file           : 次のファイルへ（リセット）
    _clear_log              : ログをクリア
    _log                    : ログ追記（スレッドセーフ）
    _set_progress           : 進捗更新（スレッドセーフ）
    _set_ui_enabled         : UI 有効/無効
    _redraw_drop_zone       : ドロップゾーン再描画
    _make_slider_row        : スライダー行ヘルパー
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(WINDOW_TITLE)
        self.root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.root.minsize(WINDOW_MIN_WIDTH, WINDOW_MIN_HEIGHT)
        self.root.configure(bg=BG_BASE)
        self.root.resizable(True, True)

        # 状態変数
        self.is_processing    = False
        self.output_path      = None   # 最後に保存した出力ファイルのパス
        self.last_dir         = ""     # 最後に開いたフォルダ（次回ダイアログの初期位置）

        # ウィジェット変数
        self.file_path_var       = tk.StringVar()
        self.silence_thresh_var  = tk.DoubleVar(value=float(DEFAULT_SILENCE_THRESH_DBFS))
        self.min_silence_var     = tk.DoubleVar(value=DEFAULT_MIN_SILENCE_DURATION_MS / 1000.0)
        self.padding_var         = tk.DoubleVar(value=DEFAULT_PADDING_MS / 1000.0)

        # ドロップゾーンの状態
        self._drop_zone_hovered  = False
        self._drop_zone_has_file = False

        self._build_ui()
        self._setup_drag_and_drop()
        self._setup_shortcuts()

        if not LIBRARIES_OK:
            self._log(f"ライブラリ未インストール: {IMPORT_ERROR_MESSAGE}", "error")
            self._log("pip install moviepy pydub を実行してください。", "warning")

        self._log("起動しました。動画ファイルをドラッグするか、クリックして選択してください。", "info")
        self._log("対応形式: mp4 / avi / mov / mkv / wmv / flv", "info")
        self._log("ショートカット: Space=実行  O=開く  R=再生  E=フォルダ  N=次へ  Ctrl+L=ログクリア", "info")

    # ─────────────────────────────────────
    # UIの組み立て
    # ─────────────────────────────────────

    def _build_ui(self):
        """全体のレイアウトを組み立てる"""
        outer = tk.Frame(self.root, bg=BG_BASE)
        outer.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

        self._build_header(outer)
        self._build_drop_zone(outer)
        self._build_params_section(outer)
        self._build_run_button(outer)
        self._build_progress_area(outer)
        self._build_result_panel(outer)   # 完了後に表示するパネル
        self._build_log_section(outer)

    def _build_header(self, parent):
        """タイトル・サブタイトル・ショートカットヒント行"""
        header = tk.Frame(parent, bg=BG_BASE)
        header.pack(fill=tk.X, pady=(0, 16))

        # ── タイトル行 ──
        title_row = tk.Frame(header, bg=BG_BASE)
        title_row.pack(anchor=tk.W)

        tk.Frame(title_row, bg=ACCENT, width=4, height=32).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        tk.Label(
            title_row, text="SilenceCut",
            font=("Yu Gothic UI", 22, "bold"),
            fg=TEXT_PRIMARY, bg=BG_BASE,
        ).pack(side=tk.LEFT)
        tk.Label(
            title_row, text="  v1.3",
            font=("Yu Gothic UI", 10),
            fg=TEXT_MUTED, bg=BG_BASE,
        ).pack(side=tk.LEFT, anchor=tk.S, pady=(8, 0))

        # ── サブタイトル ──
        tk.Label(
            header,
            text="動画の無音部分を自動検出・削除。編集時間を一瞬で短縮。",
            font=("Yu Gothic UI", 10),
            fg=TEXT_SECONDARY, bg=BG_BASE,
        ).pack(anchor=tk.W, pady=(4, 0))

    def _build_drop_zone(self, parent):
        """
        ファイル選択ゾーン

        - クリック: ダイアログを開く
        - ドラッグ＆ドロップ: Windows はネイティブ対応、他は tkinter DND
        """
        self.drop_zone_canvas = tk.Canvas(
            parent, bg=BG_BASE, highlightthickness=0, height=88,
        )
        self.drop_zone_canvas.pack(fill=tk.X, pady=(0, 12))
        self.drop_zone_canvas.bind("<Configure>", self._redraw_drop_zone)
        self.drop_zone_canvas.bind("<Button-1>",  self._on_select_file)
        self.drop_zone_canvas.bind("<Enter>",     lambda _: self._set_dz_hover(True))
        self.drop_zone_canvas.bind("<Leave>",     lambda _: self._set_dz_hover(False))

    def _set_dz_hover(self, hovered: bool):
        self._drop_zone_hovered = hovered
        self._redraw_drop_zone()

    def _redraw_drop_zone(self, _=None):
        """ドロップゾーンの見た目を再描画する"""
        c = self.drop_zone_canvas
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        r = 10

        has_file = self._drop_zone_has_file
        hovered  = self._drop_zone_hovered

        border_color = ACCENT if (hovered or has_file) else BORDER_COLOR
        dash         = () if has_file else (6, 4)

        pts = [r,0, w-r,0, w,0, w,r, w,h-r, w,h,
               w-r,h, r,h, 0,h, 0,h-r, 0,r, 0,0, r,0]
        c.create_polygon(pts, smooth=True, fill=BG_ELEVATED, outline="")
        c.create_polygon(pts, smooth=True, fill="",
                         outline=border_color, width=1, dash=dash)

        cx, cy = w // 2, h // 2

        if has_file:
            # ── ファイル選択済み ──
            fname = os.path.basename(self.file_path_var.get())
            # ファイル名が長すぎる場合は省略
            if len(fname) > 60:
                fname = fname[:28] + "…" + fname[-28:]
            c.create_text(cx, cy - 14, text="▶",
                          fill=ACCENT, font=("Yu Gothic UI", 13), anchor="center")
            c.create_text(cx, cy + 8, text=fname,
                          fill=TEXT_PRIMARY, font=("Yu Gothic UI", 11, "bold"),
                          anchor="center")
            c.create_text(cx, cy + 26, text="クリックで変更 / ドラッグで上書き",
                          fill=TEXT_MUTED, font=FONT_HINT, anchor="center")
        else:
            # ── 未選択 ──
            txt_color = ACCENT if hovered else TEXT_SECONDARY
            c.create_text(cx, cy - 10, text="ここに動画ファイルをドラッグ  または  クリックして選択",
                          fill=txt_color, font=("Yu Gothic UI", 11), anchor="center")
            c.create_text(cx, cy + 12, text="mp4 / avi / mov / mkv / wmv / flv",
                          fill=TEXT_MUTED, font=FONT_HINT, anchor="center")

    def _build_params_section(self, parent):
        """スライダー3本の設定エリア"""
        section = tk.Frame(parent, bg=BG_BASE)
        section.pack(fill=tk.X, pady=(0, 12))
        section.columnconfigure(0, weight=1)
        section.columnconfigure(1, weight=1)
        section.columnconfigure(2, weight=1)

        self._make_slider_row(section, 0, "無音しきい値", "dBFS",
                              self.silence_thresh_var, -80, 0, 1,
                              "小さい値ほど敏感に検出", (0, 6))
        self._make_slider_row(section, 1, "最小無音時間", "秒",
                              self.min_silence_var, 0.1, 5.0, 0.1,
                              "短いほど細かくカット", (6, 6))
        self._make_slider_row(section, 2, "前後の余白", "秒",
                              self.padding_var, 0.0, 1.0, 0.05,
                              "自然な間を残す", (6, 0))

        # [拡張ポイント] ここにプリセットボタンを追加できます

    def _make_slider_row(self, parent, column, label, unit, variable,
                         from_, to, resolution, hint, padx=(0, 0)):
        """スライダー1本分のUI（ラベル・値・スライダー・ヒント）を作るヘルパー"""
        card = tk.Frame(parent, bg=BG_SURFACE, padx=12, pady=10)
        card.grid(row=0, column=column, sticky="nsew", padx=padx)
        card.configure(highlightbackground=BORDER_COLOR, highlightthickness=1)

        row = tk.Frame(card, bg=BG_SURFACE)
        row.pack(fill=tk.X)

        tk.Label(row, text=label, font=FONT_LABEL,
                 fg=TEXT_SECONDARY, bg=BG_SURFACE).pack(side=tk.LEFT)

        val_lbl = tk.Label(row, text=self._fmt(variable.get(), unit),
                           font=FONT_VALUE, fg=ACCENT, bg=BG_SURFACE)
        val_lbl.pack(side=tk.RIGHT)

        slider = tk.Scale(
            card, variable=variable,
            from_=from_, to=to, resolution=resolution,
            orient=tk.HORIZONTAL, showvalue=False,
            bg=BG_SURFACE, fg=ACCENT, troughcolor=BG_ELEVATED,
            activebackground=ACCENT_DARK,
            highlightthickness=0, bd=0,
            sliderlength=14, sliderrelief=tk.FLAT,
        )
        slider.pack(fill=tk.X, pady=(4, 0))

        tk.Label(card, text=hint, font=FONT_HINT,
                 fg=TEXT_MUTED, bg=BG_SURFACE).pack(anchor=tk.W, pady=(2, 0))

        def _refresh(*_):
            val_lbl.config(text=self._fmt(variable.get(), unit))

        variable.trace_add("write", _refresh)

    @staticmethod
    def _fmt(value: float, unit: str) -> str:
        """スライダー値を表示用文字列に変換"""
        return f"{int(value)} dBFS" if unit == "dBFS" else f"{value:.2f} s"

    def _build_run_button(self, parent):
        """実行ボタン"""
        self.run_button = FlatButton(
            parent,
            text="▶   無音をカットして保存  [Space]",
            command=self._on_start_process,
            bg=ACCENT, fg=BG_BASE,
            hover_bg=ACCENT_DARK,
            height=50,
        )
        self.run_button.pack(fill=tk.X, pady=(0, 10))

    def _build_progress_area(self, parent):
        """進捗バーとステータスラベル"""
        area = tk.Frame(parent, bg=BG_BASE)
        area.pack(fill=tk.X, pady=(0, 8))

        self.progress_bar = SlimProgressBar(area, height=3)
        self.progress_bar.pack(fill=tk.X)

        self.status_label = tk.Label(
            area, text="",
            font=("Yu Gothic UI", 9),
            fg=TEXT_MUTED, bg=BG_BASE, anchor=tk.E,
        )
        self.status_label.pack(fill=tk.X, pady=(3, 0))

    def _build_result_panel(self, parent):
        """
        処理完了後に表示するアクションパネル

        ── ボタン3つ ──
        ▶ 再生  [R]  : 出力ファイルをデフォルトプレイヤーで再生
        📁 フォルダ [E]: 出力先フォルダをエクスプローラーで開く
        ➜ 次のファイル [N]: UIをリセットして次の作業へ
        """
        self.result_panel = tk.Frame(parent, bg=BG_BASE)
        # 初期状態は非表示（完了後に pack する）

        # ── ラベル ──
        tk.Label(
            self.result_panel,
            text="出力ファイル",
            font=FONT_LABEL,
            fg=TEXT_SECONDARY, bg=BG_BASE,
        ).pack(anchor=tk.W, pady=(0, 4))

        # ── ボタン行 ──
        btn_row = tk.Frame(self.result_panel, bg=BG_BASE)
        btn_row.pack(fill=tk.X)
        btn_row.columnconfigure(0, weight=2)
        btn_row.columnconfigure(1, weight=2)
        btn_row.columnconfigure(2, weight=3)

        self.btn_play = FlatButton(
            btn_row,
            text="▶  再生  [R]",
            command=self._on_play_output,
            bg=BG_SURFACE, fg=ACCENT,
            hover_bg=BG_ELEVATED,
            height=40,
        )
        self.btn_play.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.btn_folder = FlatButton(
            btn_row,
            text="📁  フォルダを開く  [E]",
            command=self._on_open_folder,
            bg=BG_SURFACE, fg=TEXT_PRIMARY,
            hover_bg=BG_ELEVATED,
            height=40,
        )
        self.btn_folder.grid(row=0, column=1, sticky="ew", padx=(0, 6))

        self.btn_next = FlatButton(
            btn_row,
            text="➜  次のファイルへ  [N]",
            command=self._on_next_file,
            bg=ACCENT, fg=BG_BASE,
            hover_bg=ACCENT_DARK,
            height=40,
        )
        self.btn_next.grid(row=0, column=2, sticky="ew")

    def _build_log_section(self, parent):
        """ログ表示エリア"""
        log_frame = tk.Frame(
            parent, bg=LOG_BG,
            highlightbackground=BORDER_COLOR,
            highlightthickness=1,
        )
        log_frame.pack(fill=tk.BOTH, expand=True)

        # ── ログヘッダー（タイトル + クリアボタン）──
        log_header = tk.Frame(log_frame, bg=LOG_BG)
        log_header.pack(fill=tk.X, padx=12, pady=(8, 0))

        tk.Label(
            log_header, text="LOG",
            font=("Consolas", 9),
            fg=TEXT_MUTED, bg=LOG_BG,
        ).pack(side=tk.LEFT)

        # ログクリアボタン（小さく目立たない）
        clear_btn = tk.Label(
            log_header, text="クリア  [Ctrl+L]",
            font=FONT_HINT, fg=TEXT_MUTED, bg=LOG_BG,
            cursor="hand2",
        )
        clear_btn.pack(side=tk.RIGHT)
        clear_btn.bind("<Button-1>", lambda _: self._clear_log())
        clear_btn.bind("<Enter>",    lambda _: clear_btn.config(fg=TEXT_SECONDARY))
        clear_btn.bind("<Leave>",    lambda _: clear_btn.config(fg=TEXT_MUTED))

        # ── テキストエリア ──
        scrollbar = tk.Scrollbar(
            log_frame, width=6,
            bg=BG_ELEVATED, troughcolor=LOG_BG,
            activebackground=BORDER_COLOR,
            relief=tk.FLAT, bd=0,
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=(0, 4))

        self.log_text = tk.Text(
            log_frame,
            font=FONT_LOG, bg=LOG_BG, fg=LOG_INFO,
            insertbackground=TEXT_PRIMARY,
            selectbackground="#2E3140",
            relief=tk.FLAT, bd=0,
            padx=12, pady=6,
            yscrollcommand=scrollbar.set,
            state=tk.DISABLED,
            wrap=tk.WORD,
            spacing1=2,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)

        self.log_text.tag_config("info",    foreground=LOG_INFO)
        self.log_text.tag_config("success", foreground=LOG_SUCCESS)
        self.log_text.tag_config("warning", foreground=LOG_WARNING)
        self.log_text.tag_config("error",   foreground=LOG_ERROR)

    # ─────────────────────────────────────
    # ドラッグ＆ドロップのセットアップ
    # ─────────────────────────────────────

    def _setup_drag_and_drop(self):
        """
        ドラッグ＆ドロップを初期化する（tkinterdnd2 を使用）

        tkinterdnd2 は tkinter に DnD 機能を追加する軽量ライブラリ。
        インストール: pip install tkinterdnd2

        tkinterdnd2 がない場合はクリック選択のみ有効になる。
        ログにインストール案内を表示する。
        """
        try:
            # tkinterdnd2 が使えるかチェック
            # main() で TkinterDnD.Tk() を使って起動している場合のみ有効
            self.drop_zone_canvas.drop_target_register("DND_Files")
            self.drop_zone_canvas.dnd_bind("<<Drop>>", self._on_dnd2_drop)
            self._log("ドラッグ＆ドロップ: 有効", "info")
        except Exception:
            self._log("ドラッグ＆ドロップ: 無効（クリックで選択してください）", "info")
            self._log("  有効にするには: pip install tkinterdnd2", "info")

    def _on_drop_file(self, file_path: str):
        """ドロップされたファイルパスを受け取る共通メソッド"""
        if not self.is_processing:
            self._set_file(file_path)

    def _on_dnd2_drop(self, event):
        """
        tkinterdnd2 からのドロップイベント処理

        tkinterdnd2 の event.data は以下のいずれかの形式で届く:
          ① スペースなしパス  : C:/Users/user/video.mp4
          ② スペースありパス  : {C:/Users/my name/my video.mp4}
          ③ 複数ファイル      : {C:/path/a.mp4} {C:/path/b.mp4}
                             または C:/path/a.mp4 C:/path/b.mp4

        波括弧はパスにスペースが含まれるときの区切り文字。
        単純に split() すると「C:/Users/my」で切れてしまうため、
        波括弧の有無を判定しながら正しく1つ目のパスだけを取り出す。
        """
        path = _parse_first_dnd_path(event.data)
        if path:
            self._on_drop_file(path)

    # ─────────────────────────────────────
    # ショートカットキーの登録
    # ─────────────────────────────────────

    def _setup_shortcuts(self):
        """
        キーボードショートカットを root ウィンドウに登録する

        bind_all を使うと、フォーカスがどのウィジェットにあっても動く。
        """
        # Space / Return → 実行
        self.root.bind_all("<space>",  lambda _: self._on_start_process())
        self.root.bind_all("<Return>", lambda _: self._on_start_process())

        # O → ファイルを開く
        self.root.bind_all("<o>",      lambda _: self._on_select_file())
        self.root.bind_all("<O>",      lambda _: self._on_select_file())

        # R → 再生
        self.root.bind_all("<r>",      lambda _: self._on_play_output())
        self.root.bind_all("<R>",      lambda _: self._on_play_output())

        # E → フォルダを開く
        self.root.bind_all("<e>",      lambda _: self._on_open_folder())
        self.root.bind_all("<E>",      lambda _: self._on_open_folder())

        # N → 次のファイルへ
        self.root.bind_all("<n>",      lambda _: self._on_next_file())
        self.root.bind_all("<N>",      lambda _: self._on_next_file())

        # Ctrl+L → ログクリア
        self.root.bind_all("<Control-l>", lambda _: self._clear_log())
        self.root.bind_all("<Control-L>", lambda _: self._clear_log())

    # ─────────────────────────────────────
    # イベントハンドラ
    # ─────────────────────────────────────

    def _on_select_file(self, _event=None):
        """ファイル選択ダイアログを開く"""
        if self.is_processing:
            return

        # 最後に開いたフォルダから始める（初回はホームディレクトリ）
        initial_dir = self.last_dir or os.path.expanduser("~")

        path = filedialog.askopenfilename(
            title="動画ファイルを選択してください",
            filetypes=SUPPORTED_VIDEO_EXTENSIONS,
            initialdir=initial_dir,
        )
        if path:
            self._set_file(path)

    def _set_file(self, file_path: str):
        """
        ファイルをセットして UIに反映する共通処理

        クリック選択・ドラッグ＆ドロップ、どちらからでもここに来る。
        """
        self.file_path_var.set(file_path)
        self._drop_zone_has_file = True

        # 次回ダイアログ用にフォルダを記憶
        self.last_dir = os.path.dirname(file_path)

        self._redraw_drop_zone()
        self._hide_result_panel()    # 前回の完了パネルを隠す
        self._set_progress(0)
        self.status_label.config(text="", fg=TEXT_MUTED)
        self._log(f"選択: {os.path.basename(file_path)}", "info")

    def _on_start_process(self):
        """実行ボタン / Space キーで処理を開始"""
        if self.is_processing:
            return
        if not LIBRARIES_OK:
            self._log("ライブラリ未インストール。pip install moviepy pydub を実行してください。", "error")
            return

        err = validate_inputs(
            file_path=self.file_path_var.get(),
            silence_thresh_str=str(int(self.silence_thresh_var.get())),
            min_silence_str=str(self.min_silence_var.get()),
            padding_str=str(self.padding_var.get()),
        )
        if err:
            self._log(f"入力エラー: {err}", "warning")
            self.status_label.config(text=err, fg=LOG_WARNING)
            return

        self._hide_result_panel()
        self._set_ui_enabled(False)
        self._set_progress(0)
        self._log("─" * 36, "info")
        self._log("処理を開始します...", "info")

        threading.Thread(target=self._run_process, daemon=True).start()

    def _run_process(self):
        """バックグラウンドスレッドで動画処理を実行"""
        self.is_processing = True
        try:
            result = cut_silence_from_video(
                input_path=self.file_path_var.get(),
                silence_thresh_dbfs=float(self.silence_thresh_var.get()),
                min_silence_duration_ms=int(self.min_silence_var.get() * 1000),
                padding_ms=int(self.padding_var.get() * 1000),
                seek_step_ms=DEFAULT_SEEK_STEP_MS,
                log_callback=self._log,
                progress_callback=self._set_progress,
            )
            if result:
                self.root.after(0, self._on_process_complete, result)
            else:
                self.root.after(0, self._on_process_failed)
        except Exception as e:
            self._log(f"予期せぬエラー: {e}", "error")
            self.root.after(0, self._on_process_failed)
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self._set_ui_enabled(True))

    def _on_process_complete(self, result):
        """処理完了後のUI更新（メインスレッドで実行）"""
        output_path, orig_sec, out_sec = result
        cut_sec = orig_sec - out_sec
        cut_pct = (cut_sec / orig_sec * 100) if orig_sec > 0 else 0

        self.output_path = output_path

        # ステータスバーに結果サマリー
        self.status_label.config(
            text=(f"完了  {orig_sec:.1f}s → {out_sec:.1f}s  |  "
                  f"{cut_sec:.1f}s 削減  ({cut_pct:.1f}% 短縮)"),
            fg=ACCENT,
        )

        # 完了後アクションパネルを表示（スライドイン風）
        self._show_result_panel()

    def _on_process_failed(self):
        """処理失敗後のUI更新（メインスレッドで実行）"""
        self.status_label.config(
            text="処理に失敗しました。ログを確認してください。",
            fg=LOG_ERROR,
        )
        self._set_progress(0)

    # ─────────────────────────────────────
    # 完了後アクションパネルの表示・非表示
    # ─────────────────────────────────────

    def _show_result_panel(self):
        """完了後パネルを進捗バーの直下に表示する"""
        self.result_panel.pack(
            fill=tk.X,
            pady=(0, 10),
            before=self.log_text.master,  # ログエリアの直前に挿入
        )

    def _hide_result_panel(self):
        """完了後パネルを非表示にする"""
        self.result_panel.pack_forget()
        self.output_path = None

    # ─────────────────────────────────────
    # 完了後アクション
    # ─────────────────────────────────────

    def _on_play_output(self):
        """
        出力動画をデフォルトプレイヤーで再生する  [R]

        出力ファイルがなければ何もしない。
        """
        if not self.output_path or not os.path.exists(self.output_path):
            self._log("再生できる出力ファイルがありません。先に処理を実行してください。", "warning")
            return
        self._log(f"再生: {os.path.basename(self.output_path)}", "info")
        play_video_in_default_player(self.output_path)

    def _on_open_folder(self):
        """
        出力先フォルダをエクスプローラーで開く  [E]

        出力ファイルが選択された状態でフォルダが開く（Windows）。
        """
        if not self.output_path:
            # 出力ファイルがなければ入力ファイルのフォルダを開く
            path = self.file_path_var.get()
            if path:
                open_path_in_explorer(os.path.dirname(path))
            return
        self._log(f"フォルダを開く: {os.path.dirname(self.output_path)}", "info")
        open_path_in_explorer(self.output_path)

    def _on_next_file(self):
        """
        次のファイル処理へリセットする  [N]

        - 完了パネルを隠す
        - ドロップゾーンをリセット
        - 進捗・ステータスをクリア
        - ログには区切り線を入れる（作業履歴として残す）
        """
        self._hide_result_panel()
        self._drop_zone_has_file = False
        self.file_path_var.set("")
        self._redraw_drop_zone()
        self._set_progress(0)
        self.status_label.config(text="", fg=TEXT_MUTED)
        self._log("─" * 36, "info")
        self._log("次のファイルを選択してください。", "info")

    def _clear_log(self):
        """ログエリアをクリアする  [Ctrl+L]"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    # ─────────────────────────────────────
    # UI更新ヘルパー（スレッドセーフ）
    # ─────────────────────────────────────

    def _log(self, message: str, level: str = "info"):
        """ログエリアにメッセージを追記する（別スレッドから呼んでも安全）"""
        def _update():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{message}\n", level)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _update)

    def _set_progress(self, value: float):
        """進捗バーとステータスを更新する（別スレッドから呼んでも安全）"""
        def _update():
            self.progress_bar.set_value(value / 100.0)
            if value > 0:
                self.status_label.config(
                    text=f"処理中  {int(value)}%",
                    fg=TEXT_SECONDARY,
                )
            else:
                self.status_label.config(text="", fg=TEXT_MUTED)
        self.root.after(0, _update)

    def _set_ui_enabled(self, enabled: bool):
        """処理中はUIを無効化、完了後に再度有効化する"""
        if enabled:
            self.run_button.set_enabled(True)
            self.run_button.set_text("▶   無音をカットして保存  [Space]")
            self.drop_zone_canvas.config(cursor="hand2")
        else:
            self.run_button.set_enabled(False)
            self.run_button.set_text("⏳   処理中...")
            self.drop_zone_canvas.config(cursor="arrow")


# =============================================
# アプリの起動エントリーポイント
# =============================================

def main():
    """
    アプリを起動するメイン関数

    ドラッグ＆ドロップに tkinterdnd2 を使う。
    tkinterdnd2 がない場合は通常の tk.Tk で起動し、
    クリック選択のみ有効になる。

    インストール: pip install tkinterdnd2
    """
    # ── FFmpeg のパスを最初に設定する（EXE化対応） ──────────
    # EXE と同じフォルダ or ffmpeg サブフォルダに ffmpeg.exe があれば
    # 自動で認識する。PATH が通っていれば何もしない。
    setup_ffmpeg_path()

    # Windows: 高解像度ディスプレイ（4K/2K）で文字が小さくならないようにする
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    # ── tkinterdnd2 があれば使う、なければ通常の Tk で起動 ──
    # TkinterDnD.Tk は tk.Tk のサブクラスで、DnD 機能が追加されている。
    # どちらの場合もアプリとしては動作する（DnD の有無だけが違う）。
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()

    root.configure(bg=BG_BASE)
    SilenceCutterApp(root)

    # [拡張ポイント] アプリ終了時の処理（設定保存など）
    # root.protocol("WM_DELETE_WINDOW", app.on_closing)

    root.mainloop()


if __name__ == "__main__":
    main()
