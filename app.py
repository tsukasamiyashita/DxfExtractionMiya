# -*- coding: utf-8 -*-
"""
DxfExtractionMiya v1.0.0
機能：DXFファイルからのプレビュー＆テキスト抽出（マウスで範囲指定・連続追加）、設定保存
メインアプリケーションモジュール
"""

import os
import sys
import re
import traceback
import json
import threading
import queue
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk

from ui_preview import PreviewDialog
from app_logic import run_extract_dxf

# ==========================
# 共通変数
# ==========================
APP_TITLE = "DxfExtractionMiya"
VERSION = "v1.0.0"

selected_files = []
selected_folder = ""
current_mode = None  # "file" or "folder"

# UI変数のプレースホルダー（UI構築時に初期化）
base_kw_var = None
base_kw2_var = None
base_dist_var = None

# ==========================
# プログレスダイアログ
# ==========================
class ProgressWindow(Toplevel):
    def __init__(self, parent, title="処理中"):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x180")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # キャンセル用フラグ
        self.cancel_event = threading.Event()
        
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (450 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (180 // 2)
        self.geometry(f"+{x}+{y}")
        
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        
        self.lbl_status = Label(self, text="準備中...", font=("Meiryo UI", 9))
        self.lbl_status.pack(pady=(15, 5))
        
        self.progress = ttk.Progressbar(self, orient=HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(pady=10)
        
        self.btn_cancel = Button(self, text="中止", command=self.on_cancel, bg="#DC3545", fg="white", font=("Meiryo UI", 10, "bold"), padx=20)
        self.btn_cancel.pack(pady=10)
        
    def on_cancel(self):
        if self.cancel_event.is_set():
            return
        if messagebox.askyesno("確認", "処理を中断しますか？", parent=self):
            self.cancel_event.set()
            self.lbl_status.config(text="停止しています...")
            self.btn_cancel.config(state=DISABLED, text="停止中")

    def update_progress(self, current, total, text):
        if self.cancel_event.is_set():
            return
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.lbl_status.config(text=text)

# ==========================
# 選択処理
# ==========================
def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(
        title="ファイルを選択",
        filetypes=[("DXFファイル", "*.dxf"), ("すべてのファイル", "*.*")]
    )
    if files:
        selected_files = list(files)
        selected_folder = ""
        current_mode = "file"
        update_path_display()
        update_button_state()

def select_folder():
    global selected_folder, selected_files, current_mode
    folder = filedialog.askdirectory(title="フォルダを選択")
    if folder:
        selected_folder = folder
        selected_files = []
        current_mode = "folder"
        update_path_display()
        update_button_state()

def update_path_display():
    text_paths.delete(1.0, END)
    if current_mode == "file":
        text_paths.insert(END, "\n".join(selected_files))
    elif current_mode == "folder":
        text_paths.insert(END, f"フォルダ: {selected_folder}")

def update_button_state():
    state = NORMAL if current_mode else DISABLED
    try:
        btn_extract.config(state=state)
        
        if current_mode == "file":
            btn_file.config(bg="#0D6EFD", fg="white", font=("Meiryo UI", 9, "bold"))
            btn_folder.config(bg="#E9ECEF", fg="black", font=("Meiryo UI", 9))
        elif current_mode == "folder":
            btn_folder.config(bg="#0D6EFD", fg="white", font=("Meiryo UI", 9, "bold"))
            btn_file.config(bg="#E9ECEF", fg="black", font=("Meiryo UI", 9))
        else:
            btn_file.config(bg="#E9ECEF", fg="black", font=("Meiryo UI", 9))
            btn_folder.config(bg="#E9ECEF", fg="black", font=("Meiryo UI", 9))
    except NameError:
        pass # UI構築前（起動時）のエラー回避

# ==========================
# 設定の保存と読み込み処理群
# ==========================
def _get_current_settings_dict():
    global current_mode, selected_files, selected_folder
    kw_list = []
    for row in keyword_entries:
        try:
            replaces = json.loads(row["replaces_var"].get())
        except:
            replaces = []
            
        kw_list.append({
            "col": row["col_var"].get(),
            "format": row["format_var"].get(),
            "xmin": row["xmin_var"].get(),
            "xmax": row["xmax_var"].get(),
            "ymin": row["ymin_var"].get(),
            "ymax": row["ymax_var"].get(),
            "sample": row.get("sample_text", ""),
            "exclude": row["exclude_var"].get(),
            "replaces": replaces
        })

    geom = root.geometry()
    m = re.match(r"(\d+)x(\d+)([-+]\d+)([-+]\d+)", geom)
    if m:
        w, h, x, y = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
    else:
        w = root.winfo_width()
        h = root.winfo_height()
        x = root.winfo_x()
        y = root.winfo_y()

    if w < 100 or h < 100:
        w, h = 860, 850

    return {
        "window_width": w,
        "window_height": h,
        "window_x": x,
        "window_y": y,
        "current_mode": current_mode,
        "selected_files": selected_files,
        "selected_folder": selected_folder,
        "mode": mode_var.get(),
        "save_option": save_option.get(),
        "threshold": threshold_var.get(),
        "base_kw": base_kw_var.get(),
        "base_kw2": base_kw2_var.get(),
        "base_dist": base_dist_var.get(),
        "keywords": kw_list
    }

def _apply_settings_dict(settings, apply_env=False):
    global current_mode, selected_files, selected_folder
    
    if apply_env:
        w = settings.get("window_width")
        h = settings.get("window_height")
        x = settings.get("window_x")
        y = settings.get("window_y")
        
        if w is not None and h is not None and w >= 100 and h >= 100:
            if x is not None and y is not None:
                try:
                    sw = root.winfo_screenwidth()
                    sh = root.winfo_screenheight()
                    if x > sw - 100 or y > sh - 100 or x < -w + 100 or y < -h + 100:
                        x, y = 100, 100
                except:
                    pass
                root.geometry(f"{w}x{h}{x:+d}{y:+d}")
            else:
                root.geometry(f"{w}x{h}")
                
        if "current_mode" in settings:
            current_mode = settings["current_mode"]
        if "selected_files" in settings:
            selected_files = settings["selected_files"]
        if "selected_folder" in settings:
            selected_folder = settings["selected_folder"]
            
        update_path_display()
        update_button_state()

    if "mode" in settings:
        mode_var.set(settings["mode"])
        try: toggle_mode()
        except: pass
    if "save_option" in settings:
        save_option.set(settings["save_option"])
    if "threshold" in settings:
        threshold_var.set(settings["threshold"])
        
    if "base_kw" in settings:
        base_kw_var.set(settings["base_kw"])
    elif "keywords" in settings and len(settings["keywords"]) > 0 and "kw" in settings["keywords"][0]:
        base_kw_var.set(settings["keywords"][0].get("kw", ""))
        
    if "base_kw2" in settings:
        base_kw2_var.set(settings["base_kw2"])
    elif "keywords" in settings and len(settings["keywords"]) > 0 and "kw2" in settings["keywords"][0]:
        base_kw2_var.set(settings["keywords"][0].get("kw2", ""))

    if "base_dist" in settings:
        base_dist_var.set(settings["base_dist"])
        
    if "keywords" in settings:
        for row in list(keyword_ui_frames):
            row.destroy()
        keyword_entries.clear()
        keyword_ui_frames.clear()
        
        for kw_data in settings["keywords"]:
            # 古い設定ファイルからの互換性対応
            replaces = kw_data.get("replaces")
            if replaces is None:
                rb = kw_data.get("replace_before", "")
                ra = kw_data.get("replace_after", "")
                if rb:
                    replaces = [{"before": rb, "after": ra}]
                else:
                    replaces = []
                    
            add_keyword_row(
                col=kw_data.get("col", ""),
                format_type=kw_data.get("format", "標準"),
                xmin=kw_data.get("xmin", 0.0),
                xmax=kw_data.get("xmax", 0.0),
                ymin=kw_data.get("ymin", 0.0),
                ymax=kw_data.get("ymax", 0.0),
                sample_text=kw_data.get("sample", ""),
                exclude_text=kw_data.get("exclude", ""),
                replaces=replaces
            )
    
    root.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

def save_env_settings():
    try:
        user_dir = os.path.expanduser('~')
        app_dir = os.path.join(user_dir, 'DxfExtractionMiya')
        os.makedirs(app_dir, exist_ok=True)
        settings_file = os.path.join(app_dir, 'settings.json')
        
        settings = _get_current_settings_dict()
        
        with open(settings_file, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
            
        messagebox.showinfo("保存完了", f"環境設定を保存しました。\n次回起動時に自動で適用されます。")
    except Exception as e:
        messagebox.showerror("保存エラー", f"環境設定の保存中にエラーが発生しました。\n{e}")

def load_env_settings():
    try:
        user_dir = os.path.expanduser('~')
        app_dir = os.path.join(user_dir, 'DxfExtractionMiya')
        settings_file = os.path.join(app_dir, 'settings.json')
        
        if not os.path.exists(settings_file):
            return
            
        with open(settings_file, 'r', encoding='utf-8') as f:
            settings = json.load(f)
            
        _apply_settings_dict(settings, apply_env=True)
    except Exception as e:
        pass # 起動時の自動読込はエラーを出さずにスキップ

def save_extract_settings():
    try:
        user_dir = os.path.expanduser('~')
        app_dir = os.path.join(user_dir, 'DxfExtractionMiya')
        os.makedirs(app_dir, exist_ok=True)
        
        settings_file = filedialog.asksaveasfilename(
            title="現在のDXF抽出詳細設定を保存",
            initialdir=app_dir,
            initialfile="extraction_settings.json",
            defaultextension=".json",
            filetypes=[("JSONファイル", "*.json"), ("すべてのファイル", "*.*")]
        )
        
        if not settings_file:
            return
            
        settings = _get_current_settings_dict()
        
        with open(settings_file, 'w', encoding='utf-8') as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
            
        messagebox.showinfo("保存完了", f"抽出設定を保存しました。\n保存先: {settings_file}")
    except Exception as e:
        messagebox.showerror("保存エラー", f"抽出設定の保存中にエラーが発生しました。\n{e}")

def load_extract_settings():
    try:
        user_dir = os.path.expanduser('~')
        app_dir = os.path.join(user_dir, 'DxfExtractionMiya')
        
        settings_file = filedialog.askopenfilename(
            title="保存したDXF抽出詳細設定を読込",
            initialdir=app_dir,
            filetypes=[("JSONファイル", "*.json"), ("すべてのファイル", "*.*")]
        )
        
        if not settings_file:
            return
            
        with open(settings_file, 'r', encoding='utf-8') as f:
            settings = json.load(f)
            
        _apply_settings_dict(settings, apply_env=False)
        messagebox.showinfo("読込完了", "抽出設定を読み込みました。")
    except Exception as e:
        messagebox.showerror("読込エラー", f"抽出設定の読込中にエラーが発生しました。\n{e}")

def get_save_dir(original_path=None):
    if save_option.get() == 1 and original_path:
        return os.path.dirname(original_path)
    else:
        folder = filedialog.askdirectory(title="保存先フォルダを選択")
        return folder

def open_preview():
    target_dxf = None
    if current_mode == "file":
        for f in selected_files:
            if f.lower().endswith('.dxf'):
                target_dxf = f
                break
    elif current_mode == "folder" and selected_folder:
        try:
            for f in os.listdir(selected_folder):
                if f.lower().endswith('.dxf'):
                    target_dxf = os.path.join(selected_folder, f)
                    break
        except Exception:
            pass
        
    if not target_dxf:
        messagebox.showwarning("警告", "対象内にDXFファイルが見つかりません。")
        return
        
    def on_preview_complete(kw_text, kw2_text, base_dist, col_name, format_type, xmin, xmax, ymin, ymax, ext_text, exclude_text, replaces):
        base_kw_var.set(kw_text)
        base_kw2_var.set(kw2_text)
        base_dist_var.set(base_dist)
        if not col_name:
            col_name = f"抽出列{len(keyword_entries) + 1}"
        add_keyword_row(col_name, format_type, xmin, xmax, ymin, ymax, ext_text, exclude_text, replaces)

    PreviewDialog(root, target_dxf, on_preview_complete, base_kw_var.get(), base_kw2_var.get())

def extract_dxf_text():
    target_files = []
    if current_mode == "file":
        target_files = [f for f in selected_files if f.lower().endswith('.dxf')]
    elif current_mode == "folder" and selected_folder:
        target_files = [os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith('.dxf')]

    if not target_files:
        messagebox.showwarning("警告", "対象のDXFファイルがありません。")
        return

    save_dir = get_save_dir(target_files[0])
    if not save_dir: return 

    is_keyword_mode = (mode_var.get() == 1)
    y_threshold = threshold_var.get()
    base_kw_str = base_kw_var.get().replace(" ", "").replace("　", "").lower()
    base_kw2_str = base_kw2_var.get().replace(" ", "").replace("　", "").lower()
    base_dist = base_dist_var.get()
    
    keyword_settings = []
    for cfg in keyword_entries:
        try:
            xmin, xmax = float(cfg["xmin_var"].get()), float(cfg["xmax_var"].get())
            ymin, ymax = float(cfg["ymin_var"].get()), float(cfg["ymax_var"].get())
        except:
            xmin, xmax, ymin, ymax = 0.0, 0.0, 0.0, 0.0
            
        col_name = cfg["col_var"].get().strip()
        exclude_str = cfg["exclude_var"].get().strip()
        
        try:
            reps = json.loads(cfg["replaces_var"].get())
        except:
            reps = []
            
        if not col_name:
            col_name = f"列{keyword_entries.index(cfg)+1}"
            
        keyword_settings.append({
            "col_name": col_name,
            "format": cfg["format_var"].get(),
            "xmin": xmin, "xmax": xmax,
            "ymin": ymin, "ymax": ymax,
            "exclude": exclude_str,
            "replaces": reps
        })

    prog_win = ProgressWindow(root, "DXFテキスト抽出")
    result_queue = queue.Queue()

    def callback(current, total, text):
        root.after(0, lambda: prog_win.update_progress(current, total, text))
        
    def task():
        try:
            res = run_extract_dxf(
                target_files, save_dir, is_keyword_mode, y_threshold, 
                base_kw_str, base_kw2_str, base_dist, keyword_settings, 
                progress_callback=callback, cancel_check=prog_win.cancel_event.is_set
            )
            result_queue.put(("success", res))
        except Exception as e:
            result_queue.put(("error", traceback.format_exc()))

    def check_thread():
        if thread.is_alive():
            root.after(100, check_thread)
        else:
            prog_win.destroy()
            status, res = result_queue.get()
            if status == "success":
                success, msg = res
                if success:
                    messagebox.showinfo("完了", msg)
                else:
                    messagebox.showwarning("結果", msg)
            else:
                messagebox.showerror("抽出エラー", res)

    thread = threading.Thread(target=task, daemon=True)
    thread.start()
    check_thread()

# ==========================
# メニュー用関数群
# ==========================
def show_readme():
    readme_path = None
    # PyInstaller環境の展開先(_MEIPASS)か、開発環境のカレントディレクトリを基準にする
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    
    for fname in ["README.md", "readme.md"]:
        p = os.path.join(base_path, fname)
        if os.path.exists(p):
            readme_path = p
            break
            
    if not readme_path:
        messagebox.showerror("エラー", "READMEファイルが見つかりません。")
        return
        
    try:
        with open(readme_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except Exception as e:
        messagebox.showerror("エラー", f"READMEの読み込みに失敗しました。\n{e}")
        return
        
    readme_win = Toplevel(root)
    readme_win.title("README (使い方)")
    readme_win.geometry("700x600")
    
    text_widget = Text(readme_win, wrap=WORD, font=("Meiryo UI", 10))
    scrollbar = ttk.Scrollbar(readme_win, command=text_widget.yview)
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    scrollbar.pack(side=RIGHT, fill=Y)
    text_widget.pack(side=LEFT, fill=BOTH, expand=True, padx=(10, 0), pady=10)
    
    text_widget.insert(END, content)
    text_widget.config(state=DISABLED) # 読み取り専用

def show_version():
    messagebox.showinfo(
        "バージョン情報", 
        f"{APP_TITLE}\nバージョン: {VERSION}\n\nDXFテキスト抽出 ＆ Excel自動集約ツール"
    )

# ==========================
# UI構築とヘルパー関数
# ==========================
def open_replace_dialog(parent_win, replaces_var, on_update_cb, btn_widget):
    dlg = Toplevel(parent_win)
    dlg.title("詳細置換設定 (最大10件)")
    dlg.geometry("320x400")
    dlg.grab_set()
    dlg.resizable(False, False)
    
    Label(dlg, text="置換前 ⇒ 置換後（上から順に適用されます）", font=("Meiryo UI", 9)).pack(pady=(10, 5))
    
    try:
        current_replaces = json.loads(replaces_var.get())
    except:
        current_replaces = []
        
    entries = []
    frame = Frame(dlg)
    frame.pack(fill=BOTH, expand=True, padx=10)
    
    for i in range(10):
        row_f = Frame(frame)
        row_f.pack(fill=X, pady=2)
        Label(row_f, text=f"{i+1}:", font=("Meiryo UI", 9), width=3).pack(side=LEFT)
        v_b = StringVar()
        v_a = StringVar()
        if i < len(current_replaces):
            v_b.set(current_replaces[i].get("before", ""))
            v_a.set(current_replaces[i].get("after", ""))
            
        Entry(row_f, textvariable=v_b, width=12).pack(side=LEFT, padx=2)
        Label(row_f, text="⇒", font=("Meiryo UI", 9)).pack(side=LEFT)
        Entry(row_f, textvariable=v_a, width=12).pack(side=LEFT, padx=2)
        entries.append((v_b, v_a))
        
    def save():
        new_replaces = []
        for v_b, v_a in entries:
            b = v_b.get()
            a = v_a.get()
            if b:
                new_replaces.append({"before": b, "after": a})
        replaces_var.set(json.dumps(new_replaces))
        if btn_widget:
            btn_widget.config(text=f"⚙ 設定 ({len(new_replaces)})")
        if on_update_cb:
            on_update_cb()
        dlg.destroy()
        
    Button(dlg, text="保存して閉じる", command=save, bg="#0D6EFD", fg="white", font=("Meiryo UI", 9, "bold")).pack(pady=10)

root = Tk()
root.title(f"{APP_TITLE} {VERSION}")
root.geometry("860x800")
root.minsize(800, 600)
root.configure(bg="#F8F9FA")

# メニューバーの構築
menu_bar = Menu(root)
root.config(menu=menu_bar)

help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="ヘルプ", menu=help_menu)
help_menu.add_command(label="README (使い方)", command=show_readme)
help_menu.add_command(label="バージョン情報", command=show_version)

style = ttk.Style()
style.theme_use("clam")

# UI変数の初期化
base_kw_var = StringVar()
base_kw2_var = StringVar()
base_dist_var = DoubleVar(value=0.0)

header_frame = Frame(root, bg="#0D6EFD", pady=15)
header_frame.pack(fill=X)
Label(header_frame, text=f"{APP_TITLE} {VERSION}", font=("Meiryo UI", 16, "bold"), bg="#0D6EFD", fg="white").pack()
Label(header_frame, text="DXFテキスト抽出 ＆ Excel自動集約ツール", font=("Meiryo UI", 10), bg="#0D6EFD", fg="white").pack()

container = Frame(root, bg="#F8F9FA")
container.pack(fill=BOTH, expand=True)
canvas = Canvas(container, bg="#F8F9FA", highlightthickness=0)
scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
scrollable_frame = Frame(canvas, bg="#F8F9FA", padx=20, pady=10)

scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_width())
canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas.find_withtag("all")[0], width=e.width))
canvas.configure(yscrollcommand=scrollbar.set)
canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

def on_main_mousewheel(e):
    # e.widgetのトップレベルウィンドウがメインウィンドウ(root)の場合のみスクロールさせる
    if e.widget.winfo_toplevel() == root:
        canvas.yview_scroll(int(-1*(e.delta/120)), "units")

root.bind_all("<MouseWheel>", on_main_mousewheel)

main_frame = scrollable_frame

# --- 1. 対象の選択 ---
Label(main_frame, text="1. 対象の選択", font=("Meiryo UI", 11, "bold"), bg="#F8F9FA").pack(anchor=W, pady=(10, 5))
btn_frame = Frame(main_frame, bg="#F8F9FA")
btn_frame.pack(fill=X)
btn_file = Button(btn_frame, text="📄 DXFファイルを選択", command=select_files, width=20, bg="#E9ECEF", relief=GROOVE)
btn_file.pack(side=LEFT, padx=5, pady=5)
btn_folder = Button(btn_frame, text="📁 DXFフォルダを選択", command=select_folder, width=20, bg="#E9ECEF", relief=GROOVE)
btn_folder.pack(side=LEFT, padx=5, pady=5)

Label(main_frame, text="選択中のパス:", bg="#F8F9FA").pack(anchor=W, pady=(5, 0))
text_paths = Text(main_frame, height=3, font=("Meiryo UI", 9))
text_paths.pack(fill=X, pady=5)

# --- 2. 設定 ---
Label(main_frame, text="2. 設定", font=("Meiryo UI", 11, "bold"), bg="#F8F9FA").pack(anchor=W, pady=(15, 5))

# 環境設定保存ボタン
env_btn_frame = Frame(main_frame, bg="#F8F9FA")
env_btn_frame.pack(fill=X, pady=(0, 5))
Button(env_btn_frame, text="💾 現在のウィンドウサイズなどの環境設定を保存", command=save_env_settings, bg="#E9ECEF", relief=GROOVE).pack(side=LEFT, padx=5)

settings_frame = Frame(main_frame, bg="#FFFFFF", padx=15, pady=15, relief=SOLID, bd=1)
settings_frame.pack(fill=X, pady=5)

Label(settings_frame, text="【保存先】", font=("Meiryo UI", 9, "bold"), bg="#FFFFFF").pack(anchor=W)
save_option = IntVar(value=1)
Radiobutton(settings_frame, text="元のファイルと同じフォルダに保存", variable=save_option, value=1, bg="#FFFFFF").pack(anchor=W)
Radiobutton(settings_frame, text="実行時に任意のフォルダを指定して保存", variable=save_option, value=2, bg="#FFFFFF").pack(anchor=W)

Frame(settings_frame, height=1, bg="#DEE2E6").pack(fill=X, pady=10)

Label(settings_frame, text="【DXF抽出 詳細設定】", font=("Meiryo UI", 9, "bold"), bg="#FFFFFF").pack(anchor=W)

# --- 抽出設定 保存・読込エリア ---
settings_btn_frame = Frame(settings_frame, bg="#FFFFFF")
settings_btn_frame.pack(fill=X, pady=(0, 10))
Button(settings_btn_frame, text="💾 現在の抽出設定を保存", command=save_extract_settings, bg="#E9ECEF", relief=GROOVE).pack(side=RIGHT, padx=2)
Button(settings_btn_frame, text="📂 保存した抽出設定を読込", command=load_extract_settings, bg="#E9ECEF", relief=GROOVE).pack(side=RIGHT, padx=2)
# ------------------------

mode_var = IntVar(value=1)
def toggle_mode():
    state = NORMAL if mode_var.get() == 1 else DISABLED
    for child in kw_container.winfo_children():
        try: child.configure(state=state)
        except: pass
    for row in keyword_ui_frames:
        for child in row.winfo_children():
            try: child.configure(state=state)
            except: pass

Radiobutton(settings_frame, text="1. プレビュー画面から抽出範囲を選択して抽出（複数ファイル時は自動集約）", variable=mode_var, value=1, bg="#FFFFFF", command=toggle_mode).pack(anchor=W, pady=(5, 0))

kw_container = Frame(settings_frame, bg="#F0F4F8", padx=10, pady=10)
kw_container.pack(fill=X, pady=5, padx=20)

# 共通基準文字フレーム
base_frame = Frame(kw_container, bg="#F0F4F8")
base_frame.pack(fill=X, pady=(0, 5))
Label(base_frame, text="▼ 図面全体の共通基準文字", font=("Meiryo UI", 9, "bold"), bg="#F0F4F8").pack(anchor=W)

bk_inner = Frame(base_frame, bg="#F0F4F8")
bk_inner.pack(fill=X, pady=2)
Label(bk_inner, text="第1基準:", font=("Meiryo UI", 9), bg="#F0F4F8").pack(side=LEFT)
Entry(bk_inner, textvariable=base_kw_var, width=15).pack(side=LEFT, padx=5)

Label(bk_inner, text="第2基準(任意):", font=("Meiryo UI", 9), bg="#F0F4F8").pack(side=LEFT, padx=5)
Entry(bk_inner, textvariable=base_kw2_var, width=15).pack(side=LEFT, padx=5)

def clear_base_keywords():
    base_kw_var.set("")
    base_kw2_var.set("")
    base_dist_var.set(0.0)

Button(bk_inner, text="クリア", command=clear_base_keywords, bg="#6C757D", fg="white", font=("Meiryo UI", 8), padx=5).pack(side=LEFT, padx=(10, 0))

Label(kw_container, text="▼ 抽出項目（プレビュー画面で連続追加可能）", font=("Meiryo UI", 9, "bold"), bg="#F0F4F8").pack(anchor=W, pady=(10,0))

kw_btn_frame = Frame(kw_container, bg="#F0F4F8")
kw_btn_frame.pack(fill=X, pady=5)
Button(kw_btn_frame, text="＋ プレビューを開いて基準文字と抽出範囲を設定", command=open_preview, bg="#0D6EFD", fg="white", font=("Meiryo UI", 9, "bold"), padx=10).pack(side=LEFT)

def clear_all_keywords():
    if not keyword_entries:
        return
    if messagebox.askyesno("確認", "追加されたすべての抽出項目をクリアしますか？"):
        for row in list(keyword_ui_frames):
            row.destroy()
        keyword_entries.clear()
        keyword_ui_frames.clear()
        root.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

Button(kw_btn_frame, text="🗑 すべてクリア", command=clear_all_keywords, bg="#DC3545", fg="white", font=("Meiryo UI", 9, "bold"), padx=10).pack(side=LEFT, padx=10)


kw_list_frame = Frame(kw_container, bg="#F0F4F8")
kw_list_frame.pack(fill=X, pady=5)

keyword_entries = []     # 内部データ用
keyword_ui_frames = []   # UI削除用

def add_keyword_row(col="", format_type="標準", xmin=0.0, xmax=0.0, ymin=0.0, ymax=0.0, sample_text="", exclude_text="", replaces=None):
    if replaces is None:
        replaces = []

    row_frame = Frame(kw_list_frame, bg="#FFFFFF", pady=6, padx=10, relief=SOLID, bd=1)
    row_frame.pack(fill=X, pady=3)

    top_frame = Frame(row_frame, bg="#FFFFFF")
    top_frame.pack(fill=X, pady=(0, 3))
    
    Label(top_frame, text="出力列名:", font=("Meiryo UI", 8, "bold"), bg="#FFFFFF", fg="#495057").pack(side=LEFT)
    col_var = StringVar(value=col)
    Entry(top_frame, textvariable=col_var, width=12).pack(side=LEFT, padx=2)

    Label(top_frame, text="表示形式:", font=("Meiryo UI", 8), bg="#FFFFFF", fg="#495057").pack(side=LEFT, padx=(5,0))
    format_var = StringVar(value=format_type)
    format_cb = ttk.Combobox(top_frame, textvariable=format_var, values=["標準", "数値", "通貨", "会計", "日付", "時刻", "パーセンテージ", "分数", "指数", "文字列"], width=8, state="readonly")
    format_cb.pack(side=LEFT, padx=2)

    Label(top_frame, text="置換:", font=("Meiryo UI", 8), bg="#FFFFFF", fg="#495057").pack(side=LEFT, padx=(5,0))
    replaces_var = StringVar(value=json.dumps(replaces))
    btn_replace = Button(top_frame, text=f"⚙ 設定 ({len(replaces)})", font=("Meiryo UI", 8), bg="#E9ECEF",
                         command=lambda: open_replace_dialog(root, replaces_var, update_sample_label, btn_replace))
    btn_replace.pack(side=LEFT, padx=2)

    Label(top_frame, text="除外:", font=("Meiryo UI", 8), bg="#FFFFFF", fg="#495057").pack(side=LEFT, padx=(5,0))
    exclude_var = StringVar(value=exclude_text)
    Entry(top_frame, textvariable=exclude_var, width=10).pack(side=LEFT, padx=2)

    sample_label = Label(top_frame, text="", font=("Meiryo UI", 8, "bold"), bg="#FFFFFF", fg="#198754")
    if sample_text:
        sample_label.pack(side=LEFT, padx=5)

    def update_sample_label(*args):
        if not sample_text:
            return
            
        res_text = sample_text
        try:
            reps = json.loads(replaces_var.get())
        except:
            reps = []
            
        for rep in reps:
            rep_before = rep.get("before", "")
            rep_after = rep.get("after", "")
            if rep_before:
                try:
                    res_text = re.sub(rep_before, rep_after, res_text)
                except re.error:
                    res_text = res_text.replace(rep_before, rep_after)
                
        excludes = [x.strip() for x in exclude_var.get().split(",") if x.strip()]
        for ex in excludes:
            try:
                res_text = re.sub(ex, "", res_text)
            except re.error:
                res_text = res_text.replace(ex, "")
                
        sample_label.config(text=f"【元】{sample_text} ⇒ 【後】{res_text}")
        
    replaces_var.trace_add("write", update_sample_label)
    exclude_var.trace_add("write", update_sample_label)
    update_sample_label()

    bottom_frame = Frame(row_frame, bg="#FFFFFF")
    bottom_frame.pack(fill=X)

    xmin_var = DoubleVar(value=xmin)
    xmax_var = DoubleVar(value=xmax)
    ymin_var = DoubleVar(value=ymin)
    ymax_var = DoubleVar(value=ymax)

    Label(bottom_frame, text="X範囲:", bg="#FFFFFF", fg="#0D6EFD").pack(side=LEFT)
    Spinbox(bottom_frame, textvariable=xmin_var, from_=-5000, to=5000, increment=0.1, width=6, format="%.3f").pack(side=LEFT)
    Label(bottom_frame, text="~", bg="#FFFFFF").pack(side=LEFT)
    Spinbox(bottom_frame, textvariable=xmax_var, from_=-5000, to=5000, increment=0.1, width=6, format="%.3f").pack(side=LEFT)

    Frame(bottom_frame, width=1, bg="#DEE2E6").pack(side=LEFT, fill=Y, padx=10)

    Label(bottom_frame, text="Y範囲:", bg="#FFFFFF", fg="#198754").pack(side=LEFT)
    Spinbox(bottom_frame, textvariable=ymin_var, from_=-5000, to=5000, increment=0.1, width=6, format="%.3f").pack(side=LEFT)
    Label(bottom_frame, text="~", bg="#FFFFFF").pack(side=LEFT)
    Spinbox(bottom_frame, textvariable=ymax_var, from_=-5000, to=5000, increment=0.1, width=6, format="%.3f").pack(side=LEFT)

    def remove_row():
        row_frame.destroy()
        for item in keyword_entries:
            if item["frame"] == row_frame:
                keyword_entries.remove(item)
                break
        if row_frame in keyword_ui_frames:
            keyword_ui_frames.remove(row_frame)
        root.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    Button(top_frame, text="削除", command=remove_row, bg="#DC3545", fg="white", padx=10).pack(side=RIGHT, padx=5)

    row_data = {
        "col_var": col_var, "format_var": format_var, "xmin_var": xmin_var, "xmax_var": xmax_var,
        "ymin_var": ymin_var, "ymax_var": ymax_var, "exclude_var": exclude_var, "frame": row_frame, "sample_text": sample_text,
        "replaces_var": replaces_var
    }
    keyword_entries.append(row_data)
    keyword_ui_frames.append(row_frame)
    
    root.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))

Radiobutton(settings_frame, text="2. 図面内のすべてのテキストを抽出（複数ファイル時は自動集約）", variable=mode_var, value=2, bg="#FFFFFF", command=toggle_mode).pack(anchor=W, pady=(15, 0))
Label(settings_frame, text="※全体抽出モード用のテキストズレ許容値（閾値）:", bg="#FFFFFF").pack(anchor=W, pady=(5, 0))
threshold_var = DoubleVar(value=20.0)
Spinbox(settings_frame, from_=0.0, to=500.0, increment=1.0, textvariable=threshold_var, width=10).pack(anchor=W, pady=2)


# --- 3. 実行 ---
Label(main_frame, text="3. 実行", font=("Meiryo UI", 11, "bold"), bg="#F8F9FA").pack(anchor=W, pady=(20, 5))
action_frame = Frame(main_frame, bg="#F8F9FA")
action_frame.pack(fill=X)

btn_extract = Button(action_frame, text="🚀 選択対象のDXFを処理してExcelに出力", command=extract_dxf_text, height=2, bg="#198754", fg="white", font=("Meiryo UI", 10, "bold"), state=DISABLED)
btn_extract.pack(fill=X, pady=5)


if __name__ == "__main__":
    # 起動時の自動読込処理 (UI構築直後に実行)
    root.after(100, load_env_settings)
    root.mainloop()