# -*- coding: utf-8 -*-
"""
プレビュー＆範囲選択ダイアログモジュール
"""
import os
import math
import re
import json
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from dxf_core import get_all_elements_from_dxf, zen_to_han_alnum

class PreviewDialog(Toplevel):
    def __init__(self, parent, dxf_path, on_complete, current_base_kw, current_base_kw2):
        super().__init__(parent)
        self.title(f"プレビュー＆範囲選択 - {os.path.basename(dxf_path)}")
        self.geometry("1200x800")
        
        try: self.state('zoomed')
        except:
            try: self.attributes('-zoomed', True)
            except: pass

        self.grab_set()
        
        self.on_complete = on_complete
        self.texts, self.shapes = get_all_elements_from_dxf(dxf_path)
        
        self.scale = 1.0
        self.base_scale = 1.0
        self.min_x = 0.0
        self.max_x = 0.0
        self.min_y = 0.0
        self.max_y = 0.0
        
        self.anchor = None
        self.anchor2 = None
        
        if current_base_kw:
            kw_c = zen_to_han_alnum(current_base_kw).replace(" ", "").replace("　", "").lower()
            for t in self.texts:
                tc = t['text'].replace(" ", "").replace("　", "").lower()
                if tc == kw_c:
                    self.anchor = t
                    break
                    
        if current_base_kw2:
            kw2_c = zen_to_han_alnum(current_base_kw2).replace(" ", "").replace("　", "").lower()
            for t in self.texts:
                tc = t['text'].replace(" ", "").replace("　", "").lower()
                if tc == kw2_c:
                    self.anchor2 = t
                    break

        if self.anchor:
            self.mode = StringVar(value="rect")
        else:
            self.mode = StringVar(value="anchor")
            
        self.mode.trace_add("write", self.update_radio_colors)
            
        self.rect_dxf_start = None
        self.rect_dxf_end = None
        self.is_dragging = False
        
        self.setup_ui()
        self.init_transform()
        self.draw()
        self.update_radio_colors()
        
    def setup_ui(self):
        toolbar = Frame(self, bg="#E9ECEF", pady=10, padx=10)
        toolbar.pack(fill=X)
        
        Label(toolbar, text="手順: ", font=("Meiryo UI", 10, "bold"), bg="#E9ECEF").pack(side=LEFT)
        self.rb1 = Radiobutton(toolbar, text="1. 第1基準文字を選択", variable=self.mode, value="anchor", indicatoron=0, bg="#F8D7DA", selectcolor="#DC3545", padx=10, pady=5)
        self.rb1.pack(side=LEFT, padx=5)
        self.rb2 = Radiobutton(toolbar, text="2. 第2基準文字を選択(任意)", variable=self.mode, value="anchor2", indicatoron=0, bg="#FFE5D0", selectcolor="#FD7E14", padx=10, pady=5)
        self.rb2.pack(side=LEFT, padx=5)
        self.rb3 = Radiobutton(toolbar, text="3. 抽出範囲をドラッグ", variable=self.mode, value="rect", indicatoron=0, bg="#CFE2FF", selectcolor="#0D6EFD", padx=10, pady=5)
        self.rb3.pack(side=LEFT, padx=5)
        
        Label(toolbar, text="※ ホイール:ズーム  |  右ドラッグ:移動  |  Shift+ホイール:横スクロール  |  Ctrl+ホイール:縦スクロール", fg="#6C757D", bg="#E9ECEF").pack(side=LEFT, padx=20)
        
        btn_frame = Frame(toolbar, bg="#E9ECEF")
        btn_frame.pack(side=RIGHT, padx=5)
        Button(btn_frame, text="＋ 設定を追加して次へ", command=self.confirm, bg="#0D6EFD", fg="white", font=("Meiryo UI", 10, "bold"), padx=10).pack(side=LEFT, padx=5)
        Button(btn_frame, text="完了して閉じる", command=self.finish_and_close, bg="#6C757D", fg="white", font=("Meiryo UI", 10, "bold"), padx=10).pack(side=LEFT, padx=5)
        Button(btn_frame, text="中止", command=self.cancel_and_close, bg="#DC3545", fg="white", font=("Meiryo UI", 10, "bold"), padx=10).pack(side=LEFT, padx=5)
        
        # オプション設定バーを追加
        options_bar = Frame(self, bg="#F8F9FA", pady=5, padx=15, relief=SOLID, bd=1)
        options_bar.pack(fill=X, pady=(0, 5))

        Label(options_bar, text="出力列名:", font=("Meiryo UI", 9), bg="#F8F9FA").pack(side=LEFT)
        self.col_name_var = StringVar(value="")
        Entry(options_bar, textvariable=self.col_name_var, width=15).pack(side=LEFT, padx=(2, 15))
        
        Label(options_bar, text="表示形式:", font=("Meiryo UI", 9), bg="#F8F9FA").pack(side=LEFT)
        self.format_var = StringVar(value="標準")
        format_cb = ttk.Combobox(options_bar, textvariable=self.format_var, values=["標準", "数値", "通貨", "会計", "日付", "時刻", "パーセンテージ", "分数", "指数", "文字列"], width=8, state="readonly")
        format_cb.pack(side=LEFT, padx=(2, 15))
        self.format_var.trace_add("write", lambda *args: self.update_preview_text())
        
        Label(options_bar, text="置換:", font=("Meiryo UI", 9), bg="#F8F9FA").pack(side=LEFT)
        self.replaces_var = StringVar(value="[]")
        self.replaces_var.trace_add("write", lambda *args: self.update_preview_text())
        self.btn_replace = Button(options_bar, text="⚙ 設定 (0)", font=("Meiryo UI", 8), bg="#E9ECEF", command=self.open_replace_dialog)
        self.btn_replace.pack(side=LEFT, padx=(2, 15))

        Label(options_bar, text="除外文字(カンマ区切/正規表現可):", font=("Meiryo UI", 9), bg="#F8F9FA").pack(side=LEFT)
        self.exclude_var = StringVar(value="")
        self.exclude_var.trace_add("write", lambda *args: self.update_preview_text())
        Entry(options_bar, textvariable=self.exclude_var, width=15).pack(side=LEFT, padx=2)
        
        preview_bar = Frame(self, bg="#D1E7DD", pady=8, padx=15)
        preview_bar.pack(fill=X)
        self.preview_text_var = StringVar()
        self.preview_entry = Entry(preview_bar, textvariable=self.preview_text_var, font=("Meiryo UI", 12, "bold"), bg="#D1E7DD", fg="#0F5132", readonlybackground="#D1E7DD", relief=FLAT)
        self.preview_entry.pack(side=LEFT, fill=X, expand=True)
        self.preview_entry.config(state='readonly')
        
        self.update_preview_text()
        
        canvas_frame = Frame(self, bg="white")
        canvas_frame.pack(fill=BOTH, expand=True)

        self.canvas = Canvas(canvas_frame, bg="white", cursor="crosshair")
        
        self.vbar = ttk.Scrollbar(canvas_frame, orient=VERTICAL, command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(canvas_frame, orient=HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vbar.set, xscrollcommand=self.hbar.set)

        self.vbar.pack(side=RIGHT, fill=Y)
        self.hbar.pack(side=BOTTOM, fill=X)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_left_press)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)
        self.canvas.bind("<ButtonPress-3>", self.on_right_press)
        self.canvas.bind("<B3-Motion>", self.on_right_drag)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.canvas.bind("<Shift-MouseWheel>", self.on_shift_mousewheel)
        self.canvas.bind("<Control-MouseWheel>", self.on_ctrl_mousewheel)
        
    def open_replace_dialog(self):
        dlg = Toplevel(self)
        dlg.title("詳細置換設定 (最大10件)")
        dlg.geometry("320x400")
        dlg.grab_set()
        dlg.resizable(False, False)
        
        Label(dlg, text="置換前 ⇒ 置換後（上から順に適用されます）", font=("Meiryo UI", 9)).pack(pady=(10, 5))
        
        try:
            current_replaces = json.loads(self.replaces_var.get())
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
            self.replaces_var.set(json.dumps(new_replaces))
            self.btn_replace.config(text=f"⚙ 設定 ({len(new_replaces)})")
            dlg.destroy()
            
        Button(dlg, text="保存して閉じる", command=save, bg="#0D6EFD", fg="white", font=("Meiryo UI", 9, "bold")).pack(pady=10)

    def update_radio_colors(self, *args):
        mode = self.mode.get()
        if mode == "anchor":
            self.rb1.config(fg="white")
        else:
            self.rb1.config(fg="black")
            
        if mode == "anchor2":
            self.rb2.config(fg="white")
        else:
            self.rb2.config(fg="black")
            
        if mode == "rect":
            self.rb3.config(fg="white")
        else:
            self.rb3.config(fg="black")

    def init_transform(self):
        xs, ys = [], []
        for t in self.texts:
            xs.append(t["x"])
            ys.append(t["y"])
        for s in self.shapes:
            if s['type'] == 'line':
                xs.extend([s['start'][0], s['end'][0]])
                ys.extend([s['start'][1], s['end'][1]])
            elif s['type'] == 'polyline':
                for pt in s['points']:
                    xs.append(pt[0]); ys.append(pt[1])
            elif s['type'] in {'circle', 'arc'}:
                r, c = s['radius'], s['center']
                xs.extend([c[0] - r, c[0] + r])
                ys.extend([c[1] - r, c[1] + r])

        if not xs or not ys:
            messagebox.showwarning("警告", "描画可能なテキストや図形が見つかりませんでした。", parent=self)
            return

        self.min_x, self.max_x = min(xs), max(xs)
        self.min_y, self.max_y = min(ys), max(ys)
        
        w = self.max_x - self.min_x
        h = self.max_y - self.min_y
        if w == 0: w = 1
        if h == 0: h = 1
        
        self.base_scale = min(900 / w, 700 / h) * 0.9
        self.scale = self.base_scale
        
    def d2c(self, x, y):
        cx = (x - self.min_x) * self.scale + 50
        cy = (self.max_y - y) * self.scale + 50
        return cx, cy
        
    def c2d(self, cx, cy):
        x = (cx - 50) / self.scale + self.min_x
        y = self.max_y - (cy - 50) / self.scale
        return x, y
        
    def _transform_point(self, px, py):
        ax, ay = self.anchor["x"], self.anchor["y"]
        return px - ax, py - ay

    def get_base_dist(self):
        if self.anchor and self.anchor2:
            dx = self.anchor2["x"] - self.anchor["x"]
            dy = self.anchor2["y"] - self.anchor["y"]
            return math.sqrt(dx**2 + dy**2)
        return 0.0

    def get_extracted_text(self):
        if not self.anchor or not self.rect_dxf_start or not self.rect_dxf_end:
            return ""
            
        x1, y1 = self.rect_dxf_start
        x2, y2 = self.rect_dxf_end
        
        pts = [
            self._transform_point(x1, y1), self._transform_point(x2, y1),
            self._transform_point(x1, y2), self._transform_point(x2, y2)
        ]
        xmin = min(p[0] for p in pts)
        xmax = max(p[0] for p in pts)
        ymin = min(p[1] for p in pts)
        ymax = max(p[1] for p in pts)
        
        matched_texts = []
        for t in self.texts:
            if t == self.anchor or (self.anchor2 and t == self.anchor2): continue
            
            t_clean = t['text'].strip()
            if t_clean == self.anchor['text'].strip() or (self.anchor2 and t_clean == self.anchor2['text'].strip()): continue
            
            rep_x = t['x']
            rep_y = t['y'] + (t.get('h', 2.5) * 0.5)
            
            tx, ty = self._transform_point(rep_x, rep_y)
            if xmin <= tx <= xmax and ymin <= ty <= ymax:
                matched_texts.append(t)
        
        if matched_texts:
            matched_texts.sort(key=lambda item: (-item['y'], item['x']))
            return "/".join([t['text'] for t in matched_texts])
        return ""

    def update_preview_text(self):
        if not self.anchor:
            self.preview_text_var.set("【抽出プレビュー】 1. 図面内の「第1基準文字」をクリックしてください")
        elif self.mode.get() == "anchor2":
            self.preview_text_var.set(f"【第1基準: {self.anchor['text']}】 2. スケール・回転補正用の「第2基準文字」をクリックしてください（不要なら次へ）")
        elif not self.rect_dxf_start or not self.rect_dxf_end:
            kw2_txt = f" / 第2: {self.anchor2['text']}" if self.anchor2 else ""
            self.preview_text_var.set(f"【基準設定済: {self.anchor['text']}{kw2_txt}】 3. 抽出したい値の範囲を左ドラッグで囲んでください")
        else:
            ext_text = self.get_extracted_text()
            if ext_text:
                # 置換文字のリアルタイム適用
                try:
                    reps = json.loads(self.replaces_var.get())
                except:
                    reps = []
                for rep in reps:
                    rep_before = rep.get("before", "")
                    rep_after = rep.get("after", "")
                    if rep_before:
                        try:
                            ext_text = re.sub(rep_before, rep_after, ext_text)
                        except re.error:
                            ext_text = ext_text.replace(rep_before, rep_after)
                        
                # 除外文字のリアルタイム適用
                excludes = [x.strip() for x in self.exclude_var.get().split(",") if x.strip()]
                for ex in excludes:
                    try:
                        ext_text = re.sub(ex, "", ext_text)
                    except re.error:
                        ext_text = ext_text.replace(ex, "")
                        
                self.preview_text_var.set(f"【抽出テスト ({self.format_var.get()})】 {ext_text}")
            else:
                self.preview_text_var.set(f"【抽出テスト】 (※選択範囲内にテキストが見つかりません)")

    def draw(self, update_text=True):
        self.canvas.delete("all")
        shape_color = "#C0C0C0"
        
        for s in self.shapes:
            try:
                if s['type'] == 'line':
                    cx1, cy1 = self.d2c(*s['start'])
                    cx2, cy2 = self.d2c(*s['end'])
                    self.canvas.create_line(cx1, cy1, cx2, cy2, fill=shape_color)
                elif s['type'] == 'polyline':
                    pts = s['points']
                    if not pts: continue
                    c_pts = [self.d2c(*pt) for pt in pts]
                    if s.get('closed') and len(c_pts) > 2: c_pts.append(c_pts[0])
                    flat_pts = [coord for pt in c_pts for coord in pt]
                    self.canvas.create_line(flat_pts, fill=shape_color)
                elif s['type'] == 'circle':
                    cx, cy = self.d2c(*s['center'])
                    cr = s['radius'] * self.scale
                    self.canvas.create_oval(cx - cr, cy - cr, cx + cr, cy + cr, outline=shape_color)
                elif s['type'] == 'arc':
                    cx, cy = self.d2c(*s['center'])
                    cr = s['radius'] * self.scale
                    tk_extent = s['end_angle'] - s['start_angle']
                    if tk_extent < 0: tk_extent += 360
                    tk_start = 360 - s['end_angle']
                    self.canvas.create_arc(cx - cr, cy - cr, cx + cr, cy + cr, start=tk_start, extent=tk_extent, outline=shape_color, style=ARC)
            except: pass

        for t in self.texts:
            cx, cy = self.d2c(t["x"], t["y"])
            color = "black"
            px_height = int(t.get("h", 2.5) * self.scale)
            f_size = max(2, px_height)
            
            if self.anchor and self.anchor == t:
                color = "#DC3545"
                f_size = int(f_size * 1.5)
                self.canvas.create_oval(cx-5, cy-5, cx+5, cy+5, fill="#DC3545", outline="#DC3545")
            elif self.anchor2 and self.anchor2 == t:
                color = "#FD7E14"
                f_size = int(f_size * 1.5)
                self.canvas.create_oval(cx-5, cy-5, cx+5, cy+5, fill="#FD7E14", outline="#FD7E14")
                
            self.canvas.create_text(cx, cy, text=t["text"], anchor="sw", fill=color, font=("Meiryo UI", -f_size))
            
        self.draw_rect()
        
        # 固定領域によるスクロール領域設定（図形の有無によらず一定空間を確保）
        cx1, cy1 = self.d2c(self.min_x, self.max_y)
        cx2, cy2 = self.d2c(self.max_x, self.min_y)
        margin = 200
        sr = (min(cx1, cx2) - margin, min(cy1, cy2) - margin, max(cx1, cx2) + margin, max(cy1, cy2) + margin)
        self.canvas.configure(scrollregion=sr)
        
        if update_text:
            self.update_preview_text()

    def draw_rect(self):
        self.canvas.delete("selection_rect")
        if self.rect_dxf_start and self.rect_dxf_end:
            cx1, cy1 = self.d2c(*self.rect_dxf_start)
            cx2, cy2 = self.d2c(*self.rect_dxf_end)
            self.canvas.create_rectangle(cx1, cy1, cx2, cy2, outline="#0D6EFD", dash=(4, 4), width=2, fill="#CFE2FF", stipple="gray25", tags="selection_rect")
            
    def on_left_press(self, event):
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        dx, dy = self.c2d(cx, cy)
        
        if self.mode.get() in ("anchor", "anchor2"):
            best_t, best_dist = None, float('inf')
            for t in self.texts:
                dist = (t["x"] - dx)**2 + (t["y"] - dy)**2
                if dist < best_dist:
                    best_dist = dist; best_t = t
            if best_t:
                if self.mode.get() == "anchor":
                    self.anchor = best_t
                    self.mode.set("anchor2")
                else:
                    self.anchor2 = best_t
                    self.mode.set("rect")
                self.draw()
        elif self.mode.get() == "rect":
            if not self.anchor:
                messagebox.showinfo("案内", "先に「1. 第1基準文字」を選択してください。", parent=self)
                self.mode.set("anchor")
                return
            self.rect_dxf_start = (dx, dy)
            self.rect_dxf_end = (dx, dy)
            self.is_dragging = True
            self.draw_rect()
            
    def on_left_drag(self, event):
        if self.mode.get() == "rect" and self.is_dragging:
            cx = self.canvas.canvasx(event.x)
            cy = self.canvas.canvasy(event.y)
            dx, dy = self.c2d(cx, cy)
            self.rect_dxf_end = (dx, dy)
            self.draw_rect()
            
    def on_left_release(self, event):
        if self.mode.get() == "rect" and self.is_dragging:
            self.is_dragging = False
            self.update_preview_text()
            
    def on_right_press(self, event):
        self.canvas.scan_mark(event.x, event.y)
        
    def on_right_drag(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)
        
    def on_mousewheel(self, event):
        zoom = 1.2 if event.delta > 0 else 0.8
        
        # ズーム前のマウス直下のCanvas座標とDXF座標を取得
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        dx, dy = self.c2d(cx, cy)
        
        # スケールを更新して再描画
        self.scale *= zoom
        self.draw(update_text=False)
        self.canvas.update_idletasks() # scrollregionの更新を確実に反映
        
        # ズーム後のマウス直下に来るべき新しいCanvas座標
        new_cx, new_cy = self.d2c(dx, dy)
        
        sr_val = self.canvas.cget("scrollregion")
        if sr_val:
            sr_x, sr_y, sr_x2, sr_y2 = [float(v) for v in sr_val.split()]
            sr_w = sr_x2 - sr_x
            sr_h = sr_y2 - sr_y
            
            # マウスポインタ(event.x, event.y)に新しいCanvas座標が来るように
            # ウィンドウ左上(0, 0)に来るべきCanvas座標を逆算
            target_left = new_cx - event.x
            target_top = new_cy - event.y
            
            # scrollregionに対する割合を計算してスクロールを移動
            fx = (target_left - sr_x) / sr_w if sr_w > 0 else 0
            fy = (target_top - sr_y) / sr_h if sr_h > 0 else 0
            
            self.canvas.xview_moveto(fx)
            self.canvas.yview_moveto(fy)
            
        return "break"
        
    def on_shift_mousewheel(self, event):
        self.canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        return "break"

    def on_ctrl_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        return "break"
        
    def confirm(self):
        if not self.anchor:
            messagebox.showwarning("エラー", "第1基準文字が選択されていません。", parent=self)
            return
        if not self.rect_dxf_start or not self.rect_dxf_end:
            messagebox.showwarning("エラー", "抽出範囲が選択されていません。\n左ドラッグで範囲を四角く囲んでください。", parent=self)
            return
            
        x1, y1 = self.rect_dxf_start
        x2, y2 = self.rect_dxf_end
        
        pts = [
            self._transform_point(x1, y1), self._transform_point(x2, y1),
            self._transform_point(x1, y2), self._transform_point(x2, y2)
        ]
        xmin = min(p[0] for p in pts)
        xmax = max(p[0] for p in pts)
        ymin = min(p[1] for p in pts)
        ymax = max(p[1] for p in pts)
        
        ext_text = self.get_extracted_text()
        
        kw_text = self.anchor["text"]
        kw2_text = self.anchor2["text"] if self.anchor2 else ""
        base_dist = self.get_base_dist()
        format_type = self.format_var.get()
        exclude_text = self.exclude_var.get()
        try:
            replaces = json.loads(self.replaces_var.get())
        except:
            replaces = []
        col_name = self.col_name_var.get()
        
        self.on_complete(kw_text, kw2_text, base_dist, col_name, format_type, xmin, xmax, ymin, ymax, ext_text, exclude_text, replaces)
        
        self.rect_dxf_start = None
        self.rect_dxf_end = None
        self.mode.set("rect")
        self.exclude_var.set("") # 連続追加のためクリア
        self.replaces_var.set("[]")
        self.col_name_var.set("")
        self.btn_replace.config(text="⚙ 設定 (0)")
        
        self.draw(update_text=False)
        self.preview_text_var.set(f"✅ 抽出項目を追加しました！続けて次の抽出範囲をドラッグしてください。")

    def finish_and_close(self):
        if self.anchor and self.rect_dxf_start and self.rect_dxf_end:
            x1, y1 = self.rect_dxf_start
            x2, y2 = self.rect_dxf_end
            
            pts = [
                self._transform_point(x1, y1), self._transform_point(x2, y1),
                self._transform_point(x1, y2), self._transform_point(x2, y2)
            ]
            xmin = min(p[0] for p in pts)
            xmax = max(p[0] for p in pts)
            ymin = min(p[1] for p in pts)
            ymax = max(p[1] for p in pts)
            
            ext_text = self.get_extracted_text()
                
            kw_text = self.anchor["text"]
            kw2_text = self.anchor2["text"] if self.anchor2 else ""
            base_dist = self.get_base_dist()
            format_type = self.format_var.get()
            exclude_text = self.exclude_var.get()
            try:
                replaces = json.loads(self.replaces_var.get())
            except:
                replaces = []
            col_name = self.col_name_var.get()
            
            self.on_complete(kw_text, kw2_text, base_dist, col_name, format_type, xmin, xmax, ymin, ymax, ext_text, exclude_text, replaces)
            
        self.destroy()

    def cancel_and_close(self):
        self.destroy()