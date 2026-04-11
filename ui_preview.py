# -*- coding: utf-8 -*-
"""
プレビュー＆範囲選択ダイアログモジュール
"""
import os
import math
from tkinter import *
from tkinter import messagebox
from dxf_core import get_all_elements_from_dxf

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
        self.offset_x = 0.0
        self.offset_y = 0.0
        self.min_x = 0.0
        self.max_y = 0.0
        
        self.anchor = None
        self.anchor2 = None
        
        if current_base_kw:
            kw_c = current_base_kw.replace(" ", "").replace("　", "").lower()
            for t in self.texts:
                tc = t['text'].replace(" ", "").replace("　", "").lower()
                if tc == kw_c:
                    self.anchor = t
                    break
                    
        if current_base_kw2:
            kw2_c = current_base_kw2.replace(" ", "").replace("　", "").lower()
            for t in self.texts:
                tc = t['text'].replace(" ", "").replace("　", "").lower()
                if tc == kw2_c:
                    self.anchor2 = t
                    break

        if self.anchor:
            self.mode = StringVar(value="rect")
        else:
            self.mode = StringVar(value="anchor")
            
        self.rect_dxf_start = None
        self.rect_dxf_end = None
        self.is_dragging = False
        self.pan_start_x = 0
        self.pan_start_y = 0
        
        self.setup_ui()
        self.init_transform()
        self.draw()
        
    def setup_ui(self):
        toolbar = Frame(self, bg="#E9ECEF", pady=10, padx=10)
        toolbar.pack(fill=X)
        
        Label(toolbar, text="手順: ", font=("Meiryo UI", 10, "bold"), bg="#E9ECEF").pack(side=LEFT)
        self.rb1 = Radiobutton(toolbar, text="1. 第1基準文字を選択", variable=self.mode, value="anchor", indicatoron=0, bg="#FFF3CD", selectcolor="#FFC107", padx=10, pady=5)
        self.rb1.pack(side=LEFT, padx=5)
        self.rb2 = Radiobutton(toolbar, text="2. 第2基準文字を選択(任意)", variable=self.mode, value="anchor2", indicatoron=0, bg="#CFF4FC", selectcolor="#0DCAF0", padx=10, pady=5)
        self.rb2.pack(side=LEFT, padx=5)
        self.rb3 = Radiobutton(toolbar, text="3. 抽出範囲をドラッグ", variable=self.mode, value="rect", indicatoron=0, bg="#D1E7DD", selectcolor="#198754", padx=10, pady=5)
        self.rb3.pack(side=LEFT, padx=5)
        
        Label(toolbar, text="※ ホイール:ズーム  |  右ドラッグ:移動", fg="#6C757D", bg="#E9ECEF").pack(side=LEFT, padx=20)
        
        btn_frame = Frame(toolbar, bg="#E9ECEF")
        btn_frame.pack(side=RIGHT, padx=5)
        Button(btn_frame, text="＋ 設定を追加して次へ", command=self.confirm, bg="#0D6EFD", fg="white", font=("Meiryo UI", 10, "bold"), padx=10).pack(side=LEFT, padx=5)
        Button(btn_frame, text="完了して閉じる", command=self.finish_and_close, bg="#6C757D", fg="white", font=("Meiryo UI", 10, "bold"), padx=10).pack(side=LEFT, padx=5)
        
        preview_bar = Frame(self, bg="#D1E7DD", pady=8, padx=15)
        preview_bar.pack(fill=X)
        self.preview_label = Label(preview_bar, text="", font=("Meiryo UI", 12, "bold"), bg="#D1E7DD", fg="#0F5132")
        self.preview_label.pack(side=LEFT)
        self.update_preview_text()
        
        self.canvas = Canvas(self, bg="white", cursor="crosshair")
        self.canvas.pack(fill=BOTH, expand=True)
        
        self.canvas.bind("<ButtonPress-1>", self.on_left_press)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_release)
        self.canvas.bind("<ButtonPress-3>", self.on_right_press)
        self.canvas.bind("<B3-Motion>", self.on_right_drag)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        
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

        self.min_x, max_x = min(xs), max(xs)
        min_y, self.max_y = min(ys), max(ys)
        
        w = max_x - self.min_x
        h = self.max_y - min_y
        if w == 0: w = 1
        if h == 0: h = 1
        
        self.base_scale = min(900 / w, 700 / h) * 0.9
        self.scale = self.base_scale
        self.offset_x = 50
        self.offset_y = 50
        
    def d2c(self, x, y):
        cx = (x - self.min_x) * self.scale + self.offset_x
        cy = (self.max_y - y) * self.scale + self.offset_y
        return cx, cy
        
    def c2d(self, cx, cy):
        x = (cx - self.offset_x) / self.scale + self.min_x
        y = self.max_y - (cy - self.offset_y) / self.scale
        return x, y
        
    def _transform_point(self, px, py):
        ax, ay = self.anchor["x"], self.anchor["y"]
        # 抽出枠の意図しない肥大化を防ぐため、回転・スケール補正を無効化
        # 第1基準文字からの単純な平行移動のみで相対座標を計算する
        return px - ax, py - ay

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
            
            # アライメントによる横ズレを防ぐためXは挿入点のまま、Yのみ文字の高さの中央付近を代表点とする
            rep_x = t['x']
            rep_y = t['y'] + (t.get('h', 2.5) * 0.5)
            
            tx, ty = self._transform_point(rep_x, rep_y)
            if xmin <= tx <= xmax and ymin <= ty <= ymax:
                matched_texts.append(t)
        
        if matched_texts:
            matched_texts.sort(key=lambda item: (-item['y'], item['x']))
            return "".join([t['text'] for t in matched_texts])
        return ""

    def update_preview_text(self):
        if not self.anchor:
            self.preview_label.config(text="【抽出プレビュー】 1. 図面内の「第1基準文字」をクリックしてください")
        elif self.mode.get() == "anchor2":
            self.preview_label.config(text=f"【第1基準: {self.anchor['text']}】 2. スケール・回転補正用の「第2基準文字」をクリックしてください（不要なら次へ）")
        elif not self.rect_dxf_start or not self.rect_dxf_end:
            kw2_txt = f" / 第2: {self.anchor2['text']}" if self.anchor2 else ""
            self.preview_label.config(text=f"【基準設定済: {self.anchor['text']}{kw2_txt}】 3. 抽出したい値の範囲を左ドラッグで囲んでください")
        else:
            ext_text = self.get_extracted_text()
            if ext_text:
                self.preview_label.config(text=f"【抽出テスト】 {ext_text}")
            else:
                self.preview_label.config(text=f"【抽出テスト】 (※選択範囲内にテキストが見つかりません)")

    def draw(self, update_text=True):
        self.canvas.delete("all")
        cw = max(10, self.canvas.winfo_width())
        ch = max(10, self.canvas.winfo_height())
        margin = 100
        shape_color = "#C0C0C0"
        
        for s in self.shapes:
            try:
                if s['type'] == 'line':
                    cx1, cy1 = self.d2c(*s['start'])
                    cx2, cy2 = self.d2c(*s['end'])
                    if max(cx1, cx2) < -margin or min(cx1, cx2) > cw + margin or max(cy1, cy2) < -margin or min(cy1, cy2) > ch + margin:
                        continue
                    self.canvas.create_line(cx1, cy1, cx2, cy2, fill=shape_color)
                elif s['type'] == 'polyline':
                    pts = s['points']
                    if not pts: continue
                    c_pts = [self.d2c(*pt) for pt in pts]
                    if s.get('closed') and len(c_pts) > 2: c_pts.append(c_pts[0])
                    
                    cxs = [pt[0] for pt in c_pts]; cys = [pt[1] for pt in c_pts]
                    if max(cxs) < -margin or min(cxs) > cw + margin or max(cys) < -margin or min(cys) > ch + margin: continue
                    flat_pts = [coord for pt in c_pts for coord in pt]
                    self.canvas.create_line(flat_pts, fill=shape_color)
                elif s['type'] == 'circle':
                    cx, cy = self.d2c(*s['center'])
                    cr = s['radius'] * self.scale
                    if cx + cr < -margin or cx - cr > cw + margin or cy + cr < -margin or cy - cr > ch + margin: continue
                    self.canvas.create_oval(cx - cr, cy - cr, cx + cr, cy + cr, outline=shape_color)
                elif s['type'] == 'arc':
                    cx, cy = self.d2c(*s['center'])
                    cr = s['radius'] * self.scale
                    if cx + cr < -margin or cx - cr > cw + margin or cy + cr < -margin or cy - cr > ch + margin: continue
                    tk_extent = s['end_angle'] - s['start_angle']
                    if tk_extent < 0: tk_extent += 360
                    tk_start = 360 - s['end_angle']
                    self.canvas.create_arc(cx - cr, cy - cr, cx + cr, cy + cr, start=tk_start, extent=tk_extent, outline=shape_color, style=ARC)
            except: pass

        for t in self.texts:
            cx, cy = self.d2c(t["x"], t["y"])
            if cx < -margin or cx > cw + margin or cy < -margin or cy > ch + margin: continue
                
            color = "black"
            px_height = int(t.get("h", 2.5) * self.scale)
            f_size = max(2, px_height)
            
            if self.anchor and self.anchor == t:
                color = "red"
                f_size = int(f_size * 1.5)
                self.canvas.create_oval(cx-5, cy-5, cx+5, cy+5, fill="red", outline="red")
            elif self.anchor2 and self.anchor2 == t:
                color = "#FF8C00"
                f_size = int(f_size * 1.5)
                self.canvas.create_oval(cx-5, cy-5, cx+5, cy+5, fill="#FF8C00", outline="#FF8C00")
                
            self.canvas.create_text(cx, cy, text=t["text"], anchor="sw", fill=color, font=("Meiryo UI", -f_size))
            
        if self.rect_dxf_start and self.rect_dxf_end:
            cx1, cy1 = self.d2c(*self.rect_dxf_start)
            cx2, cy2 = self.d2c(*self.rect_dxf_end)
            self.canvas.create_rectangle(cx1, cy1, cx2, cy2, outline="blue", dash=(4, 4), width=2, fill="#e6f2ff", stipple="gray25")
            
        if update_text:
            self.update_preview_text()
            
    def on_left_press(self, event):
        dx, dy = self.c2d(event.x, event.y)
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
            
    def on_left_drag(self, event):
        if self.mode.get() == "rect" and self.is_dragging:
            dx, dy = self.c2d(event.x, event.y)
            self.rect_dxf_end = (dx, dy)
            self.draw()
            
    def on_left_release(self, event):
        if self.mode.get() == "rect" and self.is_dragging:
            self.is_dragging = False
            self.draw()
            
    def on_right_press(self, event):
        self.pan_start_x, self.pan_start_y = event.x, event.y
        
    def on_right_drag(self, event):
        self.offset_x += (event.x - self.pan_start_x)
        self.offset_y += (event.y - self.pan_start_y)
        self.pan_start_x, self.pan_start_y = event.x, event.y
        self.draw()
        
    def on_mousewheel(self, event):
        zoom = 1.2 if event.delta > 0 else 0.8
        mx, my = event.x, event.y
        dx, dy = self.c2d(mx, my)
        self.scale *= zoom
        self.offset_x = mx - (dx - self.min_x) * self.scale
        self.offset_y = my - (self.max_y - dy) * self.scale
        self.draw()
        
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
        col_name = "" # 空文字を渡すと、親ウィンドウ側で自動採番される
        
        self.on_complete(kw_text, kw2_text, col_name, xmin, xmax, ymin, ymax, ext_text)
        
        # 連続設定のためにアンカーは保持し、ドラッグ範囲だけクリアする
        self.rect_dxf_start = None
        self.rect_dxf_end = None
        self.mode.set("rect")
        
        self.draw(update_text=False)
        self.preview_label.config(text=f"✅ 抽出項目を追加しました！続けて次の抽出範囲をドラッグしてください。")

    def finish_and_close(self):
        # 閉じる前に、未追加のドラッグ範囲があれば設定に追加する
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
            col_name = "" 
            
            self.on_complete(kw_text, kw2_text, col_name, xmin, xmax, ymin, ymax, ext_text)
            
        self.destroy()