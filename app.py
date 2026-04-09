# -*- coding: utf-8 -*-
"""
DxfExtractionMiya v1.0.0
機能：DXFファイルからのテキスト抽出（全体／キーワード指定）、複数Excelファイルのデータ集約
1ファイル完結版
"""

import os
import re
import ezdxf
import openpyxl
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter import ttk

# ==========================
# 共通変数
# ==========================
APP_TITLE = "DxfExtractionMiya"
VERSION = "v1.0.0"

selected_files = []
selected_folder = ""
current_mode = None  # "file" or "folder"

# ==========================
# 選択処理
# ==========================
def select_files():
    global selected_files, selected_folder, current_mode
    files = filedialog.askopenfilenames(
        title="ファイルを選択",
        filetypes=[("すべての対応ファイル", "*.dxf;*.xlsx;*.xlsm;*.xls"), ("DXFファイル", "*.dxf"), ("Excelファイル", "*.xlsx;*.xlsm;*.xls")]
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
    btn_extract.config(state=state)
    btn_aggregate.config(state=state)

# ==========================
# 共通ユーティリティ
# ==========================
def get_save_dir(original_path=None):
    if save_option.get() == 1 and original_path:
        return os.path.dirname(original_path)
    else:
        folder = filedialog.askdirectory(title="保存先フォルダを選択")
        return folder

def sanitize_text(text):
    """Excelに書き込む際にエラーとなる制御文字などを除去"""
    if text is None:
        return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(text)).strip()

def apply_text_inheritance(final_aggregated_data):
    """「〃」や「同上」などの記号がある場合、上のセルのテキストを自動でコピーして補完する"""
    if len(final_aggregated_data) <= 1: return
    
    def is_text_to_inherit(text):
        s = str(text).strip()
        if not s or s in ["〃", "”", "\"", "''", "””", "''", "同上", "...", "…"]: return False
        return bool(re.search(r'[a-zA-Zａ-ｚＡ-Ｚぁ-んァ-ン一-龥0-9０-９]', s))
    
    header = final_aggregated_data[0]
    skip_cols = {idx for idx, h in enumerate(header) if "備考" in str(h)}
    
    for col_idx in range(1, len(header)):
        if col_idx in skip_cols: continue
        last_text = ""
        for row_idx in range(1, len(final_aggregated_data)):
            cell_val = str(final_aggregated_data[row_idx][col_idx]).strip()
            
            if cell_val == "None":
                cell_val = ""
                final_aggregated_data[row_idx][col_idx] = ""
                
            if cell_val in ["〃", "”", "\"", "''", "””", "''", "同上", "...", "…"]:
                if last_text: final_aggregated_data[row_idx][col_idx] = last_text
            elif cell_val == "":
                pass
            else:
                last_text = cell_val if is_text_to_inherit(cell_val) else ""

# ==========================
# DXFテキスト抽出処理
# ==========================
def extract_dxf_text():
    try:
        # 処理対象ファイルの取得
        target_files = []
        if current_mode == "file":
            target_files = [f for f in selected_files if f.lower().endswith('.dxf')]
        elif current_mode == "folder" and selected_folder:
            target_files = [os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith('.dxf')]

        if not target_files:
            messagebox.showwarning("警告", "選択された対象内にDXFファイルが見つかりません。")
            return

        save_dir = get_save_dir(target_files[0])
        if not save_dir:
            return # キャンセル時

        # キーワードのパース
        keywords_str = keywords_var.get().strip()
        keywords = [k.strip() for k in keywords_str.split(',') if k.strip()]
        is_keyword_mode = len(keywords) > 0

        y_threshold = threshold_var.get()
        processed_count = 0

        if is_keyword_mode:
            # ---------------------------------------------------------
            # モードA: キーワード抽出モード (全ファイルを1つの表にまとめる)
            # ---------------------------------------------------------
            out_wb = openpyxl.Workbook()
            ws = out_wb.active
            ws.title = "キーワード抽出結果"
            
            # ヘッダーの作成
            header = ["元ファイル名"] + keywords
            ws.append(header)

            for file_path in target_files:
                try:
                    doc = ezdxf.readfile(file_path)
                    msp = doc.modelspace()
                except Exception as e:
                    print(f"読み込みエラー: {os.path.basename(file_path)} ({e})")
                    continue

                texts = []
                
                # 1. モデルスペースの全エンティティを走査
                for entity in msp:
                    # 通常のテキスト
                    if entity.dxftype() in {'TEXT', 'MTEXT'}:
                        try:
                            text_val = entity.plain_text() if entity.dxftype() == 'MTEXT' else entity.dxf.text
                            insert_pt = entity.dxf.insert
                            clean_text = sanitize_text(text_val)
                            if clean_text:
                                texts.append({"text": clean_text, "x": insert_pt.x, "y": insert_pt.y, "tag": ""})
                        except: pass
                        
                    # ブロック（INSERT）内の属性と固定テキストを展開して取得
                    elif entity.dxftype() == 'INSERT':
                        # 属性（ATTRIB）の取得
                        if hasattr(entity, 'attribs') and entity.attribs:
                            for attrib in entity.attribs:
                                try:
                                    text_val = attrib.dxf.text
                                    tag_val = attrib.dxf.tag
                                    insert_pt = attrib.dxf.insert
                                    clean_text = sanitize_text(text_val)
                                    if clean_text:
                                        texts.append({"text": clean_text, "x": insert_pt.x, "y": insert_pt.y, "tag": tag_val})
                                except: pass
                        
                        # ブロック内の固定テキストの取得（仮想エンティティとして展開）
                        try:
                            for v_entity in entity.virtual_entities():
                                if v_entity.dxftype() in {'TEXT', 'MTEXT'}:
                                    text_val = v_entity.plain_text() if v_entity.dxftype() == 'MTEXT' else v_entity.dxf.text
                                    try:
                                        insert_pt = v_entity.dxf.insert
                                    except:
                                        continue
                                    clean_text = sanitize_text(text_val)
                                    if clean_text:
                                        texts.append({"text": clean_text, "x": insert_pt.x, "y": insert_pt.y, "tag": ""})
                        except: pass

                if not texts:
                    continue

                # 1ファイル分のデータ（行）を生成
                file_row = [os.path.basename(file_path)]
                for kw in keywords:
                    found_val = ""
                    kw_clean = kw.replace(" ", "").replace("　", "")
                    
                    # アプローチ1: 属性(ATTRIB)のTAGから直接探す (最も確実)
                    for t in texts:
                        if t['tag']:
                            tag_clean = t['tag'].replace(" ", "").replace("　", "")
                            if kw_clean in tag_clean or tag_clean in kw_clean:
                                found_val = t['text']
                                break
                    
                    if found_val:
                        file_row.append(found_val)
                        continue
                    
                    # アプローチ2: テキスト内に「キーワード＋値」が含まれているか探す (例: "図番：M-1234")
                    kw_entity = None
                    for t in texts:
                        t_clean = t['text'].replace(" ", "").replace("　", "")
                        if kw_clean in t_clean:
                            kw_entity = t
                            # 元のテキストからキーワードと記号を取り除いた実際の値を抽出
                            val_orig = re.sub(re.escape(kw), "", t['text'], flags=re.IGNORECASE)
                            if val_orig != t['text']:
                                val_orig = val_orig.strip(" :：=＝-ー_　\t\n")
                                if val_orig:
                                    found_val = val_orig
                                    break
                            else:
                                val_clean = re.sub(re.escape(kw_clean), "", t_clean, flags=re.IGNORECASE)
                                val_clean = val_clean.strip(" :：=＝-ー_　\t\n")
                                if val_clean:
                                    found_val = val_clean
                                    break
                    
                    if found_val:
                        file_row.append(found_val)
                        continue

                    # アプローチ3: 見つかったキーワードテキストの「右」または「下」を探す
                    if kw_entity:
                        kw_x, kw_y = kw_entity['x'], kw_entity['y']
                        best_dist = float('inf')
                        best_text = ""
                        
                        for t in texts:
                            # 自分自身や、キーワードと全く同じテキスト（重複配置）は除外
                            if t == kw_entity or t['text'].strip() == kw_entity['text'].strip():
                                continue
                                
                            dx = t['x'] - kw_x
                            dy = t['y'] - kw_y
                            
                            # 距離が0に近すぎるものは重なっているテキストとみなして除外
                            if abs(dx) < 0.1 and abs(dy) < 0.1:
                                continue
                            
                            # 右側にあるか判定 (Yのズレは閾値以内、Xは右方向。文字基点のズレを考慮して少しのマイナスは許容)
                            is_right = (dx > -y_threshold) and (dx < y_threshold * 15) and (abs(dy) <= y_threshold)
                            
                            # 下側にあるか判定 (Xのズレはある程度許容、Yは下方向)
                            is_below = (dy < -y_threshold * 0.2) and (dy > -y_threshold * 5) and (abs(dx) <= y_threshold * 3)
                            
                            if is_right or is_below:
                                dist = (dx**2 + dy**2)**0.5
                                if dist < best_dist:
                                    best_dist = dist
                                    best_text = t['text']
                        
                        found_val = best_text
                    
                    file_row.append(found_val)
                
                ws.append(file_row)
                processed_count += 1

            # 列幅の自動調整
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        if cell.value:
                            length = sum(2 if ord(c) > 255 else 1 for c in str(cell.value))
                            if length > max_length: max_length = length
                    except: pass
                ws.column_dimensions[col_letter].width = min(max_length + 2, 60)

            import time
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(save_dir, f"DXFキーワード抽出_{timestamp}.xlsx")
            out_wb.save(output_path)
            
            messagebox.showinfo("完了", f"{processed_count} 個のファイルから指定キーワードを抽出し、\n1つのExcelファイルにまとめました。\n\n保存先: {output_path}")

        else:
            # ---------------------------------------------------------
            # モードB: 全体抽出モード (ファイルごとにすべてのテキストを出力)
            # ---------------------------------------------------------
            for file_path in target_files:
                try:
                    doc = ezdxf.readfile(file_path)
                    msp = doc.modelspace()
                except Exception as e:
                    print(f"読み込みエラー: {os.path.basename(file_path)} ({e})")
                    continue

                texts = []
                
                for entity in msp:
                    if entity.dxftype() in {'TEXT', 'MTEXT'}:
                        try:
                            text_val = entity.plain_text() if entity.dxftype() == 'MTEXT' else entity.dxf.text
                            insert_pt = entity.dxf.insert
                            clean_text = sanitize_text(text_val)
                            if clean_text:
                                texts.append({"text": clean_text, "x": insert_pt.x, "y": insert_pt.y})
                        except: pass
                        
                    elif entity.dxftype() == 'INSERT':
                        if hasattr(entity, 'attribs') and entity.attribs:
                            for attrib in entity.attribs:
                                try:
                                    text_val = attrib.dxf.text
                                    insert_pt = attrib.dxf.insert
                                    clean_text = sanitize_text(text_val)
                                    if clean_text:
                                        texts.append({"text": clean_text, "x": insert_pt.x, "y": insert_pt.y})
                                except: pass
                        
                        try:
                            for v_entity in entity.virtual_entities():
                                if v_entity.dxftype() in {'TEXT', 'MTEXT'}:
                                    text_val = v_entity.plain_text() if v_entity.dxftype() == 'MTEXT' else v_entity.dxf.text
                                    try:
                                        insert_pt = v_entity.dxf.insert
                                    except:
                                        continue
                                    clean_text = sanitize_text(text_val)
                                    if clean_text:
                                        texts.append({"text": clean_text, "x": insert_pt.x, "y": insert_pt.y})
                        except: pass

                if not texts:
                    continue

                texts.sort(key=lambda item: (-item['y'], item['x']))

                rows = []
                current_row = []
                current_y = None

                for item in texts:
                    if current_y is None:
                        current_y = item['y']
                        current_row.append(item)
                    elif abs(current_y - item['y']) <= y_threshold:
                        current_row.append(item)
                    else:
                        current_row.sort(key=lambda i: i['x'])
                        rows.append([i['text'] for i in current_row])
                        current_y = item['y']
                        current_row = [item]
                
                if current_row:
                    current_row.sort(key=lambda i: i['x'])
                    rows.append([i['text'] for i in current_row])

                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "抽出データ"
                for row_data in rows: ws.append(row_data)

                for col in ws.columns:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        try:
                            if cell.value:
                                length = sum(2 if ord(c) > 255 else 1 for c in str(cell.value))
                                if length > max_length: max_length = length
                        except: pass
                    ws.column_dimensions[col_letter].width = min(max_length + 2, 60)

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                output_path = os.path.join(save_dir, f"{base_name}_テキスト抽出.xlsx")
                wb.save(output_path)
                processed_count += 1

            messagebox.showinfo("完了", f"{processed_count} 個のDXFファイルからの全体抽出が完了しました。")

    except Exception as e:
        messagebox.showerror("抽出エラー", f"テキスト抽出処理中にエラーが発生しました。\n詳細: {e}")

# ==========================
# データ集約処理
# ==========================
def aggregate_data():
    try:
        # 対象ファイルの取得
        target_files = []
        if current_mode == "file":
            target_files = [f for f in selected_files if f.lower().endswith(('.xlsx', '.xlsm', '.xls'))]
        elif current_mode == "folder" and selected_folder:
            # 「集約」という文字が含まれるファイルは再集約防止のため除外
            target_files = [os.path.join(selected_folder, f) for f in os.listdir(selected_folder) if f.lower().endswith(('.xlsx', '.xlsm', '.xls')) and "集約" not in f]
            
        if not target_files:
            messagebox.showwarning("警告", "選択された対象内に集約可能なExcelファイルが見つかりません。")
            return

        save_dir = get_save_dir(target_files[0])
        if not save_dir:
            return

        agg_header = ["元ファイル名"]
        agg_rows = []

        # マスター結合ロジック：列名を自動認識してマッピングする
        for f in target_files:
            fname = os.path.basename(f)
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                sheet = wb.worksheets[0]
                rows = list(sheet.iter_rows(values_only=True))
                wb.close()

                if not rows: continue

                # 空行は除外し、有効な行のみを抽出
                valid_rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
                if not valid_rows: continue

                curr_header = valid_rows[0]
                curr_data = valid_rows[1:]

                # ヘッダーが空の場合はダミー列名を付与
                safe_header = [str(h).strip() if h is not None and str(h).strip() else f"列{i+1}" for i, h in enumerate(curr_header)]
                col_mapping = {}

                # マスターヘッダー（agg_header）を更新しつつ、列の対応表を作成
                for i, h in enumerate(safe_header):
                    if h not in agg_header:
                        agg_header.append(h)
                    col_mapping[i] = agg_header.index(h)

                # データをマッピング表に合わせて配置
                for r in curr_data:
                    row = [""] * len(agg_header)
                    row[0] = fname
                    for i, val in enumerate(r):
                        if i in col_mapping:
                            idx = col_mapping[i]
                            # 必要に応じて行の長さを拡張
                            if idx >= len(row):
                                row.extend([""] * (idx - len(row) + 1))
                            if val is None or str(val).strip() == "None":
                                row[idx] = ""
                            else:
                                row[idx] = str(val).strip()
                    agg_rows.append(row)
            except Exception as e:
                print(f"スキップ: {fname} ({e})")

        if not agg_rows:
            messagebox.showinfo("結果", "集約できるデータがありませんでした。")
            return

        out_wb = openpyxl.Workbook()
        ws = out_wb.active
        ws.title = "集約"

        # ヘッダーとデータを結合
        final_data = [agg_header] + [r + [""] * (len(agg_header) - len(r)) for r in agg_rows]

        # 同上コピー機能（〃 などの記号を補完）を適用
        apply_text_inheritance(final_data)

        # Excelへ書き込み
        for row_data in final_data:
            ws.append(row_data)

        # 列幅調整
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        length = sum(2 if ord(c) > 255 else 1 for c in str(cell.value))
                        if length > max_length: max_length = length
                except: pass
            ws.column_dimensions[col_letter].width = min(max_length + 2, 60)

        import time
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(save_dir, f"データ集約_{timestamp}.xlsx")
        out_wb.save(output_path)

        messagebox.showinfo("完了", f"{len(target_files)}個のファイルデータをスマートに集約しました。\n保存先: {output_path}")

    except Exception as e:
        messagebox.showerror("集約エラー", f"データ集約処理中にエラーが発生しました。\n詳細: {e}")

# ==========================
# UI構築
# ==========================
root = Tk()
root.title(f"{APP_TITLE} {VERSION}")
root.geometry("640x780")
root.minsize(600, 750)
root.configure(bg="#F8F9FA")

style = ttk.Style()
style.theme_use("clam")

# ヘッダー
header_frame = Frame(root, bg="#0D6EFD", pady=15)
header_frame.pack(fill=X)
Label(header_frame, text=f"{APP_TITLE} {VERSION}", font=("Meiryo UI", 16, "bold"), bg="#0D6EFD", fg="white").pack()
Label(header_frame, text="DXFテキスト抽出 ＆ Excelスマート集約ツール", font=("Meiryo UI", 10), bg="#0D6EFD", fg="white").pack()

main_frame = Frame(root, bg="#F8F9FA", padx=20, pady=10)
main_frame.pack(fill=BOTH, expand=True)

# 選択セクション
Label(main_frame, text="1. 対象の選択", font=("Meiryo UI", 11, "bold"), bg="#F8F9FA").pack(anchor=W, pady=(10, 5))

btn_frame = Frame(main_frame, bg="#F8F9FA")
btn_frame.pack(fill=X)
Button(btn_frame, text="📄 ファイルを選択", command=select_files, width=20, bg="#E9ECEF", relief=GROOVE).pack(side=LEFT, padx=5, pady=5)
Button(btn_frame, text="📁 フォルダを選択", command=select_folder, width=20, bg="#E9ECEF", relief=GROOVE).pack(side=LEFT, padx=5, pady=5)

Label(main_frame, text="選択中のパス:", bg="#F8F9FA").pack(anchor=W, pady=(5, 0))
text_paths = Text(main_frame, height=5, width=70, font=("Meiryo UI", 9))
text_paths.pack(pady=5)

# 設定セクション
Label(main_frame, text="2. 設定", font=("Meiryo UI", 11, "bold"), bg="#F8F9FA").pack(anchor=W, pady=(15, 5))

settings_frame = Frame(main_frame, bg="#FFFFFF", padx=15, pady=15, relief=SOLID, bd=1)
settings_frame.pack(fill=X, pady=5)

Label(settings_frame, text="【保存先】", font=("Meiryo UI", 9, "bold"), bg="#FFFFFF").pack(anchor=W)
save_option = IntVar(value=1)
Radiobutton(settings_frame, text="元のファイルと同じフォルダに保存", variable=save_option, value=1, bg="#FFFFFF").pack(anchor=W)
Radiobutton(settings_frame, text="実行時に任意のフォルダを指定して保存", variable=save_option, value=2, bg="#FFFFFF").pack(anchor=W)

Frame(settings_frame, height=1, bg="#DEE2E6").pack(fill=X, pady=10)

Label(settings_frame, text="【DXF抽出 詳細設定】", font=("Meiryo UI", 9, "bold"), bg="#FFFFFF").pack(anchor=W)

Label(settings_frame, text="抽出キーワード (カンマ区切りで複数指定):", bg="#FFFFFF").pack(anchor=W, pady=(10, 0))
keywords_var = StringVar(value="図番, 品名")
Entry(settings_frame, textvariable=keywords_var, width=70).pack(anchor=W, pady=2)
Label(settings_frame, text="※キーワードを指定すると、その文字の「右」または「下」にある値を抽出し、\n　指定キーワードを列（カラム）とした1つのExcel表に全ファイルをまとめます。\n※空欄の場合は、ファイルごとに図面内のすべてのテキストを抽出します。", font=("Meiryo UI", 8), fg="#6C757D", bg="#FFFFFF", justify=LEFT).pack(anchor=W, pady=(0, 10))

Label(settings_frame, text="テキストのズレ許容値（閾値）:", bg="#FFFFFF").pack(anchor=W)
threshold_var = DoubleVar(value=20.0)
Spinbox(settings_frame, from_=0.0, to=500.0, increment=1.0, textvariable=threshold_var, width=10).pack(anchor=W, pady=2)
Label(settings_frame, text="※「右」や「下」にあると判定する範囲の調整に使用します。", font=("Meiryo UI", 8), fg="#6C757D", bg="#FFFFFF").pack(anchor=W)


# 実行セクション
Label(main_frame, text="3. 実行", font=("Meiryo UI", 11, "bold"), bg="#F8F9FA").pack(anchor=W, pady=(20, 5))

action_frame = Frame(main_frame, bg="#F8F9FA")
action_frame.pack(fill=X)

btn_extract = Button(action_frame, text="🚀 選択対象のDXFからテキストを抽出", command=extract_dxf_text, height=2, bg="#198754", fg="white", font=("Meiryo UI", 10, "bold"), state=DISABLED)
btn_extract.pack(fill=X, pady=5)

btn_aggregate = Button(action_frame, text="🧩 選択対象のExcelを1つにスマート集約", command=aggregate_data, height=2, bg="#6F42C1", fg="white", font=("Meiryo UI", 10, "bold"), state=DISABLED)
btn_aggregate.pack(fill=X, pady=5)

if __name__ == "__main__":
    root.mainloop()