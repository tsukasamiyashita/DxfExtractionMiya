# -*- coding: utf-8 -*-
"""
抽出・集約メインロジックモジュール
"""
import os
import math
import time
import traceback
import openpyxl
from tkinter import messagebox
from dxf_core import get_all_elements_from_dxf, apply_text_inheritance

def run_extract_dxf(target_files, save_dir, is_keyword_mode, y_threshold, base_kw_str, base_kw2_str, keyword_settings):
    try:
        processed_count = 0
        error_logs = []

        if is_keyword_mode:
            if not base_kw_str:
                messagebox.showwarning("警告", "第1基準文字が設定されていません。")
                return
            if not keyword_settings:
                messagebox.showwarning("警告", "抽出設定（抽出範囲）がありません。「＋ プレビューを開いて基準文字と抽出範囲を設定」から設定を行ってください。")
                return

            out_wb = openpyxl.Workbook()
            ws = out_wb.active
            ws.title = "プレビュー抽出結果"
            
            ws_coord = out_wb.create_sheet(title="座標リスト(設定確認用)")
            ws_coord.append(["元ファイル名", "テキスト", "X座標", "Y座標"])
            ws_coord.column_dimensions['A'].width = 25
            ws_coord.column_dimensions['B'].width = 40
            ws_coord.column_dimensions['C'].width = 15
            ws_coord.column_dimensions['D'].width = 15
            
            header = ["元ファイル名"]
            for cfg in keyword_settings:
                header.append(cfg["col_name"])
            ws.append(header)

            for file_path in target_files:
                texts, _ = get_all_elements_from_dxf(file_path)
                if not texts:
                    error_logs.append(f"【{os.path.basename(file_path)}】テキストが見つかりません。")
                    continue

                for t in texts:
                    ws_coord.append([os.path.basename(file_path), t['text'], round(t['x'], 3), round(t['y'], 3)])

                file_row = [os.path.basename(file_path)]

                kw_candidates = [t for t in texts if base_kw_str == t['text'].replace(" ", "").replace("　", "").lower()]
                if not kw_candidates:
                    kw_candidates = [t for t in texts if base_kw_str in t['text'].replace(" ", "").replace("　", "").lower()]
                    
                kw2_candidates = []
                if base_kw2_str:
                    kw2_candidates = [t for t in texts if base_kw2_str == t['text'].replace(" ", "").replace("　", "").lower()]
                    if not kw2_candidates:
                        kw2_candidates = [t for t in texts if base_kw2_str in t['text'].replace(" ", "").replace("　", "").lower()]
                else:
                    kw2_candidates = [None]
                
                # 座標変換マトリクスを取得
                transform_params = None
                found_anchor1 = None
                found_anchor2 = None
                
                for kw_entity in kw_candidates:
                    for kw2_entity in kw2_candidates:
                        ax, ay = kw_entity['x'], kw_entity['y']
                        a2x, a2y = None, None
                        
                        if kw2_entity:
                            a2x, a2y = kw2_entity['x'], kw2_entity['y']
                            if ax == a2x and ay == a2y: continue
                            
                        if a2x is not None and a2y is not None:
                            dx, dy = a2x - ax, a2y - ay
                            L = math.hypot(dx, dy)
                            if L < 1e-6: L, theta = 1.0, 0.0
                            else: theta = math.atan2(dy, dx)
                        else:
                            L, theta = 1.0, 0.0
                            
                        cos_t = math.cos(-theta)
                        sin_t = math.sin(-theta)
                        
                        transform_params = (ax, ay, L, cos_t, sin_t)
                        found_anchor1 = kw_entity
                        found_anchor2 = kw2_entity
                        break
                    if transform_params: break
                
                if transform_params:
                    ax, ay, L, cos_t, sin_t = transform_params
                    
                    for cfg in keyword_settings:
                        x_min, x_max = cfg["xmin"], cfg["xmax"]
                        y_min, y_max = cfg["ymin"], cfg["ymax"]
                            
                        matched_texts = []
                        for t in texts:
                            if t == found_anchor1 or t == found_anchor2: continue
                            
                            t_clean_str = t['text'].strip()
                            if t_clean_str == found_anchor1['text'].strip() or (found_anchor2 and t_clean_str == found_anchor2['text'].strip()): continue
                            
                            # 文字の左端・中央付近を判定用の代表点とする
                            rep_x = t['x']
                            rep_y = t['y'] + (t.get('h', 2.5) * 0.4)
                            
                            px, py = rep_x - ax, rep_y - ay
                            nx = (px * cos_t - py * sin_t) / L
                            ny = (px * sin_t + py * cos_t) / L
                            
                            if x_min <= nx <= x_max and y_min <= ny <= y_max:
                                matched_texts.append(t)
                        
                        if matched_texts:
                            matched_texts.sort(key=lambda item: (-item['y'], item['x']))
                            found_val = "".join([t['text'] for t in matched_texts])
                        else:
                            found_val = ""
                            
                        file_row.append(found_val)
                else:
                    for _ in keyword_settings:
                        file_row.append("")
                    error_logs.append(f"【{os.path.basename(file_path)}】基準文字が見つかりません。")
                
                ws.append(file_row)
                processed_count += 1

            for col in ws.columns:
                try:
                    max_length = 0
                    col_letter = col[0].column_letter
                    for cell in col:
                        if cell.value:
                            length = sum(2 if ord(c) > 255 else 1 for c in str(cell.value))
                            if length > max_length: max_length = length
                    ws.column_dimensions[col_letter].width = min(max_length + 2, 60)
                except: pass

            output_path = os.path.join(save_dir, f"DXFプレビュー抽出_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
            try: out_wb.save(output_path)
            except PermissionError:
                messagebox.showerror("保存エラー", f"Excelファイルが開かれているため保存できません。\nパス: {output_path}")
                return
            except Exception as e:
                messagebox.showerror("保存エラー", f"エラー: {e}")
                return
            
            msg = f"{processed_count} 個のファイルから抽出を完了しました。\n保存先: {output_path}"
            if error_logs: msg += "\n\n※一部エラーあり:\n" + "\n".join(error_logs[:5])
            messagebox.showinfo("完了", msg)

        else:
            # ---------------------------------------------------------
            # モードB: 全体抽出モード
            # ---------------------------------------------------------
            for file_path in target_files:
                texts, _ = get_all_elements_from_dxf(file_path)
                if not texts:
                    error_logs.append(f"【{os.path.basename(file_path)}】読込エラー。")
                    continue

                texts.sort(key=lambda item: (-item['y'], item['x']))
                rows, current_row, current_y = [], [], None

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
                    try:
                        max_length = 0
                        col_letter = col[0].column_letter
                        for cell in col:
                            if cell.value:
                                length = sum(2 if ord(c) > 255 else 1 for c in str(cell.value))
                                if length > max_length: max_length = length
                        ws.column_dimensions[col_letter].width = min(max_length + 2, 60)
                    except: pass

                ws_coord = wb.create_sheet(title="座標リスト(設定確認用)")
                ws_coord.append(["テキスト", "X座標", "Y座標"])
                for item in texts: ws_coord.append([item['text'], round(item['x'], 3), round(item['y'], 3)])
                ws_coord.column_dimensions['A'].width, ws_coord.column_dimensions['B'].width, ws_coord.column_dimensions['C'].width = 40, 15, 15

                output_path = os.path.join(save_dir, f"{os.path.splitext(os.path.basename(file_path))[0]}_テキスト抽出.xlsx")
                try: wb.save(output_path)
                except PermissionError:
                    messagebox.showerror("保存エラー", f"Excelが開かれています。\nパス: {output_path}")
                    return
                except: return
                processed_count += 1

            msg = f"{processed_count} 個のDXFファイルからの全体抽出が完了しました。"
            if error_logs: msg += "\n\n※一部エラーあり:\n" + "\n".join(error_logs[:5])
            messagebox.showinfo("完了", msg)

    except Exception as e:
        messagebox.showerror("抽出エラー", f"予期せぬエラーが発生しました。\n\n【詳細】\n{traceback.format_exc()}")

def run_aggregate_data(target_files, save_dir):
    try:
        agg_header = ["元ファイル名"]
        agg_rows, error_logs = [], []

        for f in target_files:
            fname = os.path.basename(f)
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                sheet = wb.worksheets[0]
                rows = list(sheet.iter_rows(values_only=True))
                wb.close()

                if not rows: continue
                valid_rows = [r for r in rows if any(c is not None and str(c).strip() != "" for c in r)]
                if not valid_rows: continue

                curr_header, curr_data = valid_rows[0], valid_rows[1:]
                safe_header = [str(h).strip() if h is not None and str(h).strip() else f"列{i+1}" for i, h in enumerate(curr_header)]
                col_mapping = {}

                for i, h in enumerate(safe_header):
                    if h not in agg_header: agg_header.append(h)
                    col_mapping[i] = agg_header.index(h)

                for r in curr_data:
                    row = [""] * len(agg_header)
                    row[0] = fname
                    for i, val in enumerate(r):
                        if i in col_mapping:
                            idx = col_mapping[i]
                            if idx >= len(row): row.extend([""] * (idx - len(row) + 1))
                            row[idx] = "" if val is None or str(val).strip() == "None" else str(val).strip()
                    agg_rows.append(row)
            except Exception as e:
                error_logs.append(f"【{fname}】スキップ: {e}")

        if not agg_rows:
            messagebox.showinfo("結果", "集約できるデータがありませんでした。")
            return

        out_wb = openpyxl.Workbook()
        ws = out_wb.active
        ws.title = "集約"

        final_data = [agg_header] + [r + [""] * (len(agg_header) - len(r)) for r in agg_rows]
        apply_text_inheritance(final_data)

        for row_data in final_data: ws.append(row_data)

        for col in ws.columns:
            try:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value:
                        length = sum(2 if ord(c) > 255 else 1 for c in str(cell.value))
                        if length > max_length: max_length = length
                ws.column_dimensions[col_letter].width = min(max_length + 2, 60)
            except: pass

        output_path = os.path.join(save_dir, f"データ集約_{time.strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        try: out_wb.save(output_path)
        except PermissionError:
            messagebox.showerror("保存エラー", f"Excelファイルが開かれているため保存できません。\nパス: {output_path}")
            return
        except Exception as e:
            messagebox.showerror("保存エラー", f"エラー: {e}")
            return

        msg = f"{len(target_files)}個のファイルデータを集約しました。\n保存先: {output_path}"
        if error_logs: msg += f"\n\n※一部スキップ:\n" + "\n".join(error_logs[:3])
        messagebox.showinfo("完了", msg)

    except Exception as e:
        messagebox.showerror("集約エラー", f"エラーが発生しました。\n\n【詳細】\n{traceback.format_exc()}")