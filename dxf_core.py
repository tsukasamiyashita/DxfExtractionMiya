# -*- coding: utf-8 -*-
"""
DXF解析・共通ユーティリティモジュール
"""
import re
import ezdxf
from ezdxf import recover

def zen_to_han(text):
    if not isinstance(text, str): return text
    # 英数字・記号の変換
    zen = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ"
    han = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    zen += "！＂＃＄％＆＇（）＊＋，－．／：；＜＝＞？［＼］＾＿｀｛｜｝"
    han += "!\"#$%&'()*+,-./:;<=>?[\\]^_`{|}"
    zen += "　"
    han += " "
    res = text.translate(str.maketrans(zen, han))
    
    # 全角カタカナ -> 半角カタカナの変換
    # 2文字（濁点・半濁点）になるものを先に処理
    zen_kana_d = "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポヴ"
    han_kana_d = ["ｶﾞ","ｷﾞ","ｸﾞ","ｹﾞ","ｺﾞ","ｻﾞ","ｼﾞ","ｽﾞ","ｾﾞ","ｿﾞ","ﾀﾞ","ﾁﾞ","ﾂﾞ","ﾃﾞ","ﾄﾞ","ﾊﾞ","ﾋﾞ","ﾌﾞ","ﾍﾞ","ﾎﾞ","ﾊﾟ","ﾋﾟ","ﾌﾟ","ﾍﾟ","ﾎﾟ","ｳﾞ"]
    for z, h in zip(zen_kana_d, han_kana_d):
        res = res.replace(z, h)
        
    zen_kana_s = "ァアィイゥウェエォオカキクケコサシスセソタチッツテトナニヌネノハヒフヘホマミムメモャヤュユョヨラリルレロヮワヰヱヲンヵヶー。、・「」"
    han_kana_s = "ｧｱｨｲｩｳｪｴｫｵｶｷｸｹｺｻｼｽｾｿﾀﾁｯﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓｬﾔｭﾕｮﾖﾗﾘﾙﾚﾛﾜﾜｲｴｦﾝｶｹｰ｡､･｢｣"
    res = res.translate(str.maketrans(zen_kana_s, han_kana_s))
    
    return res

# 後方互換性のため
zen_to_han_alnum = zen_to_han

def sanitize_text(text):
    if text is None: return ""
    text_str = str(text)
    illegal_chars = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
    # 制御文字の削除のみ行い、スペースは保持する（前後の空白のみ削除）
    cleaned = illegal_chars.sub('', text_str).strip()
    return zen_to_han(cleaned)

def get_point(pt):
    try:
        if pt is None: return 0.0, 0.0
        if hasattr(pt, 'x') and hasattr(pt, 'y'):
            return float(pt.x), float(pt.y)
        elif hasattr(pt, '__getitem__') and len(pt) >= 2:
            return float(pt[0]), float(pt[1])
    except: pass
    return 0.0, 0.0

def get_text_dimensions(entity, text_val):
    height = 2.5
    try:
        if hasattr(entity, 'dxf') and hasattr(entity, 'dxftype') and callable(entity.dxftype):
            if entity.dxftype() == 'MTEXT' and hasattr(entity.dxf, 'char_height'):
                height = float(entity.dxf.char_height)
            elif hasattr(entity.dxf, 'height'):
                height = float(entity.dxf.height)
    except: pass
    width = sum((height * 1.0) if ord(char) > 255 else (height * 0.6) for char in str(text_val))
    return height, width

def extract_text_from_entity(entity):
    text_val = ""
    try:
        if hasattr(entity, 'dxftype') and callable(entity.dxftype):
            if entity.dxftype() == 'MTEXT' and hasattr(entity, 'plain_text'):
                text_val = entity.plain_text()
            elif hasattr(entity, 'dxf') and hasattr(entity.dxf, 'text'):
                text_val = entity.dxf.text
    except: pass

    x, y = 0.0, 0.0
    try:
        if hasattr(entity, 'dxf') and hasattr(entity.dxf, 'insert'):
            x, y = get_point(entity.dxf.insert)
    except: pass
    return text_val, x, y

def get_all_elements_from_dxf(file_path):
    try:
        doc = ezdxf.readfile(file_path)
    except ezdxf.DXFStructureError:
        try: doc, _ = recover.readfile(file_path)
        except: return [], []
    except: return [], []
    
    if doc is None: return [], []
    texts, seen_texts = [], set()
    shapes = []
    
    layouts = []
    try: layouts.append(doc.modelspace())
    except: pass
    try: layouts.extend(list(doc.layouts()))
    except: pass
    
    def process_entity(entity):
        try:
            if entity is None or not hasattr(entity, 'dxftype') or not callable(entity.dxftype): return
            etype = entity.dxftype()
            
            if etype in {'TEXT', 'MTEXT'}:
                text_val, x, y = extract_text_from_entity(entity)
                clean_text = sanitize_text(text_val)
                if clean_text:
                    key = f"{clean_text}_{round(x,2)}_{round(y,2)}"
                    if key not in seen_texts:
                        h, w = get_text_dimensions(entity, clean_text)
                        texts.append({"text": clean_text, "x": x, "y": y, "w": w, "h": h})
                        seen_texts.add(key)
            elif etype == 'LINE':
                shapes.append({'type': 'line', 'start': get_point(entity.dxf.start), 'end': get_point(entity.dxf.end)})
            elif etype in {'LWPOLYLINE', 'POLYLINE'}:
                points = []
                if etype == 'LWPOLYLINE':
                    for pt in entity: points.append(get_point(pt))
                else:
                    for vertex in entity.vertices: points.append(get_point(vertex.dxf.location))
                if points:
                    shapes.append({'type': 'polyline', 'points': points, 'closed': getattr(entity, 'is_closed', False)})
            elif etype == 'CIRCLE':
                shapes.append({'type': 'circle', 'center': get_point(entity.dxf.center), 'radius': float(entity.dxf.radius)})
            elif etype == 'ARC':
                shapes.append({
                    'type': 'arc', 'center': get_point(entity.dxf.center), 'radius': float(entity.dxf.radius),
                    'start_angle': float(entity.dxf.start_angle), 'end_angle': float(entity.dxf.end_angle)
                })
        except: pass

    for layout in layouts:
        try:
            for entity in layout:
                process_entity(entity)
                try:
                    if entity.dxftype() == 'INSERT':
                        if hasattr(entity, 'attribs') and entity.attribs:
                            for attrib in entity.attribs: process_entity(attrib)
                        if hasattr(entity, 'virtual_entities') and callable(entity.virtual_entities):
                            for v_entity in entity.virtual_entities(): process_entity(v_entity)
                except: pass
        except: pass
        
    return texts, shapes

def apply_text_inheritance(final_aggregated_data):
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
                cell_val, final_aggregated_data[row_idx][col_idx] = "", ""
            if cell_val in ["〃", "”", "\"", "''", "””", "''", "同上", "...", "…"]:
                if last_text: final_aggregated_data[row_idx][col_idx] = last_text
            elif cell_val != "":
                last_text = cell_val if is_text_to_inherit(cell_val) else ""