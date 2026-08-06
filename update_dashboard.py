import pandas as pd
import json
import os
import sys
import re

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

print("=== HỆ THỐNG CẬP NHẬT DỮ LIỆU TỰ ĐỘNG - LÀM SẠCH TÊN BỎ MR/MS ===")

VALID_PEOPLE = ['Hà', 'Diệu Anh', 'Hường', 'Hải', 'Hảo', 'Dũng', 'Vân Anh', 'Chi', 'Thảo Nguyên', 'Hòa', 'Nam', 'Lộc', 'Toàn bộ nhân viên', 'GVHD', 'CTV', 'Học sinh']

def clean_and_extract_names(text):
    if not text or str(text) == 'nan': return []
    s = str(text)
    # Fix typos like "Hòa. Mr" or "Chi. Ms"
    s = re.sub(r'(\w+)\.\s*(Mr|Ms|mr|ms)\b', r'\1, \2', s)
    s = s.replace(';', ',')
    
    parts = s.split(',')
    names = []
    for p in parts:
        p = p.strip()
        if not p: continue
        # Strip Mr / Ms prefix
        p_clean = re.sub(r'^(mr|ms)\.?\s*', '', p, flags=re.IGNORECASE).strip()
        
        # Match against standard names
        if p_clean in ['Hà', 'Ms Hà', 'Ms. Hà']: p_clean = 'Hà'
        elif p_clean in ['Diệu Anh', 'Ms Diệu Anh']: p_clean = 'Diệu Anh'
        elif p_clean in ['Hường', 'Ms Hường']: p_clean = 'Hường'
        elif p_clean in ['Hải', 'Ms Hải']: p_clean = 'Hải'
        elif p_clean in ['Hảo', 'Mr Hảo']: p_clean = 'Hảo'
        elif p_clean in ['Dũng', 'Mr Dũng']: p_clean = 'Dũng'
        elif p_clean in ['Vân Anh', 'Ms Vân Anh']: p_clean = 'Vân Anh'
        elif p_clean in ['Chi', 'Ms Chi']: p_clean = 'Chi'
        elif p_clean in ['Thảo Nguyên', 'Ms Thảo Nguyên']: p_clean = 'Thảo Nguyên'
        elif p_clean in ['Hòa', 'Ms Hòa']: p_clean = 'Hòa'
        elif p_clean in ['Nam', 'Mr Nam']: p_clean = 'Nam'
        elif p_clean in ['Lộc', 'Mr Lộc']: p_clean = 'Lộc'

        if p_clean and p_clean in VALID_PEOPLE:
            if p_clean not in names:
                names.append(p_clean)

    return names

# 1. AMSIO
amsio_tasks = []
if os.path.exists('AMSIO.xlsx'):
    df_amsio = pd.read_excel('AMSIO.xlsx')
    current_cat = "A. TRIỂN KHAI HỆ THỐNG AMSIO VIỆT NAM"
    for row in df_amsio.to_dict(orient='records'):
        v0 = row.get("CHECKLIST TRIỂN KHAI HỆ THỐNG & TRUYỀN THÔNG KỲ THI XÃ HỘI HÓA")
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2")
        v3 = row.get("Unnamed: 3")
        v4 = row.get("Unnamed: 4")
        v5 = row.get("Unnamed: 5")
        v6 = row.get("Unnamed: 6")
        v7 = row.get("Unnamed: 7")

        if v0 == "STT" or (v0 is None and v1 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            current_cat = str(v0).strip()
            continue

        if v1 and str(v1).strip() != 'nan':
            dirs = clean_and_extract_names(v2)
            coords = clean_and_extract_names(v3)
            execs = clean_and_extract_names(v4)

            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(amsio_tasks) + 1
            amsio_tasks.append({
                "stt": stt_val,
                "category": current_cat,
                "task_name": str(v1).strip(),
                "director": ', '.join(dirs),
                "coordinator": ', '.join(coords),
                "executor": ', '.join(execs),
                "product": str(v5).strip() if v5 and str(v5) != 'nan' else '',
                "deadline": str(v6).split(' ')[0] if v6 and str(v6) != 'nan' else '',
                "status": str(v7).strip() if v7 and str(v7) != 'nan' else '□'
            })

# 2. AP
ap_tasks = []
ap_meta = {
    "title": "KẾ HOẠCH TỔNG THỂ TUYỂN SINH CHƯƠNG TRÌNH AP (ADVANCED PLACEMENT)",
    "timeframe": "05/08/2026 – 31/08/2026",
    "opening": "05/09/2026",
    "quota": "30 học sinh (03 lớp, 10 học sinh/lớp)",
    "target": "Học sinh lớp 8–11 có định hướng du học và săn học bổng."
}

if os.path.exists('AP.xlsx'):
    df_ap = pd.read_excel('AP.xlsx')
    current_cat = "A. CHUẨN BỊ CHƯƠNG TRÌNH & HỒ SƠ"
    for row in df_ap.to_dict(orient='records'):
        v0 = row.get("Unnamed: 0")
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2")
        v3 = row.get("Unnamed: 3")
        v4 = row.get("Unnamed: 4")
        v5 = row.get("Unnamed: 5")
        v6 = row.get("Unnamed: 6")

        if v0 == "STT" or (v0 is None and v1 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            if str(v0).startswith(('A.', 'B.', 'C.', 'D.')):
                current_cat = str(v0).strip()
            continue

        if v1 and str(v1).strip() != 'nan':
            dirs = clean_and_extract_names(v2)
            coords = clean_and_extract_names(v3)
            execs = clean_and_extract_names(v4)

            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(ap_tasks) + 1
            ap_tasks.append({
                "stt": stt_val,
                "category": current_cat,
                "task_name": str(v1).strip(),
                "director": ', '.join(dirs),
                "coordinator": ', '.join(coords),
                "executor": ', '.join(execs),
                "product": str(v5).strip() if v5 and str(v5) != 'nan' else '',
                "deadline": str(v6).split(' ')[0] if v6 and str(v6) != 'nan' else '',
                "status": '□'
            })

# 3. IENA
iena_tasks = []
iena_meta = {
    "title": "KẾ HOẠCH TRIỂN KHAI ĐỘI TUYỂN IENA 2026 (ĐỨC)",
    "timeframe": "Tháng 10/2026",
    "location": "Đức (Germany)"
}

if os.path.exists('IENA.xlsx'):
    df_iena = pd.read_excel('IENA.xlsx')
    col_name = "KẾ HOẠCH TRIỂN KHAI ĐỘI TUYỂN IENA 2026 (ĐỨC)"
    current_cat = "I. TUYỂN SINH"
    for row in df_iena.to_dict(orient='records'):
        v0 = row.get(col_name)
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2")
        v3 = row.get("Unnamed: 3")
        v4 = row.get("Unnamed: 4")
        v5 = row.get("Unnamed: 5")
        v6 = row.get("Unnamed: 6")

        if v0 == "STT" or (v0 is None and v1 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            if str(v0).startswith(('I.', 'II.', 'III.', 'IV.')):
                current_cat = str(v0).strip()
            continue

        if v1 and str(v1).strip() != 'nan':
            dirs = clean_and_extract_names(v2)
            coords = clean_and_extract_names(v3)
            execs = clean_and_extract_names(v4)

            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(iena_tasks) + 1
            iena_tasks.append({
                "stt": stt_val,
                "category": current_cat,
                "task_name": str(v1).strip(),
                "director": ', '.join(dirs),
                "coordinator": ', '.join(coords),
                "executor": ', '.join(execs),
                "product": str(v5).strip() if v5 and str(v5) != 'nan' else '',
                "deadline": str(v6).split(' ')[0] if v6 and str(v6) != 'nan' else '',
                "status": '□'
            })

# 4. IPITEX
ipitex_tasks = []
ipitex_meta = {
    "title": "KẾ HOẠCH TRIỂN KHAI ĐỘI TUYỂN IPITEX 2027 (THÁI LAN)",
    "timeframe": "Tháng 02/2027",
    "location": "Thái Lan (Thailand)"
}

if os.path.exists('IPITEX.xlsx'):
    df_ipitex = pd.read_excel('IPITEX.xlsx')
    col_name = "KẾ HOẠCH TRIỂN KHAI ĐỘI TUYỂN IPITEX 2027 (THÁI LAN)"
    current_cat = "I. TUYỂN SINH"
    for row in df_ipitex.to_dict(orient='records'):
        v0 = row.get(col_name)
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2")
        v3 = row.get("Unnamed: 3")
        v4 = row.get("Unnamed: 4")
        v5 = row.get("Unnamed: 5")
        v6 = row.get("Unnamed: 6")

        if v0 == "STT" or (v0 is None and v1 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            if str(v0).startswith(('I.', 'II.', 'III.', 'IV.')):
                current_cat = str(v0).strip()
            continue

        if v1 and str(v1).strip() != 'nan':
            dirs = clean_and_extract_names(v2)
            coords = clean_and_extract_names(v3)
            execs = clean_and_extract_names(v4)

            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(ipitex_tasks) + 1
            ipitex_tasks.append({
                "stt": stt_val,
                "category": current_cat,
                "task_name": str(v1).strip(),
                "director": ', '.join(dirs),
                "coordinator": ', '.join(coords),
                "executor": ', '.join(execs),
                "product": str(v5).strip() if v5 and str(v5) != 'nan' else '',
                "deadline": str(v6).split(' ')[0] if v6 and str(v6) != 'nan' else '',
                "status": '□'
            })

# 5. NCKH
nckh_tasks = []
nckh_meta = {
    "title": "KẾ HOẠCH TRIỂN KHAI CHƯƠNG TRÌNH NGHIÊN CỨU KHOA HỌC (NCKH) 2026–2027",
    "timeframe": "2026 - 2027",
    "scope": "Tuyển sinh, Đào tạo & Hướng dẫn NCKH, Thi đấu & Công bố"
}

if os.path.exists('NCKH.xlsx'):
    df_nckh = pd.read_excel('NCKH.xlsx')
    current_cat = "I. TUYỂN SINH"
    for row in df_nckh.to_dict(orient='records'):
        v0 = row.get("Unnamed: 0")
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2")
        v3 = row.get("Unnamed: 3")
        v4 = row.get("Unnamed: 4")
        v5 = row.get("Unnamed: 5")
        v6 = row.get("Unnamed: 6")
        v7 = row.get("Unnamed: 7")

        if v0 == "STT" or (v0 is None and v1 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            if str(v0).startswith(('I.', 'II.', 'III.', 'IV.', 'V.')):
                current_cat = str(v0).strip()
            continue

        if v1 and str(v1).strip() != 'nan':
            dirs = clean_and_extract_names(v2)
            coords = clean_and_extract_names(v3)
            execs = clean_and_extract_names(v4)

            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(nckh_tasks) + 1
            nckh_tasks.append({
                "stt": stt_val,
                "category": current_cat,
                "task_name": str(v1).strip(),
                "director": ', '.join(dirs),
                "coordinator": ', '.join(coords),
                "executor": ', '.join(execs),
                "product": str(v5).strip() if v5 and str(v5) != 'nan' else '',
                "deadline": str(v6).split(' ')[0] if v6 and str(v6) != 'nan' else '',
                "status": str(v7).strip() if v7 and str(v7) != 'nan' else '□'
            })

# 6. SẢN PHẨM TRUYỀN THÔNG
media_tasks = []
media_file = 'SẢN PHẨM TRUYỀN THÔNG.xlsx'
if os.path.exists(media_file):
    df_media = pd.read_excel(media_file)
    col_title = "DANH MỤC SẢN PHẨM TRUYỀN THÔNG VÀ PHÂN CÔNG NHÂN SỰ"
    current_cat = "DANH MỤC TRUYỀN THÔNG DÙNG CHUNG"
    for row in df_media.to_dict(orient='records'):
        v0 = row.get(col_title)
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2")
        v3 = row.get("Unnamed: 3")
        v4 = row.get("Unnamed: 4")
        v5 = row.get("Unnamed: 5")

        if v0 == "STT" or (v0 is None and v1 is None and v2 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            if str(v0).startswith(('I.', 'II.', 'III.', 'IV.', 'V.', 'Gợi ý')):
                current_cat = str(v0).strip()
            continue

        v2_str = str(v2).strip() if v2 and str(v2) != 'nan' else ''
        v3_str = str(v3).strip() if v3 and str(v3) != 'nan' else ''

        if any(prefix in v2_str for prefix in ['Ms', 'Mr']):
            task_title = str(v1).strip() if v1 else ''
            dirs = clean_and_extract_names(v2)
            execs = clean_and_extract_names(v3)
            prod_output = str(v4).strip() if v4 and str(v4) != 'nan' else ''
        elif any(prefix in v3_str for prefix in ['Ms', 'Mr']):
            task_title = str(v2).strip() if v2 else (str(v1).strip() if v1 else '')
            dirs = clean_and_extract_names(v3)
            execs = clean_and_extract_names(v4)
            prod_output = str(v5).strip() if v5 and str(v5) != 'nan' else ''
        else:
            continue

        if task_title:
            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(media_tasks) + 1
            media_tasks.append({
                "stt": stt_val,
                "category": current_cat,
                "task_name": task_title,
                "director": ', '.join(dirs),
                "coordinator": '',
                "executor": ', '.join(execs),
                "product": prod_output,
                "deadline": "Theo tiến độ",
                "status": '□'
            })

def build_people_stats(task_list):
    people = {}
    for t in task_list:
        dirs = [p.strip() for p in t['director'].split(',') if p.strip()]
        coords = [p.strip() for p in t['coordinator'].split(',') if p.strip()]
        execs = [p.strip() for p in t['executor'].split(',') if p.strip()]

        all_persons = set(dirs + coords + execs)

        for p in all_persons:
            if p not in people:
                people[p] = {
                    "name": p,
                    "directorCount": 0,
                    "coordinatorCount": 0,
                    "executorCount": 0,
                    "totalTasks": 0,
                    "tasks": []
                }
            is_dir = p in dirs
            is_coord = p in coords
            is_exec = p in execs

            if is_dir: people[p]["directorCount"] += 1
            if is_coord: people[p]["coordinatorCount"] += 1
            if is_exec: people[p]["executorCount"] += 1

            people[p]["totalTasks"] += 1
            people[p]["tasks"].append({
                "stt": t["stt"],
                "task_name": t["task_name"],
                "category": t["category"],
                "roles": [r for r, flags in [("Điều hành", is_dir), ("Điều phối", is_coord), ("Thực hiện", is_exec)] if flags],
                "deadline": t["deadline"],
                "status": t["status"]
            })
    return list(people.values())

final_dataset = {
    "amsio": {
        "title": "KỲ THI AMSIO VIỆT NAM",
        "tasks": amsio_tasks,
        "people": build_people_stats(amsio_tasks)
    },
    "ap": {
        "title": "TUYỂN SINH CHƯƠNG TRÌNH AP",
        "meta": ap_meta,
        "tasks": ap_tasks,
        "people": build_people_stats(ap_tasks)
    },
    "iena": {
        "title": "ĐỘI TUYỂN IENA 2026 (ĐỨC)",
        "meta": iena_meta,
        "tasks": iena_tasks,
        "people": build_people_stats(iena_tasks)
    },
    "ipitex": {
        "title": "ĐỘI TUYỂN IPITEX 2027 (THÁI LAN)",
        "meta": ipitex_meta,
        "tasks": ipitex_tasks,
        "people": build_people_stats(ipitex_tasks)
    },
    "nckh": {
        "title": "NGHIÊN CỨU KHOA HỌC (NCKH) 2026–2027",
        "meta": nckh_meta,
        "tasks": nckh_tasks,
        "people": build_people_stats(nckh_tasks)
    },
    "media": {
        "title": "SẢN PHẨM TRUYỀN THÔNG & ẤN PHẨM",
        "meta": {"title": "DANH MỤC SẢN PHẨM TRUYỀN THÔNG VÀ PHÂN CÔNG NHÂN SỰ", "scope": "Ấn phẩm truyền thông, bộ nhận diện, brochure, poster & video"},
        "tasks": media_tasks,
        "people": build_people_stats(media_tasks)
    }
}

js_content = f"const MULTI_PROJECT_DATA = {json.dumps(final_dataset, ensure_ascii=False, indent=2)};"

with open('dashboard_data.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

print(f"OK! AMSIO: {len(amsio_tasks)} | AP: {len(ap_tasks)} | IENA: {len(iena_tasks)} | IPITEX: {len(ipitex_tasks)} | NCKH: {len(nckh_tasks)} | MEDIA: {len(media_tasks)}.")
print("-> Đã làm sạch toàn bộ danh sách nhân sự (loại bỏ hoàn toàn xưng danh Mr/Ms).")
