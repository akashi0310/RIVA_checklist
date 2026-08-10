import pandas as pd
import json
import os
import sys
import re

if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

print("=== HỆ THỐNG CẬP NHẬT DỮ LIỆU DASHBOARD TUẤN (CHECKLIST ẤN PHẨM TRUYỀN THÔNG) ===")

VALID_PEOPLE = ['Dũng', 'Ánh', 'Phương', 'Lộc', 'Vân Anh', 'Nam', 'Chi', 'Hà', 'Diệu Anh', 'Hường', 'Hảo']

def find_excel_file(filename):
    paths = [
        os.path.join('Excel_file', filename),
        filename
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None

def clean_and_extract_names(text):
    if not text or str(text) == 'nan': return []
    s = str(text)
    s = re.sub(r'(\w+)\.\s*(Mr|Ms|mr|ms)\b', r'\1, \2', s)
    s = s.replace(';', ',')
    
    parts = s.split(',')
    names = []
    for p in parts:
        p = p.strip()
        if not p: continue
        p_clean = re.sub(r'^(mr|ms)\.?\s*', '', p, flags=re.IGNORECASE).strip()
        
        if p_clean in ['Dũng', 'Mr Dũng', 'Mr. Dũng']: p_clean = 'Dũng'
        elif p_clean in ['Ánh', 'Ms Ánh', 'Ms. Ánh']: p_clean = 'Ánh'
        elif p_clean in ['Phương', 'Ms Phương', 'Ms. Phương']: p_clean = 'Phương'
        elif p_clean in ['Lộc', 'Mr Lộc', 'Mr. Lộc']: p_clean = 'Lộc'
        elif p_clean in ['Vân Anh', 'Ms Vân Anh', 'Ms. Vân Anh']: p_clean = 'Vân Anh'
        elif p_clean in ['Nam', 'Mr Nam', 'Mr. Nam']: p_clean = 'Nam'
        elif p_clean in ['Chi', 'Ms Chi', 'Ms. Chi']: p_clean = 'Chi'
        elif p_clean in ['Hà', 'Ms Hà', 'Ms. Hà']: p_clean = 'Hà'
        elif p_clean in ['Diệu Anh', 'Ms Diệu Anh']: p_clean = 'Diệu Anh'
        elif p_clean in ['Hường', 'Ms Hường']: p_clean = 'Hường'
        elif p_clean in ['Hảo', 'Mr Hảo']: p_clean = 'Hảo'

        if p_clean and p_clean in VALID_PEOPLE:
            if p_clean not in names:
                names.append(p_clean)

    return names

tuan_tasks = []
file_path = find_excel_file('CHECKLIST ẤN PHẨM DU HỌC.xlsx')

if file_path:
    df = pd.read_excel(file_path)
    col_title = df.columns[0]
    current_cat_name = "ẤN PHẨM TRUYỀN THÔNG"
    current_cat_code = "all"
    
    for row in df.to_dict(orient='records'):
        v0 = row.get(col_title)
        v1 = row.get("Unnamed: 1")
        v2 = row.get("Unnamed: 2") # Brief
        v3 = row.get("Unnamed: 3") # Nguồn dữ liệu / Tài liệu tham khảo
        v4 = row.get("Unnamed: 4") # Đầu ra cần bàn giao / Thành phẩm
        v5 = row.get("Unnamed: 5") # PIC (Director)
        v6 = row.get("Unnamed: 6") # Executor
        v7 = row.get("Unnamed: 7") # Deadline
        v8 = row.get("Unnamed: 8") # Status

        if v0 == "STT" or (v0 is None and v1 is None): continue
        if isinstance(v0, str) and not str(v0).isdigit():
            cat_letter = str(v0).strip()
            cat_title = str(v1).strip() if v1 and str(v1) != 'nan' else cat_letter
            current_cat_name = f"{cat_letter}. {cat_title}"
            current_cat_code = f"cat_{cat_letter.lower()}"
            continue

        if v1 and str(v1).strip() != 'nan':
            dirs = clean_and_extract_names(v5)
            execs = clean_and_extract_names(v6)
            stt_val = v0 if (v0 is not None and str(v0) != 'nan') else len(tuan_tasks) + 1
            
            # Status: 2 states: ☑ Done, □ Pending
            status_val = '☑' if (v8 and 'Hoàn thành' in str(v8)) else '□'
            
            tuan_tasks.append({
                "stt": stt_val,
                "category": current_cat_name,
                "category_code": current_cat_code,
                "task_name": str(v1).strip(),
                "brief": str(v2).strip() if v2 and str(v2) != 'nan' else '',
                "ref_doc": str(v3).strip() if v3 and str(v3) != 'nan' else '',
                "product": str(v4).strip() if v4 and str(v4) != 'nan' else '',
                "director": ', '.join(dirs),
                "executor": ', '.join(execs),
                "deadline": str(v7).split(' ')[0] if v7 and str(v7) != 'nan' else 'Theo tiến độ',
                "status": status_val
            })

def build_people_stats(task_list):
    people = {}
    for t in task_list:
        dirs = [p.strip() for p in t['director'].split(',') if p.strip()]
        execs = [p.strip() for p in t['executor'].split(',') if p.strip()]
        all_persons = set(dirs + execs)

        for p in all_persons:
            if p not in people:
                people[p] = {
                    "name": p,
                    "directorCount": 0,
                    "executorCount": 0,
                    "totalTasks": 0,
                    "tasks": []
                }
            is_dir = p in dirs
            is_exec = p in execs

            if is_dir: people[p]["directorCount"] += 1
            if is_exec: people[p]["executorCount"] += 1

            people[p]["totalTasks"] += 1
            people[p]["tasks"].append({
                "stt": t["stt"],
                "task_name": t["task_name"],
                "category": t["category"],
                "roles": [r for r, flags in [("Phụ trách (PIC)", is_dir), ("Thực hiện", is_exec)] if flags],
                "deadline": t["deadline"],
                "status": t["status"]
            })
    return list(people.values())

final_dataset = {
    "title": "CHECKLIST CÔNG VIỆC ẤN PHẨM TRUYỀN THÔNG",
    "tasks": tuan_tasks,
    "people": build_people_stats(tuan_tasks),
    "categories": [
        {"code": "all", "name": "Tất Cả Hạng Mục"},
        {"code": "cat_a", "name": "A. Ấn Phẩm Thương Hiệu"},
        {"code": "cat_b", "name": "B. Ấn Phẩm Quảng Bá"},
        {"code": "cat_c", "name": "C. Thông Tin Du Học"},
        {"code": "cat_d", "name": "D. Social Media"},
        {"code": "cat_e", "name": "E. Ấn Phẩm Sự Kiện"},
        {"code": "cat_f", "name": "F. Ấn Phẩm Du Học Sinh"}
    ]
}

js_content = f"const TUAN_PROJECT_DATA = {json.dumps(final_dataset, ensure_ascii=False, indent=2)};"

with open('dashboard_data_tuan.js', 'w', encoding='utf-8') as f:
    f.write(js_content)

print(f"OK! Loaded {len(tuan_tasks)} publication tasks from CHECKLIST ẤN PHẨM DU HỌC.xlsx.")
print("-> dashboard_data_tuan.js generated successfully!")
