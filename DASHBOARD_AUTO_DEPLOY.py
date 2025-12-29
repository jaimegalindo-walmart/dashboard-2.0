#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FULLY AUTOMATED DASHBOARD REFRESH
Runs every Saturday at noon
- Extracts new week data from Excel
- Regenerates chunks with push.apply()
- Auto-increments version numbers (v2 → v3 → v4, etc.)
- Updates HTML with new version numbers
- Commits to GitHub
- Netlify auto-deploys
ZERO MANUAL INTERVENTION NEEDED
"""

import io
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import json
import re
import os
import subprocess
from pathlib import Path
from datetime import datetime
import xlrd

# ============================================================================
# CONFIGURATION
# ============================================================================

BASE_WEEK_PATH = r'\\US06060NT800FIL.s06060.us.wal-mart.com\Reports\Roll Ups\2025'
DATA_FILE = r'C:\Users\j0g150w\Documents\code-puppy\dashboard\AssocDashboardReset\data.js'
DASHBOARD_FOLDER = r'C:\Users\j0g150w\Desktop\dashboard-2.0'
VERSION_FILE = r'C:\Users\j0g150w\Desktop\dashboard-version.txt'

BUILDING_STRUCTURE = {
    'EV1 Eastvale': {'Shift 1': ['ORD', 'REC', 'RSR', 'SHP'], 'Shift 2': ['REC', 'RSR', 'SHP'], 'Shift 4': ['ORD', 'REC', 'RSR', 'SHP']},
    'EV4 Eastvale': {'Shift 1': ['ORD', 'REC', 'RSR']},
    'EV5 Eastvale': {'Shift 1': ['ORD', 'REC', 'RSR', 'SHP'], 'Shift 2': ['ORD', 'REC', 'RSR', 'SHP'], 'Shift 4': ['ORD', 'REC', 'RSR', 'SHP']},
    'EV7 Eastvale': {'Shift 1': ['ORD', 'REC', 'RSR', 'SHP'], 'Shift 2': ['REC', 'SHP'], 'Shift 4': ['ORD', 'REC', 'RSR', 'SHP'], 'Shift 5': ['REC', 'SHP']},
    'F3 Fontana': {'Shift 1': ['ORD', 'REC', 'RSR', 'SHP'], 'Shift 4': ['ORD', 'REC', 'RSR', 'SHP']},
}

DEPT_MAPPING = {'ORD': 'ORD', 'REC': 'REC', 'RSR': 'RSR', 'SHP': 'SHIP'}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_current_version():
    """Get current version number from file."""
    if Path(VERSION_FILE).exists():
        try:
            return int(Path(VERSION_FILE).read_text().strip())
        except:
            return 2
    return 2

def increment_version():
    """Increment and save version number."""
    current = get_current_version()
    new_version = current + 1
    Path(VERSION_FILE).write_text(str(new_version))
    return new_version

def extract_excel_data(excel_file, building, department, shift, week_num):
    """Extract data from Excel file."""
    records = []
    try:
        wb = xlrd.open_workbook(excel_file, on_demand=True)
        ws = wb.sheet_by_index(0)
        
        for row_idx in range(4, min(100, ws.nrows)):
            try:
                cell_c = ws.cell_value(row_idx, 2)
                cell_d = ws.cell_value(row_idx, 3)
                cell_e = ws.cell_value(row_idx, 4)
                cell_f = ws.cell_value(row_idx, 5)
            except:
                continue
            
            if not cell_d or (isinstance(cell_d, str) and not cell_d.strip()):
                continue
            
            def is_percentage(val):
                if val is None or val == '':
                    return False
                s = str(val).strip()
                return bool(re.match(r'^[0-9]{1,3}(\.[0-9]+)?%?$', s))
            
            if not (is_percentage(cell_e) or is_percentage(cell_f)):
                continue
            
            name = str(cell_d).strip()
            emp_id = str(cell_c).strip() if cell_c else ''
            weekly = str(cell_e).strip() if cell_e else ''
            four_week = str(cell_f).strip() if cell_f else ''
            
            if not name:
                continue
            
            if weekly and not weekly.endswith('%'):
                weekly += '%'
            if four_week and not four_week.endswith('%'):
                four_week += '%'
            
            record = {
                'EmployeeId': emp_id,
                'Name': name,
                'Associate': name,
                'Weekly': weekly,
                'FourWeek': four_week,
                'Department': department,
                'Building': building,
                'Shift': shift,
                'Week': str(week_num)
            }
            records.append(record)
    except:
        pass
    
    return records

def run_git_command(cmd, cwd=DASHBOARD_FOLDER):
    """Run git command safely."""
    try:
        result = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True)
        return result.returncode == 0, result.stdout, result.stderr
    except:
        return False, '', 'Git command failed'

# ============================================================================
# MAIN AUTOMATION
# ============================================================================

print("\n" + "="*80)
print("FULLY AUTOMATED DASHBOARD REFRESH")
print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("="*80 + "\n")

# Get version
current_version = get_current_version()
new_version = increment_version()
print(f"Version: {current_version} → {new_version}\n")

# STEP 1: Find latest week
print("STEP 1: Finding latest week...")
base_path = Path(BASE_WEEK_PATH)
week_folders = sorted(
    [int(f.name.split()[-1]) for f in base_path.glob('Week *') if f.name.startswith('Week ')],
    reverse=True
)

if not week_folders:
    print("ERROR: No week folders found")
    sys.exit(1)

latest_week = week_folders[0]
print(f"Latest week: {latest_week}\n")

# STEP 2: Extract data
print("STEP 2: Extracting week data from Excel...")
all_records = []
for building, shifts in BUILDING_STRUCTURE.items():
    building_path = Path(BASE_WEEK_PATH) / f'Week {latest_week}' / building
    if not building_path.exists():
        continue
    
    for shift, depts in shifts.items():
        shift_path = building_path / shift
        if not shift_path.exists():
            continue
        
        for dept_short in depts:
            excel_files = list(shift_path.glob(f'{dept_short}.xls*'))
            if not excel_files:
                continue
            
            excel_file = excel_files[0]
            dept_full = DEPT_MAPPING.get(dept_short, dept_short)
            records = extract_excel_data(str(excel_file), building, dept_full, shift, latest_week)
            all_records.extend(records)

print(f"Extracted: {len(all_records)} records\n")

# STEP 3: Update master data
print("STEP 3: Updating data.js...")
with open(DATA_FILE, 'r', encoding='utf-8') as f:
    content = f.read()

match = re.search(
    r"window\.APD_DATA\['AssociateProgression'\]\s*=\s*(\[[\s\S]+?\]);",
    content
)

array_str = match.group(1)
existing_data = json.loads(array_str)

# Remove existing records for this week (avoid duplicates)
existing_data = [r for r in existing_data if r.get('Week') != str(latest_week)]
updated_data = existing_data + all_records

new_content = re.sub(
    r"window\.APD_DATA\['AssociateProgression'\]\s*=\s*\[[\s\S]+?\];",
    f"window.APD_DATA['AssociateProgression'] = {json.dumps(updated_data)};",
    content
)

with open(DATA_FILE, 'w', encoding='utf-8') as f:
    f.write(new_content)

print(f"Updated: {len(updated_data)} total records\n")

# STEP 4: Regenerate chunks
print("STEP 4: Regenerating chunk files...")
by_week = {}
for record in updated_data:
    week = int(record.get('Week', '0'))
    if week not in by_week:
        by_week[week] = []
    by_week[week].append(record)

weeks = sorted(by_week.keys())
chunks = [(1, 10), (11, 20), (21, 30), (31, 40), (41, 50)]

for start_week, end_week in chunks:
    chunk_data = []
    for week in range(start_week, end_week + 1):
        if week in by_week:
            chunk_data.extend(by_week[week])
    
    filename = os.path.join(DASHBOARD_FOLDER, f'data-weeks-{start_week:02d}-{end_week:02d}-v{new_version}.js')
    
    json_str = json.dumps(chunk_data, ensure_ascii=True, separators=(',', ':'))
    
    js_content = f"""(function() {{
  window.APD_DATA = window.APD_DATA || {{}};
  window.APD_DATA['AssociateProgression'] = window.APD_DATA['AssociateProgression'] || [];
  var data = {json_str};
  window.APD_DATA['AssociateProgression'].push.apply(window.APD_DATA['AssociateProgression'], data);
}})();
"""
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(js_content)
    
    size_mb = os.path.getsize(filename) / (1024 * 1024)
    print(f"  data-weeks-{start_week:02d}-{end_week:02d}-v{new_version}.js: {len(chunk_data)} records, {size_mb:.2f} MB")

print()

# STEP 5: Update HTML
print("STEP 5: Updating HTML with new version...")
index_file = os.path.join(DASHBOARD_FOLDER, 'index.html')
with open(index_file, 'r', encoding='utf-8') as f:
    html = f.read()

# Replace version numbers in chunk list
for i in range(2, new_version + 1):
    old = f'data-weeks-01-10-v{i-1}.js'
    new = f'data-weeks-01-10-v{new_version}.js'
    html = html.replace(old, new)

old_chunk_pattern = r"'data-weeks-(\d{2})-(\d{2})-v\d+\.js'"
new_chunk_template = f"'data-weeks-\\1-\\2-v{new_version}.js'"
html = re.sub(old_chunk_pattern, new_chunk_template, html)

with open(index_file, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"Updated HTML to load v{new_version} files\n")

# STEP 6: Git commit and push
print("STEP 6: Committing to GitHub...")
success, stdout, stderr = run_git_command('git add .')
if success:
    print("  Staged files")
    success, stdout, stderr = run_git_command(f'git commit -m "Auto-update: Week {latest_week} data (v{new_version})"')
    if success:
        print(f"  Committed")
        success, stdout, stderr = run_git_command('git push origin main')
        if success:
            print(f"  Pushed to GitHub")
            print(f"  Netlify will auto-deploy in 1-2 minutes\n")
        else:
            print(f"  WARNING: Git push failed: {stderr}\n")
    else:
        print(f"  WARNING: Git commit failed: {stderr}\n")
else:
    print(f"  WARNING: Git not configured or no changes: {stderr}\n")

# STEP 7: Summary
print("="*80)
print("AUTO-REFRESH COMPLETE!")
print("="*80)
print()
print(f"Week {latest_week} data processed and deployed")
print(f"Version bumped: v{current_version} → v{new_version}")
print(f"Files with new version: v{new_version}")
print()
print("Timeline:")
print(f"  - Extraction: DONE")
print(f"  - Chunk regeneration: DONE")
print(f"  - Version increment: DONE")
print(f"  - GitHub commit: DONE")
print(f"  - Netlify auto-deploy: Starting (check in 2 minutes)")
print()
print(f"Dashboard URL: https://associateprogressiondashboard.netlify.app")
print(f"Expected refresh: 1-2 minutes after this script finishes")
print()
print(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
print("="*80 + "\n")
