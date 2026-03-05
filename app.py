from flask import Flask, render_template, request, jsonify, redirect, url_for
import pandas as pd
import os
import json

app = Flask(__name__)

# 파일 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PLAN = os.path.join(BASE_DIR, '(반출승인)생산2팀_생산진도표(TC)_260223.xlsx')
EXCEL_FILE_DAILY = os.path.join(BASE_DIR, 'TC 대조립 일일 진도현황(260304).xlsx')

# 전역 변수 (캐싱용)
_SHEET_CACHE = {"data": [], "last_mtime": 0}
import threading
import time

def update_cache_bg():
    global _SHEET_CACHE
    while True:
        try:
            if not os.path.exists(EXCEL_FILE_DAILY):
                time.sleep(10)
                continue
                
            mtime = os.path.getmtime(EXCEL_FILE_DAILY)
            if _SHEET_CACHE["last_mtime"] == mtime and _SHEET_CACHE["data"]:
                time.sleep(10)
                continue

            from openpyxl import load_workbook
            wb = load_workbook(EXCEL_FILE_DAILY, read_only=True)
            
            valid_sheets = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                if ws.sheet_state == 'visible' and (sheet_name.startswith('TC 호기별') or sheet_name == 'TC 대조립'):
                    valid_sheets.append(sheet_name)
            wb.close()
            
            formatted_sheets = []
            
            if 'TC 대조립' in valid_sheets:
                formatted_sheets.append({
                    "type": "daily",
                    "sheet": "TC 대조립",
                    "display": "TC 대조립"
                })
                valid_sheets.remove('TC 대조립')

            daily_sheets = []
            import re
            for s in valid_sheets:
                match = re.search(r'(\d+)년\s*(\d+)월', s)
                sort_key = ""
                if match:
                    year, month = match.groups()
                    full_year = f"20{year}" if len(year) == 2 else year
                    sort_key = f"{full_year}{int(month):02d}"
                
                daily_sheets.append({
                    "type": "daily",
                    "sheet": s,
                    "display": s,
                    "sort_key": sort_key
                })
            
            daily_sheets.sort(key=lambda x: x['sort_key'], reverse=True)
            formatted_sheets.extend([{k: v for k, v in d.items() if k != 'sort_key'} for d in daily_sheets])
            
            _SHEET_CACHE["data"] = formatted_sheets
            _SHEET_CACHE["last_mtime"] = mtime
            print("Background cache updated successfully.")

        except Exception as e:
            print(f"Error in bg cache worker: {e}")
        
        # 10초마다 엑셀 파일 변경 여부 확인
        time.sleep(10)

# 백그라운드 스레드 시작
threading.Thread(target=update_cache_bg, daemon=True).start()

def get_excel_sheets():
    # 백그라운드 스레드에 의해 즉각적으로 준비된 캐시 반환
    return _SHEET_CACHE["data"] if _SHEET_CACHE["data"] else [
        { "type": "plan", "sheet": None, "display": "출하 계획표 (기본)" }
    ]

def parse_production_data(file_path, sheet_name=0):
    try:
        print(f"Parsing File: {file_path}, Sheet: {sheet_name}")
        # 1. 먼저 헤더가 1(2번째 줄)인 경우 시도
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
        
        # 'TC 호기별' 시트가 아니면서, 컬럼명이 비정상적인 경우에만 폴백
        is_monthly_sheet = str(sheet_name).startswith('TC 호기별')
        if not is_monthly_sheet:
            if len(df.columns) < 5 or ("Unnamed" in str(df.columns[0]) and "Unnamed" in str(df.columns[1])):
                print("Trying header=0 as fallback...")
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

        col_count = len(df.columns)
        row_count = len(df)
        print(f"Excel Structure: {row_count} rows, {col_count} columns")

        # 데이터가 너무 적으면 빈 결과 반환
        if col_count < 10:
            print(f"Error: Too few columns ({col_count}). Expected at least 10.")
            return []

        # 헤더 인덱스 찾기 및 컬럼 동적 매핑
        col_map = {}
        header_idx = -1
        
        # 1. df.columns가 이미 헤더인지 확인 (header=0, 1 옵션 적용 시)
        columns_str = " ".join([str(c) for c in df.columns]).upper()
        if 'NO.' in columns_str or 'NO ' in columns_str or '생산월' in columns_str:
            for col_idx, val in enumerate(df.columns):
                clean_name = str(val).replace('\n', '').replace(' ', '').upper()
                col_map[clean_name] = col_idx
            # 데이터는 df의 모든 행이므로 header_idx = -1 유지
        else:
            # 2. 데이터 중에 헤더가 있는지 스캔
            for idx, row in df.iterrows():
                row_str = " ".join([str(val) for val in row if pd.notna(val)]).upper()
                if 'NO.' in row_str or 'NO ' in row_str or '생산월' in row_str:
                    header_idx = idx
                    break
            
            if header_idx != -1:
                header_row = df.iloc[header_idx]
                for col_idx, val in enumerate(header_row):
                    if pd.notna(val):
                        clean_name = str(val).replace('\n', '').replace(' ', '').upper()
                        col_map[clean_name] = col_idx
                    
        def get_idx(*possible_names, default=None):
            for name in possible_names:
                clean_target = name.upper().replace(' ', '')
                if clean_target in col_map:
                    return col_map[clean_target]
            return default

        no_idx = get_idx('NO.', 'NO', default=1)
        month_idx = get_idx('생산월', default=2)
        section_idx = get_idx('생산직', default=3)
        model_idx = get_idx('기종', '전산기종', default=5)
        serial_idx_mapped = get_idx('호기', default=-1)
        order_idx = get_idx('오더', default=7)
        customer_idx = get_idx('출하처', '목적지', default=8)
        first_shipment_idx = get_idx('최초출하일', '최조출하일', default=9)
        target_idx = get_idx('개정출하일', default=10)
        base_idx = get_idx('BASE시작일', 'BASE작업일', default=11)
        first_start_idx = get_idx('최초시작일', default=12)
        revised_start_idx = get_idx('개정시작일', default=13)
        nc_idx = get_idx('NC', default=14)
        status_idx = get_idx('현공정', '진행상태', default=15)
        issue_idx = get_idx('ISSUE사항', '비고', default=16)

        # 날짜 형식 최적화 헬퍼 함수
        def format_date_string(date_str):
            if pd.isna(date_str) or str(date_str).lower() == 'nan' or not str(date_str).strip():
                return ''
            s = str(date_str).replace('\n', ' ').strip()
            if ' 00:00:00' in s:
                s = s.split(' ')[0]
            elif 'T00:00:00' in s:
                s = s.split('T')[0]
            return s
            
        def get_val(row, col_index, default_val=''):
            if col_index is not None and col_index != -1 and col_index < len(row) and pd.notna(row.iloc[col_index]):
                v = str(row.iloc[col_index]).strip()
                if v.lower() == 'nan': return default_val
                if v.endswith('.0') and v[:-2].isdigit(): v = v[:-2]
                return v
            return default_val

        data = []
        for idx, row in df.iterrows():
            if idx <= header_idx:
                continue
            if len(row) < 10:
                continue

            serial = ''
            if serial_idx_mapped != -1 and serial_idx_mapped < len(row):
                v = str(row.iloc[serial_idx_mapped]).strip()
                if v.lower() not in ['nan', 'none', '-', '']:
                    serial = v
            
            if not serial:
                potential_indices = [6, 5, 7, 8, 4]
                for p_idx in potential_indices:
                    if p_idx < len(row):
                        val = str(row.iloc[p_idx]).strip()
                        if val and val.lower() not in ['nan', 'none', '-', '']:
                            serial = val
                            break

            if not serial:
                continue

            first_val = str(row.iloc[1]).strip().upper() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
            second_val = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ''
            if first_val == 'NO.' or first_val == 'NO' or second_val == '생산월':
                continue

            item = {
                'no': get_val(row, no_idx),
                'month': get_val(row, month_idx),
                'section': get_val(row, section_idx),
                'model': get_val(row, model_idx).replace('\n', ' '),
                'serial': str(serial).replace('\n', ' '),
                'order': get_val(row, order_idx),
                'customer': get_val(row, customer_idx).replace('\n', ' '),
                'first_shipment': format_date_string(get_val(row, first_shipment_idx)),
                'target': format_date_string(get_val(row, target_idx)),
                'base': format_date_string(get_val(row, base_idx)),
                'first_start': format_date_string(get_val(row, first_start_idx)),
                'revised_start': format_date_string(get_val(row, revised_start_idx)),
                'nc': format_date_string(get_val(row, nc_idx)),
                'status': get_val(row, status_idx, '대기'),
                'issue': get_val(row, issue_idx).replace('\n', '<br>')
            }
            if not item['status']:
                item['status'] = '대기'
                
            data.append(item)
        
        print(f"Successfully parsed {len(data)} items.")
        return data
    except Exception as e:
        print(f"Error parsing excel: {e}")
        import traceback
        traceback.print_exc()
        return []

@app.errorhandler(404)
def page_not_found(e):
    # 사용자가 잘못된 주소(예: /1직 등)로 접속 시 404 에러 화면 대신 메인 화면으로 돌려보냅니다.
    return redirect(url_for('index'))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/work-report')
@app.route('/work_report')
def work_report():
    return render_template('work_report.html')

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file:
        # 임시 파일로 저장 (또는 메모리 데이터 읽기 가능하지만 pandas는 파일 경로 선호)
        temp_path = os.path.join(BASE_DIR, 'temp_upload.xlsx')
        file.save(temp_path)
        try:
            data = parse_production_data(temp_path)
            return jsonify(data)
        finally:
            # 처리 후 임시 파일 삭제
            if os.path.exists(temp_path):
                os.remove(temp_path)

@app.route('/api/sheets')
def get_sheets():
    sheets = get_excel_sheets()
    return jsonify(sheets)

@app.route('/api/data')
def get_data():
    file_type = request.args.get('type', 'plan') # 'plan' or 'daily'
    sheet_name = request.args.get('sheet', None)
    
    if file_type == 'daily' and sheet_name:
        data = parse_production_data(EXCEL_FILE_DAILY, sheet_name=sheet_name)
    else:
        # 기본값은 출하 계획표 시트
        data = parse_production_data(EXCEL_FILE_PLAN, sheet_name=0)
    return jsonify(data)

if __name__ == '__main__':
    # 5000번 포트에서 실행
    print("공정 관리 시스템 서버를 시작합니다...")
    # 포트를 5003로 변경하여 기존 프로세스와의 충돌 방지 및 브라우저 캐시 완벽 우회
    app.run(debug=True, port=5003, threaded=True)
