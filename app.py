from flask import Flask, render_template, request, jsonify, redirect, url_for
import pandas as pd
import os
import json

app = Flask(__name__)

# 파일 경로 설정
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, '생산2팀_생산진도표(TC)_260304_매출.xlsx')

def parse_production_data(file_path):
    try:
        # 1. 먼저 헤더가 1(2번째 줄)인 경우 시도
        df = pd.read_excel(file_path, sheet_name=0, header=1)
        
        # 만약 컬럼명이 'Unnamed'로 시작하거나 데이터가 너무 적으면 헤더 0으로 재시도
        if len(df.columns) < 5 or "Unnamed" in str(df.columns[0]):
            print("Trying header=0 as fallback...")
            df = pd.read_excel(file_path, sheet_name=0, header=0)

        col_count = len(df.columns)
        row_count = len(df)
        print(f"Excel Structure: {row_count} rows, {col_count} columns")

        # 데이터가 너무 적으면 빈 결과 반환
        if col_count < 10:
            print(f"Error: Too few columns ({col_count}). Expected at least 10.")
            return []

        # 날짜 형식 변환 및 JSON 준비
        data = []
        for _, row in df.iterrows():
            # 최소한의 데이터가 있는지 확인
            if len(row) < 10:
                continue

            # 호기번호(S/N) 위치 탐색 (기본은 6번 인덱스)
            # 파일 형식에 따라 5~8번 사이에 위치할 가능성이 높음
            serial = ''
            potential_indices = [6, 5, 7, 8, 4]
            for idx in potential_indices:
                if idx < len(row):
                    val = str(row.iloc[idx]).strip()
                    if val and val.lower() not in ['nan', 'none', '-', '']:
                        # 숫자/문자가 섞인 호기번호 형태인지 대략적으로 확인 (예: 1234, ABC-123 등)
                        serial = val
                        break
            
            if not serial:
                continue
                
            # 'NO.' 등 헤더 문구가 데이터로 인식된 경우 건너뜀
            first_val = str(row.iloc[0]).upper() if pd.notna(row.iloc[0]) else ''
            if 'NO' in first_val or '생산' in first_val:
                continue

            # 컬럼 매핑 (데이터 구조에 따라 유연하게 대응)
            item = {
                'no': str(row.iloc[1]).replace('.0', '') if len(row) > 1 and pd.notna(row.iloc[1]) else '',          
                'month': str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else '-', 
                'section': str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else '-', 
                'model': str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else '-',   
                'serial': serial,                
                'order': str(row.iloc[7]).replace('.0', '') if len(row) > 7 and pd.notna(row.iloc[7]) else '-',
                'customer': str(row.iloc[8]).strip() if len(row) > 8 and pd.notna(row.iloc[8]) else '-',
                'first_shipment': str(row.iloc[9]).strip() if len(row) > 9 and pd.notna(row.iloc[9]) else '-',
                'target': str(row.iloc[10]).split(' ')[0] if len(row) > 10 and pd.notna(row.iloc[10]) else '-', 
                'base': str(row.iloc[11]).split(' ')[0] if len(row) > 11 and pd.notna(row.iloc[11]) else '-',   
                'first_start': str(row.iloc[12]).split(' ')[0] if len(row) > 12 and pd.notna(row.iloc[12]) else '-',   
                'revised_start': str(row.iloc[13]).split(' ')[0] if len(row) > 13 and pd.notna(row.iloc[13]) else '-',     
                'nc': str(row.iloc[14]).split(' ')[0] if len(row) > 14 and pd.notna(row.iloc[14]) else '-',
                'status': str(row.iloc[15]).strip() if len(row) > 15 and pd.notna(row.iloc[15]) else '대기',
                'issue': str(row.iloc[16]).strip() if len(row) > 16 and pd.notna(row.iloc[16]) else '-'
            }
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

@app.route('/api/data')
def get_data():
    data = parse_production_data(EXCEL_FILE)
    return jsonify(data)

if __name__ == '__main__':
    # 5000번 포트에서 실행
    print("공정 관리 시스템 서버를 시작합니다...")
    # 포트를 5003로 변경하여 기존 프로세스와의 충돌 방지 및 브라우저 캐시 완벽 우회
    app.run(debug=True, port=5003, threaded=True)
