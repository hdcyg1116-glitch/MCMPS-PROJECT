import sys
import os

# app.py가 있는 경로를 sys.path에 추가
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from app import parse_production_data
    file_path = "TC 대조립 일일 진도현황(260304).xlsx"
    print(f"Testing parse_production_data on {file_path}")
    data = parse_production_data(file_path, sheet_name=0)
    print(f"Parsed {len(data)} rows.")
    if data:
        print("First row parsed:")
        print(data[0])
except Exception as e:
    print(f"Error: {e}")
