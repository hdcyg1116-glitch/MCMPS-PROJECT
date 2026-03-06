import time
import pandas as pd
from app import parse_production_data

start = time.time()
file_path = "TC 대조립 일일 진도현황(260304).xlsx"
print(f"Testing {file_path}")
data = parse_production_data(file_path)
end = time.time()
print(f"Time taken: {end - start:.2f} seconds")
print(f"Rows parsed: {len(data)}")
