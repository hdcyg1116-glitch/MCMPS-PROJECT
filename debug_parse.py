import pandas as pd
import os

file_path = '생산2팀_생산진도표(TC)_260304_매출.xlsx'

def test_parse():
    try:
        df = pd.read_excel(file_path, sheet_name=0, header=1)
        print(f"Total rows in DF: {len(df)}")
        print(f"Columns: {df.columns.tolist()}")
        
        data = []
        for i, row in df.iterrows():
            serial = str(row.iloc[6]).strip()
            if not serial or serial.lower() == 'nan' or serial == 'None':
                continue
            
            item = {
                'no': str(row.iloc[1]),
                'serial': serial,
                'model': str(row.iloc[5])
            }
            data.append(item)
            if len(data) <= 5:
                print(f"Parsed item {len(data)}: {item}")
        
        print(f"Total items parsed: {len(data)}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == '__main__':
    test_parse()
