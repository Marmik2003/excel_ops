import json
from pathlib import Path

from src.excel import GetExcel

BASE_DIR = Path(__file__).resolve().parent


def load_data():
    with open(BASE_DIR / 'raw_data/sample_data.json') as f:
        data = json.load(f)
        return data


excel = GetExcel(load_data())
wb = excel.create()
wb.save(BASE_DIR / 'generated_excel/sample01.xlsx')
