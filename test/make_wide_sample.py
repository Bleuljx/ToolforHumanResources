# -*- coding: utf-8 -*-
"""从 sample.xlsx 派生宽表测试样例：给李四的备注列填上与批注相同的文字（验证去重）"""
import openpyxl
import sys, os

src = os.path.join(os.path.dirname(__file__), "sample.xlsx")
out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "wide_sample.xlsx")

wb = openpyxl.load_workbook(src)
ws = wb["新公式版本3.1开始"]
# 找到李四所在行（D 列姓名）
row = None
for r in range(5, ws.max_row + 1):
    if ws.cell(r, 4).value == "李四":
        row = r
        break
assert row, "样例里没找到李四"
# 李四备注列（Z）= 与 N6 批注相同的文字 → 宽表备注应自动去重只保留一条
ws.cell(row, 26).value = "迟到2次，每次100"
wb.save(out)
print("saved", out)
