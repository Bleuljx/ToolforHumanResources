# -*- coding: utf-8 -*-
"""第二个测试表：花城蝶外（含脏数据：缺姓名、金额是文字、无法映射的项目）"""
import openpyxl
from openpyxl.comments import Comment
import sys, datetime

out = sys.argv[1]
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "新公式版本3.1"

ws["A1"] = "花城蝶外 2026年11月 算薪表"
headers = ["序号", "蝶城", "项目", "姓名", "员工编码", "供应商", "备注栏", "岗位", "入职日期", "离职日期",
           "基本薪资", "WJ租房补贴", "工资补发", "WY考核罚款", "WY其他扣款",
           "五险一金合计扣款（单位）", "增值税", "五险一金合计扣款（个人）",
           "WY当月应休月休天数", "WY应出勤天数", "WY实际出勤天数", "WY定额值班费",
           "WJ定额小时加班费", "WJ加班时长（小时）", "WY加班费", "WY备注", "XX特殊补贴"]
for i, h in enumerate(headers, start=1):
    ws.cell(3, i, h)

rows = [
    ("钱七", "00975551", "工程主管", 7000, 300, 0, -150, 0, 1600, 260, -700, 4, 26, 26, 0, 22, 3, 66, "", 400),
    (None, "00975552", "电工", 5500, 0, 0, 0, 0, 1200, 200, -550, 4, 26, 25, 0, 20, 0, 0, "", 0),  # 缺姓名但有金额
    ("孙八", "00975553", "保安", 4000, 0, "待定", 0, 0, 900, 150, -400, 4, 25, 25, 0, 16, 0, 0, "", 0),  # 金额是文字
]
r = 5
for row in rows:
    ws[f"B{r}"] = "花城蝶外"
    ws[f"C{r}"] = "花城花园"
    ws[f"D{r}"] = row[0]
    ws[f"E{r}"] = row[1]
    ws[f"F{r}"] = "成都XX人力资源有限公司"
    ws[f"H{r}"] = row[2]
    ws[f"K{r}"] = row[3]
    ws[f"L{r}"] = row[4]
    ws[f"M{r}"] = row[5]
    ws[f"N{r}"] = row[6]
    ws[f"O{r}"] = row[7]
    ws[f"P{r}"] = row[8]
    ws[f"Q{r}"] = row[9]
    ws[f"R{r}"] = row[10]
    ws[f"S{r}"] = row[11]
    ws[f"T{r}"] = row[12]
    ws[f"U{r}"] = row[13]
    ws[f"V{r}"] = row[14]
    ws[f"W{r}"] = row[15]
    ws[f"X{r}"] = row[16]
    ws[f"Y{r}"] = row[17]
    ws[f"Z{r}"] = row[18]
    ws[f"AA{r}"] = row[19]
    r += 1

ws["AA5"].comment = Comment("钱七:特殊补贴，标准未定", "钱七")
ws["N6"].comment = Comment("李四:代填", "李四")
wb.save(out)
print("saved", out)
