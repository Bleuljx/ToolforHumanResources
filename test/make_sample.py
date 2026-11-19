# -*- coding: utf-8 -*-
"""生成测试用算薪表（模拟真实结构：表头第3行、数据第5行起、带批注）"""
import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill, Alignment
import sys, datetime

out = sys.argv[1] if len(sys.argv) > 1 else "sample.xlsx"

wb = openpyxl.Workbook()
ws0 = wb.active
ws0.title = "工资汇总"
ws0["A1"] = "这是汇总表，应被跳过"

ws = wb.create_sheet("新公式版本3.1开始")
ws["A1"] = "麓湖蝶城 2026年11月 算薪表（新公式版本3.1）"
ws["A2"] = "制表：张三    制表日期：2026-11-15"

headers = {
    "A": "序号", "B": "蝶城", "C": "项目", "D": "姓名", "E": "员工编码",
    "F": "供应商", "G": "备注栏", "H": "岗位", "I": "入职日期", "J": "离职日期",
    "K": "基本薪资",
    "L": "WJ租房补贴", "M": "工资补发", "N": "WY考核罚款", "O": "WY其他扣款",
    "P": "五险一金合计扣款（单位）", "Q": "增值税", "R": "五险一金合计扣款（个人）",
    "S": "WY当月应休月休天数", "T": "WY应出勤天数", "U": "WY实际出勤天数",
    "V": "WY定额值班费", "W": "WJ定额小时加班费", "X": "WJ加班时长（小时）",
    "Y": "WY加班费", "Z": "WY备注",
    "AA": "其他补款", "AB": "工时", "AC": "WY应发合计", "AD": "服务费", "AE": "WJ成本合计",
}
# 1~31 日考勤明细列（AF..BJ 共 31 列）：应被自动识别为考勤明细区域
DAY_COLS = ["AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS",
            "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG",
            "BH", "BI", "BJ"]
for i, col in enumerate(DAY_COLS, start=1):
    headers[col] = str(i)
YELLOW_COLS = ["AC", "AE"]   # 表头标黄的列：自动忽略，不提示区间外项目
for col, name in headers.items():
    c = ws[f"{col}3"]
    c.value = name
    c.font = Font(name="宋体", size=12, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="FFF2CC" if col in YELLOW_COLS else "0070C0")
    c.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

rows = [
    # 姓名, 编码, 岗位, 基本薪资, 租房, 补发, 考核罚款, 其他扣款, 单位五险, 增值税, 个人五险, 应休, 应出勤, 实际出勤, 值班费, 时薪, 加班时长, 加班费, 备注, 其他补款, 工时, 应发合计, 服务费, 成本合计
    ("张三", "00975540", "秩序主管", 8000, 500, 300, -200, 0, 1800, 300, -800, 4, 26, 28, 0, 25, 6, 150, "3月调薪补差", 0, 176, 6000, 35, 6322),
    ("李四", "00975541", "保洁员", 4200, 0, 300, 0, -50, 950, 160, -420, 4, 26, 24, 200, 18, 0, 0, "", 0, 168, 5000, 28, 5100),
    ("王五", "00975542", "绿化员", 4600, 300, 0, 0, 0, 1000, 180, -460, 4, 26, 27, 0, 20, 2, 40, "", 0, 172, 5300, 30, 5500),
    (None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None),  # 空行应跳过
    ("赵六", "00975543", "客服", 5200, 0, 0, -100, -80, 1200, 200, -520, 4, 25, 25, 0, 22, 0, 0, "含餐补", 60, 160, 6100, 26, 6400),
]

r = 5
for row in rows:
    if row[0] is None:
        r += 1
        continue
    (name, code, post, base, rent, backpay, fine, other_ded, unit_si, vat, per_si,
     rest_days, should, actual, duty, hourly, ot_hours, ot_fee, remark,
     extra_pay, work_hours, gross_sum, service_fee, cost_sum) = row
    ws[f"B{r}"] = "麓湖蝶城"
    ws[f"C{r}"] = "麓湖生态城A区"
    ws[f"D{r}"] = name
    ws[f"E{r}"] = code
    ws[f"F{r}"] = "成都XX人力资源有限公司"
    ws[f"H{r}"] = post
    ws[f"I{r}"] = datetime.date(2024, 3, 1)
    ws[f"K{r}"] = base
    ws[f"L{r}"] = rent
    ws[f"M{r}"] = backpay
    ws[f"N{r}"] = fine
    ws[f"O{r}"] = other_ded
    ws[f"P{r}"] = unit_si
    ws[f"Q{r}"] = vat
    ws[f"R{r}"] = per_si
    ws[f"S{r}"] = rest_days
    ws[f"T{r}"] = should
    ws[f"U{r}"] = actual
    ws[f"V{r}"] = duty
    ws[f"W{r}"] = hourly
    ws[f"X{r}"] = ot_hours
    ws[f"Y{r}"] = ot_fee
    ws[f"Z{r}"] = remark
    ws[f"AA{r}"] = extra_pay
    ws[f"AB{r}"] = work_hours
    ws[f"AC{r}"] = gross_sum
    ws[f"AD{r}"] = service_fee
    ws[f"AE{r}"] = cost_sum
    # 考勤明细：张三 1~3 日各 8 小时，赵六 5 日 8 小时
    if name == "张三":
        for i in range(3):
            ws[f"{DAY_COLS[i]}{r}"] = 8
    if name == "赵六":
        ws[f"{DAY_COLS[4]}{r}"] = 8
    r += 1

# 批注（模拟算薪人员在各工资项目上写的说明）
ws["M5"].comment = Comment("B10WHMAG05.工靓: 3月调薪差额，补发200", "")  # 作者名长且含数字，前缀必须去掉
ws["N6"].comment = Comment("李四:迟到2次，每次100", "李四")
ws["L7"].comment = Comment("王五:员工本人租房补贴", "王五")
ws["R5"].comment = Comment("张三:含补缴医保80", "张三")

for col, w in {"A": 6, "B": 12, "C": 20, "D": 10, "E": 12, "F": 24, "H": 12,
               "K": 10, "L": 12, "M": 12, "N": 12, "O": 12, "P": 22, "Q": 10,
               "R": 22, "S": 18, "T": 14, "U": 14, "V": 14, "W": 16, "X": 16,
               "Y": 12, "Z": 24, "AA": 12, "AB": 10, "AC": 12, "AD": 10, "AE": 12}.items():
    ws.column_dimensions[col].width = w

wb.save(out)
print("saved", out)
