# -*- coding: utf-8 -*-
"""verify_wide.py — 校验宽表导出内容（表头/冻结/取数/备注汇总/去重）"""
import openpyxl, glob, sys, os

out = os.path.join(os.path.dirname(__file__), "out_wide")
files = glob.glob(os.path.join(out, "*非循环项导入(宽表).xlsx"))
assert files, "找不到宽表导出文件: " + os.listdir(out)
path = files[0]

failed = 0
def check(c, m):
    global failed
    print(("✓ " if c else "✗ ") + m)
    if not c: failed += 1

wb = openpyxl.load_workbook(path)
check(wb.sheetnames == ["非循环项导入模板"], "单表「非循环项导入模板」: " + str(wb.sheetnames))
ws = wb.active

HDR = ["员工编码", "员工姓名", "业务日期(yyyy-MM-dd)", "基本工资", "应休月休天数", "应出勤天数",
    "实际出勤天数", "定额值班费", "定额小时加班费", "加班时长（小时）", "租房补贴", "高温补贴",
    "交通住宿补助", "专项补贴", "垃圾点位专项补贴", "法定假补贴", "其他补款", "推荐奖励", "考核激励",
    "其他奖金", "全勤奖励", "工资补发", "考核罚款", "考勤扣款", "上月扣回", "其他扣款",
    "五险一金合计扣款（个人）", "五险一金合计扣款（单位）", "商业险", "服务费", "增值税", "备注"]
hdr = [ws.cell(1, c).value for c in range(1, 33)]
check(hdr == HDR, "32 列表头与模板一致")
check(ws.freeze_panes == "C2", f"冻结 C2（实际 {ws.freeze_panes}）")
check(ws.max_row == 5, f"表头 + 4 名员工（实际 max_row={ws.max_row}）")

rows = {}
for r in range(2, ws.max_row + 1):
    rows[ws.cell(r, 2).value] = [ws.cell(r, c).value for c in range(1, 33)]

# 取数：张三（基本工资@baseSalary=8000；应休月休=4；定额值班费=round(8000/26)=308；租房补贴=500；工资补发=300；考核罚款=-200；个人五险=-800；服务费 35）
zs = rows.get("张三")
check(zs is not None, "张三行存在")
check(zs[3] == 8000, f"张三 基本工资=8000（实际 {zs[3]}）")
check(zs[4] == 4, f"张三 应休月休天数=4（实际 {zs[4]}）")
check(zs[7] == 308, f"张三 定额值班费=308（自动计算 8000/26，实际 {zs[7]}）")
check(zs[10] == 500, f"张三 租房补贴=500（实际 {zs[10]}）")
check(zs[21] == 300, f"张三 工资补发=300（实际 {zs[21]}）")
check(zs[22] == -200, f"张三 考核罚款=-200 源值原样（实际 {zs[22]}）")
check(zs[26] == -800, f"张三 个人五险一金=-800（实际 {zs[26]}）")
check(zs[28] in (None, ""), f"张三 商业险未匹配留空（实际 {zs[28]!r}）")
check(zs[29] == 35, f"张三 服务费=35（区间外列也可映射，实际 {zs[29]}）")
check(zs[2] == "2026-11-16" or str(zs[2]).startswith("2026-11-16"), f"张三 业务日期=2026-11-16（实际 {zs[2]!r}）")
# 备注：批注内容本身含数字（金额已写在批注里）→ 不再拼单元格金额；源表备注列在后
expect_zs = "3月调薪差额，补发200；含补缴医保80；3月调薪补差"
check(zs[31] == expect_zs, f"张三 备注=批注段+备注列（实际 {zs[31]!r}）")

# 李四：批注金额为0不拼数字；备注列与批注相同 → 去重只保留一条
ls = rows.get("李四")
check(ls is not None, "李四行存在")
check(ls[31] == "迟到2次，每次100", f"李四 备注=去重后一条（实际 {ls[31]!r}）")
check(ls[7] == 162, f"李四 定额值班费=round(4200/26)=162（实际 {ls[7]}）")

# 王五：租房补贴批注+金额
ww = rows.get("王五")
check(ww is not None, "王五行存在")
check(ww[31] == "员工本人租房补贴300", f"王五 备注=批注+金额（实际 {ww[31]!r}）")
check(ww[7] == 177, f"王五 定额值班费=round(4600/26)=177（实际 {ww[7]}）")

# 赵六：无批注 → 备注=源表备注列
zl = rows.get("赵六")
check(zl is not None, "赵六行存在")
check(zl[31] == "含餐补", f"赵六 备注=源表备注列（实际 {zl[31]!r}）")
check(zl[22] == -100 and zl[25] == -80, f"赵六 考核罚款=-100 其他扣款=-80（实际 {zl[22]}/{zl[25]}）")

# 表头样式：Calibri 加粗，A/C 列红字
c1 = ws.cell(1, 1)
check(c1.font.bold and c1.font.name == "Calibri", f"表头 Calibri 加粗（实际 {c1.font.name} bold={c1.font.bold}）")
check(str(c1.font.color.rgb) == "FFFF0000", f"员工编码表头红字（实际 {c1.font.color.rgb}）")
check(ws.cell(1, 4).font.bold and str(ws.cell(1, 4).font.color.rgb) != "FFFF0000", "基本工资表头加粗黑字")

print("\n" + ("ALL PASS" if not failed else f"{failed} 项失败"))
sys.exit(1 if failed else 0)
