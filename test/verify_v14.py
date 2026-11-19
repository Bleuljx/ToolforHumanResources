# -*- coding: utf-8 -*-
"""verify_v14.py — 校验 v1.4 导出文件：
1) 三张表所有单元格（含空单元格）都有细框线
2) 一键表/导入模板备注 = 批注优先 + 兜底区间源表备注
3) 考勤报送备注 = 源表备注列
4) 问题清单：区间外提示只出现白列（服务费/其他补款/XX特殊补贴），不出现 工时/应发合计/成本合计
用法：python verify_v14.py <out_dir>
"""
import sys, os
import openpyxl

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(__file__), "out")
failed = []

def check(c, m):
    print(("✓ " if c else "✗ ") + m)
    if not c:
        failed.append(m)

def grid_bordered(ws, max_r, max_c, label):
    bad = []
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            b = ws.cell(row=r, column=c).border
            for side in ("left", "right", "top", "bottom"):
                if getattr(b, side).style != "thin":
                    bad.append(f"{label}!{ws.cell(row=r, column=c).coordinate}({side})")
    check(not bad, f"{label}：全表 {max_r}×{max_c} 细框线" + ("，缺框线 " + ", ".join(bad[:8]) if bad else ""))

def find_row(ws, name_col, name, key_col, key):
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=name_col).value).strip() == name and \
           str(ws.cell(row=r, column=key_col).value).strip() == key:
            return r
    return None

# ---- 考勤报送 ----
p = os.path.join(out, "麓湖蝶城考勤报送.xlsx")
wb = openpyxl.load_workbook(p)
ws = wb["考勤报送"]
check(ws.cell(row=1, column=17).value == "WY备注", "考勤报送表头第17列=WY备注")
check(ws.cell(row=2, column=4).value == "张三", "考勤报送第1行是张三")
check(ws.cell(row=2, column=17).value == "3月调薪补差", f"考勤报送张三WY备注=源表备注（实际:{ws.cell(row=2, column=17).value!r}）")
grid_bordered(ws, ws.max_row, 17, "考勤报送")

# ---- 工资类目一键表 ----
p = os.path.join(out, "麓湖蝶城工资类目一键表.xlsx")
wb = openpyxl.load_workbook(p)
ws = wb["工资类目一键表"]
rm = lambda name, key: ws.cell(row=find_row(ws, 4, name, 8, key), column=10).value if find_row(ws, 4, name, 8, key) else None
check(rm("张三", "工资补发") == "3月调薪差额，补发200", f"一键表 张三/工资补发=批注（实际:{rm('张三', '工资补发')!r}）")
check(rm("赵六", "WY其他扣款") == "含餐补", f"一键表 赵六/WY其他扣款=兜底源表备注（实际:{rm('赵六', 'WY其他扣款')!r}）")
check(rm("李四", "WY其他扣款") in ("", None), f"一键表 李四/WY其他扣款=空（实际:{rm('李四', 'WY其他扣款')!r}）")
check(rm("王五", "WJ租房补贴") == "员工本人租房补贴", f"一键表 王五/租房补贴=批注（实际:{rm('王五', 'WJ租房补贴')!r}）")
grid_bordered(ws, ws.max_row, 11, "一键表")

# 扣款不取反：源表原值（正负号不改）
def amt(name, key):
    rr = find_row(ws, 4, name, 8, key)
    return ws.cell(row=rr, column=9).value if rr else None
check(amt("张三", "WY考核罚款") == -200, f"一键表 张三/WY考核罚款=源表原值 -200（实际:{amt('张三','WY考核罚款')!r}）")
check(amt("李四", "WY其他扣款") == -50, f"一键表 李四/WY其他扣款=源表原值 -50（实际:{amt('李四','WY其他扣款')!r}）")
check(amt("赵六", "WY其他扣款") == -80, f"一键表 赵六/WY其他扣款=源表原值 -80（实际:{amt('赵六','WY其他扣款')!r}）")
check(amt("张三", "五险一金合计扣款（个人）") == -800, f"一键表 张三/个人五险=源表原值 -800（实际:{amt('张三','五险一金合计扣款（个人）')!r}）")

# ---- 非循环项导入 ----
p = os.path.join(out, "麓湖蝶城非循环项导入.xlsx")
wb = openpyxl.load_workbook(p)
ws = wb["非循环项导入模板"]
r = find_row(ws, 2, "张三", 5, "工资补发")
if r:
    check(ws.cell(row=r, column=7).value == "3月调薪差额，补发200",
          f"导入模板 张三/工资补发 备注=批注（实际:{ws.cell(row=r, column=7).value!r}）")
else:
    check(False, "导入模板缺 张三/工资补发 行")
check(ws.sheet_state == "visible" and "_opt0" in wb.sheetnames, "导入模板含隐藏 _opt0")
grid_bordered(ws, ws.max_row, 7, "导入模板")

# ---- 考勤字段进导入模板（非 0 才导） ----
def imp_val(name, item):
    rr = find_row(ws, 2, name, 5, item)
    return ws.cell(row=rr, column=6).value if rr else None

check(imp_val("张三", "基本工资") == 8000, f"导入 张三/基本工资=8000（实际:{imp_val('张三','基本工资')!r}）")
check(imp_val("张三", "应休月休天数") == 4, f"导入 张三/应休月休天数=4（实际:{imp_val('张三','应休月休天数')!r}）")
check(imp_val("张三", "应出勤天数") == 26, f"导入 张三/应出勤天数=26（实际:{imp_val('张三','应出勤天数')!r}）")
check(imp_val("张三", "实际出勤天数") == 28, f"导入 张三/实际出勤天数=28（实际:{imp_val('张三','实际出勤天数')!r}）")
check(imp_val("张三", "定额值班费") == 308, f"导入 张三/定额值班费=308（自动计算 8000/26，实际:{imp_val('张三','定额值班费')!r}）")
check(imp_val("张三", "定额小时加班费") == 25, f"导入 张三/定额小时加班费=25（实际:{imp_val('张三','定额小时加班费')!r}）")
check(imp_val("张三", "加班时长（小时）") == 6, f"导入 张三/加班时长=6（实际:{imp_val('张三','加班时长（小时）')!r}）")
check(imp_val("李四", "加班时长（小时）") is None, "导入 李四/加班时长=0 未导出（非 0 才导）")

# ---- 问题清单 ----
p = os.path.join(out, "问题清单.xlsx")
wb = openpyxl.load_workbook(p)
ws = wb["问题清单"]
details = [str(ws.cell(row=r, column=6).value) for r in range(2, ws.max_row + 1)]
oob = [d for d in details if "区间外" in d or ("区间" in d and "没有导出" in d)]
oob_all = details  # 类型列在5，说明在6；区间外类型判断用第5列
types = [str(ws.cell(row=r, column=5).value) for r in range(2, ws.max_row + 1)]
oob_details = [details[i] for i, t in enumerate(types) if t == "区间外项目"]
joined = " || ".join(oob_details)
check("服务费" in joined, "区间外提示包含 服务费（白列应提示）")
check("其他补款" in joined, "区间外提示包含 其他补款（白列应提示）")
check("工时" not in joined, "区间外提示不含 工时（忽略列）")
check("应发合计" not in joined, "区间外提示不含 应发合计（黄列自动忽略）")
check("成本合计" not in joined, "区间外提示不含 成本合计（黄列自动忽略）")
check("没有导出到一键表" in joined, "提示文案已改为易懂版本")
check(not any("考勤字段" in d for d in details), "7 个考勤字段全部映射成功，无未映射提示")
check("「1」" not in joined and "「31」" not in joined, "考勤明细区域（1~31 日）不触发区间外提示")

print("\n" + ("有 %d 项失败" % len(failed) if failed else "校验全部通过"))
sys.exit(1 if failed else 0)
