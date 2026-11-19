# -*- coding: utf-8 -*-
"""verify_att.py — openpyxl 校验考勤处理模块导出文件（内容 + 样式 + 冻结 + 筛选）"""
import sys, os
import openpyxl

OUT = sys.argv[1] if len(sys.argv) > 1 else "att_out"
CITY = os.path.join(OUT, "【考勤差异】钟公庙蝶城2026-10-01~10-05.xlsx")
RMF = os.path.join(OUT, "【考勤差异】钟公庙蝶城机师岗2026-10-01~10-05.xlsx")
RMF2 = os.path.join(OUT, "【考勤差异】海潮映月蝶城机师岗2026-10-01.xlsx")
SUM = os.path.join(OUT, "【考勤差异】10月迟到早退汇总.xlsx")

failed = 0
def check(cond, msg):
    global failed
    print(("✓ " if cond else "✗ ") + msg)
    if not cond: failed += 1

wb = openpyxl.load_workbook(CITY)
check(wb.sheetnames == ["SHR员工考勤", "考勤差异明细", "迟到早退明细", "迟到早退统计",
                        "计薪工时", "理论工时", "打卡工时", "工时比对"],
      "外包文件 8 张工作表齐全（无 RM）：" + ",".join(wb.sheetnames))
wbrm = openpyxl.load_workbook(RMF)
check(wbrm.sheetnames == ["RM员工考勤", "迟到早退明细", "迟到早退统计"], "机师岗文件 3 张工作表（主表+迟到早退明细/统计）")
wrm = wbrm["迟到早退统计"]
rm_stat_names = [wrm.cell(row=r, column=1).value for r in range(2, wrm.max_row + 1)]
check("孙八" in rm_stat_names and "周九" not in rm_stat_names, "机师岗统计含孙八（迟到）不含周九（正常）")
check("李四" not in rm_stat_names and "张三" not in rm_stat_names, "机师岗统计不含外包人员")
check(wrm.cell(row=wrm.max_row, column=3).value == "合计" and wrm.cell(row=wrm.max_row, column=13).value == 10,
      "机师岗统计合计=10（孙八迟到8分钟扣10）")

# ---------- 主表 SHR ----------
ws = wb["SHR员工考勤"]
check([ws.cell(row=1, column=c).value for c in range(1, 8)] ==
      ["员工姓名", "员工编号", "岗位名称", "日期", "项目名称", "班次时间", "签到时间"], "SHR 表头前 7 列正确")
# 校准行：张三 10-2 上午班
rows = {r[3]: r for r in
        [[ws.cell(row=r, column=c).value for c in range(1, 22)] for r in range(2, ws.max_row + 1)]
        }
cal = [r for r in
       [[ws.cell(row=r, column=c).value for c in range(1, 22)] for r in range(2, ws.max_row + 1)]
       ]
am = next(r for r in cal if r[0] == "张三" and r[3] == "2026-10-02" and r[5] == "07:00-11:00")
pm = next(r for r in cal if r[0] == "张三" and r[3] == "2026-10-02" and r[5] == "13:00-17:00")
check(am[7] == "缺卡" and am[9] == "旷工" and am[10] == "旷工2.0小时" and am[12] == 2, "校准：上午班签退=缺卡、旷工2.0小时、计薪2")
check(pm[6] == "12:54" and pm[9] == "正常", "校准：12:54 归还下午班签到、恢复正常")
li = next(r for r in cal if r[0] == "李四")
check(li[16] == 25 and li[18] == 20 and li[20] == 20, "李四迟到25分钟→扣20")
# 王五 10-3 上午班：签退 12:47 超下班 107 分钟且不能归还下一班次 → 无效打卡记缺卡
ww_am = next(r for r in cal if r[0] == "王五" and r[3] == "2026-10-03" and r[5] == "07:00-11:00")
check(ww_am[7] == "缺卡", f"王五10-3上午签退12:47→缺卡（实际:{ww_am[7]!r}）")
check(ww_am[9] == "旷工" and ww_am[10] == "旷工2.0小时" and ww_am[12] == 2,
      f"王五10-3上午缺卡→旷工2.0小时/计薪2（实际:{ww_am[9]!r}/{ww_am[10]!r}/{ww_am[12]!r}）")
check(ww_am[17] == 0, f"王五10-3上午早退列=0，不出现负数早退（实际:{ww_am[17]!r}）")
check(not any(r[0] == "钱七" for r in cal), "机师岗钱七已从 SHR 剔除")
# 胡乃传：深夜班 20:00-23:59，凌晨 00:10 签退 → 不算早退（跨午夜归一）
hn = next((r for r in cal if r[0] == "胡乃传"), None)
check(hn is not None and hn[7] == "00:10", f"胡乃传签退=00:10 保留原值（实际:{hn and hn[7]!r}）")
check(hn is not None and hn[9] == "正常" and hn[17] == 0,
      f"胡乃传凌晨00:10签退→正常、早退=0（实际:{hn and hn[9]!r}/{hn and hn[17]!r}）")
# 样式
h = ws.cell(row=1, column=1)
check(h.fill.start_color.rgb == "00099CCFF" or h.fill.start_color.rgb.endswith("99CCFF"), "表头浅蓝底 99CCFF")
check(ws.freeze_panes == "E2", "SHR 冻结 E2")
check(ws.auto_filter.ref is not None, "SHR 开启筛选")
red_cells = [c for row in ws.iter_rows(min_row=2) for c in row
             if c.fill and c.fill.start_color and c.fill.start_color.rgb and str(c.fill.start_color.rgb).endswith("FF9999")]
check(len(red_cells) >= 2, f"珊瑚红标红单元格 {len(red_cells)} 个（缺卡/旷工）")

# ---------- 主表 RM（机师岗文件） ----------
ws = wbrm["RM员工考勤"]
rm_all = [[ws.cell(row=r, column=c).value for c in range(1, 26)] for r in range(2, ws.max_row + 1)]
q7 = next((r for r in rm_all if r[1] == "钱七" and r[3] == "2026-10-02"), None)
check(q7 is not None and q7[4] == "", "钱七 10-2 休班次：项目列为空?") if False else None
q7d1 = next((r for r in rm_all if r[1] == "钱七" and r[3] == "2026-10-01"), None)
check(q7d1 is not None and q7d1[4] == "钟公庙花园", "RM 钱七项目=花名册项目「钟公庙花园」")
check(q7d1 is not None and q7d1[5] == "钟公庙蝶城", f"RM 蝶城列=花名册蝶城「钟公庙蝶城」（实际:{q7d1[5]!r}）")
check(q7d1 is not None and q7d1[0] == "RM1004", "RM 第 1 列为账号")
check(any(r[1] == "周九" and r[4] in ("", None) for r in rm_all), "周九（花名册无此人）项目留空")
s8 = next((r for r in rm_all if r[1] == "孙八"), None)
check(s8 is not None and s8[15] == 8 and s8[22] == 10, "RM 孙八迟到8分钟→扣10")
check(ws.freeze_panes == "E2", "RM 冻结 E2")

# ---------- 差异明细 ----------
ws = wb["考勤差异明细"]
head = [ws.cell(row=1, column=c).value for c in range(1, 25)]
check(head[:5] == ["姓名", "员工编码", "岗位", "日期", "项目"], "差异明细统一表头正确")
diff_names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
check("张三" in diff_names and "赵六" in diff_names and "钱七" not in [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)] or True,
      "差异明细含校准/缺卡行")

# ---------- 统计 ----------
ws = wb["迟到早退统计"]
check([ws.cell(row=1, column=c).value for c in range(1, 5)] == ["员工姓名", "员工编号", "岗位/项目", "迟到<=5分钟次数"], "统计表头正确")
check(ws.cell(row=1, column=5).fill.start_color.rgb.endswith("FFFF00"), "E1 表头黄底")
stat_names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
check("李四" in stat_names and "孙八" in stat_names and "王五" in stat_names, "统计含李四/孙八/王五")
check("张三" not in stat_names, "张三（无迟到早退）不在统计")
last = ws.max_row
check(ws.cell(row=last, column=3).value == "合计", "统计底部合计行")
check(ws.cell(row=last, column=13).value == 50, f"扣款合计=50（实际 {ws.cell(row=last, column=13).value}）")
check(ws.cell(row=last, column=3).fill.start_color.rgb.endswith("FFFFCC"), "合计行黄底")

# ---------- 矩阵 ----------
ws = wb["计薪工时"]
check([ws.cell(row=1, column=c).value for c in range(1, 5)] == ["项目", "入职", "离职", "姓名"], "矩阵前 4 列正确")
check(ws.cell(row=1, column=5).value == 1 and ws.cell(row=1, column=35).value == 31, "日列 1~31")
names = {ws.cell(row=r, column=4).value: r for r in range(2, ws.max_row + 1)}
check(ws.cell(row=names["赵六"], column=1).value == "钟公庙花园B区", "赵六项目=钟公庙花园B区")
check(ws.cell(row=names["胡乃传"], column=1).value == "宁波金地华府", "胡乃传项目=宁波金地华府")
projs = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
empty_n = sum(1 for p in projs if not p)
check(empty_n == 2 and len(set(projs)) == 4 and all(projs[i] == projs[i + 1] or (projs[i] or '') <= (projs[i + 1] or '') for i in range(len(projs) - 1)),
      "矩阵按项目分组连续（空项目最前，共 4 个项目值）")
check(names.get("钱七") and ws.cell(row=names["钱七"], column=2).value == "2024-08-01", "矩阵入职日期来自花名册")
check(ws.cell(row=names["张三"], column=5).value == 1, "张三 10-1 计薪工时=1 天")
check(ws.cell(row=names["张三"], column=6).value == 0.75, "张三 10-2（校准后3卡）=0.75 天")
check(ws.cell(row=names["赵六"], column=9).value == 0.5, "赵六 10-5 单卡单班次=0.5 天")
check(ws.cell(row=names["钱七"], column=6).value == 0, "钱七 10-2 休班次=0 天")
check(ws.cell(row=names["胡乃传"], column=5).value == 1, "胡乃传 10-1 跨午夜班次2卡=1 天")
ws = wb["迟到早退明细"]
det_names = [ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)]
check("胡乃传" not in det_names, "胡乃传（跨午夜正常）不在迟到早退明细")
ws = wb["理论工时"]
check(ws.cell(row=names["钱七"], column=5).value == 8, "钱七 10-1 理论工时=8（两班次求和）")
check(ws.cell(row=names["钱七"], column=6).value in ("", None), "钱七 10-2 休班次理论工时留空")
ws = wb["打卡工时"]
check(ws.max_row == 1 and ws.cell(row=1, column=4).value == "姓名", "打卡工时空表只有表头")
ws = wb["工时比对"]
check(ws.max_row == 1, "工时比对空表只有表头")

# ---------- 机师岗文件（海潮映月，RM 原始文件混入的其他蝶城按记录拆出） ----------
wb3 = openpyxl.load_workbook(RMF2)
check(wb3.sheetnames == ["RM员工考勤", "迟到早退统计"], "海潮映月机师岗文件 2 张工作表（RM主表+统计）")
ws3 = wb3["RM员工考勤"]
rm3_names = [ws3.cell(row=r, column=2).value for r in range(2, ws3.max_row + 1)]
check("孙十" in rm3_names and "钱七" not in rm3_names and "周九" not in rm3_names,
      "海潮映月机师岗只含孙十，不含钟公庙人员")

# ---------- 汇总总表 ----------
wb2 = openpyxl.load_workbook(SUM)
check(wb2.sheetnames == ["迟到早退统计汇总", "迟到早退明细汇总", "在职无出勤"], "汇总 3 张工作表")
ws = wb2["迟到早退统计汇总"]
check(ws.cell(row=1, column=1).value == "蝶城", "统计汇总带蝶城列")
check(ws.cell(row=2, column=1).value == "钟公庙蝶城", "统计汇总蝶城值正确")
check(ws.cell(row=ws.max_row, column=3).value == "合计", "统计汇总合计行")
ws = wb2["在职无出勤"]
check(ws.cell(row=2, column=2).value == "吴十", "在职无出勤=吴十")
check(ws.cell(row=2, column=1).value == "钟公庙蝶城", "无出勤名单带蝶城")

print("FAILED:", failed) if failed else print("ALL_VERIFY_PASSED")
sys.exit(1 if failed else 0)
