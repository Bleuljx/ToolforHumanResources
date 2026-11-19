# make_ren_sample.py — 造文件重命名模块的样例（花名册 + 3 个模块 zip）
# 覆盖：正常匹配 / 编码匹配 / 同名同姓按项目消歧 / 同名按调度中心消歧 /
#       未匹配(周九) / 目标重名冲突(保单2个张三) / 缺失检查(王五无身份证正面)
import os
from openpyxl import Workbook

OUT = os.path.join(os.path.dirname(__file__))
MOD = '身份证正面'

# 1. 花名册
wb = Workbook()
ws = wb.active
ws.title = '花名册'
ws.append(['调度中心', '员工项目', '员工编码', '员工姓名'])
ws.append(['甬江蝶城', '项目甲', 'R001', '张三'])
ws.append(['甬江蝶城', '项目甲', 'R002', '李四'])
ws.append(['海潮蝶城', '项目乙', 'R003', '张三'])   # 同名同姓
ws.append(['甬江蝶城', '项目甲', 'R004', '王五'])
wb.save(os.path.join(OUT, 'ren_roster.xlsx'))

import zipfile

def make_zip(name, items):
    with zipfile.ZipFile(os.path.join(OUT, name), 'w', zipfile.ZIP_STORED) as z:
        for arc, data in items:
            z.writestr(arc, data)

# 2. 身份证正面.zip：撞名(张三.jpg 无项目段→失败消歧) + 正常 + 按项目消歧 + 匹配不上
make_zip('ren_身份证正面.zip', [
    ('张三.jpg', b'zhang-A-id-front'),          # 两名张三且无消歧段 → 失败
    ('李四.jpg', b'li-si-id-front'),            # 唯一匹配 → R002
    ('张三-项目乙.jpg', b'zhang-B-id-front'),   # 撞名 → 按项目段消歧 → R003
    ('周九.jpg', b'zhou-jiu-id-front'),         # 花名册无此人 → 失败
])

# 3. 劳务合同协议.zip：按调度中心消歧 + 按项目消歧
make_zip('ren_劳务合同协议.zip', [
    ('甬江蝶城-张三-劳务合同协议.pdf', b'xieyi-A'),       # 撞名 → 调度中心段 → R001
    ('张三-项目乙-劳务合同协议.pdf', b'xieyi-B'),          # 撞名 → 项目段 → R003
])

# 4. 投保证明.zip：2 个文件 → 特殊模块阻断（只允许 1 个）
make_zip('ren_投保证明.zip', [
    ('R001-投保证明首页.jpg', b'baodan-1'),
    ('副本/R001-投保证明首页.jpg', b'baodan-2'),
])

# 5. 投保证明1.zip：1 个原始文件 → 按花名册全员复制重命名
make_zip('ren_投保证明1.zip', [
    ('投保证明原件.jpg', b'baodan-src'),
])

print('ren samples written to', OUT)
