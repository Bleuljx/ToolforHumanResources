# 人力工具集成平台（salary-tool）

离线单文件网页工具，面向人力/薪酬场景，把多个 Excel VBA 宏工具改造为纯前端实现。**双击 `dist/人力工具集成平台.html` 即可使用，无需安装任何环境。**

## 功能模块

| 模块 | 说明 |
| --- | --- |
| 超7天未出勤 | 考勤统计结果 × 花名册 × 上期反馈，筛出无出勤天数 ≥ N 的员工，导出异常明细（支持反馈回导） |
| 考勤处理 | 外包 SHR / 机师岗 RM 原始导出 + 花名册，打卡校准、迟到早退扣款重算，按蝶城输出考勤差异表 + 迟到早退汇总 |
| 薪资报送 | 算薪表 → 考勤报送 + 工资类目一键表 + 非循环项导入模板（含 662 项薪酬项目 E 列自动匹配）+ 宽表模式 |
| 文件重命名 | 按花名册批量重命名档案资料（9 个模块），自定义分段改名规则（工号-类型 固定尾段），自动按模块拆包 zip |

支持多文件多蝶城批处理、问题清单（xlsx 下载）、配置导出/导入（覆盖全部模块）、localStorage 记忆。

## 目录结构

```
salary-tool/
├── index.html          # GitHub Pages 直达入口（= dist 工具的副本，根目录必须有它，链接打开即是工具）
├── build.js            # 构建脚本：内联所有 JS/数据 → 单文件 HTML
├── src/
│   ├── index.html      # 界面骨架 + 样式
│   ├── core.js         # 薪资报送业务逻辑
│   ├── att-core.js     # 考勤处理业务逻辑
│   ├── idle-core.js    # 超7天未出勤业务逻辑
│   ├── ren-core.js     # 文件重命名业务逻辑
│   ├── *-app.js        # 各模块界面层
│   ├── xlsxw.js        # 自研 xlsx 写入器（样式/冻结/筛选/批注）
│   ├── zip.js          # 自研 zip 打包/解包（CompressionStream）
│   └── options.json    # 非循环项 662 项薪酬项目字典
├── vendor/xlsx.js      # SheetJS（读取 Excel 用）
├── dist/               # 构建产物（双击即用的单文件）
└── test/               # jsdom 端到端 + openpyxl 校验 + 假数据样本
```

## 构建与测试

```bash
# 构建单文件（需要 Node.js ≥ 18）
node build.js

# 测试：先生成假数据样本（需要 Python + openpyxl）
cd test
python make_sample.py && python make_att_sample.py && python make_idle_sample.py && python make_ren_sample.py && python make_wide_sample.py

# jsdom 端到端（需要 NODE_PATH 指向装了 jsdom 的 node_modules）
node ui.js && node ui_att.js && node ui_idle.js && node ui_ren.js && node ui_summary.js && node ui_wide.js

# openpyxl 导出校验
python verify_v14.py && python verify_att.py && python verify_wide.py
```

> test 目录中的 xlsx/zip 均为脚本生成的假数据（张三/李四/WY1001），不含任何真实员工信息。

## 使用

- **本地**：直接双击 `dist/人力工具集成平台.html`（推荐 Chrome / Edge），完全离线运行，数据不出本机。
- **在线**：本仓库已开启 GitHub Pages，访问 `https://<你的用户名>.github.io/ToolforHumanResources/` 即可直接使用（根目录 `index.html` 是直达入口）。工具本身仍在浏览器本地处理数据，不上传服务器。

首次使用请点左下角 **📖 用户手册**（点击会在新标签页打开独立教程页），内含四个模块的逐步操作指南和常见问题解答。

> 工具更新后：重新运行 `node build.js`，把生成的 `dist/人力工具集成平台.html` 同步覆盖到根目录 `index.html` 再提交即可。
