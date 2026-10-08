---
name: software-copyright-dbdb
description: 生成中国计算机软件著作权（软著）登记所需的两份 Word 鉴别材料——操作手册（软件说明书）docx 和源代码页 code.docx，人工审核通过后转换为 PDF；并可按版权中心在线申请表页面截图生成逐字段复制粘贴的“申请表填写内容”Markdown。当用户提到“软著”“软件著作权”“软著申请/登记”“源代码文档”“代码页 code.docx”“前30页后30页”“操作手册/软件说明书 软著”“软著申请表/在线填报/申请表填写内容”，或要求把一个 Python 桌面/软件项目整理成版权中心提交材料时使用。内置一套已授权成功案例的固定标准版式（A4、封面目录、宋体正文、代码页连续行号、页眉名称版本、右上角页码），运行于 Windows + WPS 或 MS Word。
---

# 软著材料生成（操作手册 + 源代码页 + 申请表填写内容）

生成版权中心要求的两份 Word 鉴别材料：**软件全称.docx（操作手册/说明书）** 与 **code.docx（源代码页）**；
用户在线填报时，另生成一份逐字段可复制的 **软著申请表填写内容 Markdown**（见第 5 节）。
版式为内置标准版式（来自已授权成功案例），规格见 [references/format-spec.md](references/format-spec.md)，
不得随意改字体字号；手册内容写法见 [references/manual-content-guide.md](references/manual-content-guide.md)；
在线申请表填写内容的字段、字数上限与取材规则见
[references/application-form-guide.md](references/application-form-guide.md)。

## 0. 开工前必须先问清并确认（一条消息问完）

缺少以下任一信息不要动手；用户已给的直接采用：

1. **输出目标文件夹**：最终 docx（及审核后 PDF）放在哪个目录。所有中间产物放在该目录下的工作子目录，
   不要写到项目源码目录、桌面或临时目录之外。
2. **源程序位置与范围**：项目根目录；程序入口文件（如 run.py / main.py）和主体源码包（如 app/）；
   与用户逐一确认**排除项**（tests、tools、打包脚本、.venv/venv、模型、数据库、照片、导出包、
   发布版 exe/build/dist 等非申请人编写或不应提交的内容）。
3. **手册材料**：操作手册草稿、需求/完整方案、README 等文件路径。
4. **软件全称、版本号、完成日期**：全称必须与将来申请表、两份文档页眉逐字一致。
5. 运行环境：Windows + 已安装 WPS 或 MS Word；可用的 Python 3（优先项目自带 .venv）。

> 开工时**不要**询问申请表填写内容 MD：它是两份 PDF 完成后的可选交付物，到第 5 步再主动询问。

确认后再开始；用户指定的目录和范围是硬约束，**不得删除、移动、改名项目中的任何文件**，
源码与材料一律只读使用。

## 1. 环境准备（隔离安装，不污染全局/项目环境）

脚本依赖 `python-docx`（必需）；PDF 侧校验可选 `pymupdf`。分页、目录页码、PDF 转换依赖
WPS（COM 对象 `KWPS.Application`）或 MS Word（`Word.Application`），脚本会自动二选一。

在输出目录下建工作目录并隔离安装依赖（PowerShell；`<work>` 为工作目录，`<py>` 为 Python 路径）：

```powershell
mkdir <work>\pylibs
<py> -m pip install --target <work>\pylibs python-docx pymupdf
# 之后每次运行 Python 脚本前：
$env:PYTHONPATH="<work>\pylibs"
```

- 编排逻辑写入 .py；调用 .ps1 用 `powershell -NoProfile -ExecutionPolicy Bypass -File ...`。
- .ps1 脚本只含 ASCII（PowerShell 5.1 对无 BOM 中文脚本会乱码）；含中文的 JSON 用 UTF-8。
- 没有 WPS/Word 时无法测量分页与转 PDF：明确告知用户安装，不要用估算页数冒充实测。

## 2. 源代码页 code.docx

```powershell
# (1) 扫描源码清单，按用户确认的范围编辑 files.txt（每行一个相对路径，顺序即拼接顺序）
<py> scripts/build_code.py discover --project <项目根> --out <work>\files.txt
# (2) 全量构建（含 TAB 的源码会报错退出，先让用户在项目内统一为空格缩进）
<py> scripts/build_code.py full --project <项目根> --files <work>\files.txt `
     --header "软件全称 V1.0" --out <work>\code_full.docx
# (3) 后台测量切点（不生成 PDF）
powershell -ExecutionPolicy Bypass -File scripts/find_code_cuts.ps1 `
     -Path <work>\code_full.docx -Out <work>\code_cuts.json
# (4) 切分为前30页+后30页
<py> scripts/build_code.py cut --project <项目根> --files <work>\files.txt `
     --header "软件全称 V1.0" --cuts <work>\code_cuts.json --out <work>\code.docx
# (5) 后台校验：恰好60页、除末页外每页≥50个行号（不生成 PDF）
powershell -ExecutionPolicy Bypass -File scripts/verify_code_com.ps1 -Path <work>\code.docx
```

规则与异常：

- 全量 ≤60 页时不切分，直接提交全量（find_code_cuts 会提示）。
- 源码**逐字保留**：不改写、不格式化、不删空行/注释，严禁把 Python 字符串里的英文双引号
  改成中文引号；word 技能 audit 对代码页报的 `E_CHINESE_ASCII_DOUBLE_QUOTE` 是误报，忽略。
- 前 30 页第一页必须是程序入口（run.py/main.py 放 files.txt 最前）；
  后 30 页结束于最后一个主体源文件的末行。
- verify 失败必须重测切点重建，不得手工凑行。

## 3. 操作手册 软件全称.docx

1. 通读第 0 步收集的全部材料（docx 用 word 技能的 read.py，md/txt 直接读），
   按 [manual-content-guide.md](references/manual-content-guide.md) 编写 `<work>\manual_spec.json`。
2. 渲染并回填目录页码：

```powershell
<py> scripts/build_manual.py --spec <work>\manual_spec.json --out <work>\软件全称.docx
powershell -ExecutionPolicy Bypass -File scripts/get_heading_pages.ps1 `
     -Path <work>\软件全称.docx -Out <work>\heading_pages.json
<py> scripts/patch_toc.py --docx <work>\软件全称.docx --headings <work>\heading_pages.json
```

3. 目录必须恰好 1 页；回填后再次后台重分页确认总页数与页码（不生成 PDF）。
4. 截图位置全部是灰框占位，提示语写明应截什么界面，由申请人自行贴图。
5. 可用 word 技能 `scripts/audit.py audit <docx>` 做格式体检（手册应 errors=0；
   代码页只看 font_audit，引号误报忽略）。

## 4. 人工审核闸门（硬性流程，不得跳过）

1. 两份 docx 自检通过后，**先交付 docx 给申请人人工审核**，明确告知：截图占位需贴图、
   请核对软件全称/版本/功能描述/代码范围。
2. **在申请人明确回复“审核通过/可以转 PDF”之前，禁止运行 export_pdf.ps1。**
3. 申请人提出修改：只改对应 spec/源码清单后重建 docx，重做第 2/3 步自检，再次回到本闸门等待。
4. 仅当两份 docx 都获确认后，才分别导出 PDF，并回读 PDF 页数（代码 60 页、
   满页≥50 行；手册与 docx 页数一致）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/export_pdf.ps1 -Path <work>\code.docx    -Pdf <输出目录>\code.pdf
powershell -ExecutionPolicy Bypass -File scripts/export_pdf.ps1 -Path <work>\软件全称.docx -Pdf <输出目录>\软件全称.pdf
```

两份 PDF 导出并回读核对后，**主动询问使用者**：“两份鉴别材料 PDF 已完成，是否需要我按版权
中心在线申请表，再生成一份逐字段可复制粘贴的申请表填写内容 Markdown？”使用者同意后进入第 5 步；
使用者也可在任意时刻主动要求，此时直接进入第 5 步。

## 5. 在线申请表填写内容 MD（PDF 完成后，主动询问再生成）

使用者确认需要、并提供在线填报页面截图后，按
[application-form-guide.md](references/application-form-guide.md) 生成
`软著申请表填写内容_<软件全称>_<版本号>.md`：

1. 按在线步骤页组织，每个文本字段给可整段复制的代码块，选项题标明点选，上传项对应两份 PDF。
2. 内容全部以实际为准：源程序量对确认范围实测行数；开发机硬件/系统实测；依赖版本取自
   requirements/打包配置；功能与技术特点只写材料中真实存在的内容；身份信息、完成/发表日期、
   联系方式留醒目占位由用户填写，不代填、不编造。
3. 交付前逐字段实测字数：50 限字段 ≤50、技术特点 ≤100、主要功能 500–1300，留 1–2 字余量。
4. MD 先生成交使用者核对；在线页面字段/上限以实时标注为准，截图未覆盖的身份、日期等页面
   只列占位清单，由使用者本人填报。

## 6. 交付前自检清单

- [ ] code.docx：60 页（或全量≤60页）、满页≥50 行、行号连续、页眉全称+版本、右上角页码连续、
  首页为程序入口、末页为源码结尾；无 tests/工具/依赖目录代码。
- [ ] 手册：封面/单页目录/正文两节；正文页码从 1；页眉全称版本；标题黑体居中、正文宋体小四
  1.5 倍行距；表格与图题规范；截图均为占位框；内容均可溯源到材料。
- [ ] 两份文档软件全称、版本号完全一致。
- [ ] 申请表 MD（仅在两份 PDF 完成、主动询问并获使用者确认后生成）：字段内容均可溯源、
  源程序量为确认范围实测、各字段字数合规、身份/日期为醒目占位未代填。
- [ ] 未经申请人确认不产出 PDF；最终文件按“软件全称.docx / code.docx”命名放入输出目录。

## 7. 脚本清单（scripts/）

| 脚本 | 作用 |
|---|---|
| build_manual.py | 由 manual_spec.json 渲染固定版式手册 docx |
| build_code.py | discover/full/cut 三子命令：发现源码、全量代码页、60 页切分 |
| find_code_cuts.ps1 | WPS/Word 后台二分测量前后 30 页切点 |
| verify_code_com.ps1 | 不导出 PDF，后台校验页数与每页行号数 |
| get_heading_pages.ps1 | 后台测量手册标题页码 |
| patch_toc.py | 回填手册目录页码 |
| export_pdf.ps1 | **人工审核通过后**才使用：docx 导出 PDF |
| verify_code.py | （可选）导出 PDF 后用 pymupdf 复核行号连续性并渲染抽查页 |
