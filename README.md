# software-copyright-dbdb

为中国计算机软件著作权（软著）登记生成两份 Word 鉴别材料的 AI 助手技能：

- **操作手册（软件说明书）.docx**：封面、单页目录、正文，固定标准版式，截图位置自动留灰框；
- **源代码页 code.docx**：源程序前 30 页 + 后 30 页（不足 60 页则交全部），连续行号、页眉软件名称与版本、右上角连续页码。

两份 docx 先生成交付人工审核，**经使用者明确确认后才转换 PDF**。

## 特性

- 版式固化为脚本：A4、页边距、字体字号、行距、分页、目录页码回填全部自动完成，不依赖手工排版；
- 源代码逐字保真：不改写、不格式化、不删空行与注释，严禁用空行/注释凑行数；
- 切点自动测量：调用 WPS 或 MS Word 后台重分页，二分定位前 30 页 / 后 30 页边界，并逐页校验每页不少于 50 行（末页除外）；
- 开工前强制确认：输出目标文件夹、源程序位置与提交范围、手册材料、软件全称与版本号；
- 全流程本地运行：不联网、不上传、无遥测；唯一网络行为是首次使用时按需 `pip install python-docx`；
- 自动兼容 **WPS（KWPS）与 Microsoft Word（Word.Application）**。

## 运行环境

- Windows 10 / 11
- 已安装 WPS Office 或 Microsoft Word（分页测量、目录页码、PDF 转换依赖其 COM 组件）
- Python 3.10+（仅需标准库 + python-docx；PDF 复核可选 pymupdf）

## 安装

将本仓库（或 Release 中的 zip 解压后）整个文件夹放入 AI 助手的用户技能目录，使路径形如：

```
<助手工作目录>/workspace/.user_skills/software-copyright-dbdb/SKILL.md
```

重启或刷新助手后，对助手说“帮我做软著材料 / 生成 code.docx 和操作手册”即可触发。首次使用时助手会在你指定的输出目录下隔离安装依赖，不污染全局 Python 环境。

## 使用流程

1. 助手一次性问清：输出目录、项目源码根目录与提交范围（入口文件、主体包、排除 tests/tools/.venv 等）、手册材料、软件全称与版本；
2. 生成源代码页：扫描源码清单（需你确认范围）→ 全量构建 → 测量切点 → 切分 60 页 → 后台校验；
3. 生成操作手册：助手阅读你提供的材料，编写 `manual_spec.json`，渲染并回填目录页码；
4. **人工审核闸门**：先拿到两份 docx，自行贴图、核对名称版本与功能描述；你明确回复“审核通过”后，助手才导出 PDF；
5. 导出后回读 PDF 核对页数与行数。

详细版式规则见 `references/format-spec.md`，手册内容规格见 `references/manual-content-guide.md`。

## 仓库结构

```
software-copyright-dbdb/
├── SKILL.md                          # 技能主文件（工作流与命令）
├── LICENSE                           # MIT
├── references/
│   ├── format-spec.md                # 标准版式与 60 页规则
│   └── manual-content-guide.md       # 手册内容与 manual_spec.json 规格
└── scripts/
    ├── build_manual.py               # 由 spec 渲染手册 docx
    ├── build_code.py                 # discover / full / cut：发现源码、全量页、60 页切分
    ├── find_code_cuts.ps1            # WPS/Word 后台测量切点
    ├── verify_code_com.ps1           # 不导出 PDF，后台校验页数与每页行数
    ├── get_heading_pages.ps1         # 测量手册标题页码
    ├── patch_toc.py                  # 回填目录页码
    ├── export_pdf.ps1                # 人工审核通过后才使用：docx 导出 PDF
    └── verify_code.py                # 可选：导出 PDF 后用 pymupdf 复核
```

## 合规说明

脚本内置版式依据中国版权保护中心公开的登记材料要求（源程序与文档前后各连续 30 页、程序每页不少于 50 行、页眉标注软件名称与版本、右上角连续页码等）整理。登记要求可能调整，正式提交前请以中国版权保护中心（ccopyright.com.cn）当期要求为准。本工具只负责排版辅助，材料内容的真实性、合法性由使用者负责。

## 许可证

[MIT](LICENSE) © wyfanfanfan57-bit
