# -*- coding: utf-8 -*-
"""软著源代码页 docx 构建器（Python 项目，宋体五号、连续行号、页眉名称+页码）。

三个子命令:

  1) discover  扫描项目生成有序文件清单（可手工增删/调序）
     python build_code.py discover --project <project_root> --out files.txt
     默认跳过 .venv/venv/env/__pycache__/build/dist/tests 等非交付源码目录。

  2) full      按清单拼接全部源码，生成全量代码文档（供分页测量）
     python build_code.py full --files files.txt --header "软件名 V1.0" --out code_full.docx

  3) cut       按 cuts.json 输出“前30页 + 分页符 + 后30页”（约60页）
     python build_code.py cut --files files.txt --header "软件名 V1.0" \
         --cuts code_cuts.json --out code.docx

cuts.json 由 find_code_cuts.ps1 对 full 文档测量得到，字段:
  total_pages, paragraphs, front_end, back_start（行索引从 0 开始）。

源码逐字保留：禁止把源码中的英文双引号改成中文引号；含 TAB 缩进的行会报错退出，
需先在项目中统一为空格缩进（软著代码页不接受 TAB 造成的错位）。
"""
import argparse, json, os, sys
from docx import Document
from docx.shared import Pt, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# 默认跳过：依赖/缓存/构建产物/测试/辅助工具与打包脚本目录。
# 发现结果只是初稿，必须由使用者结合实际确认范围后再编辑 files.txt
# （个别项目若把业务主体放在 tools 等目录，可手工加回）。
SKIP_DIRS = {'.venv', 'venv', 'env', '.env', '__pycache__', 'build', 'dist',
             'tests', 'test', 'tools', 'packaging', 'docs', 'examples',
             'fixtures', '.git', '.idea', '.vscode', 'node_modules', 'migrations'}


def discover(project, out):
    root = os.path.abspath(project)
    chosen = []
    # 根目录 run.py / main.py 优先，随后根目录其余 .py
    root_pys = sorted(f for f in os.listdir(root)
                      if f.endswith('.py') and os.path.isfile(os.path.join(root, f)))
    priority = [f for f in ('run.py', 'main.py', 'app.py') if f in root_pys]
    ordered_root = priority + [f for f in root_pys if f not in priority]
    chosen += [f for f in ordered_root]
    # 包目录（含 __init__.py），包内 __init__ 优先
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith('.')]
        if dirpath == root:
            continue
        if '__init__.py' not in filenames:
            continue
        pys = sorted(f for f in filenames if f.endswith('.py'))
        pys = (['__init__.py'] if '__init__.py' in pys else []) + \
              [f for f in pys if f != '__init__.py']
        rel_dir = os.path.relpath(dirpath, root)
        chosen += [os.path.join(rel_dir, f).replace('/', '\\') for f in pys]
    with open(out, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(chosen))
    print('discovered %d files -> %s' % (len(chosen), out))
    for f in chosen:
        print('  ', f)


def read_lines(project, files_list):
    lines, tab_hits = [], []
    with open(files_list, encoding='utf-8') as fh:
        rels = [ln.strip() for ln in fh if ln.strip()]
    for rel in rels:
        path = os.path.join(project, rel)
        with open(path, encoding='utf-8') as fh:
            content = fh.read()
        ls = content.split('\n')
        while ls and ls[-1] == '':
            ls.pop()
        for i, l in enumerate(ls, 1):
            if '\t' in l:
                tab_hits.append('%s:%d' % (rel, i))
        lines.extend(ls)
    if tab_hits:
        print('发现 TAB 缩进行，软著代码页需统一空格缩进，请先处理：', file=sys.stderr)
        for h in tab_hits[:50]:
            print('  ', h, file=sys.stderr)
        sys.exit(2)
    return lines, rels


def _new_doc(header_text, line_count_hint=None):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Twips(11906); sec.page_height = Twips(16838)
    sec.top_margin = Twips(1440); sec.bottom_margin = Twips(1440)
    sec.left_margin = Twips(1800); sec.right_margin = Twips(1800)
    sec.header_distance = Twips(851); sec.footer_distance = Twips(992)

    normal = doc.styles['Normal']
    normal.font.name = '宋体'; normal.font.size = Pt(10.5)
    rpr = normal.element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts'); rpr.insert(0, rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), '宋体')
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    sectPr = sec._sectPr
    ln = OxmlElement('w:lnNumType'); ln.set(qn('w:countBy'), '1'); ln.set(qn('w:restart'), 'continuous')
    sectPr.append(ln)
    grid = OxmlElement('w:docGrid'); grid.set(qn('w:type'), 'lines'); grid.set(qn('w:linePitch'), '312')
    sectPr.append(grid)

    hdr = sec.header; hdr.is_linked_to_previous = False
    hp = hdr.paragraphs[0]
    for old in list(hp.runs):
        old._element.getparent().remove(old._element)
    hp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    hp.paragraph_format.line_spacing = 1.0
    tabs = hp.paragraph_format.tab_stops
    tabs.add_tab_stop(Twips(4153), WD_TAB_ALIGNMENT.CENTER)
    tabs.add_tab_stop(Twips(8306), WD_TAB_ALIGNMENT.RIGHT)
    pPr = hp._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr'); bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '4')
    bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), 'auto')
    pbdr.append(bottom); pPr.append(pbdr)

    def hrun(text=None):
        r = hp.add_run()
        rPr = r._r.get_or_add_rPr()
        rfonts = OxmlElement('w:rFonts')
        for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            rfonts.set(qn(a), '宋体')
        rPr.append(rfonts)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '18'); rPr.append(sz)
        if text is not None:
            t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text; r._r.append(t)
        return r

    def field(child):
        r = hrun(); r._r.append(child)
    hrun('\t' + header_text + '\t')
    b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin'); field(b)
    ri = hrun()
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    ri._r.append(instr)
    sep = OxmlElement('w:fldChar'); sep.set(qn('w:fldCharType'), 'separate'); field(sep)
    rc = hrun('1')
    e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end'); field(e)

    ftr = sec.footer; ftr.is_linked_to_previous = False
    ftr.paragraphs[0].paragraph_format.line_spacing = 1.0
    return doc


def append_code_line(doc, text):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:after'), '46'); sp.set(qn('w:line'), '210'); sp.set(qn('w:lineRule'), 'exact')
    pPr.append(sp)
    r = p.add_run()
    rPr = r._r.get_or_add_rPr()
    rfonts = OxmlElement('w:rFonts')
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rfonts.set(qn(a), '宋体')
    rPr.append(rfonts)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '21'); rPr.append(sz)
    szcs = OxmlElement('w:szCs'); szcs.set(qn('w:val'), '21'); rPr.append(szcs)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text; r._r.append(t)
    return p


def build_full(lines, header_text, out):
    doc = _new_doc(header_text)
    for text in lines:
        append_code_line(doc, text)
    doc.save(out)
    print('full saved: %d lines -> %s' % (len(lines), out))


def build_cut(lines, cuts, header_text, out):
    fe, bs = cuts['front_end'], cuts['back_start']
    front, back = lines[:fe], lines[bs:]
    doc = _new_doc(header_text)
    for text in front:
        append_code_line(doc, text)
    br_run = doc.paragraphs[-1].add_run()
    br = OxmlElement('w:br'); br.set(qn('w:type'), 'page'); br_run._r.append(br)
    for text in back:
        append_code_line(doc, text)
    doc.save(out)
    print('cut saved: front %d + back %d = %d paragraphs -> %s'
          % (len(front), len(back), len(front) + len(back), out))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest='cmd', required=True)

    d = sub.add_parser('discover')
    d.add_argument('--project', required=True); d.add_argument('--out', required=True)

    f = sub.add_parser('full')
    f.add_argument('--project', required=True); f.add_argument('--files', required=True)
    f.add_argument('--header', required=True); f.add_argument('--out', required=True)

    c = sub.add_parser('cut')
    c.add_argument('--project', required=True); c.add_argument('--files', required=True)
    c.add_argument('--header', required=True); c.add_argument('--cuts', required=True)
    c.add_argument('--out', required=True)

    args = ap.parse_args()
    if args.cmd == 'discover':
        discover(args.project, args.out)
    else:
        lines, rels = read_lines(args.project, args.files)
        print('files: %d, source lines: %d' % (len(rels), len(lines)))
        if args.cmd == 'full':
            build_full(lines, args.header, args.out)
        else:
            cuts = json.load(open(args.cuts, encoding='utf-8'))
            build_cut(lines, cuts, args.header, args.out)


if __name__ == '__main__':
    main()
