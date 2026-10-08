# -*- coding: utf-8 -*-
"""回填手册目录页码：把目录条目末尾的占位“\\t0”替换为 get_heading_pages.ps1 测得的页码。

用法:
  python patch_toc.py --docx 手册.docx --headings heading_pages.json
目录条目数必须与标题数一致，否则报错（说明目录与正文标题不匹配）。
"""
import argparse, json
from docx import Document


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--docx', required=True)
    ap.add_argument('--headings', required=True)
    args = ap.parse_args()

    meta = json.load(open(args.headings, encoding='utf-8-sig'))
    pages = [h['pageAdj'] for h in meta['headings']]

    doc = Document(args.docx)
    idx = 0
    for p in doc.paragraphs:
        runs = p.runs
        if runs and runs[-1].text.startswith('\t0') and len(runs[-1].text) <= 3:
            if idx >= len(pages):
                raise SystemExit('目录条目多于标题数：%d' % idx)
            runs[-1].text = '\t' + str(pages[idx])
            idx += 1
    if idx != len(pages):
        raise SystemExit('目录条目数 %d 与标题数 %d 不一致' % (idx, len(pages)))
    doc.save(args.docx)
    print('TOC patched, entries:', idx)


if __name__ == '__main__':
    main()
