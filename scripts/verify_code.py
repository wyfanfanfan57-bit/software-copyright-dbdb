# -*- coding: utf-8 -*-
"""校验代码页 PDF：总页数、每页行号数量（满页应 >=50）、行号全局连续；
并渲染首页、前后分界页与末页 PNG 供人工目检。

用法:
  python verify_code.py --pdf code.pdf --imgdir img [--expect-pages 60]
退出码非 0 表示校验失败。
"""
import argparse, os, re, sys
import fitz  # PyMuPDF


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pdf', required=True)
    ap.add_argument('--imgdir', required=True)
    ap.add_argument('--expect-pages', type=int, default=0)
    args = ap.parse_args()

    d = fitz.open(args.pdf)
    os.makedirs(args.imgdir, exist_ok=True)
    summary, prev_last = [], None
    failures = []
    for i in range(d.page_count):
        nums = set()
        for w in d[i].get_text('words'):
            x0, word = w[0], w[4]
            if x0 < 95 and re.fullmatch(r'\d{1,5}', word):
                nums.add(int(word))
        nums = sorted(nums)
        if not nums:
            failures.append('page %d: no line numbers' % (i + 1))
        else:
            if prev_last is not None and nums[0] != prev_last + 1:
                failures.append('page %d: line numbers not continuous (%d -> %d)'
                                % (i + 1, prev_last, nums[0]))
            prev_last = nums[-1]
        summary.append((i + 1, nums[0] if nums else None, nums[-1] if nums else None, len(nums)))

    for s in summary:
        print('page %2d first=%s last=%s count=%s' % s)
    last_page = d.page_count
    for pno, s in enumerate(summary):
        if s[3] < 50 and (pno + 1) != last_page:
            failures.append('page %d has only %d numbered lines (<50)' % (pno + 1, s[3]))
    if args.expect_pages and d.page_count != args.expect_pages:
        failures.append('page count %d != expected %d' % (d.page_count, args.expect_pages))

    picks = sorted({0, min(29, last_page - 1), max(0, last_page - 30), last_page - 1})
    for i in picks:
        d[i].get_pixmap(dpi=110).save(os.path.join(args.imgdir, 'p%02d.png' % (i + 1)))
    print('rendered pages:', [i + 1 for i in picks])

    if failures:
        print('FAILURES:', file=sys.stderr)
        for fmsg in failures:
            print(' -', fmsg, file=sys.stderr)
        sys.exit(1)
    print('VERIFY OK: %d pages, line numbers 1..%s continuous' % (last_page, prev_last))


if __name__ == '__main__':
    main()
