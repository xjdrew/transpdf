# -*- coding: utf8 -*-
import sys, pathlib, csv,re
import logging

import fitz

# logging
logging.basicConfig(
    level=logging.INFO, 
    format="[%(levelname)s] %(message)s")


def doc_get_text(pdf_file):
    doc = fitz.open(pdf_file)
    words = []
    # https://pymupdf.readthedocs.io/en/latest/app1.html#text-extraction-flags
    for page in doc:
        for t in page.get_text("words", flags = fitz.TEXTFLAGS_SEARCH & ~fitz.TEXT_DEHYPHENATE):
            # (x0, y0, x1, y1, "word", block_no, line_no, word_no)
            words.append(t[4])
    return words

def tidy_words(words):
    output = []

    for w in words:
        w = w.strip()
        if len(w) == 0:
            continue
        if re.match("^[-—）]*$", w):
            continue
        if not w.isascii():
            output.append(w)
    return output

def unique_words(words):
    words = list(set(words))
    words.sort()
    return words

def _color_output(msg): return f'\033[1;36m{msg}\033[0m'

def _guess_file_name():
    prefix = 'dict'
    suffix = '.csv'
    for i in range(100):
        fname = '{0}{1}{2}'.format(prefix, i, suffix)
        if pathlib.Path(fname).is_file():
            continue
        return fname
    return ''

def _default_files():
    p = pathlib.Path(".")
    pdf_files = list(p.glob('*.pdf'))
    suffix = "_new.pdf"

    ret = []
    for f in pdf_files:
        if len(f.name) > len(suffix) and f.name[-len(suffix):] == suffix:
            continue
        ret.append(f.name)
    return ret

if __name__ == '__main__':
    pdf_files = None
    if len(sys.argv) == 1:
        # 没有输入任何pdf文件
        pdf_files = _default_files()
    else:
        pdf_files = sys.argv[1:]

    total_files = len(pdf_files)
    if total_files < 1:
        logging.error("请提供待提取文本的pdf文件")
        exit(1)

    logging.info("%d 个文件待提取", total_files)
    logging.info(_color_output("开始提取文本"))
    words = []
    for i in range(0, len(pdf_files)):
        pdf_file = pathlib.Path(pdf_files[i])
        if pdf_file.suffix != ".pdf": # 检查参数，避免输错文件
            logging.warning("{{{0}/{1}}} 忽略非pdf文件 {2}".format(i+1, total_files, pdf_file))
            continue
        logging.info("{{{0}/{1}}} 提取文件 {2}".format(i+1, total_files, pdf_file))
        words.extend(doc_get_text(pdf_file))

    logging.info("共提取文本 %d 条", len(words))
    logging.info(_color_output("开始清理文本"))
    w1 = tidy_words(words)
    logging.info("去除非中文后，剩余%d条", len(w1))

    w2 = unique_words(w1)
    logging.info("去除重复后，剩余%d条", len(w2))

    logging.info(_color_output("开始写入文件"))

    fname = _guess_file_name()
    if fname == '':
        logging.error('生成文件过多，请先删除一些')
    
    logging.info("file: %s", fname)

    with open(fname, 'w', encoding='utf-8-sig', newline='') as csvfile:
        writer = csv.writer(csvfile, dialect='excel')
        for word in w2:
            writer.writerow([word])
    logging.info(_color_output("写入文件成功"))
