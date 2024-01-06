# -*- coding: utf8 -*-
import sys, pathlib, csv
import logging

import fitz
#from pdf2docx import Converter

# logging
logging.basicConfig(
    level=logging.INFO, 
    format="[%(levelname)s] %(message)s")

def _color_output(msg): return f'\033[1;36m{msg}\033[0m'

# 替换后的字体
font_name = "helv"
font_size = 5
font_height = None

def init_env(ename):
    global font_size, font_height

    logging.info("int env: %s", ename)
    if ename=="PINGAN":
        font_size = 6
    elif ename=="CMB":
        # font_name = "SimSun"
        font_size = 5
    elif ename=="ABC": # 农业银行
        font_size = 5
    else:
        assert False, "uknown bank " + ename
    
    font = fitz.Font(fontname=font_name)
    font_height = (font.ascender - font.descender)*font_size

class Dictionary:
    def __init__(self, keys, table):
        self.keys = keys
        self.table = table
        
        # 按照长度排序，替换时优先替换长字符串，避免短字符串是长字符串的子串
        self.keys.sort(key=lambda s:[len(s), s], reverse=True)
        pass

def read_dict(fname):
    keys = list()
    table = dict()
    with open(fname, encoding='utf-8-sig', newline='') as csvfile:
        reader = csv.reader(csvfile, dialect='excel')
        for row in reader:
            key = row[0].strip()
            value = row[1].strip()
            if len(key) > 0:
                if len(value) > 0:
                    table[key] = value
                    keys.append(key)
                    # logging.info("dict %s %s", key, value)
                else:
                    logging.warning("no translate: %s", key)
    return Dictionary(keys, table)
    
def calc_new_rect(old_rect, text):
    text_length = fitz.get_text_length(text, fontname=font_name, fontsize=font_size)
    text_height = font_height
    old_width = old_rect.width
    old_height = old_rect.height
    if old_height < text_height:
        logging.error("new text height(%d) > old text height(%d)", text_height, old_height)

    # 默认扩充列宽
    width_up = 20

    y_diff = 0 # 起始高度微调
    new_width = None
    new_height = None
    if text_length - old_width <= width_up:
        new_width = text_length + 3 # 长度调整
        new_height = text_height
    elif text_length < (old_width + width_up)*2:
        new_width = old_width + width_up
        new_height = text_height * 2
        # y_diff = -(text_height/2)
    elif text_length < (old_width + width_up*3)*2:
        new_width = old_width + width_up*3
        new_height = text_height * 2
        # y_diff = -(text_height/2)
    else:
        logging.warning("[%s] is too long", text)
        new_width = old_width + width_up*5
        new_height = text_height * 2
        # y_diff = -(text_height/2)

    return fitz.Rect(old_rect.x0, old_rect.y0+y_diff, old_rect.x0 + new_width, old_rect.y0 + new_height + y_diff)

def delete_all_image(page):
    pix = fitz.Pixmap(fitz.csGRAY, (0, 0, 1, 1), 1)  # small pixmap
    pix.clear_with(255)  # empty its content

    img_list = page.get_images()
    for img in img_list:
        page.replace_image(img[0], pixmap=pix)
    return len(img_list)

def delete_all_signatures(page):
    widgets = page.widgets([fitz.PDF_WIDGET_TYPE_SIGNATURE])
    count = 0
    for widget in widgets:
        count = count + 1
        page.delete_widget(widget)
    return count

def translate(dictionary, pdf_file, new_file):
    keys = dictionary.keys
    table = dictionary.table

    doc = fitz.open(pdf_file)

    total_images = 0
    total_signatures = 0
    for page in doc:
        # 移除图片若有
        total_images = total_images + delete_all_image(page)
        total_signatures = total_signatures + delete_all_signatures(page)
    logging.info("remove %d images", total_images)
    logging.info("remove %d signatures", total_signatures)

    # search all hits
    search_results = {} # key -> [(page1, hits), (page2, hits)]
    for key in keys:
        # 依次处理每页
        for page in doc:
            # 搜索原值
            hits = page.search_for(key, quads=True, flags=fitz.TEXT_PRESERVE_WHITESPACE | fitz.TEXT_PRESERVE_LIGATURES)
            if not (key in search_results):
                search_results[key] = []
            search_results[key].append((page, hits))

            # 涂掉，避免被后面的短字符串搜索到
            for quad in hits:
                page.add_redact_annot(quad)
            page.apply_redactions()

    # clear
    # for key in keys:
    #     results = search_results[key]
    #     for (page, hits) in results:
    #         # 清除原值
    #         for quad in hits:
    #             page.add_redact_annot(quad)
    #         page.apply_redactions()
    
    # replace
    repeated = {}
    for key in keys:
        results = search_results[key]
        for (page, hits) in results:
            # 填充新值
            for quad in hits:
                text = table[key]
                if text == '-': # - 表示留空
                    continue
                newrect = calc_new_rect(quad.rect, text)
                page.add_freetext_annot(newrect, text=text, fontname=font_name, fontsize=font_size, align=fitz.TEXT_ALIGN_LEFT)
                # logging.info("++ %s - %s", key, text)
                # if key == '交易时间' or key == '交易摘要' or key == '交易⽇期':
                # logging.info("------------------ %s - %s - %s - %s", key, text, str(newrect), str(annot))
                # logging.info("------------------ %s - %s", key, text)

    doc.save(new_file)

def convert_to_docx(pdf_file, docx_file):
    # convert pdf to docx
    cv = Converter(pdf_file)
    cv.convert(docx_file)      # all pages by default
    cv.close()

# 生成待翻译文件列表
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
    if len(sys.argv) < 3:
        print("Usage: {0} bank dict.csv origin.pdf".format(sys.argv[0]))
        exit(1)
    
    init_env(sys.argv[1])

    if len(sys.argv) == 3:
        # 没有输入任何pdf文件
        pdf_files = _default_files()
    else:
        pdf_files = sys.argv[2:]

    dict_fname = pathlib.Path(sys.argv[2])
    if dict_fname.suffix != ".csv":
        logging.error("dict must be csv")
        exit(2)
    
    dictionary = read_dict(dict_fname)
    logging.info("%d valid keys in dictionary", len(dictionary.keys))
    #print("\n".join(dictionary.keys))

    total_files = len(pdf_files)
    logging.info("开始翻译文件, 总数 %d", total_files)

    for i in range(0, total_files):
        pdf_file = pathlib.Path(pdf_files[i])
        if pdf_file.suffix != ".pdf": # 检查参数，避免输错文件
            logging.warning("{{{0}/{1}}} 忽略非pdf文件 {2}".format(i+1, total_files, pdf_file))
            continue
        
        new_file = pdf_file.with_name(pdf_file.name[:-4] + "_new.pdf")
        logging.info(_color_output("{{{0}/{1}}} {2} 翻译为 {3}".format(i+1, total_files, pdf_file, new_file)))
        translate(dictionary, pdf_file, new_file)

        # docx_file = new_file.with_suffix(".docx")
        # logging.info(_color_output("{{{0}/{1}}} {2} 转换为 {3}".format(i-1, total_files, new_file, docx_file)))
        # convert_to_docx(new_file, docx_file)
    
    logging.info(_color_output("完成"))
    
    