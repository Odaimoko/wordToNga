#!/usr/bin/env python3
# coding=utf-8
import argparse
import pathlib
import re
import subprocess
import zipfile

from bs4 import BeautifulSoup
from docx import Document

working_dir = "temp/"

HEADER_TO_SIZE = {
    "h1": "150%",
    "h2": "130%",
    "h3": "120%",
    "h4": "110%",
}
STYLE_NAME_TO_COLOR = {
    "酒红": "deeppink",
    "蓝色": "blue"
}


# 纯文本，暂不支持表格
# 支持样式
# 支持颜色
# 支持加粗
# 支持无序列表
# 不支持图片，有谁告诉我怎么自动上传到nga服务器？？

def parse_args():
    parser = argparse.ArgumentParser(description = 'World To Nga document converter')
    parser.add_argument("-i", "--i", help = "Input document path")
    args = parser.parse_args()
    return args


def get_nga_color_name_by_style_name(style_name: str) -> str:
    """将样式名中的颜色转换为nga的颜色tag"""
    for style_color in STYLE_NAME_TO_COLOR:
        if style_color in style_name:
            return STYLE_NAME_TO_COLOR[style_color]
    print("[Error] No color matched for {0}.".format(style_name))
    return "black"


def is_custom_style(style_name: str) -> bool:
    """自定义的命名规范： B站-颜色+ 其他样式"""
    return "B站" in style_name


# https://stackoverflow.com/questions/60201419/extract-all-the-images-in-a-docx-file-using-python
def extract_images(input_doc):
    """将word文档的所有图片按名字提取到文档同名文件夹下"""
    archive = zipfile.ZipFile(input_doc)
    img_path = input_doc.with_name(input_doc.name + "_img")
    for file in archive.filelist:
        if file.filename.startswith('word/media/'):
            archive.extract(file, img_path)


def add_doc_styles_tag(input_doc):
    document = Document(input_doc)
    paragraph_styles = [
        s for s in document.styles if is_custom_style(s.name)
    ]
    for para in document.paragraphs:
        # print(para.style.name, '\t', para.text, )
        for run in para.runs:  # 遍历所有的，至少文本可以
            if run.style in paragraph_styles:
                new_text = "[color={ngacolor}]{content}[/color]".format(
                    ngacolor = get_nga_color_name_by_style_name(run.style.name), content = run.text)
                print('\t', run.style.name, '\t', run.text, "->", new_text)
                run.text = new_text
    output_doc = input_doc.with_name(input_doc.stem + "_Mod").with_suffix(".docx")
    document.save(output_doc)
    return output_doc


def html_as_intermediate(input_doc):
    output_doc = input_doc.with_suffix(".html")
    subprocess.call("pandoc -i {0} -o {1}".format(input_doc, output_doc))
    with open(output_doc, "r", encoding = "utf8") as f:
        html_str = f.read()
    soup = BeautifulSoup(html_str, 'html.parser')
    
    for tag in HEADER_TO_SIZE:  # Header 替换成字体大小+加粗
        for h1 in soup.select(tag):
            h1.string = "[b][size={size}]{content}[/size][/b]".format(size = HEADER_TO_SIZE[tag], content = h1.string)
    # 转化为NGA格式
    for tag in soup.select("a"):  # 链接
        tag.string = "[url={url}]{content}[/url]".format(url = tag.attrs['href'], content = tag.string)
    for tag in soup.select("strong"):  # 加粗
        tag.string = "[b]{content}[/b]".format(content = tag.string)
    
    for tag in soup.select("ul"):  # 无序列表
        tag.insert_before("[list]")
        for item in tag.select("li"):
            item = item.next  # <li>里面还有个<p>，这个<p>会在后面变成br，所以这里不用额外加换行。
            item.insert_before("[*]")
        tag.insert_after(BeautifulSoup("<p>[/list]</p>", "html.parser"))  # 主动加一个<p>，最后变成换行
    
    for tag in soup.select("p"):  # 将HTML里的 paragraph/段落换成 linebreak/换行，免得复制的时候一句话占两行。
        tag.name = "br"
    
    # 图片，无法自动插入
    # 故意将图片换回段落，以在编辑器中有换行，更明显看出这里需要插入图片。
    for tag in soup.select("img"):
        tag.string = "【我去这里需要插入图片：{0}】".format(tag.attrs["src"])
        tag.attrs = {}
        tag.name = "p"
    
    html_str = str(soup)  # 不能prettify。 prettify会换行。
    html_str = html_str.replace("<br>", "")
    
    # 找到所有设置图片大小的，删除 {width="3.38125in" height="3.807638888888889in"}
    width_height_pattern = r"""\{width="\d.*in"\sheight=".*in"\}"""  # width和height中间可能是空格可能是回车
    # print(re.findall(width_height_pattern, html_str))
    html_str = re.sub(width_height_pattern, "", html_str)
    with open(output_doc.with_suffix(".html"), "w", encoding = "utf8") as f:
        f.write(html_str)


def main():
    args = parse_args()
    input_doc = pathlib.PurePosixPath(args.i)
    extract_images(input_doc)
    output_doc = add_doc_styles_tag(input_doc)
    html_as_intermediate(output_doc)


if __name__ == '__main__':
    main()
