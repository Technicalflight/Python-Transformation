#!/usr/bin/python
# -*- coding:utf-8 -*-
# @author  : Technicalflight
# @time    : 2022.7
# @version : 
#   V1 
#

from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM
import cv2
import execjs
import aspose.words as aw
import js2py
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPDF
#svg to png
def svg_to_png():
    svg = input("请输入svg文件的路径:")
    png = input("请输入保存为png格式的文件名:")
    pic = svg2rlg(svg)
    renderPM.drawToFile(pic,png + '.png')

#png to svg
def png_to_svg():
    png = input("请输入png文件的路径:")
    fileNames = [png]
    svg = input("请输入保存为svg格式的文件名:")
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    shapes = [builder.insert_image(fileName) for fileName in fileNames]
    pageSetup = builder.page_setup
    pageSetup.page_width = max(shape.width for shape in shapes)
    pageSetup.page_height = sum(shape.height for shape in shapes)
    pageSetup.top_margin = 0
    pageSetup.left_margin = 0
    pageSetup.bottom_margin = 0
    pageSetup.right_margin = 0
    doc.save(svg + ".svg")

#svg to jpg
def svg_to_jpg():
    svg = input("请输入svg文件的路径:")
    jpg = input("请输入保存为jpg格式的文件名:")
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    shape = builder.insert_image(svg)
    shape.image_data.save(jpg + ".jpg")

#jpg to svg
def jpg_to_svg():
    jpg = input("请输入jpg文件的路径:")
    fileNames = [jpg]
    svg = input("请输入保存为svg格式的文件名:")
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    shapes = [builder.insert_image(fileName) for fileName in fileNames]
    pageSetup = builder.page_setup
    pageSetup.page_width = max(shape.width for shape in shapes)
    pageSetup.page_height = sum(shape.height for shape in shapes)
    pageSetup.top_margin = 0
    pageSetup.left_margin = 0
    pageSetup.bottom_margin = 0
    pageSetup.right_margin = 0
    doc.save(svg + ".svg")

#pdf to svg
def pdf_to_svg():
    pdf = input("请输入pdf文件的路径:")
    svg = input("请输入保存为svg格式的文件名:")
    doc = aw.Document(pdf)
    for page in range(0, doc.page_count):
        extractedPage = doc.extract_pages(page, 1)
        extractedPage.save(f"Output_{page + 1}" + svg +".svg")
    #从1开始遍历所有文件
    print("按任意键继续...")
    input()
    for i in range(1, doc.page_count):
        #打开文件并读取代码
        with open(f"Output_{i}" + svg +".svg", "r",encoding='utf-8') as f:
            code = f.read()
        #删除标签内的内容
        code = code.replace("<image", "<text")
        #保存文件
        with open(f"Output_{i}" + svg +".svg", "w",encoding='utf-8') as f:
            f.write(code)
    print("如果想删除2段红色的文字，请在浏览器内F12打开控制台，输入以下命令：")
    print( 
    '''
      document.querySelector("svg > g > g:nth-child(5) > g > g:nth-child(1)").remove();
      document.querySelector("svg > g > g:nth-child(4) > g:nth-child(1)").remove();
      如果有残留的水印，请删除svg文件中的水印<image>标签
    '''
    )


#svg to pdf
def svg_to_pdf():
    svg = input("请输入svg文件的路径:")
    pdf = input("请输入保存为pdf格式的文件名:")
    drawing = svg2rlg(svg)
    renderPDF.drawToFile(drawing, pdf + '.pdf')


#pdf to word
def pdf_to_word():
    print('此方法生成的word含有水印，请自行删除')
    pdf = input("请输入pdf文件的路径:")
    word = input("请输入保存为word格式的文件名:")
    doc = aw.Document(pdf)
    doc.save(word + ".docx")

#word to pdf
def word_to_pdf():
    print('此方法生成的pdf含有水印，请自行删除')
    word = input("请输入word文件的路径:")
    pdf = input("请输入保存为pdf格式的文件名:")
    doc = aw.Document(word)
    doc.save(pdf + ".pdf")






if __name__ == '__main__':
    #死循环
    while True:
        cz = input("""\033[1;32m
                            Technicalflight
                                yyds
                            --------------
                            1.svg to png
                            2.png to svg
                            3.svg to jpg
                            4.jpg to svg
                            5.pdf to svg
                            6.svg to pdf
                            7.pdf to word
                            8.word to pdf
                            9.exit
                  .-~~~~~~~~~-._       _.-~~~~~~~~~-.
              __.'              ~.   .~              `.__
            .'//                  \./                  \\`.
          .'//                     |                     \\`.
        .'// .-~"""""""~~~~-._     |     _,-~~~~"""""""~-. \\`.
      .'//.-"                 `-.  |  .-'                 "-.\\`.
    .'//______.============-..   \ | /   ..-============.______\\`.
  .'______________________________\|/______________________________`.
        \033[0m""")
        if cz == '1':
            svg_to_png()
            print("图片已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '2':
            png_to_svg()
            print("svg已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '3':
            svg_to_jpg()
            print("jpg已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '4':
            jpg_to_svg()
            print("svg已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '5':
            pdf_to_svg()
            print("svg已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '6':
            svg_to_pdf()
            print("pdf已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '7':
            pdf_to_word()
            print("word已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '8':
            word_to_pdf()
            print("pdf已保存")
            print('按任意键返回操作菜单')
            input()
        elif cz == '9':
            print("退出")
            break