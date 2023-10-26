# -*- coding: utf-8 -*-
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import mm
from reportlab.lib import colors

# 初期設定
def make(filename="resume"): # ファイル名
    pdf_canvas = set_info(filename) # キャンバス名
    print_string(pdf_canvas)
    pdf_canvas.save() # 保存

def set_info(filename):
    pdf_canvas = canvas.Canvas("./{0}.pdf".format(filename)) # 保存先
    pdf_canvas.setAuthor("") # 作者
    pdf_canvas.setTitle("") # 表題
    pdf_canvas.setSubject("") # 件名
    return pdf_canvas

#履歴書フォーマット作成
def print_string(pdf_canvas):
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5')) # フォント
    width, height = A4 # 用紙サイズ


    # (1)履歴書 タイトル
    font_size = 24 # フォントサイズ
    pdf_canvas.setFont('HeiseiKakuGo-W5', font_size)
    pdf_canvas.drawString(60, 770, '給与明細') # 書き出し(横位置, 縦位置, 文字)

    # (2)作成日
    font_size = 10
    pdf_canvas.setFont('HeiseiKakuGo-W5', font_size)
    pdf_canvas.drawString(285, 770,  '    年         月         日')

    # (6)学歴・職歴
    data = [
            ['        項目', '   時間', '                                            金額'],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' '],
            [' ', ' ', ' ']
        ]
    table = Table(data, colWidths=(25*mm, 14*mm, 121*mm), rowHeights=7.5*mm)
    table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'HeiseiKakuGo-W5', 11),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
    table.wrapOn(pdf_canvas, 20*mm, 20*mm)
    table.drawOn(pdf_canvas, 20*mm, 20*mm)
    # 1枚目終了
    pdf_canvas.showPage()
    print("over")


# 作成
make()