# -*- coding: utf-8 -*-
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
import openpyxl

def read_excel(title = "インプットサンプル社員1(1).xlsx"):
    wb = openpyxl.load_workbook(title)
    ws = wb["社員"]
    sdata = {}
    # cはセル番地でA1のセルを取得する

    for row in ws.iter_rows():
        key = ''
        for cell in row:
            if cell.column == 1:
                key = cell.value
                sdata[key] = []
            else:
                value = str(cell.value)
                if "." in value:
                    value =format(int(cell.value),",")
                sdata[key].append(value)
    return sdata


# 初期設定
def make(filename="resume", sdata= '', index =''): # ファイル名
    pdf_canvas = set_info(filename) # キャンバス名
    print_string(pdf_canvas, sdata, index)
    pdf_canvas.save() # 保存

def set_info(filename):
    pdf_canvas = canvas.Canvas("./{0}.pdf".format(filename)) # 保存先
    pdf_canvas.setAuthor("") # 作者
    pdf_canvas.setTitle("") # 表題
    pdf_canvas.setSubject("") # 件名

    dx = 150*mm
    dy = 180*mm 
    dWidth = 40*mm
    dHeight = 10*mm

    # 画像挿入(画像パス、始点x、始点y、幅、高さ)
    pdf_canvas.drawImage("logo.png", dx, dy, dWidth, dHeight)

    return pdf_canvas

#履歴書フォーマット作成
def print_string(pdf_canvas, sdata, index):
    pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5')) # フォント

    font_size = 15 # フォントサイズ
    pdf_canvas.setFont('HeiseiKakuGo-W5', font_size)
    pdf_canvas.drawString(400, 780, '給与支払明細書') # 書き出し(横位置, 縦位置, 文字)

    #(1) 名前、個人情報
    data = [
            [' 支給年月日 ', ' 社員コード ', '   氏 名   '],
            [sdata['支給年月日'][index], sdata['社員コード'][index],  sdata['氏名'][index]]
        ]
    table = Table(data, colWidths=(30*mm, 30*mm, 25*mm), rowHeights=(7.5*mm, 12*mm))
    table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'HeiseiKakuGo-W5', 11),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
        ]))
    table.wrapOn(pdf_canvas, 20*mm, 20*mm)
    table.drawOn(pdf_canvas, 20*mm, 260*mm)


    #(2) 差引控除金額 
    data = [
            [' 差引控除金額 '],
            [sdata['差引支給金額'][index]]
        ]
    table = Table(data, colWidths=50*mm, rowHeights=6*mm)
    table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'HeiseiKakuGo-W5', 11),
            ('BOX', (0, 0), (-1, -1), 2, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ]))
    table.wrapOn(pdf_canvas, 20*mm, 20*mm)
    table.drawOn(pdf_canvas, 140*mm, 260*mm)


    #(3) 名前、個人情報
    data = [
            ['基\n本\n支\n給', '基本給','役職手当','職能手当','住宅手当','家族手当','残業時間', '残業手当','皆勤手当'],
            ['', sdata['基本給'][index],  sdata['役職手当'][index], sdata['職能手当'][index],sdata['住宅手当'][index],sdata['家族手当'][index],sdata['残業時間'][index],sdata['残業手当'][index],sdata['皆勤手当'][index]],
            ['', '','食事手当', '通勤手当', '', '' , '', '調整手当', '総支給金額'],
            ['', '', sdata['食事手当'][index], sdata['通勤手当'][index],'', '', '', sdata['調整手当'][index],sdata['総支給金額'][index],],
            ['控\n除\n項\n目', '*健康保険*','厚生年金', '雇用保険', '介護保険', '所得税', '住民税', '旅行会費',''],
            ['', sdata['健康保険'][index],sdata['厚生年金'][index],sdata['雇用保険'][index],sdata['介護保険'][index],sdata['所得税'][index],sdata['住民税'][index],sdata['旅行会費'][index],''],
            ['', '','', '', '', '' , '', '', '総控除金額'],
            ['', '','', '', '', '' , '', '',sdata['総控除金額'][index]],

        ]
    table = Table(data, colWidths=19*mm, rowHeights=7.5*mm)
    table.setStyle(TableStyle([
            ('FONT', (0, 0), (-1, -1), 'HeiseiKakuGo-W5', 10),
            ('BOX', (0, 0), (-1, -1), 1, colors.black),
            ('INNERGRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ('SPAN', (0, 0), (0, 3)),
            ('SPAN', (0, 4), (0, 7)),

        ]))
    table.wrapOn(pdf_canvas, 20*mm, 20*mm)
    table.drawOn(pdf_canvas, 20*mm, 190*mm)   
    # 1枚目終了
    pdf_canvas.showPage()
    print(read_excel())


# 作成
sdata = read_excel()
for index,key in enumerate(sdata['氏名']):
    make(filename='./salary/'+ key, sdata=sdata, index=index)