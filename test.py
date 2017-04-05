# from reportlab.pdfbase import pdfmetrics
# from reportlab.pdfbase.ttfonts import TTFont
# pdfmetrics.registerFont(TTFont('msyh', 'msyh.ttf'))
# from reportlab.lib.units import mm
# from reportlab.lib.pagesizes import A7,A4,A6
# from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer,Image,Table,TableStyle,BaseDocTemplate
#
# def rpt():
#     story=[]
#     # img = Image('1.png')
#     # img.drawHeight = 190/4
#     # img.drawWidth = 144/4
#     img=''
#
#     #表格数据：用法详见reportlab-userguide.pdf中chapter 7 Table
#     component_data= [['IT设备标识牌', '', '', ''],
#                      ['','','',img],
#                      ['1:1', '1:1', '1:2', '1:3'],
#                      ['2:0', '2:1', '2:2', '2:3'],
#                      ['3:0', '3:1', '3:2', '3:3'],
#                      ['4:0', '4:1', '4:2', '4:3'],
#     ]
#     #创建表格对象，并设定各列宽度
#     component_table = Table(component_data,rowHeights=[5*mm,None,6*mm,6*mm,6*mm,6*mm])
#     #添加表格样式
#     component_table.setStyle(TableStyle([
#         ('FONTNAME', (0, 0), (-1, -1), 'msyh'),  # 字体
#         ('FONTSIZE', (0, 0), (-1, -1), 8),  # 字体大小
#         ('SPAN', (0, 0), (-1, 0)),  # 合并第一行
#         ('SPAN', (0, 1), (-2, 1)),
#         ('ALIGN', (0, 0), (0, 0), 'CENTER'),  # 第一行水平居中对齐
#         ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 垂直居中对齐
#         ('GRID', (0, 0), (-1, -1), 0.5, 'BLACK'),  # 设置表格框线为黑色，线宽为0.5
#     ]))
#     for i in range(3):
#         story.append(component_table)
#
#
#     doc = SimpleDocTemplate('out.pdf',pagesize=A4)
#     doc._topMargin=1*mm
#     doc.build(story)
#
# if __name__ == '__main__':
#     rpt()
#     print('ok')
"""
Example how to adjust the Frame-size in a BaseDocTemplate
"""
from reportlab.platypus import BaseDocTemplate, Frame, Paragraph,\
    NextPageTemplate, PageBreak, PageTemplate,Table,TableStyle
from reportlab.lib.units import inch,mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A7
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('msyh', 'msyh.ttf'))

doc = BaseDocTemplate('out.pdf',showBoundary=0,pagesize=A7,leftMargin=0,rightMargin=0,topMargin=0,bottomMargin=0)
frame_pages = Frame(0, 0, doc.width, doc.height, id='remaining', leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0)
doc.addPageTemplates([PageTemplate(id='table', frames=frame_pages)])

Elements=[]

data = [['IT设备标识牌','','','','',''],
        ['','','','','',''],
        ['设备型号','1111','','设备状态','2222',''],
        ['设备用途','3333','','使用年限','4444',''],
        ['责任人','5555','','联系电话','6666','']
]
colwidth = [doc.width/6]*6
rowheight = [5*mm,30*mm]+[(doc.height-35*mm)/3]*3
t = Table(data,colWidths=colwidth,rowHeights=rowheight)
t.setStyle(TableStyle([('GRID',(0,0),(-1,-1),0.5,'BLACK'),
                       ('FONTNAME', (0, 0), (-1, -1), 'msyh'),
                       ('ALIGN',(0,0),(-1,0),'CENTER'),
                       ('VALIGN',(0,0),(-1,0),'TOP'),
                       ('ALIGN',(0,1),(-1,-1),'LEFT'),
                       ('SPAN',(0,0),(-1,0)),
                       ('SPAN',(0,1),(3,1)),
                       ('SPAN',(-2,1),(-1,1)),
                       ('SPAN',(1,2),(2,2)),
                       ('SPAN',(-1,2),(-2,2)),
                       ('SPAN',(1,3),(2,3)),
                       ('SPAN',(-1,3),(-2,3)),
                       ('SPAN',(1,4),(2,4)),
                       ('SPAN',(-1,4),(-2,4))
                      ]))

# Elements.append(Paragraph("Frame first page!",styles['Normal']))
# Elements.append(NextPageTemplate('remaining_pages'))  #This will load the next PageTemplate with the adjusted Frame.
for i in range(5):
    Elements.append(t)
    Elements.append(PageBreak()) # This will force a page break so you are guarented to get the next PageTemplate/Frame

# Elements.append(Paragraph("Frame remaining pages!,  "*500,styles['Normal']))

#start the construction of the pdf
doc.build(Elements)