from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('msyh', 'msyh.ttf'))
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A7,A4,A6
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer,Image,Table,TableStyle

def rpt():
    story=[]
    # img = Image('1.png')
    # img.drawHeight = 190/4
    # img.drawWidth = 144/4
    img=''

    #表格数据：用法详见reportlab-userguide.pdf中chapter 7 Table
    component_data= [['IT设备标识牌', '', '', ''],
                     ['','','',img],
                     ['1:1', '1:1', '1:2', '1:3'],
                     ['2:0', '2:1', '2:2', '2:3'],
                     ['3:0', '3:1', '3:2', '3:3'],
                     ['4:0', '4:1', '4:2', '4:3'],
    ]
    #创建表格对象，并设定各列宽度
    component_table = Table(component_data,rowHeights=[5*mm,None,6*mm,6*mm,6*mm,6*mm])
    #添加表格样式
    component_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'msyh'),  # 字体
        ('FONTSIZE', (0, 0), (-1, -1), 8),  # 字体大小
        ('SPAN', (0, 0), (-1, 0)),  # 合并第一行
        ('SPAN', (0, 1), (-2, 1)),
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),  # 第一行水平居中对齐
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # 垂直居中对齐
        ('GRID', (0, 0), (-1, -1), 0.5, 'BLACK'),  # 设置表格框线为黑色，线宽为0.5
    ]))
    for i in range(3):
        story.append(component_table)


    doc = SimpleDocTemplate('out.pdf',pagesize=A7, leftMargin=5, rightMargin=5,topMagrin=5,bottomMargin=5)
    doc.build(story)

if __name__ == '__main__':
    rpt()
    print('ok')