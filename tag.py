from reportlab.platypus import BaseDocTemplate, Frame, Paragraph,\
    NextPageTemplate, PageBreak, PageTemplate,Table,TableStyle
from reportlab.lib.units import inch,mm,cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('msyh', 'msyh.ttf'))


class Tag():
    def __init__(self,out_file='out.pdf',pagesize=(80*mm,80*mm)):
        print('Class AssetTag:init')
        self.doc = BaseDocTemplate(out_file,showBoundary=0,pagesize=pagesize,leftMargin=0,rightMargin=0,topMargin=0,bottomMargin=0)
        self.frame_pages = Frame(0, 0, self.doc.width, self.doc.height, id='remaining', leftPadding=0, bottomPadding=0, rightPadding=0, topPadding=0)
        self.doc.addPageTemplates([PageTemplate(id='table', frames=self.frame_pages)])

        self.Elements=[]
        self.arg_custom_header=None
        self.arg_data=None
        self.data =None
        self.count_of_column = 6
        self.colwidth = [self.doc.width/self.count_of_column]*self.count_of_column

        # [行1高度，行2高度]+[(文件总高度-前两行高度)/剩余的n行]×n
        self.count_of_row = 6  # 表格的总行数
        self.n = self.count_of_row-2
        self.row_one_height = 15
        self.row_two_height = self.doc.height*0.45
        self.rowheight = [self.row_one_height,self.row_two_height]+ [(self.doc.height-self.row_two_height-self.row_one_height) / self.n] * self.n

# Elements.append(Paragraph("Frame first page!",styles['Normal']))
# Elements.append(NextPageTemplate('remaining_pages'))  #This will load the next PageTemplate with the adjusted Frame.
    def addpage(self, _header, _data,_fontsize):

        self.data = [['IT设备标识牌','','','','',''],
                ['','','','','',''],
                ['{}'.format(_header[0]), '{}'.format(_data[_header[0]]), '',
                 '{}'.format(_header[1]), '{}'.format(_data[_header[1]]), ''],
                ['{}'.format(_header[2]), '{}'.format(_data[_header[2]]), '',
                 '{}'.format(_header[3]), '{}'.format(_data[_header[3]]), ''],
                ['{}'.format(_header[4]), '{}'.format(_data[_header[4]]), '',
                 '{}'.format(_header[5]), '{}'.format(_data[_header[5]]), ''],
                ['{}'.format(_header[6]), '{}'.format(_data[_header[6]]), '',
                 '{}'.format(_header[7]), '{}'.format(_data[_header[7]]), '']
                ]

        self.t = Table(data=self.data, colWidths=self.colwidth, rowHeights=self.rowheight)
        self.t.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 0.5, 'BLACK'),
                                    ('FONTNAME', (0, 0), (-1, -1), 'msyh'),
                                    ('FONTSIZE', (0, 1), (-1, -1), _fontsize),
                                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                                    ('VALIGN', (0, 0), (-1, 0), 'TOP'),
                                    ('ALIGN', (0, 1), (-1, -1), 'RIGHT'),
                                    ('SPAN', (0, 0), (-1, 0)),  # 合并第一行
                                    ('SPAN', (0, 1), (-1, 1)),  # 合并第二行
                                    ('SPAN', (-2, 1), (-1, 1)),  # 合并第二行的最后两列
                                    ('SPAN', (1, 2), (2, 2)),
                                    ('SPAN', (-1, 2), (-2, 2)),
                                    ('SPAN', (1, 3), (2, 3)),
                                    ('SPAN', (-1, 3), (-2, 3)),
                                    ('SPAN', (1, 4), (2, 4)),
                                    ('SPAN', (-1, 4), (-2, 4)),
                                    ('SPAN', (1, 5), (2, 5)),
                                    ('SPAN', (-1, 5), (-2, 5))
                                    ]))

        self.Elements.append(self.t)
        self.Elements.append(PageBreak()) # This will force a page break so you are guarented to get the next PageTemplate/Frame

# Elements.append(Paragraph("Frame remaining pages!,  "*500,styles['Normal']))

# start the construction of the pdf
    def build(self):
        self.doc.build(self.Elements)