from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A7
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from reportlab.platypus import Image
from openpyxl import load_workbook
import datetime
pdfmetrics.registerFont(TTFont('Vera', 'wqy-zenhei.ttc'))


class AssetTags():
    def __init__(self, excel_file, index_of_header, out_pdf_file, image_=None):
        self.__button0y = 0.3 * mm
        self.__button1y = 5 * mm
        self.__button2y = 10 * mm
        self.__button3y = 15 * mm
        self.__button4y = 20 * mm
        self.__button5y = 69 * mm
        self.__button6y = 74 * mm
        self.__left0x = 0.3 * mm
        self.__left1x = 12.5 * mm
        self.__left2x = 52.5 * mm
        self.__left3x = 65 * mm
        self.__left4x = 104 * mm

        self.__image = image_
        self.__header, self.__datas = self.__loadfile(excel_file, index_of_header)
        self.__c = canvas.Canvas(out_pdf_file, pagesize=A7)

    @staticmethod
    def __loadfile(excel_, header_at_row):
        """参数：excel文件；表头所在的行
           返回值：表头内容（列表）；excel内容（字典）
        """
        wb = load_workbook(excel_, read_only=True)
        st = wb.active
        header = []
        data={}

        for i in range(1, st.max_column + 1):
            header.append(st.cell(row=header_at_row, column=i).value)

        for i, row in zip(range(0, st.max_row), st.rows):
            temp_dict = {}
            for j, cell in zip(range(0, len(header)), row):
                if isinstance(cell.value,datetime.datetime):
                    temp_dict[header[j]] = str(cell.value)[:-8]
                else:
                    temp_dict[header[j]] = str(cell.value)
            data[i] = temp_dict
        return header, data

    def __set_figure(self, const_text):
        """此函数用来绘制框架"""

        const_text = [str(i) for i in const_text]
        self.__c.setFont('Vera', 12)

        # 绘制边界大方框
        self.__c.rect(0.1 * mm, 0.1 * mm, 104 * mm, 74 * mm, fill=0)

        # 加载图片
        if self.__image is not None:
            a = Image(self.__image, width=38 * mm, height=48 * mm)
            a.drawOn(self.__c, self.__left3x, self.__button4y + .5 * mm)

        # 绘制三个黄色方框
        self.__c.setFillColorRGB(150, 150, 0)
        self.__c.setStrokeColorRGB(150, 150, 0)
        self.__c.rect(self.__left0x, self.__button0y + 0.1 * mm, 12.5 * mm, 20 * mm, stroke=0, fill=1)
        self.__c.rect(self.__left2x, self.__button0y, 12.5 * mm, 20 * mm, stroke=0, fill=1)
        self.__c.rect(self.__left0x, self.__button5y, 103.6 * mm, 5 * mm, stroke=0, fill=1)

        # 绘制中间三条竖线
        self.__c.setFillColorRGB(0, 0, 0)
        self.__c.setStrokeColorRGB(0, 0, 0)
        self.__c.line(self.__left1x, self.__button0y, self.__left1x, self.__button4y)
        self.__c.line(self.__left2x, self.__button0y, self.__left2x, self.__button4y)
        self.__c.line(self.__left3x, self.__button0y, self.__left3x, self.__button5y)

        # 绘制中间五条横线
        self.__c.line(self.__left0x, self.__button1y, self.__left4x, self.__button1y)
        self.__c.line(self.__left0x, self.__button2y, self.__left4x, self.__button2y)
        self.__c.line(self.__left0x, self.__button3y, self.__left4x, self.__button3y)
        self.__c.line(self.__left0x, self.__button4y, self.__left4x, self.__button4y)
        self.__c.line(self.__left0x, self.__button5y, self.__left4x, self.__button5y)

        self.__c.drawCentredString(50 * mm, self.__button5y + mm, 'IT设备标识牌')
        self.__c.setFont('Vera', 8)
        textleftx = 6 * mm
        textleft2x = 59 * mm

        self.__c.drawCentredString(textleftx, 16 * mm, const_text[0])   # 左1
        self.__c.drawCentredString(textleft2x, 16 * mm, const_text[1])  # 右1
        self.__c.drawCentredString(textleftx, 11 * mm, const_text[2])   # 左2
        self.__c.drawCentredString(textleft2x, 11 * mm, const_text[3])  # 右2
        self.__c.drawCentredString(textleftx, 6 * mm, const_text[4])    # 左3
        self.__c.drawCentredString(textleft2x, 6 * mm, const_text[5])   # 右3
        self.__c.drawCentredString(textleftx, mm, const_text[6])        # 左4
        self.__c.drawCentredString(textleft2x, mm, const_text[7])       # 右4

    def __set_text(self, header_, s_):
        s =s_
        self.__c.setFont('Vera', 8)
        self.__c.drawString(self.__left1x + mm, self.__button3y + mm, s[header_[0]])
        self.__c.drawString(self.__left3x + mm, self.__button3y + mm, s[header_[1]])
        self.__c.drawString(self.__left1x + mm, self.__button2y + mm, s[header_[2]])
        self.__c.drawString(self.__left3x + mm, self.__button2y + mm, s[header_[3]])
        self.__c.drawString(self.__left1x + mm, self.__button1y + mm, s[header_[4]])
        self.__c.drawString(self.__left3x + mm, self.__button1y + mm, s[header_[5]])
        self.__c.drawString(self.__left1x + mm, self.__button0y + mm, s[header_[6]])
        self.__c.drawString(self.__left3x + mm, self.__button0y + mm, s[header_[7]])

    def create_page(self, header=None):
        if header is None:
            header=self.__header

        for i,d in enumerate(self.__datas):
            #因为是字典，所以d的值为与i的值一样

            self.__set_figure(header)
            self.__set_text(header, self.__datas[i])
            self.__c.showPage()

    def save(self):
        self.__c.save()

if __name__ == '__main__':

    my_custom = ['资产名称','购置日期','资产编码','所在区域','资产状态','保管人员','供应商','单价']

    pdf = AssetTags('0.xlsx', 2, 'pdf.pdf', '1.png')

    pdf.create_page(my_custom)

    pdf.save()
