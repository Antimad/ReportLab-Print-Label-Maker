from reportlab.lib.pagesizes import A4
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics import renderPDF
from reportlab.pdfgen.canvas import Canvas
import pandas as pd
import math

PAGESIZE = A4
NUM_LABELS_X = 5
NUM_LABELS_Y = 10
# BAR_WIDTH = 1.5
# BAR_HEIGHT = 51.0
ItemCode_Y = 80
ItemPrice_Y = 50

# BARCODE_Y = 17

LABEL_WIDTH = PAGESIZE[0] / NUM_LABELS_X
LABEL_HEIGHT = PAGESIZE[1] / NUM_LABELS_Y
SHEET_TOP = PAGESIZE[1]

NOLA = pd.read_excel('VA - NOLA.xlsx', skiprows=5)
Manifest = pd.read_excel('VA - NOLA.xlsx')
ShipmentNumber = Manifest['MANIFEST'][2].replace('#', '')
ContainerNumber = Manifest['Unnamed: 3'][1].replace('CONT:', '')

ShipmentNumber_Y = 60
ContainerNumber_Y = 50
ItemDescription_Y = 40


def front_label(code: str, price: str) -> Drawing:
    item_code = String(0, ItemCode_Y, code, fontName="Helvetica", fontSize=10, textAnchor="middle")
    item_code.x = LABEL_WIDTH / 2    # Centers the Text

    item_description = String(1, ItemPrice_Y, price, fontName="Helvetica", fontSize=20, textAnchor="middle")
    item_description.x = LABEL_WIDTH / 2
    """
    barcode = Ean13BarcodeWidget(ean13)
    barcode.barWidth = BAR_WIDTH
    barcode.barHeight = BAR_HEIGHT
    x0, y0, bw, bh = barcode.getBounds()
    barcode.x = (LABEL_WIDTH - bw) / 2
    barcode.y = BARCODE_Y

    """
    label_drawing = Drawing(LABEL_WIDTH, LABEL_HEIGHT)
    label_drawing.add(item_code)
    label_drawing.add(item_description)
    #  label_drawing.add(barcode)
    return label_drawing


def back_label(code: str, vessel: str, cont: str, desc: str) -> Drawing:
    item_code = String(0, ItemCode_Y - 10, code, fontName="Helvetica-Bold", fontSize=12, textAnchor="middle")
    item_code.x = LABEL_WIDTH / 2    # Centers the Text

    vessel_code = String(1, ShipmentNumber_Y, vessel, fontName="Helvetica", fontSize=10, textAnchor="middle")
    vessel_code.x = LABEL_WIDTH / 2

    container_code = String(0, ContainerNumber_Y, cont, fontName='Helvetica', fontSize=10, textAnchor='middle')
    container_code.x = LABEL_WIDTH / 2

    description = String(0, ItemDescription_Y, desc, fontName='Helvetica', fontSize=10, textAnchor='middle')
    description.x = LABEL_WIDTH / 2

    label_drawing = Drawing(LABEL_WIDTH, LABEL_HEIGHT)
    label_drawing.add(item_code)
    label_drawing.add(vessel_code)
    label_drawing.add(container_code)
    label_drawing.add(description)
    #  label_drawing.add(barcode)
    return label_drawing


"""
def fill_sheet(canvas: Canvas, label_drawing: Drawing, items, prices):
    count = 0
    for u in range(0, NUM_LABELS_Y):
        for i in range(0, NUM_LABELS_X):
            item = items[count]
            price = prices[count]
            count += 1
            x = i*LABEL_WIDTH
            y = SHEET_TOP - LABEL_HEIGHT - u * LABEL_HEIGHT
            renderPDF.draw(label_drawing, canvas, x, y)

"""


front_canvas = Canvas("FrontTag.pdf", pagesize=PAGESIZE)
back_canvas = Canvas("BackTag.pdf", pagesize=PAGESIZE)
# sticker = label(NOLA['ITEM'], NOLA['PRICE'])
count = 0
qty = list(NOLA['QTY'])

for pages in range(0, math.ceil(10)):
    for u in range(0, NUM_LABELS_Y):
        for i in range(0, NUM_LABELS_X):
            try:
                front_sticker = front_label(NOLA['ITEM'][count], '$' + str(NOLA['PRICE'][count]))
                back_sticker = back_label(NOLA['ITEM'][count], ShipmentNumber, ContainerNumber,
                                          NOLA['DESCRIPTION'][count].replace('W/N', '')[0:15])
            except KeyError:
                break
            x = i * LABEL_WIDTH
            y = SHEET_TOP - LABEL_HEIGHT - u * LABEL_HEIGHT - 15
            renderPDF.draw(front_sticker, front_canvas, x, y)
            renderPDF.draw(back_sticker, back_canvas, x, y)
            if qty[count] > 1:
                qty[count] -= 1
            else:
                count += 1
    front_canvas.showPage()
    back_canvas.showPage()

front_canvas.save()
back_canvas.save()
