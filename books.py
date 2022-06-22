import pandas as pd
import os
import qrcode
import fpdf
from bidi.algorithm import get_display
import arabic_reshaper
from warnings import filterwarnings
import zipfile


current_directory = os.path.abspath(os.path.dirname('books.xlsx'))
qrcodes = os.path.join(current_directory, 'qrcodes')
covers = os.path.join(current_directory, 'covers')

df = pd.read_excel('books.xlsx', sheet_name='books')
if not (os.path.exists(qrcodes) and os.path.exists(covers)):
    os.makedirs(qrcodes)
    os.makedirs(covers)

count = len(df['الرواية'])

for i in range(0, count):
    book_name = df['الرواية'][i]
    author_name = df['المؤلف'][i]
    data = df['صفحة_الرواية'][i]
    img = qrcode.make(data)
    img_name = f'{qrcodes}/{book_name}.png'
    img.save(img_name)
    pdf = fpdf.FPDF(format='letter')
    pdf.add_page()
    pdf.add_font('traditional arabic', '', 'trado.ttf', uni=True)
    pdf.set_font('traditional arabic', '', 43)
    pdf.image(img_name, x=47.5, y=67, w=120, h=120, type='PNG', link=data)
    pdf.cell(0, 170, ln=2, align='C')
    pdf.cell(0, 30, txt=get_display(
        arabic_reshaper.reshape(book_name)), ln=2, align='C')
    pdf.cell(0, 30, txt=get_display(
        arabic_reshaper.reshape(author_name)), ln=2, align='C')
    filterwarnings('ignore')
    pdf.output(f'{covers}/{book_name}.pdf', dest='F')
    filterwarnings('default')


def zipdir(path, ziph):
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file), os.path.relpath(
                os.path.join(root, file), os.path.join(path, '..')))


with zipfile.ZipFile('covers.zip', 'w', zipfile.ZIP_DEFLATED) as zipf:
    zipdir(covers, zipf)
