import pandas as pd
from PyPDF2 import PdfFileWriter, PdfFileReader
import io
import reportlab
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

x=input("Enter File Name containing data: ")
t=input("Enter sheet name:")
f=input("Enter Template pdf Name: ")
xc=int(input("Enter x cordinate: "))
yc=int(input("Enter y cordinate: "))
size=int(input("Enter text size: "))

try:
    w=pd.read_excel(x,t)
except:
    w=pd.rea
q=w["Full Name"].dropna()
for x in q:
    x.upper()
    existing_pdf = PdfFileReader(open(f, "rb"))
    m=existing_pdf.getPageLayout()
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=m)
    can.setFont("Times-Roman", size)
    can.drawCentredString(xc,yc, x)
    can.showPage()
    can.save()
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    output = PdfFileWriter()
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    w=x+".pdf"
    output.addPage(page)
    outputStream = open(w, "wb")
    output.write(outputStream)
    outputStream.close()
    print("Running")
print("done")