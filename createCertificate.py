import PIL
from PIL import Image, ImageFont, ImageDraw  
import xlrd 

workbook = xlrd.open_workbook('covid_quiz.xlsx')
sheet = workbook.sheet_by_index(0)
fnt = ImageFont.truetype('montserrat.ttf', 68)

for i in range(1,155) :
  im = Image.open("certificate.png")
  certificateDraw = ImageDraw.Draw(im)
  name= str(sheet.cell(i, 2).value)
  certificateDraw.text((0,830),name.center(95), font=fnt, fill =(252,248,118,255))
  im.save("certificates/"+name+".png")
print("completed")

