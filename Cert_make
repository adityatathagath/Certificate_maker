from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
import xlrd




    # to open file and read
def make_certi(namee):
    # to open image
    img=Image.open('C:/Users/aditw/Desktop/Python_Certificate/Certificate-sender-master/workshop.jpg')
    # pen for writing on image
    pen=ImageDraw.Draw(img)
    # font to wrt
    fnt=ImageFont.truetype('C:/Users/aditw/Desktop/Python_Certificate/Certificate-sender-master/a.ttf',160)
    # name to be written on certificate
    name="r"

    # size of name in pixels
    req_size=pen.textsize(text=namee,font=fnt)

    # place on certificate on which it is to be written
    place=(1847,1499)

    # writing name on certificate
    pen.text(xy=(place[0]-req_size[0]/2,place[1]-req_size[1]),text=namee,fill=(0,0,0),font=fnt)

    # saving the new image
    img.save('C:/Users/aditw/Desktop/Python_Certificate/Certificate-sender-master/cert%s.jpg'%namee)

if __name__ == "__main__":

    Book = xlrd.open_workbook('C:/Users/aditw/Desktop/Python_Certificate/Certificate-sender-master/Book1.xlsx')
    WorkSheet = Book.sheet_by_name('Sheet1')

    num_row = WorkSheet.nrows - 1
    row = 0

    while row < num_row:
        row += 1

        nameee = WorkSheet.cell_value(row, 0)

        namee=make_certi(nameee)
