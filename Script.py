from PIL import Image, ImageDraw, ImageFont
import pandas as pd
form = pd.read_excel("spreadsheet.xlsx") #location of the excel file
list = form['NAME'].to_list() # Name Column
list1 = form['EVENT'].to_list() # Event Column
for i in list:
    im = Image.open('Certificate.png') # blank certificate image
    imgh= im.height
    w=(im.width)
    d = ImageDraw.Draw(im)
    font = ImageFont.truetype("SF-Pro-Display-Medium.otf", 52) # font file and size for writing

    t= d.textsize(i,font=font)[0]
    location = (
        ((w - t)/2)-292, 670) # location of name placeholder
    text_color = (0, 0, 0)
    d.text(location,i , fill=text_color,font=font)
    fill_color = "White"  # background color
    for n in list1:
        t = d.textsize(n, font=font)[0]
        location = (
            ((w - t) / 2) - 292, 834) # location of event placeholder
        text_color = (0, 0, 0)
        d.text(location, n, fill=text_color, font=font)
        fill_color = "White"  # background color 
    if im.mode in ('RGBA'):
        background = Image.new(im.mode[:-1], im.size, fill_color)
        background.paste(im, im.split()[-1])
        im = background

    im.save("certificate_"+i+".pdf") # certificate saved as certificate_name.pdf
