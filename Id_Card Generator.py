#importing necessary modules
from PIL import Image, ImageDraw, ImageFont, ImageOps
import qrcode
import xlrd


count = int(input("Total number of employees: "))#specify the number of cantidates
wb = xlrd.open_workbook("Employees List.xlsx")#loading the excel sheet
sheet = wb.sheet_by_index(0)
i=1
print("\nStarting ID Card Generation...\n")

for counter in range(0,count):
    
    image1 = Image.new('RGB', (500,700), (255, 255, 255))#creating a plane image
    font = ImageFont.truetype('arial.ttf', size =20 )#you can use other fonts (calibre for example), but make sure you have it installed on your pc
    image = ImageOps.expand(image1,border=10,fill='brown')#adding frame
    write = ImageDraw.Draw(image)
    
    name = sheet.col_values(0,start_rowx=counter+1,end_rowx=counter+2)#reading name of cantidate(0 indicates the first row of the sheet, names are provided in the first row)
    gender = sheet.col_values(1,start_rowx=counter+1,end_rowx=counter+2)#reading gender details from the 2nd row
    dob=sheet.col_values(2,start_rowx=counter+1,end_rowx=counter+2)
    blood_group = sheet.col_values(5,start_rowx=counter+1,end_rowx=counter+2)
    phone=sheet.col_values(3,start_rowx=counter+1,end_rowx=counter+2)
    address=sheet.col_values(4,start_rowx=counter+1,end_rowx=counter+2)

    color = 'rgb(64,64,64)'
    company = "WebWare"#adding name of the company
    write.text((202,35), company, fill=color, font=ImageFont.truetype('arial.ttf', size =35))


    imgg = Image.open(str(name[0]) + ".jpg")#opening image of the cantidate, make sure it is kept in the same folder along with the excel sheet and python script, also the filename must match with the name provided in the sheet
    pic1 = imgg.resize((200, 230), Image.ANTIALIAS)
    pic = ImageOps.expand(pic1,border=4,fill='gray')
    image.paste(pic, (175,100))

    
    
    print(str(i) + "." + name[0])                                   
    color = 'rgb(0,0,0)'
    write.text((50,370), "Name:   " + str(name[0]), fill=color, font=font)#writing name into the image

                                   
    color = 'rgb(0,0,0)'
    write.text((50,410), ("Gender: "+str(gender[0])), fill=color, font=font)#writing gender
                                       

                                      
    color = 'rgb(0,0,0)'
    write.text((50,450), ("DOB:     "+str(dob[0])), fill=color, font=font)#DOB
    

                                   
    color = 'rgb(0,0,0)'
    write.text((50,490), ("Blood Group: "+str(blood_group[0])), fill=color, font=font)#blood group

                                 
    color = 'rgb(0,0,0)'
    write.text((50,530), ("Contact: "+str(phone[0])), fill=color, font=font)#contact


    color = 'rgb(0,0,0)'
    write.text((50,570), ("Address: "+str(address[0])), fill=color, font=font)#address
                                       
        
    qrc = qrcode.make(str(name[0])+ ", " + str(phone[0]) + ", " + str(address[0]))#generating qrcode with details like name, contact number and address
    qrcod = qrc.resize((100, 100), Image.ANTIALIAS)
    image.paste(qrcod, (380, 610))
    image.save(str(name[0]) + ".png")#saving the final image with the cantidate's name as filename
    i+=1
#repeating the process for all the cantidates
print("\nID Cards Successfully Generated!!")
