#import libraries
import numpy as np
import pandas as pd
from openpyxl import Workbook

# function return volume of sphere 
def volumeSphere (Sphere):
    return 4/3*np.pi*Sphere[2]**3

# function return volume of cube
def volumeCube (cube):
    return cube[2]**3

# function return volume of box
def Box_volume (box):
    return box[0]*box[1]*box[2]

# function return no.of object per box 
def no_of_objectsPerBox(box_volume,object_volume):
    return int(box_volume/object_volume)

# function return no.of boxes required
def no_of_boxes (object,no_of_objectperBox):
    if (no_of_objectperBox == 0 ):
        return 0
    else:
        return round(int(object[3])/no_of_objectperBox+0.5)

#Box size available
Box_size = [200,400,500]

#Reading data from robot that saved in dataset sheet
df = pd.read_excel('dataset.xlsx')
Data_From_Robot = []
book = Workbook()
sheet = book.active

# first row in the new worksheet that provide all information
rows = ['shape','color','Dimension cm','Number of objects','volume of Shape (cm3)','Volume of box (cm3)','Maximum number of objects in one box','Total number.of box']
sheet.append(rows)

# count the total no. of box required 
count = 0

# Reading file excel and process data to get the total no. of boxes
for index, row in df.iterrows():
    Data_From_Robot.append(row.to_list())
for i in range(len(Data_From_Robot)):
    if Data_From_Robot[i][0] == 'Box':
        volumeofobject = volumeCube(Data_From_Robot[i])
    else : 
        volumeofobject = volumeSphere(Data_From_Robot[i])
    volumeofbox = Box_volume(Box_size)
    objectsPerBox = no_of_objectsPerBox(volumeofbox,volumeofobject)
    TotalNoOfBoxes = no_of_boxes(Data_From_Robot[i],objectsPerBox)
    count +=TotalNoOfBoxes
    rows = (Data_From_Robot[i][0],Data_From_Robot[i][1],Data_From_Robot[i][2],Data_From_Robot[i][3],volumeofobject,volumeofbox,objectsPerBox,TotalNoOfBoxes)
    sheet.append(rows)

# Save data in new file called result.xlsx
book.save("Result.xlsx")


print("The total number of boxes required for all objects is "+ str(count) + " box")
print("You can check the details in excel's file called Result.xslx")
