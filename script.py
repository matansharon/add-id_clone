import os
from docx2pdf import convert
source ='C:/ePDM_Elcam/test/FE_Test'
destintion='C:/Users/matan.s/Desktop/python/add-in_clone'
files=os.listdir(destintion)
print(files)
with open('test.text',"w")as file:
    file.write(str(files))