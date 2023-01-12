import os
from docx2pdf import convert
source ='C:/ePDM_Elcam/test/FE_Test'
destintion='C:/Users/matan.s/Desktop/python/add-in_clone'
files=os.listdir(destintion)
print(files)
with open(destintion+'/test.txt','w') as f:
    f.write('test')
for file in files:
    if file.endswith('.docx'):
        convert(destintion+'/'+file, destintion+'/'+file[:-5]+'.pdf')
        print('done')
        
