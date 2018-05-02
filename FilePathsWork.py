import string,os,xlwt

while True:
    foldques = raw_input("Please input the full path to the folder that you want mapped.")
    if os.path.isdir(foldques):
        print (foldques +" will have its contents mapped.")
        path  = foldques
        break
    else:
        print("That is not a valid folder on this computer, please provide a valid input.")

wb =xlwt.Workbook()
ws = wb.add_sheet(os.path.basename(foldques)+ "  File Structure")
ws.write(0,16,path)

foldstyle = "font: bold on, color-index red"
filestyle = "font: bold off, color-index green"
foldstyle = xlwt.easyxf(foldstyle)
filestyle = xlwt.easyxf(filestyle)
ws.write(1,16,"RED, BOLD TEXT DENOTES A FOLDER",foldstyle)
ws.write(2,16,"Green, un-bolded text denotes a file",filestyle)
rowcount = 1
total = 0
for root, dirs, files in os.walk(path):
    path = root.count("\\")-2
    ws.write(rowcount,path,os.path.basename(root),foldstyle)
    rowcount+=1
    total+=1
    for filey in files:
        ws.write(rowcount,path+1,filey,filestyle)
        rowcount+=1
        total+=1
print (total)
wb.save(os.path.join(os.getcwd(),"FolderStructure.xls"))
os.startfile(os.path.join(os.getcwd(),"FolderStructure.xls"))








