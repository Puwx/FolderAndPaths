import os,xlwt,arcpy

while True:
    foldques = raw_input("Please input the full path to the folder that you want mapped.")
    if os.path.isdir(foldques):
        print (foldques +" will have its contents mapped.")
        path  = foldques
        break
    else:
        print("That is not a valid folder on this computer, please provide a valid input.")

wb =xlwt.Workbook()
ws = wb.add_sheet("File Structure")
ws.write(0,16,path)

foldstyle = "font: bold on, color-index red"
filestyle = "font: bold off, color-index green"
spatialstyle = "font: underline on,color-index orange"
foldstyle = xlwt.easyxf(foldstyle)
filestyle = xlwt.easyxf(filestyle)
spatialstyle = xlwt.easyxf(spatialstyle)
ws.write(1,16,"RED, BOLD TEXT DENOTES A FOLDER",foldstyle)
ws.write(2,16,"GREEN, UN-BOLDED TEXT DENOTES A FILE",filestyle)
ws.write(3,16,"ORANGE, UNDERLINED TEXT DENOTES A GEOSPATIAL FILE",spatialstyle)
rowcount = 1



for root, dirs, files in os.walk(path):
    if "recycle" in root.lower():
        pass
    else:
        path = root.count("\\")-1
        arcpy.env.workspace = root
        ws.write(rowcount,path,os.path.basename(root),foldstyle)
        rowcount +=1
        if len(arcpy.ListFeatureClasses()) >0:
            for fc in arcpy.ListFeatureClasses():
                ws.write(rowcount,path+1,fc,spatialstyle)
                rowcount +=1
        else:
            for filey in files:
                ws.write(rowcount,path+1,filey,filestyle)
                rowcount+=1
            
        
wb.save(os.path.join(os.getcwd(),"FolderStructure.xls"))
os.startfile(os.path.join(os.getcwd(),"FolderStructure.xls"))








