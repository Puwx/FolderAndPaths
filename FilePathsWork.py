import string,os,xlwt

while True:
    foldques = raw_input("Please input the full path to the folder that you want mapped.")
    if os.path.isdir(foldques):
        print (foldques +" will have its contents mapped.")
        basepath  = foldques
        break
    else:
        print("That is not a valid folder on this computer, please provide a valid input.")

while True:
    outques = raw_input("Would like the results on excel (type excel) or notepad (type note)?")
    if outques.lower() == "excel":
        
        wb =xlwt.Workbook()
        ws = wb.add_sheet(os.path.basename(foldques)+ "  File Structure")
        ws.write(0,16,basepath)

        foldstyle = "font: bold on, color-index red"
        filestyle = "font: bold off, color-index green"
        foldstyle = xlwt.easyxf(foldstyle)
        filestyle = xlwt.easyxf(filestyle)
        ws.write(1,16,"RED, BOLD TEXT DENOTES A FOLDER",foldstyle)
        ws.write(2,16,"Green, un-bolded text denotes a file",filestyle)
        rowcount = 1
        total = 0
        pathcount = sum(basepath.count(x) for x in ("//","\\"))
        for root, dirs, files in os.walk(basepath):
            if "recycle" in basepath.lower():
                pass
            else:
                path = sum(root.count(x) for x in ("//","\\"))-pathcount
                ws.write(rowcount,path,os.path.basename(root),foldstyle)
                rowcount+=1
                total+=1
                for filey in files:
                    ws.write(rowcount,path+1,filey,filestyle)
                    rowcount+=1
        wb.save(os.path.join(os.getcwd(),"FolderStructure.xls"))
        os.startfile(os.path.join(os.getcwd(),"FolderStructure.xls"))
        break
        

    elif outques.lower() == "note":
        text = open(os.path.join(os.getcwd(),"Folderstructure.txt"),"w")
        pathcount = sum(basepath.count(x) for x in ("//","\\"))
        for root, dirs, files in os.walk(basepath):
            if "recycle" in basepath.lower():
                pass
            else:
                path = sum(root.count(x) for x in ("//","\\")) - pathcount
                indent = "   "*path
                text.write(indent+os.path.basename(root)+"\n")
                for filey in files:
                    text.write("   "+indent+filey+"\n")
            del indent
                    
        text.close()
        os.startfile(os.path.join(os.getcwd(),"Folderstructure.txt"))
        break



    else:
        pass







