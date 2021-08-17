import xlrd
path = "C:\\Users\\omshejul\\Saved\\GitHub\\excel-to-vcf\\excel-to-vcf\\book3.xlsx"
book = xlrd.open_workbook(path)
sheet = book.sheet_by_index(0)

file_content = ""

for i in range(1,sheet.nrows):
    file_content+="BEGIN:VCARD\nVERSION:3.0\nFN:"+str(sheet.cell_value(i,1))+" "+str(sheet.cell_value(i,2))+"\nN:"+str(sheet.cell_value(i,2))+";"+str(sheet.cell_value(i,1))+";;;"+"\nEMAIL;TYPE=INTERNET;TYPE=HOME:"+str(sheet.cell_value(i,4))+"\nTEL;TYPE=CELL:"+str(sheet.cell_value(i,3))+"\nitem1.ORG:"+str(sheet.cell_value(i,5))+"\nitem1.X-ABLabel:work\nitem2.TITLE:"+str(sheet.cell_value(i,6))+"\nitem2.X-ABLabel:work\nNOTE:"+str(sheet.cell_value(i,7))+"\nEND:VCARD\n"

text_file = open("C:\\Users\\omshejul\\Saved\\GitHub\\excel-to-vcf\\excel-to-vcf\\contacts.vcf", "w")
text_file.write(file_content)
text_file.close()
print ("File Created🥳🎉")