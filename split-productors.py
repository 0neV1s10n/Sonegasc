# Python program to read an excel file

# import openpyxl module
from queue import Empty
import openpyxl
from openpyxl import Workbook

# create a workbook
wb = Workbook()
#ws = wb.active

# Give the location of the file
path = "/Users/lqa/Downloads/Commande 29 octobre 2022.xlsx"

# To open the workbook
# workbook object is created
wb_obj = openpyxl.load_workbook(path, data_only=True)

# Get workbook active sheet object
# from the active attribute
sheet_obj = wb_obj.active

# Getting the value of maximum rows
# and column
row = sheet_obj.max_row
column = sheet_obj.max_column
  
print("Total Rows:", row)
print("Total Columns:", column)


print("\nValue of first row - FAMILY NAME")
for i in range(1, column + 1): 
    cell_obj = sheet_obj.cell(row = 5, column = i) 
    if cell_obj.value:
        print(cell_obj.value, end = " ")
    

# printing the value of last row
print("\nValue of last column - PRODUCTORS LIST")

# initialize prod_name and prod_pos_row arrays
prod_name = []
prod_pos_row =[]


# loop through rows and detect productors and position in file
for a in range(1, row + 1):

    producteur = sheet_obj.cell(row = a, column = 1).value
    if producteur:
        if '"**"' in producteur:
            producteur=(producteur.lstrip('"**"'))
            prod_name.append(producteur)
            prod_pos_row.append(a)

# List productors and positions in file
#print(prod_name)
#print(prod_pos_row)

# Print headers / columns titles
#print("intitulé;description;unite;prix unité;quantite_total;montant_total")


# Loop through file to list ordered items per productor
for prod in range (0,len(prod_name)-1):
#    print("Producteur ",prod,"/",len(prod_name))
    print("*** Producteur: ", prod_name[prod]," ***")
    wb.create_sheet(prod_name[prod])
    if prod == 1: del wb['Sheet']

    ws = wb[prod_name[prod]]
    ws['A1'] = 'intitulé'
    ws['B1'] = 'description'
    ws['C1'] = 'unite'
    ws['D1'] = 'prix unitaire'
    ws['E1'] = 'quantité'
    ws['F1'] = 'total'

    for item in range (prod_pos_row[prod]+1,(prod_pos_row[prod+1])):

            intitulé = sheet_obj.cell(row = item, column = 1)
            description = sheet_obj.cell(row = item, column = 2)
            unite = sheet_obj.cell(row = item, column = 3)
            prixun = sheet_obj.cell(row = item, column = 4)
            montant_total = sheet_obj.cell(row = item, column = sheet_obj.max_column - 1)
            quantite_total = sheet_obj.cell(row = item, column = sheet_obj.max_column - 2)

            if montant_total.value: 

                #print (sheet_obj.cell(row = item, column = sheet_obj.max_column - 1).value)
                if montant_total.value: 
                        #print ("intitulé : ",intitulé.value)
                        #print ("description : ",description.value)
                        #print ("unite : ",unite.value)
                        #print ("prix unité : ",prixun.value)
                        #print ("quantite_total : ",quantite_total.value)
                        #print ("montant_total : ",montant_total.value)

                        print(intitulé.value,";",description.value,";",unite.value,";",prixun.value,";",quantite_total.value,";",montant_total.value)

                        insert_row=intitulé.value,description.value,unite.value,prixun.value,quantite_total.value,montant_total.value
                        ws.append(insert_row)
                    
wb.save('/Users/lqa/Downloads/test.xlsx')
