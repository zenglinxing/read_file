import openpyxl

'''
data in the .xlsx is recorded in a dictionary.
structure of dictionary------------------------------------------------
key : sheet name
value : the list of data
parameter--------------------------------------------------------------
container : use list/tuple to store the data
by_row : list in value's list is recorded in the row of .xlsx file(if True)
'''

def read_xlsx(file,container=list,by_row=True):
    workbook=openpyxl.load_workbook(file)
    sheets=workbook.sheetnames
    ns=len(sheets)
    d={}
    if by_row in (True,False):
        by_row=(by_row,)*ns
    for k in range(ns):
        sheet=workbook[sheets[k]]
        n1=sheet.max_row
        n2=sheet.max_column
        if by_row[k]:
            d[sheet.title]=container(container(sheet.cell(row=i+1,
                                                          column=j+1).value
                                               for j in range(n2))
                                     for i in range(n1))
        else:
            d[sheet.title]=container(container(sheet.cell(row=i+1,
                                                          column=j+1).value
                                               for i in range(n1))
                                     for j in range(n2))
    return d

def write_xlsx(d,output,by_row=True):
    workbook=openpyxl.Workbook()
    sheets=[]
    names=list(d.keys())
    ns=len(names)
    if by_row in (True,False):
        by_row=(by_row,)*ns
    for k in range(ns):
        sheet=names[k]
        lis=d[sheet]
        # if only create_sheet(), the first empty 'Sheet' is quite annoying
        sheets.append(workbook.active if k==0 else workbook.create_sheet())
        sheets[-1].title=sheet
        if by_row[k]:
            n1=len(d[sheet])
            for i in range(n1):
                n2=len(lis[i])
                for j in range(n2):
                    sheets[-1].cell(i+1,j+1,lis[i][j])
        else:
            n1=len(d[sheet])
            for i in range(n1):
                n2=len(lis[i])
                for j in range(n2):
                    sheets[-1].cell(j+1,i+1,lis[i][j])
    workbook.save(filename=output)
