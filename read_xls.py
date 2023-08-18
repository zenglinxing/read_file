import xlrd,xlwt

'''
data in the .xlsx is recorded in a dictionary.
structure of dictionary------------------------------------------------
key : sheet name
value : the list of data
parameter--------------------------------------------------------------
container : use list/tuple to store the data
by_row : list in value's list is recorded in the row of .xlsx file(if True)
'''

# container: default list. list and tuple are supported
# take
'''
1 2
3 4
'''
# as an example, if by_rows=True, return [[1,2],[3,4]]
# if by_rows=False, return [[1,3],[2,4]]
def read_xls(file,container=list,by_row=True):
    workbook=xlrd.open_workbook(file)
    names=workbook.sheet_names()
    ns=len(names)
    d={}
    if by_row in (True,False):
        by_row=(by_row,)*ns
    for k in range(ns):
        sheet=names[k]
        worksheet=workbook.sheet_by_name(sheet)
        nrows=worksheet.nrows
        ncols=worksheet.ncols
        #values=[]
        if by_row[k]:
            d[sheet]=container(container(worksheet.cell_value(i,j)
                                         for j in range(ncols))
                               for i in range(nrows))
        else:
            d[sheet]=container(container(worksheet.cell_value(i,j)
                                         for i in range(nrows))
                               for j in range(ncols))
    return d

def write_xls(d,output,encoding='utf-8',style_compression=0,by_row=True):
    workbook=xlwt.Workbook(encoding=encoding,style_compression=style_compression)
    ns=len(d.keys())
    if by_row in (True,False):
        by_row=(by_row,)*ns
    sheets=[]
    keys=list(d.keys())
    ns=len(keys)
    for k in range(ns):
        if by_row[k]:
            sheets.append(workbook.add_sheet(keys[k],cell_overwrite_ok=True))
            n1=len(d[keys[k]])
            for i in range(n1):
                n2=len(d[keys[k]][i])
                for j in range(n2):
                    sheets[-1].write(i,j,d[keys[k]][i][j])
        else:
            sheets.append(workbook.add_sheet(keys[k],cell_overwrite_ok=True))
            n1=len(d[keys[k]])
            for i in range(n1):
                n2=len(d[keys[k]][i])
                for j in range(n2):
                    sheets[-1].write(j,i,d[keys[k]][i][j])
    workbook.save(output)
