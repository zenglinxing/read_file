# read_file
Author: Lin-Xing Zeng
Email:  jasonphysics@outlook.com

The file read_xls.py depends on xlrd and xlwt. File read_xlsx.py depends on openpyxl. Here I take read_xlsx.py as an example. For me read_txt.py is not regularly used.

## read_xlsx.py

The function read_xlsx(file,container=list,by_row=True,) requires the .xlsx file name. It return a dict, in which the key points to a 2D list. Or, if you wish, you could turn it into 2D tuple by setting container=tuple. The 2D list sort by row by default, and you can set by_row=False to switch to read by column.

In function write_xlsx(d,file,by_row=True,), file and by_row work the same as those in read_xlsx. Variable d is the dict as described above.
