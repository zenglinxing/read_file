# read txt splited by \n
# encoding: utf-8 is recommended, if txt encoded in this form
# container: return form, like [line1,line2,...]. Support tuple and list
def txt_readline(file,encoding=None,container=list):
    if encoding==None:
        txt=open(file)
    else:
        txt=open(file,encoding=encoding)
    con=txt.read()
    txt.close()
    return container(con.split('\n'))

# split: string,list and tuple are supported. Default to split space
def txt_readline_split(file,encoding=None,container=list,split=None):
    if encoding==None:
        txt=open(file)
    else:
        txt=open(file,encoding=encoding)
    con=txt.read()
    con=con.split('\n')
    txt.close()
    a=[]
    if split==None:
        for i in con:
            a.append(container(i.split()))
    elif isinstance(split,(str,list,tuple)):
        if isinstance(split,str):
            split=(split,)
        for i in con: # traverse all lines in con
            b=[i]
            for s in split: # traverse split index in split
                # append split list to b, and cut it by b[n:]
                n=len(b)
                j=0
                while j<n:
                    b=b+b[j].split(s)
                    j=j+1
                b=b[n:]
            a.append(container(b))
    return container(a)

def txt_split(file,encoding=None,container=list,split=None):
    if encoding==None:
        txt=open(file)
    else:
        txt=open(file,encoding=encoding)
    con=txt.read()
    txt.close()
    if split==None:
        a=con.split()
    elif isinstance(split,(str,list,tuple)):
        if isinstance(split,str):
            split=(split,)
        b=[con]
        for s in split: # traverse split index in split
            # append split list to b, and cut it by b[n:]
            n=len(b)
            j=0
            while j<n:
                b=b+b[j].split(s)
                j=j+1
            b=b[n:]
        a=container(b)
    return container(a)
