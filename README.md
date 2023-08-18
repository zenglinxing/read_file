# read file

This is a Python project, aiming to read the files more conveniently but cost more time and RAM.

I suggest you to throw those files into your project directly so as to call them conveniently.

# quick start

## read_xls && read_xlsx

The ways to use (read/write)_(xls/xlsx) functions are the same.

Take reading an xlsx file as an example. To read directly,

```python
d = read_xlsx(filename)
```

and you get a dictionary, with the sheets' names as the key and the 2D table (row major) as the value.

If you want the table to be read as column major,

```python
d = read_xlsx(filename, by_row=False)
```

The by_row is True by default, which indicates that all the sheets will be read by row major. You can also determine which sheet to be row major or column major.

```python
d = read_xlsx(filename, by_row=(True, False))
```

So the first sheet is row major and the second one is column major.

## write_xls && write_xlsx

Similar to read_(xls/xlsx). Given the dictionary and the parameter by_row.

```python
write_xlsx(d, filename, by_row=False)
```
