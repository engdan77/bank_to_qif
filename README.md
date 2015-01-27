# edoBank2Qif

Program to process XML/XLS and generate QIF-format supported by e.g. GnuCash

---------------------
Command Line Argument
---------------------

```
usage: edoBank2Qif.py [-h] [--existing file.csv] [--verbose] file.xls file.qif

Tool to parse bank-exports into qif-format supported by GnuCash

positional arguments:
  file.xls             Input file to parse (Skandiabanken xls-file)
  file.qif             Output qif file

optional arguments:
  -h, --help           show this help message and exit
  --existing file.csv  Create/Update/Read CSV file to only create qif of new
                       changes
  --verbose            Verbose mode
```
