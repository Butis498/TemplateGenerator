## Xlsx , Docx Multy replacement ##

Program to replace words inside a .docx or an .xlsx doc filling with a csv file input , there is a dictionary of terms which can be actually replaced in the doc "Replacement_dictionary.txt" , this will generate multiple docs with diferent data replaced instead of the dictionary inside each doc or xls.

It does not support documents with macros or other types of older versions of microsoft office

### Install openpyxl-2.5.11 ###

```console
pip install openpyxl==2.5.11
```
other versions will not work properly

## Run ##

First replce the gaps in which you want to replace data inside a document after that modify the main.py code to change the directory that will be scanned with the files to be replaced and generated , run the code and a new folder with the files will be created inside the program directory with the files generated
