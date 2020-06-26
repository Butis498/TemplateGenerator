from FileScanner import FileScanner
from TempleteGenerator import TempleteGenerator


if __name__ == '__main__':

    directory = r"C:\\Users\\some\\example\\directory" 

    fileScannerXls = FileScanner(directory , ".xlsx")
    fileScannerDocs = FileScanner(directory , ".docx")

    docx_files = fileScannerDocs.files
    xlsx_files = fileScannerXls.files

    templeteGenerator = TempleteGenerator()
    
    for xls in xlsx_files:
        templeteGenerator.replaceContentXls(xls)

    for doc in docx_files:
        templeteGenerator.replaceContentDoc(doc)
