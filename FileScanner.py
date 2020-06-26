import glob
import requests
import json
import os
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed


# Class design to get all files with supported extentions in a certain directory

class FileScanner():

    # Constructor
    # Parameters: (directory:directory to scan the supported files and its subdirectories)
    # returns: non
    def __init__(self, directory , extension):

        self.glob_path = Path(directory)
        self.files = []
        self.extension = extension
        self.api_response = []  # Contains the response from the api in a list, in a json format
        self.supportedExtXls = [".xls", ".xlsx"]
        self.supportedExtDocs = [".doc", ".docx"]
        self. api_url = 'https://api-gembox.kriptos.dev/api/gembox'
        self.fileBrowser()

    # This method browses the files in te path given in the constructor
    # Paramaters: non
    # Returns: a list of paths of the files found
    def fileBrowser(self):
        
        if self.extension == ".doc" or self.extension == ".docx":
            extensionToRead = self.supportedExtDocs
        elif self.extension == ".xls" or self.extension == ".xlsx":
            extensionToRead = self.supportedExtXls

        for FileType in extensionToRead:
            for path in Path(self.glob_path).rglob(f'*{FileType}'):
                new_path = str(path)
                self.files.append(new_path)



