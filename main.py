#!/usr/bin/env python3
# -*- coding: utf8 -*-
from getpass import getpass
import os
import subprocess
import webbrowser

# Import the required libraries.
try:
    import openpyxl
    import typeform
    import requests
except:
    print('Installation of the required libraries...')
    try:
        subprocess.Popen('pip3 install -r requirements.txt')
        import openpyxl
        import typeform
        import requests
    except:
        raise ImportError('Unable to install the required libraries!')
    else:
        print('The required libraries have been successfully installed.')

###########################
# EXCEL
###########################
class Excel_File:
    def __init__(self, path:str='', filename:str=None):
        """
        Constructor for Excel files reader class.

        :param path: The path to the Excel file to read.
        :param filename: The name of the Excel file to read.
        """
        self.path = path if path not in (None, '') or os.path.exists(path) else os.path.abspath(os.getcwd())
        self.filename = self.retrieve_filename(filename)

    def retrieve_filename(self, filename) -> str:
        """
        Method to scan the path directory to find the Excel file(s).

        :param filename: The name of the file that should be found.
        :return: The name(s) of the file(s) that was found.
        """
        if filename is not None and os.path.exists(self.path + os.path.sep + filename):
            return filename
        else:
            excel_files = [
                document for document in os.listdir(os.path.abspath(os.getcwd()))
                if document[-5:] in ['.xlsx', '.xlsm', '.xltx', '.xltm']
            ]

            if len(excel_files) == 0:
                raise ValueError('Unable to locate any Excel file in this directory:', os.path.abspath(os.getcwd()))
            elif len(excel_files) > 1:
                # Ask which of the Excel documents should be opened
                name = input(f'{len(excel_files)} Excel files were found, please type the name of the one to use: ')
                corresponding_files = [document for document in excel_files if name in document]

                # As long as exactly one document is not selected, ask the user
                while not len(corresponding_files) == 1:
                    name = input('Error, please type the name of the one to use: ')
                    corresponding_files = [document for document in excel_files if name in document]

                # Once a document is selected, open it
                return corresponding_files[0]
            else:
                return excel_files[0]
        raise Exception('Unexcepted Exception.')

    def __str__(self):
        return str(self.path + os.path.sep + self.filename)

class Workbook:
    def __init__(self, filename:(str, Excel_File), read_only:bool=False, keep_vba:bool=False, data_only:bool=False, keep_links:bool=True):
        """
        Constructor for Workbook class.

        :filename: The name of the file to extract the informations from.
        :read_only: Optimised for reading, content cannot be edited.
        :keep_vba: Preseve vba content (this does NOT mean you can use it).
        :data_only: Controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet.
        :keep_links: Whether links to external workbooks should be preserved (True by default).
        """
        self.workbook = openpyxl.load_workbook(str(filename), read_only, keep_vba, data_only, keep_links)
        self.sheets_count = len(self.workbook.sheetnames)
        self.content = [
            [
                [cell.value for cell in row]
                for row in self.workbook[sheet].rows
            ]
            for sheet in self.workbook.sheetnames
        ] if self.sheets_count > 1 else [
            [cell.value for cell in row]
            for row in self.workbook[self.workbook.sheetnames[0]].rows
        ]
        self.workbook.close()
        self.rows_titles = self.content[0] if self.sheets_count == 1 else [sheet[0] for sheet in self.content]


ef = Excel_File()
wb = Workbook(filename=ef, read_only=True)

###########################
# TYPEFORM
###########################
if input('Do you alreay have a token? [yes/no]? ')[0].lower() == 'n':
    print(
        'A token with writing access is required for this program to create an API.\n',
        'Opening tokens creation web page...'
    )
    webbrowser.open('https://admin.typeform.com/account#/section/tokens', new=1)

my_typeform = typeform.Typeform(getpass('Token (text will be hidden): '))
form = my_typeform.forms

"""
# Delete all the older forms
for e in form.list()['items']:
    form.delete(e['id'])
"""

fields = [
    {
        "title": wb.content[i + 1][1],
        "ref": f"Question-{wb.content[i + 1][0]}",
        "type": "multiple_choice",
        "properties": {
            "randomize": False,
            "allow_multiple_selection": False,
            "allow_other_choice": False,
            "vertical_alignment": True,
            "choices": [
                {
                    "label": label,
                    "ref": "test"+label
                }
                for label in wb.content[i + 1][3].split('/')
            ]
        },
        "validations": {"required": True}
    }
    for i in range(len(wb.content) - 1)
]

# Create a new form
new_form = form.create({'title': ef.filename.split('.')[0], 'fields': fields})
input(
    f'New form created with ID: {new_form["id"]}!\n',
    'Press "enter" to close this window.'
)
