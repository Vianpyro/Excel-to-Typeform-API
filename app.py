#!/usr/bin/env python3
# -*- coding: utf8 -*-
import os
import subprocess
import tkinter as tk
import tkinter.filedialog as tkfiledialog
import webbrowser

typeform_types = [
    'date', 'dropdown', 'email', 'file_upload', 
    'group', 'legal', 'long_text', 'matrix', 
    'multiple_choice', 'number', 'opinion_scale', 
    'payment', 'phone_number', 'picture_choice', 
    'ranking', 'rating', 'short_text', 'statement',
    'website', 'yes_no'
]

# subprocess.Popen('pip3 install -r requirements.txt')

###########################
# Tkinter App
###########################
class Application(tk.Frame):
    def __init__(self, master: tk.Tk = None):
        """
        Constructor for the Graphic User Interface.

        :param master: The application's Tkinter root.
        """
        super().__init__(master)
        self.master = master
        self.widgets_width = 30
        self.master.title('Excel to Typeform')
        self.grid()
        self.create_widgets()

    def create_widgets(self) -> None:
        """
        Method creating and displaying each widget on the Graphic User Interface.
        """
        # Excel resource file
        self.excel_file_label = tk.Label(self.master, text='Excel resource file:')
        self.excel_file_label.grid(row=0, column=0, sticky='w')

        self.excel_file = tk.Entry(self.master, state='readonly', width=self.widgets_width)
        self.excel_file.grid(row=0, column=1, columnspan=2)

        self.excel_file_select_button = tk.Button(self.master, text='Select', padx=7, pady=2, command=self.select_excel_file)
        self.excel_file_select_button.grid(row=0, column=3)

        # Typeform title
        self.typeform_title_label = tk.Label(self.master, text='Typeform title:')
        self.typeform_title_label.grid(row=1, column=0, sticky='w')

        self.typeform_title = tk.Entry(self.master, width=self.widgets_width)
        self.typeform_title.grid(row=1, column=1, columnspan=2)
        
        # Typeform form type
        self.typeform_type_label = tk.Label(self.master, text='Typeform type:')
        self.typeform_type_label.grid(row=2, column=0, sticky='w')

        self.typeform_type_value = tk.StringVar(self.master)
        self.typeform_type_value.set('<!> auto <!>')
        self.typeform_type = tk.OptionMenu(self.master, self.typeform_type_value, *typeform_types)
        self.typeform_type.grid(row=2, column=1, columnspan=2)

        # Typeform token
        self.typeform_token_label = tk.Label(self.master, text='Typeform token:')
        self.typeform_token_label.grid(row=3, column=0, sticky='w')

        self.typeform_token = tk.Entry(self.master, width=self.widgets_width // 2, show='*')
        self.typeform_token.grid(row=3, column=1)
        self.typeform_token.focus()

        self.typeform_token_display_variable = tk.IntVar()
        self.typeform_token_display = tk.Checkbutton(self.master, text='Show token', variable=self.typeform_token_display_variable, command=self.show_token)
        self.typeform_token_display.grid(row=3, column=2)

        self.typeform_token_renew_button = tk.Button(
            self.master, text='Renew', padx=7, pady=2,
            command=lambda: webbrowser.open('https://admin.typeform.com/account#/section/tokens', new=1)
        )
        self.typeform_token_renew_button.grid(row=3, column=3)

        # Start generation button
        self.generate_button = tk.Button(self.master, text='Generate Typeform', padx=10, pady=10, bg='green', font='bold', command=self.generate_api)

        if os.name == 'nt':
            self.generate_button.grid(row=self.master.grid_size()[1], column=0, columnspan=4, padx=self.widgets_width, pady=self.widgets_width * 3)
        else:
            self.generate_button.grid(row=self.master.grid_size()[1], column=2, columnspan=2, padx=self.widgets_width, pady=self.widgets_width * 3)
            
            # Close button
            self.close_button = tk.Button(text='Close', padx=10, pady=10, bg='red', font='bold', command=self.master.destroy)
            self.close_button.grid(row=self.master.grid_size()[1] - 1, column=0)

        # Console
        self.console_label = tk.Label(self.master, text='Console:')
        self.console_label.grid(row=self.master.grid_size()[1], column=0, sticky='w')

        self.console = tk.Text(self.master, state=tk.DISABLED, width=self.widgets_width + 10, height=7)
        self.console.grid(row=self.master.grid_size()[1], column=0, columnspan=4)

    def show_token(self) -> None:
        """
        Method dealing with the visibility of the token.
        """
        if self.typeform_token_display_variable.get():
            self.typeform_token['show'] = ''
        else:
            self.typeform_token['show'] = '*'

    def select_excel_file(self) -> None:
        """
        Method for choosing the Excel file to be used
        """
        # Ask user to select the file to open
        filename = tkfiledialog.askopenfilename(initialdir=os.path.abspath(os.getcwd()), filetypes=[('Excel files', '.xlsx .xlsm .xltx .xltm .xls')])

        # Display the file name
        self.excel_file['state'] = 'normal'
        self.excel_file.delete(0, tk.END)
        self.excel_file.insert(tk.END, filename.split('/')[-1])
        self.excel_file['state'] = 'readonly'
        self.console_log(f'Opened {filename}' if filename != '' else 'No file selected')

        # Auto generate Typeform title
        self.typeform_title.delete(0, tk.END)
        self.typeform_title.insert(tk.END, filename.split('/')[-1].split('.')[0])

    def console_log(self, message:str='<message>') -> None:
        """
        Method to log a message in the console box of the Graphic User Interface.
        """
        self.console['state'] = 'normal'
        self.console.insert(tk.END, f'> {message}.\n')
        self.console.see('end')
        self.console['state'] = tk.DISABLED

    def generate_api(self) -> None:
        """
        Method to make the request to the Typeform API.
        """
        if self.verify_widgets():
            self.console_log('Generating Typeform..')
            ...
        else:
            self.console_log('Error - At least one field is not correct')

    def verify_widgets(self) -> bool:
        """
        Method to check if a request can be made to the API or if the user still has to fill in at least one field.
        
        :return: Wether a proper request can be made to the API.
        """
        return self.excel_file.get() != '' and self.typeform_title.get() != '' and self.typeform_token.get() != ''


root = tk.Tk()
app = Application(master=root)
app.mainloop()
