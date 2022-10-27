import tkinter
from tkinter import ttk
import sys, os, time
from tkinter import filedialog as fd
import pandas as pd
from windows.find_fields import FieldFinder
from windows.map_maker import MMKML
from windows.map_maker_1118 import MMCSV
from windows.full_pipeline import FullPipeline

class MainWindow:
    def __init__(self):
        self.root = tkinter.Tk() #create window
        self.root.resizable(False, False) #set parameters
        self.root.title('MAPPING TOOL')
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w, h = 450, 500
        x, y = (sw/4) - (w/2), (sh/2) - (h/2)
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y)) #place window in center of screen
        self.root.iconbitmap('data/earth_ico.ico')
        style = ttk.Style(self.root)
        self.root.tk.call('source', r'C:\Users\marcu\PycharmProjects\EXB_GUI\styles\azure\azure.tcl')
        self.root.tk.call('set_theme', 'light')

        self.header = tkinter.Label(self.root, text = 'Mapping Tool', relief = 'groove', padx = 10, pady = 10, font = ('Times New Roman', 20 , 'bold'))
        self.header.grid(row = 0, columnspan = 3, pady = 10)

        self.ffTitle = tkinter.Label(self.root, text = 'Field Finder', padx = 5, pady = 10, font = ('Times New Roman', 16, 'underline'))
        self.ffTitle.grid(row = 1, column = 0, sticky = 'w', pady = 10)
        self.ffDesc = tkinter.Label(self.root, text='Field Finder allows you to input \n'
                                                    'potential business names and search \n'
                                                    'for the associated permit information\n'
                                                    'for these business names.', padx = 5, pady = 10, font = ('Times New Roman', 12), justify = tkinter.LEFT)
        self.ffDesc.grid(row =2, columnspan = 2, sticky = 'w')
        self.openffButton = tkinter.Button(self.root, text='Open Field Finder', command=self._open_ff)
        self.openffButton.grid(row=2, column=2, sticky='s', pady=10)

        self.genTitle = tkinter.Label(self.root, text = 'KML Generator', padx = 5, pady = 10, font = ('Times New Roman', 16, 'underline'))
        self.genTitle.grid(row = 3, column = 0, sticky = 'w', pady = 10)
        self.genDesc = tkinter.Label(self.root, text='KML Generator will take your \n'
                                                     'pollination maps and add the \n'
                                                     'associated field boundaries and \n'
                                                     'bee flight radii for each hive drop.\n'
                                                     'The second option will initialize\n'
                                                     'Field Finder to generate data that \n'
                                                     'feeds into KML Generator.', padx = 5, pady = 10, font = ('Times New Roman', 12), justify = tkinter.LEFT)
        self.genDesc.grid(row =4, columnspan = 2, sticky = 'w')
        self.openGenButton = tkinter.Button(self.root, text= 'Open KML Generator (from KML)', command = self._open_kml_gen_kml)
        self.openGenButton.grid(row =4, column = 2 , sticky = 'nw', pady = 10)
        self.openGenButton = tkinter.Button(self.root, text= 'Open KML Generator (w/ FF)', command = self._open_full_pipeline)
        self.openGenButton.grid(row =4, column = 2 , sticky = 'sw', pady = 10)

        self.root.mainloop()

    def _open_ff(self):
        self.FF = FieldFinder(self.root)
    def _open_kml_gen_kml(self):
        self.MMKML = MMKML(self.root)
    def _open_full_pipeline(self):
        self.fullPipeline = FullPipeline(self.root)

if __name__ == '__main__':
    m = MainWindow()

