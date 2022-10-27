import tkinter
import sys, os, time
from tkinter import filedialog as fd
import pandas as pd
from pykml import parser
from shapely.geometry import Point, Polygon
import shapely.wkt
from shapely import ops
import pyproj
from zipfile import ZipFile
import numpy as np
from tkinter import ttk
pd.options.mode.chained_assignment = None

class FieldFinder:
    def __init__(self, root):
        self.root = tkinter.Toplevel(root)
        self.root.title('Field Boundary Finder')
        self.root.iconbitmap('data/exb_logo.ico')
        self.root.resizable(False, True)
        self.root.geometry('650x400')

        self.kmlFrame = tkinter.Frame(self.root, highlightbackground= 'darkgrey', highlightthickness=2)
        self.kmlFrame.grid(row=0, columnspan = 2, sticky = 'ew', pady= 5, padx= 5)
        self.kmlHeader = tkinter.Label(self.kmlFrame, text = 'Select KML/KMZ file with pins in center of fields', font = ('Helvetica bold', 16))
        self.kmlHeader.grid(row=0, columnspan=2, pady=5, padx=5, sticky = 'ew')
        self.kmlTxt = tkinter.StringVar()
        self.kmlTxt.set('Select a KML/KMZ file...')
        self.kmlLabel = tkinter.Label(self.kmlFrame, bd=5, bg='lightgrey', width=50, pady=5,
                                           anchor='w', relief='groove', textvariable=self.kmlTxt)

        self.kmlLabel.grid(row=1, column=0, pady=15, padx=5)

        self.kmlOpen = tkinter.Button(self.kmlFrame, text='Browse', command=self._kml_file)
        self.kmlOpen.grid(row=1, column=1, pady=15)

        self.busFrame = tkinter.Frame(self.root, highlightbackground= 'darkgrey', highlightthickness=2)
        self.busFrame.grid(row=1, columnspan = 2, sticky = 'ew', pady=5, padx=5)
        self.header = tkinter.Label(self.busFrame, text = 'Type all potential business names below \n (separated by commas)')
        self.header.config(font = ('Helvetica bold', 16))
        self.header.grid(row=0, column =0, sticky = 'ew', padx=5)

        self.nameEntry = tkinter.Entry(self.busFrame, width = 70)
        self.nameEntry.insert(tkinter.END, 'Type Here')
        self.nameEntry.bind("<FocusIn>", self._clear_placeholder)
        self.nameEntry.grid(row=1, columnspan = 2, sticky= 'ew', pady = 15, padx=5)
        for child in self.busFrame.winfo_children():
            child.config(state = 'disable')

        self.kmlButton = tkinter.Checkbutton(self.root, text = 'KML/KMZ File', font = ('Helvetica', 12),
                                             command = self._enable_kml, highlightbackground = 'darkgrey', highlightthickness=2)
        self.kmlButton.grid(row = 0, column = 2, sticky = 'nsw', padx=10)
        self.kmlButton.select()
        self.busButton = tkinter.Checkbutton(self.root, text = 'Business Names', font = ('Helvetica', 12), command = self._enable_busnames)
        self.busButton.grid(row = 1, column = 2, sticky = 'nsw', padx=10)
        self.runtype = 'kml'

        self.subButton = tkinter.Button(self.root, text = 'Submit', command = self._submit)
        self.subButton.grid(row=2, columnspan = 2, sticky=  'ew', pady = 10, padx=5)
        self.successTxt = tkinter.StringVar()
        self.successTxt.set('Waiting...')
        self.successMessage = tkinter.Label(self.root, textvariable = self.successTxt, font = ('Helvetica bold', 12), bg = 'lightgreen')
        self.successMessage.grid(row=3, columnspan = 2, sticky=  'ew', pady = 10)

        self.progress = ttk.Progressbar(self.root, orient=tkinter.HORIZONTAL, length=100, mode='determinate')

        self.root.mainloop()

    def _enable_busnames(self):
        self.busButton.select()
        self.kmlButton.deselect()
        for child in self.kmlFrame.winfo_children():
            child.config(state = 'disable')
        for child in self.busFrame.winfo_children():
            child.config(state = 'normal')
        self.runtype = 'bus'

    def _enable_kml(self):
        self.kmlButton.select()
        self.busButton.deselect()
        for child in self.busFrame.winfo_children():
            child.config(state = 'disable')
        for child in self.kmlFrame.winfo_children():
            child.config(state = 'normal')
        self.runtype = 'kml'

    def _kml_file(self):
        filetypes = (('kml/kmz files', ('*.kml', '*.kmz')),
                     ('all files', '*.*'))
        filename = fd.askopenfilename(title='Select KML File', initialdir='/', filetypes=filetypes)
        self.kmlFile = filename
        self.kmlTxt.set(filename)
        return filename

    def _submit(self):
        self.progress.grid(row=4, columnspan = 2,  sticky=  'ew')

        self.successTxt.set('Searching...')
        self.successMessage['text'] = 'Searching...'
        self.successMessage.config(bg = 'yellow')
        self.root.update_idletasks()
        self.results = []
        if self.runtype == 'bus':
            keywords = [c.strip().lower() for c in self.nameEntry.get().split(',')]
            self.results = self._bus_name_search(keywords)
        else:
            print('runnning kml name search')
            self.results = self._kml_field_search()
        if len(self.results) == 0:
            self.successTxt.set('0 self.results found')
            self.successMessage.config(bg = 'red')
        else:
            if self.runtype == 'bus':
                comps = ''.join([str(c) + '\n' for c in self.results["permittee"].unique()]).strip('\n')
                self.root.geometry(f'650x{400+len(self.results["permittee"].unique())*15}')
                self.successTxt.set(f'{len(self.results)} fields found for companies:\n{comps}')
            else:
                comps = ''.join([str(c) + '\n' for c in self.results["permittee"].unique()]).strip('\n')
                self.root.geometry(f'650x{400+len(self.results["permittee"].unique())*15}')
                self.successTxt.set(f'{len(self.results)} fields found for companies:\n{comps}')
            self.successMessage.config(bg = 'lightgreen')
            destLabel = tkinter.Label(self.root, text='Select destination for csv file.', pady=5)
            destLabel.grid(row=5, column=0, sticky='ew')

            destButton = tkinter.Button(self.root, text='Browse', command = self._dest_folder)
            destButton.grid(row=5, column=1, sticky='ew')
            self.root.update_idletasks()


    def _dest_folder(self):
        self.destFolder = fd.askdirectory(title = 'Select Destination Folder', initialdir='/')
        self.results.reset_index(drop = True, inplace = True)
        filename = f'{self.results.loc[0, "permittee"]}'.replace(' ', '_').replace('.', '').replace(',','_') + '_field_data.csv'
        hives = []
        for acres in self.results['calc_acres']:
            remainder = round(acres*2)%4
            hives.append(round(acres*2) - remainder if remainder<3 else round(acres*2)+(4-remainder))
        self.results['hives'] = hives
        self.results.to_csv(os.path.join(self.destFolder, filename), index=False)
        saveSuccess = tkinter.Label(self.root, text = f'csv file successfully saved at:\n{self.destFolder + "/" + filename}',
                                    bg='lightgreen', font = ('Helvetica bold', 12)).grid(row =6, columnspan=2, sticky = 'ew')
        self.root.geometry(f'650x{400 + len(self.results["permittee"].unique()) * 15 + 30}')
        return self.destFolder

    def _clear_placeholder(self, e):
        self.nameEntry.delete('0', 'end')

    def _bus_name_search(self, keywords, target='permittee'):
        # os.chdir(r'C:\Users\marcu\PycharmProjects\mapping')
        master = pd.DataFrame()
        for county in os.listdir(r'C:\Users\marcu\PycharmProjects\EXB_GUI\data\boundaries'):
            if '.py' in county:
                continue
            currData = pd.read_csv(rf'C:\Users\marcu\PycharmProjects\EXB_GUI\data\boundaries\{county}\crops_{county.replace(" ", "_")}.csv')
            # print(f'Starting {county} of length {len(currData)}')
            self.successTxt.set(f'Searching {county} county...')
            self.root.update_idletasks()
            interestCol = currData[target]
            self.progress.config(maximum = len(county))
            for i, name in enumerate(interestCol):
                if i%1000 == 0:
                    self.progress['value'] = i
                    self.root.update_idletasks()
                for keyword in keywords:
                    if type(name) != str:
                        continue
                    if keyword in name.lower():
                        master = master.append(currData.loc[i, :])
        return master

    def _kml_field_search(self):
        innerDrops = self._kmlfielddrops_to_csv(self.kmlFile)
        master = pd.DataFrame()
        counties = []
        for county in os.listdir(r'C:\Users\marcu\PycharmProjects\EXB_GUI\data\boundaries'):
        # for county in ['Glenn']:
            if '.py' in county:
                continue
            currData = pd.read_csv(
                rf'C:\Users\marcu\PycharmProjects\EXB_GUI\data\boundaries\{county}\crops_{county.replace(" ", "_")}.csv', low_memory = False)
            currData['f_coords'] = np.zeros(len(currData))
            currData['field name from input'] = np.zeros(len(currData))
            self.progress.config(maximum = len(currData))
            self.progress['value'] = 0
            self.successTxt.set(f'Searching {county} county...')
            self.root.update_idletasks()
            print(self.successTxt.get())
            geom = currData['geometry'].apply(self._transform_geom)
            time.sleep(1)
            for i, geo in enumerate(geom):
                if i%1000 == 0:
                    self.progress['value'] = i
                    self.root.update_idletasks()
                for n, p in enumerate(innerDrops['points']):
                    if geo.contains(p):
                        counties.append(county)
                        row = currData.loc[i,:]
                        row['fcoords'] = p
                        row['field name from input'] = innerDrops.loc[n, 'field name']
                        master = master.append(row)
        filteredMaster=pd.DataFrame()
        master=master[master['MostRecAll'] == 1]
        for group in master.groupby('field name from input'):
            g, dat = group
            dat.sort_values
            dat['p_eff_date'] = pd.to_datetime(dat['p_eff_date'], dayfirst=True)
            dat = dat.sort_values('p_eff_date', ascending=False)
            filteredMaster = filteredMaster.append(dat.iloc[0, :])
        innerDrops.drop([c for c in innerDrops.index if innerDrops.loc[c, 'points'] not in filteredMaster['fcoords'].tolist()], axis = 0, inplace = True)
        return filteredMaster

    def _kmlfielddrops_to_csv(self, filepath):
        if '.kmz' in filepath:
            kmz = ZipFile(filepath, 'r')
            f = kmz.open('doc.kml', 'r')
        else:
            f = filepath
        doc = parser.parse(f)
        doc = doc.getroot()
        nmsp = '{http://www.opengis.net/kml/2.2}'
        coords = []
        master = pd.DataFrame()
        for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
            fieldName = pm.find('{0}name'.format(nmsp)).text
            for ls in pm.iterfind(
                    '{0}Point/{0}coordinates'.format(nmsp)):
                coords = ls.text.strip().replace('\n', '').split(' ')[0].split(',')
                coords = [float(coords[0]), float(coords[1])]
                p = Point(coords)
                master = master.append({'field name': fieldName,
                                        'points': p}, ignore_index=True)

        destFile = filepath.split('.')[0][:-3] + 'csv.csv'
        return master

    def _transform_geom(self, currGeom):
        pgon = shapely.wkt.loads(currGeom)
        try:
            if np.abs(pgon.exterior.coords[0][0]) > 200:
                wgs84 = pyproj.CRS('EPSG:3310')
                utm = pyproj.CRS('EPSG:4326')
                project = pyproj.Transformer.from_crs(wgs84, utm, always_xy=True).transform
                pgon = ops.transform(project, pgon)
        except:
            if np.abs(pgon[0].exterior.coords[0][0]) > 200:
                wgs84 = pyproj.CRS('EPSG:3310')
                utm = pyproj.CRS('EPSG:4326')
                project = pyproj.Transformer.from_crs(wgs84, utm, always_xy=True).transform
                pgon = ops.transform(project, pgon)
        return pgon


# FieldFinder(tkinter.Tk())