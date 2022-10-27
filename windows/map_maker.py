import sys,os,time
import pandas as pd
import tkinter
from tkinter import ttk
from tkinter import filedialog as fd
import simplekml
import pyproj
import shapely.wkt
from shapely import ops
from shapely.ops import transform
from shapely.geometry import Polygon, Point
from polycircles import polycircles
from pykml import parser
import numpy as np
import xml.etree.ElementTree as et
from lxml import etree
from zipfile import ZipFile
import traceback
from pyproj import Geod

def _kml_type(filepath):
    if '.kmz' in filepath:
        kmz = ZipFile(filepath, 'r')
        f = kmz.open('doc.kml', 'r')
    doc = parser.parse(f)
    doc = doc.getroot()
    nmsp = '{http://www.opengis.net/kml/2.2}'
    types = []
    for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
        if len(list(
            pm.iterfind('{0}MultiGeometry/{0}Polygon/{0}outerBoundaryIs/{0}LinearRing/{0}coordinates'.format(nmsp)))) and 'boundary' not in types:
            types.append('boundary')
        for c in pm.iter():
            if 'pushpin' in str(c):
                c = str(c)
                color = c.split('-')[0].split('_')[1]
            if color == 'grn' and 'boundary' not in types:
                types.append('boundary')
            elif color == 'ylw' and 'drops' not in types:
                types.append('drops')
            else:
                print('weird color ', color)
        if len(types) == 2:
            return types
    return types

def _kmldrops_to_csv(filepath):
    if '.kmz' in filepath:
        kmz = ZipFile(filepath, 'r')
        f = kmz.open('doc.kml', 'r')
    doc = parser.parse(f)
    doc = doc.getroot()
    nmsp = '{http://www.opengis.net/kml/2.2}'
    folders = list(doc.iterfind('.//{0}Folder'.format(nmsp)))
    master = pd.DataFrame()
    for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
        folder = [c for c in folders if pm in c.getchildren()][0]
        for c in pm.iter():
            if 'pushpin' in str(c):
                c = str(c)
                color = c.split('-')[0].split('_')[1]
                if color == 'grn':
                    fieldName = pm.find('{0}name'.format(nmsp)).text
                elif color == 'ylw':
                    fieldName = folder.name
                    hives = pm.find('{0}name'.format(nmsp)).text.strip()
                    hives = str(pm.name).strip()
            else:
                print('weird color ', color)

        for ls in pm.iterfind(
                '{0}Point/{0}coordinates'.format(nmsp)):
            coords = ls.text.strip().replace('\n', '').split(' ')[0].split(',')
            coords = [float(coords[0]), float(coords[1])]
            p = Point(coords)
            master = master.append({'field': str(fieldName),
                                    'hives': hives,
                                    'points': p,
                                    'long': float(coords[0]),
                                    'lat': float(coords[1]),
                                    'type': color}, ignore_index=True)
    master.to_csv(filepath.split('.')[0] + '_drops.csv', index=False)
    return filepath.split('.')[0] + '_drops.csv'

def transform_geom(currGeom):
    pgon = shapely.wkt.loads(currGeom)
    # if np.abs(pgon.exterior.coords[0][0] )> 200:
    #     wgs84 = pyproj.CRS('EPSG:3310')
    #     utm = pyproj.CRS('EPSG:4326')
    #     project = pyproj.Transformer.from_crs(wgs84, utm, always_xy=True).transform
    #     pgon = transform(project, pgon)
    return pgon

class KML:
    def __init__(self, farmName):
        self.kml = simplekml.Kml(name = farmName)
        self.radTypes = []
        self.flightDict = {}
        self.unnamedCounter=0

    def boundaries(self, fieldNames, geometries, bdata):
        if 'fieldFolders' not in dir(self):
            self.fieldFolders = {}
        for i, geometry in enumerate(geometries):
            if fieldNames[i] in self.fieldFolders.keys():
                folder = self.fieldFolders[fieldNames[i]]
            else:
                folder = self.kml.newfolder(name=str(fieldNames[i]))
                self.fieldFolders[fieldNames[i]] = folder
            currPoly = folder.newlinestring(name=str(fieldNames[i])+' boundary')
            currPoly.style.linestyle.color = simplekml.Color.blue
            currPoly.style.linestyle.width = 5
            currPoly.coords = geometry.exterior.coords
            currPoly.description = bdata[i]

    def drops(self, dropPoints, fieldNames, labels = None):
        if 'fieldFolders' not in dir(self):
            self.fieldFolders = {}
        self.dropPoints = dropPoints
        if list(str(labels)) ==[None]:
            labels = [str(c[0]) + ', ' + str(c[1]) for c in dropPoints]
        for i, drop in enumerate(dropPoints):
            if fieldNames[i] in self.fieldFolders.keys():
                folder = self.fieldFolders[fieldNames[i]]
            else:
                folder = self.kml.newfolder(name = str(fieldNames[i]))
                self.fieldFolders[fieldNames[i]] = folder

            drop = [drop]
            currPoint = folder.newpoint(name = str(labels[i]), coords = drop)
            currPoint.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png'


    def radius(self, type='honeybee', field = 'Unnamed Field', dropPoints = None):
        if dropPoints != None:
            self.dropPoints == dropPoints
        if 'radFolder' not in dir(self):
            self.radFolder = self.kml.newfolder(name='Flight Radius')
        if type not in self.radTypes:
            self.radTypes.append(type)
            self.flightDict[type] = self.radFolder.newfolder(name=type)

        if field == 'Unnamed Field':
            field+=f' {self.unnamedCounter}'
            self.unnamedCounter+=1

        currFolder = self.flightDict[type].newfolder(name=field)
        bees = {'honeybee': [1000 * .343, simplekml.Color.yellow],
                'bumblebee': [800 * .343, simplekml.Color.green]}

        rad, color = bees[type]
        for i, drop in enumerate(self.dropPoints):
            outerCircle = polycircles.Polycircle(latitude=drop[1],
                                                 longitude=drop[0],
                                                 radius=rad,
                                                 number_of_vertices=36)
            innerCircle = polycircles.Polycircle(latitude=drop[1],
                                                 longitude=drop[0],
                                                 radius=rad - 5,
                                                 number_of_vertices=36)
            poly = currFolder.newpolygon(name=f'{type} radius at ({drop[0]}, {drop[1]})',
                                     innerboundaryis=innerCircle.to_kml(),
                                     outerboundaryis=outerCircle.to_kml())
            poly.style.polystyle.color = color

    def save(self, filepath):
        self.kml.save(filepath)
        print(f'kml saved at {filepath}')

class FarmName:
    def __init__(self,root):
        self.name = ''
        self.root = tkinter.Tk()
        self.root.iconbitmap('data/exb_logo.ico')
        self.root.iconbitmap('data/exb_logo.ico')
        self.root.resizable(False ,False)
        # self.root.geometry('300x50')
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w, h = 300, 50
        x, y = (sw / 2) - (w / 2), (sh / 6 * 3) - (h / 2)
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))
        self.farmNameEntry = tkinter.Entry(self.root, width=30)
        self.farmNameEntry.insert(tkinter.END, 'Enter Grower/Farm Name Here')
        self.farmNameEntry.bind("<FocusIn>", self._clear_placeholder)
        self.farmNameEntry.pack()
        ttk.Button(self.root, text = 'CONTINUE', command = self._close).pack()
        self.root.mainloop()

    def _clear_placeholder(self, e):
        self.farmNameEntry.delete('0', 'end')

    def _close(self):
        self.name = self.farmNameEntry.get()
        if not len(self.name) or self.name == 'Enter Grower/Farm Name Here':
            self.name = 'Unnamed Farm'
        self.root.quit()
        return



class FlightRadiusFixer:
    def __init__(self, filepath):
        self.filepath = filepath
        self._kmlboth_to_csv(self.filepath)
    def _kmlboth_to_csv(self, filepath):
        if '.kmz' in filepath:
            kmz = ZipFile(filepath, 'r')
            f = kmz.open('doc.kml', 'r')
        else:
            f = filepath

        doc = parser.parse(f)
        doc = doc.getroot()
        nmsp = '{http://www.opengis.net/kml/2.2}'
        folders = list(doc.iterfind('.//{0}Folder'.format(nmsp)))
        for f in folders:
            if f.name == 'Flight Radius':
                print(f)
                doc.Document.remove(f)

        self.flightKml = simplekml.Kml()

        self.radFolder = self.flightKml.newfolder(name=f'Flight Radius')
        self.hFolder = self.radFolder.newfolder(name='honeybees')
        self.bFolder = self.radFolder.newfolder(name='bumblebees')

        self.masterDrops = pd.DataFrame()
        for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
            folder = [c for c in folders if pm in c.getchildren()][0]

            for c in pm.iter():
                if 'pushpin' in str(c):
                    c = str(c)
                    color = c.split('-')[0].split('_')[1]
                fieldName = folder.name
                hives = pm.find('{0}name'.format(nmsp)).text.strip()
                hives = str(pm.name).strip()

            for ls in pm.iterfind('{0}Point/{0}coordinates'.format(nmsp)):
                coords = ls.text.strip().replace('\n', '').split(' ')[0].split(',')
                coords = [float(coords[0]), float(coords[1])]
                p = Point(coords)
                self.masterDrops = self.masterDrops.append({'field': str(fieldName),
                                                            'hives': hives,
                                                            'points': p,
                                                            'long': float(coords[0]),
                                                            'lat': float(coords[1])}, ignore_index=True)
        hfDict = {}
        bfDict = {}
        bees = {self.hFolder: [1000 * .343, simplekml.Color.yellow],
                self.bFolder: [800 * .343, simplekml.Color.green]}
        for field in self.masterDrops['field'].unique():
            for fold in [self.hFolder, self.bFolder]:
                x = fold.newfolder(name = field)
                if fold==self.hFolder:
                    hfDict[field] = x
                else:
                    bfDict[field] = x
        for row in self.masterDrops.iterrows():
            i, row = row
            for fold in [self.hFolder, self.bFolder]:
                rad, color = bees[fold]
                outerCircle = polycircles.Polycircle(latitude=row['lat'],
                                                     longitude=row['long'],
                                                     radius=rad,
                                                     number_of_vertices=36)
                innerCircle = polycircles.Polycircle(latitude=row['lat'],
                                                     longitude=row['long'],
                                                     radius=rad - 5,
                                                     number_of_vertices=36)
                if fold==self.hFolder:
                    poly = hfDict[row['field']].newpolygon(name=f'{type} radius at ({row["long"]}, {row["lat"]})',
                                                  innerboundaryis=innerCircle.to_kml(),
                                                  outerboundaryis=outerCircle.to_kml())
                    poly.style.polystyle.color = color
                else:
                    poly = bfDict[row['field']].newpolygon(name=f'{type} radius at ({row["long"]}, {row["lat"]})',
                                                  innerboundaryis=innerCircle.to_kml(),
                                                  outerboundaryis=outerCircle.to_kml())
                    poly.style.polystyle.color = color
        # self.flightKml.save(self.filepath.split('.')[0] + '_flightrad.kml')



class MMKML:
    def __init__(self, root):
        f = FarmName(root)
        while True:
            time.sleep(.1)
            if f.name != '':
                break
        self.farmName = f.name
        self.dropsDone = []
        self.radDone = []
        f.root.destroy()
        self.root = tkinter.Toplevel(root)
        self.root.iconbitmap('data/exb_logo.ico')
        self.root.title(f'Map Maker for {self.farmName}')
        self.root.resizable(False, False)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w, h = 840, 300
        x, y = (sw/2) - (w/2), (sh/6*3) - (h/2)
        self.root.geometry('%dx%d+%d+%d' % (w, h, x, y))

        '''FILE OPENING'''
        self.fileFrame = tkinter.Frame(self.root, highlightbackground= 'darkgrey', highlightthickness=2)
        self.fileFrame.grid(row=0, column=0, sticky = 'ew', pady= 5, padx= 5)
        tkinter.Label(self.fileFrame, text = 'Select Files', font = ('Helvetica bold', 12)).grid(row = 0, column = 0, sticky = 'ew')

        self.kmlTxt = tkinter.StringVar()
        self.kmlTxt.set('Select a kml/kmz file...')
        self.excelTxt = tkinter.StringVar()
        self.excelTxt.set('Select a csv/xlsx file...')

        self.kmlLabel = tkinter.Label(self.fileFrame, bg='lightgrey', bd=5, width=50, pady=5,
                                      anchor='w', relief='groove', textvariable=self.kmlTxt)
        self.kmlLabel.grid(row=1, column=0, pady=10, padx=10, sticky = 'ew')
        self.kmlOpen = ttk.Button(self.fileFrame, text='Browse', command=self._kml_file)
        self.kmlOpen.grid(row=1, column=1, pady=10, padx = 10, sticky = 'e')

        self.excelLabel = tkinter.Label(self.fileFrame, bd=5, bg='lightgrey', width=50, pady=5,
                                        anchor='w', relief='groove', textvariable=self.excelTxt)
        self.excelLabel.grid(row=2, column=0, pady=10, padx=10)
        self.excelOpen = ttk.Button(self.fileFrame, text='Browse', command=self._excel_file)
        self.excelOpen.grid(row=2, column=1, pady=10, padx = 10, sticky = 'e')

        '''ELEMENT SELECTION'''
        self.elementFrame = tkinter.Frame(self.root, highlightbackground= 'darkgrey', highlightthickness=2)
        self.elementFrame.grid(row=0, column=1, sticky = 'ew', pady= 5, padx= 5, stick = 'ns')
        tkinter.Label(self.elementFrame, text='Select Elements', font = ('Helvetica bold', 12)).grid(row = 0, column = 0, sticky = 'ew')
        self.boundBool = tkinter.BooleanVar()
        self.boundBool.set(False)
        self.dropBool = tkinter.BooleanVar()
        self.dropBool.set(False)
        self.radiusBool = tkinter.BooleanVar()
        self.radiusBool.set(False)
        self.boundButton = tkinter.Checkbutton(self.elementFrame, text='Boundaries', variable=self.boundBool, padx=25, pady = 5)
        self.boundButton.grid(row=1, column=0, sticky=tkinter.W)
        self.dropButton = tkinter.Checkbutton(self.elementFrame, text='Drops', variable=self.dropBool, padx=25, pady = 5, command = self._show_HPA)
        self.dropButton.grid(row=2, column=0, sticky=tkinter.W)



        tkinter.Label(self.elementFrame, text='Hives per Acre: ').grid(row=3, column=0,sticky='w')
        self.HPAfield = tkinter.Entry(self.elementFrame, text='Hives Per Acre', width=5)
        self.HPAfield.insert(tkinter.END, '2.0')
        self.HPAfield.bind("<FocusIn>", self._clear_HPA_placeholder())
        self.HPAfield.grid(row=3, column = 0, sticky='e')
        self.HPAfield.config(state='disable')



        self.radiusButton = tkinter.Checkbutton(self.elementFrame, text='Flight Radius', variable=self.radiusBool, padx=25, pady = 5,
                                                command=self._show_bees)
        self.radiusButton.grid(row=4, column=0, sticky=tkinter.W)


        '''BEE TYPE SELECTION'''
        self.beeFrame = tkinter.Frame(self.root, highlightbackground= 'darkgrey', highlightthickness=2)
        self.beeFrame.grid(row=0, column=2, sticky = 'ew', pady= 5, padx= 5, stick = 'ns')
        tkinter.Label(self.beeFrame, text='Select Bee Type', font = ('Helvetica bold', 12)).grid(row = 0, column = 0, sticky = 'ew')
        self.hbee = tkinter.BooleanVar()
        self.bbee = tkinter.BooleanVar()
        self.hbee.set(False)
        self.bbee.set(False)
        self.honeybeeButton = tkinter.Checkbutton(self.beeFrame, text = 'Honeybee', variable =  self.hbee, pady = 10, padx = 25)
        self.honeybeeButton.grid(row = 1, column = 0, sticky = tkinter.W)
        self.bumblebeeButton = tkinter.Checkbutton(self.beeFrame, text = 'Bumblebee', variable =  self.bbee, pady = 10, padx = 25)
        self.bumblebeeButton.grid(row = 2, column = 0, sticky = tkinter.W)
        for child in self.beeFrame.winfo_children():
            child.config(state = 'disable')

        '''RUN PANEL'''
        self.runFrame = tkinter.Frame(self.root, highlightbackground= 'darkgrey', highlightthickness=2)
        self.runFrame.grid(row=2, columnspan = 3, sticky = 'ew', pady= 10, padx= 5)
        self.runFrame.grid_columnconfigure(0, weight = 1)
        self.runButton = tkinter.Button(self.runFrame , text = 'ADD AT LEAST ONE FILE', command = self.run)
        self.runButton.grid(row =0, columnspan = 2, sticky = 'ew')
        self.successTxt = tkinter.StringVar()
        self.successTxt.set('Waiting...')
        self.successLabel = tkinter.Label(self.runFrame, textvariable =self.successTxt, bg = 'yellow')
        self.successLabel.grid(row = 1, columnspan = 2, sticky = 'ew')
        self.successLabel.config(font = ('Helvetica bold', 12))
        for child in self.runFrame.winfo_children():
            child.config(state = 'disable')

        self.destTxt = tkinter.StringVar()
        self.destTxt.set('Select a destination to save...')
        self.destLabel = tkinter.Label(self.runFrame, text = 'Select a destination to save...', bg='lightgrey', bd=5, width=50, pady=5,
                                      anchor='w', relief='groove')
        self.destOpen = ttk.Button(self.runFrame, text='Browse', command=self._dest_file)
        self.destSuccessLabel = tkinter.Label(self.runFrame,font = ('Helvetica bold', 12), textvariable =self.destTxt, bg = 'lightgreen')

        self.root.mainloop()

    def _clear_HPA_placeholder(self):
        self.HPAfield.delete('0', 'end')

    def _kml_file(self):
        filetypes = (('kml/kmz files', ('*.kml', '*.kmz')),
                     ('all files', '*.*'))
        filename = fd.askopenfilename(title='Select a kml/kmz file', initialdir='/Downloads', filetypes=filetypes)
        if not len(filename):
            return

        self.kmlFile = filename
        self.kmlTxt.set(filename)
        if self.runButton['text'] != 'RUN':
            self.runButton['text'] = 'RUN'
            self.successTxt.set('Ready')
            self.successLabel.config(bg = 'lightgreen')
            for child in self.runFrame.winfo_children():
                child.config(state = 'normal')

        return filename

    def _excel_file(self):
        filetypes = (('csv/xlsx files', ('*.csv', '*.xlsx')),
                     ('all files', '*.*'))
        filename = fd.askopenfilename(title='Select a csv/xlsx file', initialdir='/Downloads', filetypes=filetypes)
        if not len(filename):
            return
        self.excelFile = filename
        self.excelTxt.set(filename)
        if self.runButton['text'] != 'RUN':
            self.runButton['text'] = 'RUN'
            self.successTxt.set('Ready')
            self.successLabel.config(bg = 'lightgreen')
            for child in self.runFrame.winfo_children():
                child.config(state = 'normal')
        return filename

    # def show_bound(self):
    #     if str(self.boundBool.get()) == False:
    #         self.boundBool.set(True)
    #         self.boundButton.select()
    #     elif str(self.boundBool.get()) == True:
    #         self.boundBool.set(False)
    #         self.boundButton.deselect()

    def _show_bees(self):
        if self.radiusBool.get():
            for child in self.beeFrame.winfo_children():
                child.config(state='normal')
        else:
            for child in self.beeFrame.winfo_children():
                child.config(state='disable')
                self.hbee.set(False)
                self.bbee.set(False)

    def _show_HPA(self):
        if self.dropBool.get():
            self.HPAfield.config(state='normal')
        else:
            self.HPAfield.config(state='disable')

    def _kml_type(self, filepath):
        if '.kmz' in filepath:
            kmz = ZipFile(filepath, 'r')
            f = kmz.open('doc.kml', 'r')
        else:
            f = filepath
        doc = parser.parse(f)
        doc = doc.getroot()
        nmsp = '{http://www.opengis.net/kml/2.2}'
        types = []
        for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
            desc = pm.descendantpaths()
            if '{0}Placemark.Polygon.outerBoundaryIs.LinearRing.coordinates'.format(
                        nmsp) in desc and 'boundary' not in types:
                types.append('boundary')
                continue
            for c in pm.iter():
                if 'pushpin' in str(c):
                    c = str(c)
                    if 'ylw' in c:
                        color = 'ylw'
                    elif 'grn' in c:
                        color = 'grn'
                    else:
                        print('weird color', c)
                    if color == 'grn' and 'boundary' not in types:
                        types.append('boundary')
                    elif color == 'ylw' and 'drops' not in types:
                        types.append('drops')
            if len(types) == 2:
                return types
        return types, doc

    def _kmlboth_to_csv(self, filepath):
        if '.kmz' in filepath:
            kmz = ZipFile(filepath, 'r')
            f = kmz.open('doc.kml', 'r')
        else:
            f = filepath

        doc = parser.parse(f)
        doc = doc.getroot()
        nmsp = '{http://www.opengis.net/kml/2.2}'
        folders = list(doc.iterfind('.//{0}Folder'.format(nmsp)))
        self.masterDrops = pd.DataFrame()
        self.masterBounds = pd.DataFrame()

        for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
            folder = [c for c in folders if pm in c.getchildren()][0]

            for c in pm.iter():
                if 'pushpin' in str(c):
                    c = str(c)
                    color = c.split('-')[0].split('_')[1]
                    if color == 'grn':
                        fieldName = pm.find('{0}name'.format(nmsp)).text
                    elif color == 'ylw':
                        fieldName = folder.name
                        hives = pm.find('{0}name'.format(nmsp)).text.strip()
                        hives = str(pm.name).strip()

            for ls in pm.iterfind(
                    '{0}Point/{0}coordinates'.format(nmsp)):
                coords = ls.text.strip().replace('\n', '').split(' ')[0].split(',')
                coords = [float(coords[0]), float(coords[1])]
                p = Point(coords)
                self.masterDrops = self.masterDrops.append({'field': str(fieldName),
                                        'hives': hives,
                                        'points': p,
                                        'long': float(coords[0]),
                                        'lat': float(coords[1]),
                                        'type': color}, ignore_index=True)

        self.dropFile = filepath.split('.')[0] + '_drops.csv'
        self.boundFile = filepath.split('.')[0] + '_bounds.csv'
        self.masterDrops.to_csv(self.dropFile, index = False)
        self.masterBounds.to_csv(self.boundFile, index = False)

    def _kmlbound_to_csv(self, filepath):
        # doc = et.parse(filepath)
        # doc = doc.getroot()
        doc = self.kmlDoc
        nmsp = '{http://www.opengis.net/kml/2.2}'
        coords = []
        self.masterBounds = pd.DataFrame()
        for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
            fieldName = pm.name
            # for ls in pm.iterfind(
            #     '{0}MultiGeometry/{0}Polygon/{0}outerBoundaryIs/{0}LinearRing/{0}coordinates'.format(nmsp)):
            ls = pm.Polygon.outerBoundaryIs.LinearRing.coordinates
            coords = ls.text.strip().replace('\n', '').split(' ')
            coords = [c.split(',') for c in coords]
            coords = [[float(i), float(j)] for [i, j, k] in coords]
            poly = Polygon(coords)
            self.masterBounds = self.masterBounds.append({'field': fieldName,
                                    'geometry': poly}, ignore_index=True)

        self.boundFile = filepath.split('.')[0] + '_bounds.csv'
        self.masterBounds.to_csv(self.boundFile, index=False)

    def _kmldrops_to_csv(self, filepath):
        if '.kmz' in filepath:
            kmz = ZipFile(filepath, 'r')
            f = kmz.open('doc.kml', 'r')
        else:
            f = filepath
        doc = parser.parse(f)
        doc = doc.getroot()
        nmsp = '{http://www.opengis.net/kml/2.2}'
        folders = list(doc.iterfind('.//{0}Folder'.format(nmsp)))
        self.masterDrops = pd.DataFrame()
        for pm in doc.iterfind('.//{0}Placemark'.format(nmsp)):
            folder = [c for c in folders if pm in c.getchildren()][0]
            fieldName=folder.name
            for c in pm.iter():
                if 'pushpin' in str(c):
                    c = str(c)
                    color = c.split('-')[0].split('_')[1]
                    if color == 'grn':
                        fieldName = pm.find('{0}name'.format(nmsp)).text
                    elif color == 'ylw':
                        fieldName = folder.name
                        hives = pm.find('{0}name'.format(nmsp)).text.strip()
                        hives = str(pm.name).strip()

            for ls in pm.iterfind(
                    '{0}Point/{0}coordinates'.format(nmsp)):
                coords = ls.text.strip().replace('\n', '').split(' ')[0].split(',')
                coords = [float(coords[0]), float(coords[1])]
                p = Point(coords)
                self.masterDrops = self.masterDrops.append({'field': str(fieldName),
                                        'hives': hives,
                                        'points': p,
                                        'long': float(coords[0]),
                                        'lat': float(coords[1]),
                                        'type': color}, ignore_index=True)
        self.dropFile = filepath.split('.')[0] + '_drops.csv'
        self.masterDrops.to_csv(self.dropFile, index=False)

    def _create_drops(self, numHives = 100, hivesPerDrop=12):
        numHives = self.acreage*float(self.HPAfield.get()) if numHives==100 else numHives
        if self.acreage<150:
            hivesPerDrop = 12
        elif self.acreage<400:
            hivesPerDrop = 24
        else:
            hivesPerDrop = 48
        if numHives<=hivesPerDrop:
            if numHives==8:
                pass
            else:
                hivesPerDrop = numHives
        numHives = round(numHives/hivesPerDrop) * hivesPerDrop
        numDrops = round(int(numHives) / int(hivesPerDrop))
        self.hivesPerDrop = hivesPerDrop
        self.dropPoints = []
        distances = []
        lats, longs = [[i for i, j in self.discLL], [j for i, j in self.discLL]]

        for i, pair in enumerate(self.discLL):
            long, lat = pair
            distances.append(np.sqrt((lat - self.discLL[i - 1][1]) ** 2 + (long - self.discLL[i - 1][0]) ** 2))
        totalDistance = np.sum(distances)
        step = totalDistance / numDrops
        currDist = 0
        self.dropPoints.append([self.discLL[0][0], self.discLL[0][1]])
        for i, pair in enumerate(self.discLL):
            long, lat = pair
            currDist += distances[i]
            if currDist >= step:
                self.dropPoints.append([long, lat])
                currDist = 0
        if len(self.dropPoints) != numDrops:
            self.dropPoints=[self.dropPoints[0]]
            print(f'NUMDROPS {numDrops} DOES NOT EQUAL DROPPOINTS {len(self.dropPoints)}')

    def _poly_area(self, geom):
        geod = Geod(ellps='WGS84')
        area = abs(geod.geometry_area_perimeter(geom)[0])
        acreage = area*0.000247105
        return acreage

    def _create_glatlong(self, geom):
        self.glongs = [i for i, j in list(geom.exterior.coords)]
        self.glats = [j for i, j in list(geom.exterior.coords)]
        self.glatlongs = []
        for i in range(len(self.glongs)):
            self.glatlongs.append([self.glongs[i], self.glats[i]])
        distances = []
        discLats = []
        discLongs = []
        self.discLL = []
        self.acreage = self._poly_area(geom)
        step = .000005  # about 5 meters

        for i, latlong in enumerate(self.glatlongs):
            distance = ((latlong[0] - self.glatlongs[i - 1][0]) ** 2 + (
                        latlong[1] - self.glatlongs[i - 1][1]) ** 2) ** .5
            newPoints = int(np.ceil(distance / step))
            discLats += list(np.linspace(self.glatlongs[i - 1][0], latlong[0], newPoints))
            discLongs += list(np.linspace(self.glatlongs[i - 1][1], latlong[1], newPoints))
        for i in range(len(discLats)):
            self.discLL.append([discLats[i], discLongs[i]])

    def _create_kml(self):
        bees = {'honeybee': [1000 * .343, simplekml.Color.yellow],
                'bumblebee': [800 * .343, simplekml.Color.green]}

        self.beeType = []
        if self.hbee:
            self.beeType.append('honeybee')
        if self.bbee:
            self.beeType.append('bumblebee')

        self.kml = simplekml.Kml(name=self.farmName)
        radFolder = self.kml.newfolder(name=f'Flight Radius')
        hFolder = radFolder.newfolder(name='honeybees')
        bFolder = radFolder.newfolder(name='bumblebees')

        try:
            self.dropData = pd.read_csv(self.csvPath)
        except:
            self.kmlTxt.set(f'Cannot find or load file at {self.csvPath}')
            self.kmlLabel.config(bg='red')
            return
        try:
            for i in range(len(self.geom)):
                folder = self.kml.newfolder(name=self.fieldNames[i])
                currGeo = folder.newlinestring(name=str(self.fieldNames[i]) + ' boundary')
                currGeo.style.linestyle.color = simplekml.Color.blue
                currGeo.style.linestyle.width = 5
                currGeo.coords = self.geom[i].exterior.coords

                currDropData = self.dropData[self.dropData['field'] == self.fieldNames[i]]
                for row in currDropData.iterrows():
                    i, row = row
                    coords = [[row['long'], row['lat']]]
                    currPoint = folder.newpoint(name=str(row['hives']) + 'h', coords=coords)
                    currPoint.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png'
                    # Radius
                    for beeT in self.beeType:
                        rad, color = bees[beeT]
                        outerCircle = polycircles.Polycircle(latitude=row['lat'],
                                                             longitude=row['long'],
                                                             radius=rad,
                                                             number_of_vertices=36)
                        innerCircle = polycircles.Polycircle(latitude=row['lat'],
                                                             longitude=row['long'],
                                                             radius=rad - 5,
                                                             number_of_vertices=36)
                        if beeT == 'honeybee':
                            poly = hFolder.newpolygon(name=f'{type} radius at ({row["long"]}, {row["lat"]})',
                                                      innerboundaryis=innerCircle.to_kml(),
                                                      outerboundaryis=outerCircle.to_kml())
                        elif beeT == 'bumblebee':
                            poly = bFolder.newpolygon(name=f'{type} radius at ({row["long"]}, {row["lat"]})',
                                                      innerboundaryis=innerCircle.to_kml(),
                                                      outerboundaryis=outerCircle.to_kml())
                        poly.style.polystyle.color = color

            self.kmlPath = self.csvPath.split('.')[0].strip('_drops') + '_finalKML.kml'
            self.kml.save(self.kmlPath)
            self.kmlTxt.set(f'KML file successfully saved at \n {self.kmlPath}')
            self.kmlLabel.config(bg='lightgreen')
        except Exception as e:
            self.kmlTxt.set('Issue creating oldBoundaries and drops in KML file \n' + str(e))
            self.kmlLabel.config(bg='red')
            return

    def _dest_file(self):
        try:
            destDir = self.kmlFile.split('.')[0]
        except:
            destDir = self.excelFile.split('.')[0]
        destFile = fd.asksaveasfile(mode='w', initialfile = f'{self.farmName}', defaultextension=".kml")
        if destFile is None:
            return
        self.myKml.save(destFile.name)
        destFile.close()
        self.destSuccessLabel.grid(row = 3, columnspan = 2,sticky = 'ew')
        self.destTxt.set(f'KML saved at {destFile.name}')

    def run(self):
        errors = {'nokml' : 'ERROR: No KML file Selected',
                  'badkml' : 'ERROR: Invalid kml file -- not filetype kml',
                  'badbound' : 'ERROR: Invalid boundary file -- not filetype csv'}
        error = ''
        boundaries = [None]
        
        if self.kmlTxt.get() == 'Select a KML File...':
            error = error + errors['nokml'] + '\n'
        elif '.kml' not in self.kmlTxt.get() and '.kmz' not in self.kmlTxt.get():
            error = error + errors['badkml'] + '\n'
        if ('.csv' not in self.excelTxt.get() or '.kml' not in self.excelTxt.get()) and self.excelTxt.get() != ('Select a csv/xlsx file...' or ''):
            error = error + errors['badbound'] + '\n'

        '''HANDLE KML FILE'''
        if 'kmlFile' in dir(self):
            self.kmlTypes, self.kmlDoc = self._kml_type(self.kmlFile)
            if 'drops' in self.kmlTypes and 'boundary' not in self.kmlTypes:
                try:
                    self._kmldrops_to_csv(self.kmlFile)
                    drops =self.masterDrops
                except Exception as e:
                    eType, eObj, eTb = sys.exc_info()
                    fname = os.path.split(eTb.tb_frame.f_code.co_filename)[1]
                    self.successTxt.set(f'kml to csv conversion failed for drops in {self.kmlFile} \n'
                                        f'{e} in {fname} line {eTb.tb_lineno}')
                    self.successLabel.config(bg = 'red')
                    print(e)
                    print(eType, fname, eTb.tb_lineno)
                    print(traceback.format_exc())
                    return
            elif 'boundary' in self.kmlTypes and 'drops' not in self.kmlTypes:
                try:
                    self._kmlbound_to_csv(self.kmlFile)
                    boundaries = self.masterBounds
                except Exception as e:
                    eType, eObj, eTb = sys.exc_info()
                    fname = os.path.split(eTb.tb_frame.f_code.co_filename)[1]
                    self.successTxt.set(f'kml to csv conversion failed for boundaries in {self.kmlFile} \n'
                                        f'{e} in {fname} line {eTb.tb_lineno}')
                    self.successLabel.config(bg = 'red')
                    print(e)
                    print(eType, fname, eTb.tb_lineno)
                    print(traceback.format_exc())
                    return
            elif len(self.kmlTypes) == 2:
                try:
                    self._kmlboth_to_csv(self.kmlFile)
                    drops = self.masterDrops
                    boundaries = self.masterBounds
                except Exception as e:
                    eType, eObj, eTb = sys.exc_info()
                    fname = os.path.split(eTb.tb_frame.f_code.co_filename)[1]
                    self.successTxt.set(f'kml to csv conversion failed for drops and boundaries in {self.kmlFile}\n'
                                        f'{e} in {fname} line {eTb.tb_lineno}')
                    self.successLabel.config(bg = 'red')
                    print(e)
                    print(eType, fname, eTb.tb_lineno)
                    print(traceback.format_exc())
                    return
            else:
                self.successTxt.set(f'No valid drops or boundaries in {self.kmlFile}')
                return

        '''HANDLE EXCEL FILE'''
        if 'excelFile' in dir(self):
            types = []
            self.excelData = pd.read_csv(self.excelFile) if '.csv' in self.excelFile else pd.read_excel(self.excelFile)
            for col in self.excelData.columns:
                if 'geo' in col or 'bound' in col or 'poly' in col:
                    types.append('boundary')
                if 'hive' in col:
                    types.append('drops')
                if 'type' in col and 'grn' in self.excelData[col].tolist():
                    types.append('boundary')
            types = pd.Series(types).unique().tolist()
            if 'drops' in types:
                if 'masterDrops' in dir(self):
                    self.masterDrops = [self.masterDrops, self.excelData]
                else:
                    self.masterDrops = self.excelData
            if 'boundary' in types:
                if 'masterBounds' in dir(self):
                    self.masterBounds = [self.masterBounds, self.excelData]
                else:
                    self.masterBounds = self.excelData

        '''ADD AN INPUT FIELD FOR FARM NAME'''
        self.myKml = KML(self.farmName)
        geometries = []
        fieldNames = []
        description = []
        if 'masterBounds' in dir(self):
            boundaries = [self.masterBounds] if type(self.masterBounds) != list else self.masterBounds
            for b in boundaries:
                needDesc = True
                cols = b.columns
                for c in cols:
                    if ('field' in c.lower() or 'name' in c.lower()) and 'coord' not in c.lower():
                        self.fieldNameCol = c
                        fieldNames += b[c].tolist()
                    if 'geo' in c.lower() or 'poly' in c.lower() or 'bound' in c.lower():
                        if 'polygon' not in str(b[c][0]).lower() and 'polygon' not in str(b[c][1]).lower():
                            self.successTxt.set('Field names or geometries could not be found in boundary file.\n'
                                                'Ensure that the column title is "geometry"')
                            return
                        self.geoCol = c
                        if type(b[c]) == str:
                            b[c] = b[c].apply(transform_geom)
                        geometries += b[c].tolist()
                    if 'calc_acres' in cols and needDesc:
                        self.geoCol = [c for c in b.columns if 'geo' in c.lower()][0] \
                            if 'geo' in ''.join([c.lower() for c in b.columns]) else [c for c in b.columns if 'poly' in c.lower()][0]
                        desc = [f'<b>Acreage:</b> {round(b.loc[i, "calc_acres"], 1)} <br>' \
                                f'<b>Permit Number:</b> {b.loc[i, "permit_num"]} <br>' \
                                f'<b>Permittee:</b> {b.loc[i, "permittee"]} <br>' \
                                f'<b>Location Narrative:</b> {b.loc[i, "loc_narr"]} <br>' \
                                f'<b>Crops:</b> {b.loc[i, "crop_list"]}' for i in b.index]
                        description+=desc
                        needDesc = False
                    elif 'calc_acres' not in cols and needDesc:
                        self.geoCol = [c for c in b.columns if 'geo' in c.lower()][0] \
                            if 'geo' in ''.join([c.lower() for c in b.columns]) else [c for c in b.columns if 'poly' in c.lower()][0]
                        tempGeos = []
                        for tg in b[self.geoCol]:
                            if type(tg) == str:
                                tempGeos.append(shapely.wkt.loads(tg))
                            else:
                                tempGeos.append(tg)
                        description+=[f'Acreage: {round(self._poly_area(tempGeos[i]), 3)}' for i in range(len(tempGeos))]
                        needDesc=False
                try:
                    fieldNames == geometries
                except:
                    self.successTxt.set('Field names or geometries could not be found in boundary file.\n'
                                        'Ensure that the column title is "geometry"')
                    return
            if self.dropBool.get():
                for i,geo in enumerate(geometries):
                    try:
                        if len(list(geo)) > 1:
                            geo = geo
                    except:
                        geo = [geo]
                    for g in geo:
                        self._create_glatlong(g)
                        if self.acreage<1:
                            continue
                        self._create_drops()
                        self.myKml.drops(self.dropPoints, fieldNames=[fieldNames[i]] * len(self.dropPoints),
                                    labels=[str(self.hivesPerDrop) + 'h'] * len(self.dropPoints))
                        self.dropsDone.append(fieldNames[i])
                        if self.hbee.get():
                            self.myKml.radius('honeybee', fieldNames[i])
                        if self.bbee.get():
                            self.myKml.radius('bumblebee', fieldNames[i])
                        self.radDone.append(fieldNames[i])
            if self.boundBool.get():
                newF, newG, newD = [], [], []
                for i in range(len(geometries)):
                    try:
                        if len(list(geometries[i])) > 1:
                            newG += list(geometries[i])
                            newF += [fieldNames[i]]*len(list(geometries[i]))
                            newD += [description[i]]*len(list(geometries[i]))
                    except:
                        newG.append(geometries[i])
                        newF.append(fieldNames[i])
                        newD.append(description[i])
                self.myKml.boundaries(newF, newG, newD)

        dropFieldNames = []
        dropCoords = []
        dropPoints = []
        hives = []
        if 'masterDrops' in dir(self):
            drops = [self.masterDrops] if type(self.masterDrops) != list else self.masterDrops
            for d in drops:
                lowerCols = [col.lower() for col in d.columns]
                for c in d.columns:
                    if 'field' in c.lower():
                        dropFieldNames += d[c].tolist()
                    if 'point' in c.lower():
                        dropPoints += d[c].tolist()
                    if 'lat' in c.lower():
                        for i in d.index:
                            dropCoords.append([d.loc[i, 'long'], d.loc[i, 'lat']])
                    if 'hive' in c.lower():
                        hives += d[c].tolist()
            hives = None if not len(hives) else hives
            if not len(dropCoords) or len(dropCoords) > len(dropPoints):
                dropCoords += [[list(c.coords)[0], list(c.coords)[1]] for c in dropPoints]

        remainingDrops = []
        for f in dropFieldNames:
            if f not in self.dropsDone:
                remainingDrops.append(f)

        if self.dropBool.get():
            self.myKml.drops(dropPoints=dropCoords, fieldNames = dropFieldNames , labels=hives)
        csvDropDf = pd.DataFrame({'coords':dropCoords,
                                  'field':remainingDrops})
        for g in csvDropDf.groupby('field'):
            g, dat = g
            if self.hbee.get():
                self.myKml.radius('honeybee',field = g, dropPoints=dat['coords'].tolist())
            if self.bbee.get():
                self.myKml.radius('bumblebee',field = g, dropPoints=dat['coords'].tolist())

        self.destLabel.grid(row = 2, column=0,pady = 10, padx = 10, sticky = 'ew')
        self.destOpen.grid(row=2, column = 1, pady=10, padx = 10, sticky = 'e')

        # if 'kmlFile' in dir(self):
        #     destFile = self.kmlFile.split('.')[0] + '_new.kml'
        # else:
        #     destFile = self.excelFile.split('.')[0] + '_new.kml'
        # destFile = destFile.replace('__' ,'_')
        # self.myKml.save(destFile)
        # self.successTxt.set(f'KML saved at {destFile}')




