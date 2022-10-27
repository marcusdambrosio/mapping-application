import pandas as pd
import numpy as np
import folium
import sys, os , time
import shapely.wkt
import pyproj
import simplekml
from shapely import ops
from shapely.geometry import Point , Polygon, LineString, MultiLineString
import matplotlib.pyplot as plt
import tkinter
from tkinter import ttk
from tkinter import filedialog as fd
from polycircles import polycircles


class MapMakerCSV:
    def __init__(self):
        # self.root = tkinter.Toplevel(self.root)
        self.root = tkinter.Tk()
        self.root.title('Premap Drops')
        self.root.resizable(False, False)
        self.root.geometry('800x300')

        self.boundaryTxt = tkinter.StringVar()
        self.boundaryTxt.set('Select a Boundary File...')

        self.boundaryLabel = tkinter.Label(self.root, text='Select a Boundary File...', bd=5, bg='lightgrey', width=50,
                                           pady=5,
                                           anchor='w', relief='groove', textvariable=self.boundaryTxt)
        self.boundaryLabel.grid(row=2, column=0, pady=10, padx=10)
        self.boundaryLabelPos = self.boundaryLabel.grid_info()

        self.boundaryOpen = ttk.Button(self.root, text='Browse', command=self.boundary_file)
        self.boundaryOpen.grid(row=2, column=1, pady=10)
        self.boundaryOpenPos = self.boundaryOpen.grid_info()

        self.runButton = tkinter.Button(self.root, text='RUN', command=self.run)
        self.runButton.grid(row=4, columnspan=4, sticky='ew')
        self.successTxt = tkinter.StringVar()
        self.successLabel = tkinter.Label(self.root, textvariable=self.successTxt, bg='lightgreen')
        self.successLabel.grid(row=5, sticky='w')
        self.successLabel.config(font=('Helvetica bold', 12))

        self.root.mainloop()


    def boundary_file(self):
        filename = fd.askopenfilename(title='Select Boundary File', initialdir='/')
        self.boundaryFile = filename
        self.boundaryTxt.set(filename)
        return filename

    def run(self):
        if '.xlsx' in self.boundaryFile:
            self.boundaryData = pd.read_excel(self.boundaryFile)
        elif '.csv' in self.boundaryFile:
            self.boundaryData = pd.read_csv(self.boundaryFile)
        else:
            self.successTxt.set('Invalid file type. Must be .xlsx or .csv')
            self.successLabel.config(bg = 'red')
            return
        geoCol = False
        fieldCol = False
        hivesCol = False
        geoT=['geo', 'poly', 'bound']
        fieldT = ['field', 'name', 'loc', 'id']
        for col in self.boundaryData.columns:
            for gt in geoT:
                if gt in col.lower():
                    geoCol = col

            for ft in fieldT:
                if ft in col.lower():
                    fieldCol = col

            if 'hive' in col:
                hivesCol = col

        errors = ''
        errorMsg = {'geo': 'Boundary column not found in spreadsheet. Please title the boundary column "geometry".',
                    'field': 'Field name column not found in spreadsheet. Please title the field name column "fields".',
                    'hives': 'Hive count column not found in spreadsheet. Please title the hive count column "hives".'
                    }
        if not geoCol:
            errors = errors + errorMsg['geo'] + '\n'
        if not fieldCol:
            errors = errors + errorMsg['field'] + '\n'
        if not hivesCol:
            errors = errors + errorMsg['hives'] + '\n'

        if errors != '':
            self.successTxt.set(errors)
            self.successLabel.config(bg = 'red')
            return

        self.geom = self.boundaryData[geoCol]
        if type(self.geom[0]) == str:
            self.geom = self.geom.apply(self._transform_geom)

        self.fieldNames = self.boundaryData[fieldCol]
        self.hiveCounts = self.boundaryData[hivesCol]

        self.master = pd.DataFrame(columns = ['long', 'lat', 'hives', 'field'])
        
        for i in range(len(self.geom)):
            currDf = pd.DataFrame()
            self._create_glatlong(self.geom[i])
            self._create_drops(self.hiveCounts[i])
            currDf['long'] = [i for [i,j] in self.dropPoints]
            currDf['lat'] = [j for [i,j] in self.dropPoints]
            currDf['hives'] = [self.hivesPerDrop] * len(self.dropPoints)
            currDf['field'] = [self.fieldNames[i]] * len(self.dropPoints)
            self.master = self.master.append(currDf)

        self.csvPath = self.boundaryFile.split('.')[0] + '_drops.csv'
        self.master.to_csv(self.csvPath, index = False)
        self.successTxt.set(f'Drop file saved at {self.csvPath}')
        self.successLabel.config(bg = 'lightgreen')

        self.kmlButton = tkinter.Button(self.root, text = 'Generate KML', command = self._create_kml)
        self.kmlButton.grid(row=6, columnspan= 4, sticky = 'ew')
        self.kmlTxt = tkinter.StringVar()
        self.kmlLabel = tkinter.Label(self.root, text = '', textvariable = self.kmlTxt, bg= 'lightgreen', font= ('Helvetica bold', 12))
        self.kmlLabel.grid(row = 7, sticky = 'w')

    def _transform_geom(self, currGeom):
        pgon = shapely.wkt.loads(currGeom)
        wgs84 = pyproj.CRS('EPSG:3310')
        utm = pyproj.CRS('EPSG:4326')
        project = pyproj.Transformer.from_crs(wgs84, utm, always_xy=True).transform
        pgon = ops.transform(project, pgon)
        return pgon

    def _create_drops(self, numHives, hivesPerDrop = 12):
        numDrops = round(int(numHives)/int(hivesPerDrop))
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
            if currDist >= step - .005 * step:
                self.dropPoints.append([long, lat])
                currDist = 0
        if len(self.dropPoints) != numDrops:
            print(f'NUMDROPS {numDrops} DOES NOT EQUAL DROPPOINTS {len(self.dropPoints)}')

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
        step = .00005 #about 5 meters

        for i, latlong in enumerate(self.glatlongs):
            distance = ((latlong[0] - self.glatlongs[i - 1][0]) ** 2 + (latlong[1] - self.glatlongs[i - 1][1]) ** 2) ** .5
            newPoints = int(np.ceil(distance / step))
            discLats += list(np.linspace(self.glatlongs[i - 1][0], latlong[0], newPoints))
            discLongs += list(np.linspace(self.glatlongs[i - 1][1], latlong[1], newPoints))
        for i in range(len(discLats)):
            self.discLL.append([discLats[i], discLongs[i]])

    def _create_kml(self):
        bees = {'honeybee': [1000 * .343, simplekml.Color.yellow],
                'bumblebee': [800 * .343, simplekml.Color.green]}
        self.beeType = ['honeybee', 'bumblebee']
        if self.beeType == str:
            self.beeType = [self.beeType]

        self.kml = simplekml.Kml(name = 'temp farm name')
        radFolder = self.kml.newfolder(name=f'{type} Flight Radius')
        hFolder = radFolder.newfolder(name = 'honeybees')
        bFolder = radFolder.newfolder(name = 'bumblebees')

        try:
            self.dropData = pd.read_csv(self.csvPath)
        except:
            self.kmlTxt.set(f'Cannot find or load file at {self.csvPath}')
            self.kmlLabel.config(bg='red')
            return
        try:
            for i in range(len(self.geom)):
                folder = self.kml.newfolder(name = self.fieldNames[i])
                currGeo = folder.newlinestring(name = str(self.fieldNames[i]) + ' boundary')
                currGeo.style.linestyle.color = simplekml.Color.blue
                currGeo.style.linestyle.width = 5
                currGeo.coords = self.geom[i].exterior.coords

                currDropData = self.dropData[self.dropData['field'] == self.fieldNames[i]]
                for row in currDropData.iterrows():
                    i, row = row
                    coords = [[row['long'], row['lat']]]
                    currPoint = folder.newpoint(name=str(row['hives']) + 'h', coords=coords)
                    currPoint.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png'
                    #Radius
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
            self.kmlTxt.set(f'KML file successfully saved at {self.kmlPath}')
            self.kmlLabel.config(bg='lightgreen')
        except Exception as e:
            self.kmlTxt.set('Issue creating oldBoundaries and drops in KML file \n' + str(e))
            self.kmlLabel.config(bg = 'red')
            return

    def _radius(self, type = 'honeybee'):
        bees = {'honeybee': [1000*.343, simplekml.Color.yellow],
                'bumblebee': [800*.343, simplekml.Color.green]}

        rad, color = bees[type]
        for i, drop in enumerate(self.dropPoints):
            outerCircle = polycircles.Polycircle(latitude=drop[0],
                                             longitude=drop[1],
                                             radius=rad,
                                             number_of_vertices=36)
            innerCircle = polycircles.Polycircle(latitude=drop[0],
                                             longitude=drop[1],
                                             radius=rad-5,
                                             number_of_vertices=36)
            poly = folder.newpolygon(name = f'{type} radius at ({drop[0]}, {drop[1]})',
                                       innerboundaryis = innerCircle.to_kml(),
                                       outerboundaryis = outerCircle.to_kml())
            poly.style.polystyle.color = color



