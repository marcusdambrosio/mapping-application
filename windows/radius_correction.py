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
from zipfile import ZipFile
import traceback
from pyproj import Geod


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
        self.flightKml.save(self.filepath.split('.')[0] + '_flightrad.kml')



FlightRadiusFixer(r'C:\Users\marcu\Documents\EXB\Sandridge_maps.kmz')
