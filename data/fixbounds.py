import shapely.wkt
import pandas as pd
import pyproj
from shapely.geometry import Polygon
from shapely import ops
import os
import numpy as np

def transform_geom(currGeom):
    pgon = shapely.wkt.loads(currGeom)
    wgs84 = pyproj.CRS('EPSG:3310')
    utm = pyproj.CRS('EPSG:4326')
    project = pyproj.Transformer.from_crs(wgs84, utm, always_xy=True).transform
    pgon = ops.transform(project, pgon)

    return pgon

for county in os.listdir(r'C:\Users\marcu\PycharmProjects\EXB_GUI\data\oldBoundaries'):
    if '.py' in county:
        continue

    if county in os.listdir('boundaries'):
        continue

    currData = pd.read_csv(
        rf'C:\Users\marcu\PycharmProjects\EXB_GUI\data\oldBoundaries\{county}\crops_{county.replace(" ", "_")}.csv')
    print(county)
    currData['geometry'] = currData['geometry'].apply(transform_geom)

    os.mkdir(rf'C:\Users\marcu\PycharmProjects\EXB_GUI\data\boundaries\{county}')

    currData.to_csv(rf'C:\Users\marcu\PycharmProjects\EXB_GUI\data\boundaries\{county}\crops_{county.replace(" ", "_")}.csv', index=False)
