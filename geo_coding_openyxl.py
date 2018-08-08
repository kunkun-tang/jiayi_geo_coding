import googlemaps
import pprint
import time
from openpyxl import Workbook
from openpyxl import load_workbook


gmaps = googlemaps.Client(key='AIzaSyDIm37RAL1x7_IiAAUEYYJh9Yoh5phcdFI')

# change excel file path
wb2 = load_workbook('1994missed.xlsx')
sht = wb2['Sheet1']
# specify which column should be placed for latitude
lat_column = 'M'
# specify which column should be placed for longitude
lng_column = 'N'

sht[lat_column + '1'] = 'lat'
sht[lng_column + '1'] = 'lng'

rownum = 6563
while (sht['A'+str(rownum)].value != None):
    rownum_str = str(rownum)

    try:
        full_addr = sht['J' + rownum_str].value;
        pprint.pprint(full_addr)
    except:
        print('something converting not correctly.')

    try:
        geocode_result = gmaps.geocode(full_addr)
        loc = geocode_result[0]['geometry']['location']
        lat = loc['lat']
        lng = loc['lng']
        sht[lat_column + rownum_str] = lat
        sht[lng_column + rownum_str] = lng
    except:
        print('google geocoding API not returning correctly.')

    rownum += 1

wb2.save(filename = "1994missed_added.xlsx")

