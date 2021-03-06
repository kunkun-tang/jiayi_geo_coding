import googlemaps
import pprint
import time
from openpyxl import Workbook
from openpyxl import load_workbook


gmaps = googlemaps.Client(key='AIzaSyDIm37RAL1x7_IiAAUEYYJh9Yoh5phcdFI')

# change excel file path
wb2 = load_workbook('1998missed.xlsx')
sht = wb2['Sheet1']
# specify which column should be placed for latitude
lat_column = 'S'
# specify which column should be placed for longitude
lng_column = 'T'
# specify which column should be placed for full address
full_addr_column = 'N'



sht[lat_column + '1'] = 'lat'
sht[lng_column + '1'] = 'lng'

rownum = 2
while (sht['A'+str(rownum)].value != None):
    rownum_str = str(rownum)

    try:
        full_addr = sht[full_addr_column + rownum_str].value;
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

wb2.save(filename = "1998missed_added.xlsx")

