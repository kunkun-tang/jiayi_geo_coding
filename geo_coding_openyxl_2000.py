import googlemaps
import pprint
import time
from openpyxl import Workbook
from openpyxl import load_workbook


gmaps = googlemaps.Client(key='AIzaSyDT1QvcZS-rdZgb-nWdj2TBxl-4H6vm0F0')

# change excel file path
wb2 = load_workbook('2000missed.xlsx')
sht = wb2['Sheet1']
# specify which column should be placed for latitude
lat_column = 'R'
# specify which column should be placed for longitude
lng_column = 'S'
# specify which column should be placed for full address
full_addr_column = 'O'



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

wb2.save(filename = "2000missed_added.xlsx")

