import googlemaps
import pprint
import xlwings as xw
import time

gmaps = googlemaps.Client(key='AIzaSyDIm37RAL1x7_IiAAUEYYJh9Yoh5phcdFI')

# change excel file path
wb = xw.Book('/Users/latang/Downloads/1994missed.xlsx')  # connect to an existing file in the current working directory
sht = wb.sheets['Sheet1']

# specify which column should be placed for latitude
lat_column = 'M'
# specify which column should be placed for longitude
lng_column = 'N'

sht.range(lat_column + '1').value = 'lat'
sht.range(lng_column + '1').value = 'lng'

rownum = 4941
while (sht.range('A'+str(rownum)).value != None):
    rownum_str = str(rownum)

    try:
        #specify which column for street name and city name
        street_and_city = sht.range('C' + rownum_str + ':D' + rownum_str).value;

        #specify which column for state
        state = sht.range('F' + rownum_str).value;
        full_addr_list = [x.encode('UTF8') for x in street_and_city] + [state.encode('UTF8')]

        full_addr = ", ".join(full_addr_list)
        pprint.pprint(full_addr)
    except:
        print('something converting not correctly.')

    try:
        geocode_result = gmaps.geocode(full_addr)
        loc = geocode_result[0]['geometry']['location']
        lat = loc['lat']
        lng = loc['lng']
        sht.range(lat_column + rownum_str).value = lat
        sht.range(lng_column + rownum_str).value = lng
    except:
        print('google geocoding API not returning correctly.')

    rownum += 1



