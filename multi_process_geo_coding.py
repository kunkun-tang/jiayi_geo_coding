import googlemaps
from datetime import datetime
import pprint
import xlwings as xw
import time
from xlwings import Range
from multiprocessing import Lock, Process, Queue, current_process

gmaps = googlemaps.Client(key='AIzaSyDIm37RAL1x7_IiAAUEYYJh9Yoh5phcdFI')

# change excel file path
wb = xw.Book('/Users/latang/Downloads/1995missed.xlsx')  # connect to an existing file in the current working directory
sht = wb.sheets['Sheet1']

# specify which column should be placed for latitude
lat_column = 'S'
# specify which column should be placed for longitude
lng_column = 'T'

# sepcify which column should be placed for full address
full_addr_column = 'P'

sht.range(lat_column + '1').value = 'lat'
sht.range(lng_column + '1').value = 'lng'

rng = sht.range('A1')

# This gives the last row in column A
# last_row_num = rng.end('down').row
last_row_num = 50

queue = Queue()
def do_job(tasks_to_accomplish, tasks_that_are_done):
    while True:
        try:
            '''
                try to get task from the queue. get_nowait() function will 
                raise queue.Empty exception if the queue is empty. 
                queue(False) function would do the same task also.
            '''
            task = tasks_to_accomplish.get_nowait()
        except queue.empty():
            break
        else:
            '''
                if no exception has been raised, add the task completion 
                message to task_that_are_done queue
            '''
            rownum_str = task
            # change excel file path
            wb = xw.Book('/Users/latang/Downloads/1995missed.xlsx')  # connect to an existing file in the current working directory
            sht = wb.sheets['Sheet1']

            full_addr = sht.range(full_addr_column + rownum_str).value
            pprint.pprint(full_addr)
            try:
                geocode_result = gmaps.geocode(full_addr)
                loc = geocode_result[0]['geometry']['location']
                lat = loc['lat']
                lng = loc['lng']
                sht.range(lat_column + rownum_str).value = lat
                sht.range(lng_column + rownum_str).value = lng
            except:
                print('google geocoding API not returning correctly.')

            tasks_that_are_done.put(task + ' is done by ' + current_process().name)
    return True


def main():
    number_of_task = last_row_num
    number_of_processes = 4
    tasks_to_accomplish = Queue()
    tasks_that_are_done = Queue()
    processes = []

# from second A2 to last row
    for i in range(2, last_row_num + 1):
        tasks_to_accomplish.put(str(i))

    # creating processes
    for w in range(number_of_processes):
        p = Process(target=do_job, args=(tasks_to_accomplish, tasks_that_are_done))
        processes.append(p)
        p.start()

    # completing process
    for p in processes:
        p.join()

    # print the output
    while not tasks_that_are_done.empty():
        print(tasks_that_are_done.get())

    return True


if __name__ == '__main__':
    main()