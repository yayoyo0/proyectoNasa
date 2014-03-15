#!/usr/bin/python

##Imports
import sys
import h5py
import xlwt

if str(sys.argv[1]) == "-h" or str(sys.argv[1]) == "--help":
    print("Usage: nasa.py <filename>")
else:
    ##Obtaining the filename from the console parameters
    f = h5py.File(sys.argv[1])

    ##Entering to the folder that contains dataset
    mod04 = f['mod04']

    ##Tests for the file
    ##print(f.filename)
    ##print(f.file)
    ##print(f.keys())
    ##print(mod04.name)
    ##print(mod04.keys)
    ##print(mod04.values)
    ##print(mod04.items)
    ##print(mod04.iterkeys)

    ##Enters the geolocation folder
    geo = mod04.get('Geolocation Fields')

    ##Get the latitude dataset
    lat = geo.get('Latitude')

    ##The shape returns how many rows/columns is stored in the dataset
    ##print(lat.shape)

    ##Retriving the data from the dataset
    data = lat[0:202,0:134]

    ##To check the data obtained
    ##print(data)


    ##Creating and filling the excel file
    book = xlwt.Workbook(encoding="utf-8")

    sheet1 = book.add_sheet(f.filename)

    i = 0
    j = 0
    for n in data:
        for m in n:
            ##print("i " + str(i) + " j " + str(j))
            sheet1.write(j, i, float(m))
            i = i+1
        i = 0
        j = j+1

    book.save(f.filename + ".xls")

    ##Closing the file
    f.close()

    print("Finished OK")
