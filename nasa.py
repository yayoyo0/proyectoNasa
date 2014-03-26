#!/usr/bin/python

##Imports
import sys
import h5py
import xlwt
import os


##Creating and filling the excel file
book = xlwt.Workbook(encoding="utf-8")
        
def writeExcel(data, SheetName):
    sheet1 = book.add_sheet(SheetName)

    i = 0
    j = 0
    for n in data:
        for m in n:
            sheet1.write(j, i, float(m))
            i = i+1
        i = 0
        j = j+1
    
if len(sys.argv) == 2:
    nombre = sys.argv[1]
    lenNombre = len(sys.argv[1])
    if nombre.endswith(".hdf"):
        os.system("./converter/bin/h4toh5 " + nombre)
        temp = list(nombre)
        temp[lenNombre - 2] = '5'
        temp[lenNombre - 1] = ""
        nombre = "".join(temp)
        
        ##Obtaining the filename from the console parameters
        f = h5py.File(nombre)

        ##Entering to the folder that contains dataset
        mod04 = f['mod04']

        ##Enters the data fields folder
        dat = mod04.get('Data Fields')

        ##Working with Solar Azimuth
        solar = dat.get('Solar_Azimuth')
        dataSolar = solar[18:23,95:106]         ##Retriving the data from the dataset using Mexico coordinates
        
        ##Working with ODRSLAO
        ODRSLAO = dat.get('Optical_Depth_Ratio_Small_Land_And_Ocean')
        dataODRSLAO = ODRSLAO[18:23,95:106]
        ##Working with IODLAO 
        IODLAO = dat.get('Image_Optical_Depth_Land_And_Ocean')
        dataIODLAO = IODLAO[18:23,95:106] 
        ##Working with Scatering Angle 
        SA = dat.get('Scattering_Angle')
        dataSA= SA[18:23,95:106]
        ##Working with Solar Zenith 
        SZ = dat.get('Solar_Zenith')
        dataSZ = SZ[18:23,95:106]
        ##Working with Aerosol Type Land 
        ATL = dat.get('Aerosol_Type_Land')
        dataATL = ATL[18:23,95:106] 
        ##Working with Fitting Error Land 
        FEL= dat.get('Fitting_Error_Land')
        dataFEL = FEL[18:23,95:106]
        ##Working with Correct_Optical_Depth_Land_wav2p1 
        CODLw = dat.get('Corrected_Optical_Depth_Land_wav2p1')
        dataCODLw = CODLw[18:23,95:106] 
        ##Working with MCL 
        MCL = dat.get('Mass_Concentration_Land')
        dataMCL = MCL[18:23,95:106]
        ##Working with AEL 
        AEL = dat.get('Angstrom_Exponent_Land')
        dataAEL = AEL[18:23,95:106] 
        ##Working with DBAODL
        DBAODL = dat.get('Deep_Blue_Aerosol_Optical_Depth_550_Land')
        dataDBAODL = DBAODL[18:23,95:106] 
        ##Working with DBAODLS
        DBAODLS = dat.get('Deep_Blue_Aerosol_Optical_Depth_550_Land_STD')
        dataDBAODLS = DBAODLS[18:23,95:106] 
        ##Working with SA 
        SeA = dat.get('Sensor_Azimuth')
        dataSeA = SeA[18:23,95:106] 
        ##Working with SZ
        SeZ = dat.get('Sensor_Zenith')
        dataSeZ = SeZ[18:23,95:106] 
        ##Working with ODLAO 
        ODLAO = dat.get('Optical_Depth_Land_And_Ocean')
        dataODLAO = ODLAO[18:23,95:106] 
        
                
        ##Write data to Excel
        writeExcel(dataSolar,"Solar Azimuth")
        writeExcel(dataODRSLAO,"ODRSLAO")
        writeExcel(dataIODLAO,"IODLAO")
        writeExcel(dataSA,"Scattering Angle")
        writeExcel(dataSZ,"Solar Zenith")
        writeExcel(dataATL,"Aerosol Type Land")
        writeExcel(dataFEL,"Fitting Error Land")
        writeExcel(dataCODLw,"CODL wav2p1")
        writeExcel(dataMCL,"Mass Concentration Land")
        writeExcel(dataAEL,"Angstrom Exponent Land")
        writeExcel(dataDBAODL,"DBAODL 550")
        writeExcel(dataDBAODLS,"DBAODL 550 STD")
        writeExcel(dataSeA,"Sensor Azimuth")
        writeExcel(dataSeZ,"Sensor Zenith")
        writeExcel(dataODLAO,"Optical Depth Land And Ocean")
        
        ##Save Excel book
        book.save(f.filename + ".xls")

        ##Closing the file
        f.close()
        ##Remove converted file
        os.system("rm -f " + nombre)
        print("Finished OK")
    else:
        if nombre.endswith(".h5"):
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

            sheet1 = book.add_sheet("Latitude")

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
        else:
            print("Unexpected filetype, please use .hdf or .h5 filetypes")
else:
    print("Usage: ./nasa.py <filename>")