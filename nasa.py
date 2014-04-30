#!/usr/bin/python

##Imports
import sys
import h5py
import xlwt
import os

##Mexico coordinates
##MEXICO = 18:23,95:106

##Creating and filling the excel file
book = xlwt.Workbook(encoding="utf-8")
 
##Ensure directory exist function
def ensure_dir(f):
    ##print(os.path.dirname(f))
    d = os.path.dirname(f)
    if not os.path.exists(d):
        os.makedirs(d)

##Write excel function
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
    ##print(nombre)
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
        ##book.save(f.filename + ".xls")
       
        nombre2 = nombre[nombre.rfind("MOD04"):]
        ## print(nombre2)
        temp = list(nombre2)
        ##print(temp)
        year = temp[10] + temp[11] + temp[12] + temp[13]
        day  = temp[14] + temp[15] + temp[16]

        ##print(year)
        ##print(day)
        if(year == "2000" or year == "2004" or year == "2008" or year == "2012" or year == "2016" or year == "2020"):
            ##print("bisiesto")
            if(day >= "001" and day <= "031"):
                month = "january"
            
            if(day > "031" and day <= "060"):
                month = "february"
            
            if(day > "060" and day <= "091"):
                month = "march"
            
            if(day > "091" and day <= "121"):
                month = "april"
            
            if(day > "121" and day <= "152"):
                month = "may"
            
            if(day > "152" and day <= "182"):
                month = "june"
            
            if(day > "182" and day <= "213"):
                month = "july"
            
            if(day > "213" and day <= "244"):
                month = "august"
            
            if(day > "244" and day <= "274"):
                month = "september"
            
            if(day > "274" and day <= "305"):
                month = "october"
            
            if(day > "305" and day <= "335"):
                month = "november"
            
            if(day > "335" and day <= "366"):
                month = "december"
            
        ##2001-2003, 2005-2007, 2009-2011, 2013-2015
        if(year == "2001" or year == "2002" or year == "2003" or year == "2005" or year == "2006" or year == "2007" or year == "2009" or year == "2010" or year == "2011" or year == "2013" or year == "2014" or year == "2015"):
            ##print("normal");
            if(day >= "001" and day <= "031"):
                month = "january"
            
            if(day > "031" and day <= "059"):
                month = "february"
            
            if(day > "059" and day <= "090"):
                month = "march"
            
            if(day > "090" and day <= "120"):
                month = "april"
            
            if(day > "120" and day <= "151"):
                month = "may"
            
            if(day > "151" and day <= "181"):
                month = "june"
            
            if(day > "181" and day <= "212"):
                month = "july"
            
            if(day > "212" and day <= "243"):
                month = "august"
            
            if(day > "243" and day <= "273"):
                month = "september"
            
            if(day > "273" and day <= "304"):
                month = "october"
            
            if(day > "304" and day <= "334"):
                month = "november"
            
            if(day > "334" and day <= "365"):
                month = "december"
    
        ##print(os.getcwd())
        ensure_dir("./RESULTS/year/" + year + "/" + month + "/" +  f.filename + ".xls")
        book.save ("./RESULTS/year/" + year + "/" + month + "/" +  f.filename + ".xls")

        if(month == "march" or month == "april" or month == "may"):
            ensure_dir("./RESULTS/season/" + year + "/spring/" +  f.filename + ".xls")
            book.save ("./RESULTS/season/" + year + "/spring/" +  f.filename + ".xls")
        if (month == "june" or month == "july" or month == "august"):
            ensure_dir("./RESULTS/season/" + year + "/summer/" +  f.filename + ".xls")
            book.save ("./RESULTS/season/" + year + "/summer/" +  f.filename + ".xls")
        if (month == "september" or month == "october" or month == "november"):
            ensure_dir("./RESULTS/season/" + year + "/autumn/" +  f.filename + ".xls")
            book.save ("./RESULTS/season/" + year + "/autumn/" +  f.filename + ".xls")
        if (month == "december" or month == "january" or month == "february"):
            ensure_dir("./RESULTS/season/" + year + "/winter/" +  f.filename + ".xls")
            book.save ("./RESULTS/season/" + year + "/winter/" +  f.filename + ".xls")



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
            ##book.save(f.filename + ".xls")
            

            temp = list(nombre)
            year = temp[10] + temp[11] + temp[12] + temp[13]
            day  = temp[14] + temp[15] + temp[16]

            ##print(year)
            ##print(day)
            if(year == "2000" or year == "2004" or year == "2008" or year == "2012" or year == "2016" or year == "2020"):
                ##print("bisiesto")
                if(day >= "001" and day <= "031"):
                    month = "january"
                
                if(day > "031" and day <= "060"):
                    month = "february"
                
                if(day > "060" and day <= "091"):
                    month = "march"
                
                if(day > "091" and day <= "121"):
                    month = "april"
                
                if(day > "121" and day <= "152"):
                    month = "may"
                
                if(day > "152" and day <= "182"):
                    month = "june"
                
                if(day > "182" and day <= "213"):
                    month = "july"
                
                if(day > "213" and day <= "244"):
                    month = "august"
                
                if(day > "244" and day <= "274"):
                    month = "september"
                
                if(day > "274" and day <= "305"):
                    month = "october"
                
                if(day > "305" and day <= "335"):
                    month = "november"
                
                if(day > "335" and day <= "366"):
                    month = "december"
                
            ##2001-2003, 2005-2007, 2009-2011, 2013-2015
            if(year == "2001" or year == "2002" or year == "2003" or year == "2005" or year == "2006" or year == "2007" or year == "2009" or year == "2010" or year == "2011" or year == "2013" or year == "2014" or year == "2015"):
                ##print("normal");
                if(day >= "001" and day <= "031"):
                    month = "january"
                
                if(day > "031" and day <= "059"):
                    month = "february"
                
                if(day > "059" and day <= "090"):
                    month = "march"
                
                if(day > "090" and day <= "120"):
                    month = "april"
                
                if(day > "120" and day <= "151"):
                    month = "may"
                
                if(day > "151" and day <= "181"):
                    month = "june"
                
                if(day > "181" and day <= "212"):
                    month = "july"
                
                if(day > "212" and day <= "243"):
                    month = "august"
                
                if(day > "243" and day <= "273"):
                    month = "september"
                
                if(day > "273" and day <= "304"):
                    month = "october"
                
                if(day > "304" and day <= "334"):
                    month = "november"
                
                if(day > "334" and day <= "365"):
                    month = "december"
        

            if(month == "march" or month == "april" or month == "may"):
                ensure_dir(os.getcwd() + "/RESULTS/season/" + year + "/spring/" +  f.filename + ".xls")
                book.save (os.getcwd() + "/RESULTS/season/" + year + "/spring/" +  f.filename + ".xls")
            if (month == "june" or month == "july" or month == "august"):
                ensure_dir(os.getcwd() + "/RESULTS/season/" + year + "/summer/" +  f.filename + ".xls")
                book.save (os.getcwd() + "/RESULTS/season/" + year + "/summer/" +  f.filename + ".xls")
            if (month == "september" or month == "october" or month == "november"):
                ensure_dir(os.getcwd() + "/RESULTS/season/" + year + "/autumn/" +  f.filename + ".xls")
                book.save (os.getcwd() + "/RESULTS/season/" + year + "/autumn/" +  f.filename + ".xls")
            if (month == "december" or month == "january" or month == "february"):
                ensure_dir(os.getcwd() + "/RESULTS/season/" + year + "/winter/" +  f.filename + ".xls")
                book.save (os.getcwd() + "/RESULTS/season/" + year + "/winter/" +  f.filename + ".xls")

            ensure_dir(os.getcwd() + "/RESULTS/year/" + year + "/" + month + "/" +  f.filename + ".xls")
            book.save (os.getcwd() + "/RESULTS/year/" + year + "/" + month + "/" +  f.filename + ".xls")

            ##Closing the file
            f.close()
        else:
            print("Unexpected filetype, please use .hdf or .h5 filetypes")
else:
    print("Usage: ./nasa.py <filename>")