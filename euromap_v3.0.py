#! -*- coding: utf8 -*-
import csv, xlrd
import datetime
import io, re, os, sys


def get_data(fpath):
    offset = 2
    
    workbook = xlrd.open_workbook(fpath)
    worksheet = workbook.sheet_by_index(0)    
    
    rows = []
    for i, row in enumerate(range(worksheet.nrows)):
        if i <= offset:
            continue
        r = []
        for j, col in enumerate(range(worksheet.ncols)):
            if worksheet.cell_type(i,j) == 1: # if cell contains text
                r.append(worksheet.cell_value(i,j).encode('utf8'))
            elif worksheet.cell_type(i,j) == 3: # if cell contains a datetime object
                date = xlrd.xldate.xldate_as_datetime(worksheet.cell_value(i,j), workbook.datemode)
                if date.year == 1899:
                    date = str(date.time().strftime("%H:%M"))
                else:
                    date = str(date.day)+'.'+str(date.month)+'.'+str(date.year)
                r.append(date)
            else: # if cell contains other objects (mainly numbers)
                r.append(worksheet.cell_value(i,j))
        rows.append(r)
    return rows

def replace_all(text, dic):
    for i, j in dic.items():
        j = str(j)
        
        if "<" in j:
            j = j.replace('<', '&#60;')
            j = j.replace(',', '.')
        if r'\xc2\xb' in j:
            j = j.replace(r'\xc2\xb', 'deg')
        
        text = text.replace(i, j)
    return text
    
def make_kml(fname, outname):
    data = get_data(fname)
    done = []
    
    
    for i in range(len(data)):
        if data[i] == [u'']*25:
            continue
        
        if type(data[i][15]) is str:
            if "," in str(data[i][15]):
                data[i][15] = data[i][15].replace(",",".")
        
        template = "template.kml"
        template = open(os.path.join(basepath,template), "r")
        kml = open(outname, "a")          

        unit_corr = 1

        if "µBq" in str(data[i][17]):
            unit_corr = 1000
        elif "mBq" in str(data[i][17]):
            unit_corr = 1
        elif str(data[i][17]).split("/") == "Bq":
            unit_corr = 0.001
          
        same_name = []
        a = 0
        if data[i][1] not in done:
            done.append(data[i][1])
            same_name.append(i)
            
            b = 1
            if type(data[i][15]) is float:
                a = i            
            
            for j in range(len(data)):
                if i != j:
                    if data[j][1] == data[i][1]:
                        same_name.append(j)
                        if type(data[j][15]) is (float or int):
                            if a == 0:
                                a = j
                            if data[a][15] < data[j][15]:
                                a = j
            if a != i:
                tmp = data[a]
                data[a] = data[max(same_name)]
                data[max(same_name)] = tmp
        
        try:
            if "<" in str(data[i][15]):
                color = "01ffffff"
            else:
                z = int((data[i][15]/unit_corr)/colorn)
                if z < 256:
                    z = '{:02x}'.format(z)
                    color = "6414ff"+str(z)
                elif z > 256 and z < 512:
                    z = (z-512)*(-1)
                    z = '{:02x}'.format(z)
                    color = "6414"+str(z)+"ff"
                elif z > 512:
                    color = "641400FF"                        
        except TypeError:
            continue                
        
        repl = {"COUNTRY":data[i][0], "NAME":data[i][1], "LAT":data[i][2], "LONG":data[i][3],
                "ALT":data[i][4], "FRACTION":data[i][5], "STARTDATE":data[i][6],
                "STARTHOUR":data[i][7], "ENDDATE":data[i][8], "ENDHOUR":data[i][9],
                "TSTAMP":data[i][10], "VOLUME":data[i][11], "NORM":data[i][12],
                "TREF":data[i][13], "NUCLIDE":data[i][14], "ACTCONCENT":data[i][15],
                "UNCERT":data[i][16], "UNIT":data[i][17], "STANDDEV":data[i][18],
                "COINCORR":data[i][19], "PARDAUGH":data[i][20], "THRESH":data[i][21],
                "MESSTAND":data[i][22], "REFDATE":data[i][23], "COMMENT":data[i][24],
                "COLOR":color}
            
        for line in template:
            kml.write(replace_all(line, repl))
        kml.close()
            
def init_kml(fname):
    kml = open(fname, "w")
    kml.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    kml.write('<kml xmlns="http://www.opengis.net/kml/2.2">\n')
    kml.write('\t<Document>\n')
    kml.write('\t\t<name>Map_'+str(fname)+'</name>\n')
    kml.write('\t\t<description/>\n')
    kml.close()
    
def end_kml(fname):
    kml = open(fname, "a")
    kml.write('\t</Document>\n')
    kml.write('</kml>\n')
    kml.close()    


# MAIN PROGRAM
basepath = os.path.dirname(sys.executable)
activity_limit = 0.0

lambda: os.system('cls')
print("--- Welcome euromap_v3.0 ---\n Created by the Swiss OFSP, Section for Environmental Radiation.  \n contact: Rytz Lukas - Lukas.Rytz@bag.admin.ch \n\n This application generates KML Files out of the Ro5 Excel sheets. This lets you create maps of radiation measurements. At the moment, the KML files can only be used with Google Maps or Google Earth. Please read the Readme.txt for further information. \n\n ")

while activity_limit == 0.0:
    try:
        activity_limit = float(input('The RedLine will be used for color coding the Placemarks.\n The Placemarks will be colored from green (0 Bq) to red (RedLine) in 512 color steps. \n\n Please enter a value for RedLine in [mBq/m^3]:\n')) #in [mBq/m^3]
    except ValueError:
        print("ERROR: The Value you entered is not a number!")

colorn = activity_limit/512
source = "Excel"
source = os.path.join(basepath,source)
sink = "Kml"
sink = os.path.join(basepath,sink)

#os.chdir(basepath)
files = os.listdir(source)
targets = []

os.chdir(basepath)
for name in files:
    if ".xlsx" in name:
        targets.append(name)

print("Creating KML files...")
i = 0

for target in targets:
    i+=1
    outname = str(target).split('.')[0]+".kml"
    outname = os.path.join(basepath,sink,outname)
    target = os.path.join(basepath,source,target)
    
    init_kml(outname)
    make_kml(target, outname)
    end_kml(outname)

print("Parsing done. "+str(i)+" KML files have been created")
a = input('Press Enter to exit...')
    
