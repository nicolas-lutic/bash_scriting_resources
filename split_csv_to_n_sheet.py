#!/usr/bin/python2.7
import csv
import sys
import os
import shutil
import xlwt
from datetime import datetime
#=========================================================================
# Definition des des variables temporelles
#=========================================================================

def cutfile( filename, out_file_name, column_cutting_index ):

    if(os.stat(filename).st_size==0):
        return

    with open(filename, 'rb') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=';',quotechar='"', quoting=csv.QUOTE_MINIMAL)
        prev_file_name=""
        b_first_line=True
        a_column_name=""
        wb = xlwt.Workbook()
        rowi = 1
        nb_sheet=0
        for row in spamreader:
            if (b_first_line):
                b_first_line=False
                a_column_name = row
                continue
            if (prev_file_name!=row[column_cutting_index]):
                #create new sheet
                sheetname=row[column_cutting_index]
                if (len(row[column_cutting_index]) > 30 ):
                    sheetname = row[column_cutting_index][0:25]+str(nb_sheet)
                prev_file_name=row[column_cutting_index]
                ws = wb.add_sheet(sheetname)
                rowi = 1
                nb_sheet+=1
            for coli, value in enumerate(row):
                    ws.write(rowi,coli,value)
            rowi +=1
        wb.save(out_file_name+".xls")
                  
   

if __name__ == "__main__":
    if len( sys.argv ) <> 4:
        print 'usage : 1 file name '
        print 'usage : 2 output file name '
        print 'usage : 3 index of the cutting column to create new sheet  '
        sys.exit( 1 )
    else:
        cutfile (sys.argv[1],sys.argv[2],int(sys.argv[3]))


#with open(vlogdir+"HCHC-HCHP-"+vDATE+".csv", "a") as hchpfile:  
  #                               hchpfile.write("'"+vDATE +" "+ vHEURE+"';"+ data["HCHC"]+";"+ data["HCHP"]+"\r\n")
#=========================================================================
