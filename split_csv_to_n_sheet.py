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

def cutfile( filename, outFileName, columnCuttingIndex ):

    if(os.stat(filename).st_size==0):
        return

    with open(filename, 'rb') as csvfile:
        spamreader = csv.reader(csvfile, delimiter=',',quotechar='"', quoting=csv.QUOTE_MINIMAL)
        prevFileName=""
        bFirstLine=True
        aColumnName=""
        wb = xlwt.Workbook()
        rowi = 1
        nb_sheet=0
        for row in spamreader:
            if (bFirstLine):
                bFirstLine=False
                aColumnName = row
                continue
            if (prevFileName!=row[columnCuttingIndex]):
                #create new sheet
                sheetname=row[columnCuttingIndex]
                if (len(row[columnCuttingIndex]) > 30 ):
                    sheetname = row[columnCuttingIndex][0:25]+str(nb_sheet)
                prevFileName=row[columnCuttingIndex]
                ws = wb.add_sheet(sheetname)
                rowi = 1
                nb_sheet+=1
            for coli, value in enumerate(row):
                ws.write(rowi,coli,value.decode("utf8"))
            rowi +=1
        wb.save(outFileName+".xls")

if __name__ == "__main__":
    if len( sys.argv ) <> 4:
        print 'usage : 1 file name '
        print 'usage : 2 output file name '
        print 'usage : 3 index of the cutting column to create new sheet  '
        sys.exit( 1 )
    else:
        cutfile (sys.argv[1],sys.argv[2],int(sys.argv[3]))
