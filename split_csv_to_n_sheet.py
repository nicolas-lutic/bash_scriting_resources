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
        bFirstLine=True
        aColumnName=""
        wb = xlwt.Workbook()
        sheetRows={}
        for row in spamreader:
            if (bFirstLine):
                bFirstLine=False
                aColumnName = row
                continue
            sheetName=row[columnCuttingIndex]
            # this limit is from the max size of a sheetName
            if (len(row[columnCuttingIndex]) > 30 ):
                sheetName = row[columnCuttingIndex][0:25]+str(nb_sheet)
            if(sheetRows.has_key(sheetName)==False):
                  sheetRows[sheetName]=[]
            sheetRows[sheetName].append(row)
        
        for sheetName, rows in sorted(sheetRows.iteritems()):
            ws = wb.add_sheet(sheetName)
            for coli, value in enumerate(aColumnName):
                ws.write(0,coli,value.decode("utf8"))         
            for index, row in enumerate(rows):
                for coli, value in enumerate(row):
                    ws.write(index+1,coli,value.decode("utf8")) 

        wb.save(outFileName+".xls")

if __name__ == "__main__":
    if len( sys.argv ) <> 4:
        print 'usage : 1 file name '
        print 'usage : 2 output file name '
        print 'usage : 3 index of the cutting column to create new sheet  '
        sys.exit( 1 )
    else:
        cutfile (sys.argv[1],sys.argv[2],int(sys.argv[3]))
