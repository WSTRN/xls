import sys

import pandas as pd
import xlrd
from numpy import int32

from xls2 import Xls2


class Xls2CMid(Xls2):

    def func_xls2c(self, ModeTable, xlrdsheet, index):
        OutPut = []
        mergedCellsInf = xlrdsheet.merged_cells
        # print(mergedCellsInf)
        # print(len(mergedCellsInf))
        mergedCellsPos = []
        mergedCellsLen = []
        
        for mergedCell in mergedCellsInf:
            if ( mergedCell[2]==11 ):
                mergedCellsLen.append( mergedCell[1]-mergedCell[0] )
                mergedCellsPos.append( mergedCell[1]-3 )
        # print(xls.iloc[0,3])
        # print(mergedCellsPos,mergedCellsLen)
        OutPut.append('static const MidModeUnit_t Mode'+str(index+1)+'[] =\r')
        OutPut.append('{\r')
        
        for row in ModeTable.index:
            block = '    {CHL_TABLEX('+'{:>02}'.format(str(ModeTable.iloc[row, 0]))+'), MFE_INFO, '
            block +=    '{{'+'{:>18}'.format(ModeTable.iloc[row, 4])+', '
            block +=    '{:>5}'.format(str(ModeTable.iloc[row, 5]))+', '
            block +=    '{:>4}'.format(str(ModeTable.iloc[row, 6]))+'}, '
            block +=    '{'+'{:>18}'.format(ModeTable.iloc[row, 7])+', '
            block +=    '{:>5}'.format(str(ModeTable.iloc[row, 8]))+', '
            block +=    '{:>4}'.format(str(ModeTable.iloc[row, 9]))+'}}, '
            block +=    '{:>4}'.format(str(ModeTable.iloc[row, 10]))+'},\r'
            
            if ( row in mergedCellsPos ):
                listpop = mergedCellsLen.pop()
                # print("    REPEAT_UNIT(%3d,"%listpop,"%3d),"%xls.iloc[(row+1-listpop),7],file=outputFile)
                block += '    REPEAT_UNIT('+'{:>3}'.format(str(listpop))+','+'{:>3}'.format(str(ModeTable.iloc[row+1-listpop, 11]))+'),\r'
            OutPut.append(block)
        OutPut.append('}\r')
        # print("}\n",file=outputFile)
        OutPut = ''.join(OutPut)
        # print(OutPut)
        return OutPut
        

    def convert(self):
        self.cal_mode_number()
        for i in range(self.ModeNumber):
            sheetModeTable = 's_ModeTable#' + str(i+1)
            ModeTable = pd.read_excel(self.xlsfile,sheet_name=sheetModeTable,header=1,keep_default_na=False,usecols="A:L")
            workbook = xlrd.open_workbook(self.xlsfile,formatting_info=True)
            xlrdsheet = workbook.sheet_by_name(sheetModeTable)
            self.OutputData.append(self.func_xls2c(ModeTable,xlrdsheet,i))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Usage: python3 xls2c.py <xlsfile>')
        sys.exit(1)
    x2c = Xls2CMid(sys.argv[1])
    x2c.convert()
    output = x2c.get_data()
    with open('outputMid.c', 'w') as f:
        print(''.join(output), file=f)
        f.close()
    sys.exit(0)

