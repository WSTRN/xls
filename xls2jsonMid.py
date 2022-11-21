##
# @file xls2json.py
# @brief Convert Recipe Excel file to JSON file
# @author wangtaisheng
# @version v0.1
# @date 2022-10-11

import json
import re
import sys

import pandas as pd
import xlrd

from xls2 import Xls2


class Xls2JsonMid(Xls2):
    def __init__(self, xlsfile, register, productName='Product'):
        Xls2.__init__(self, xlsfile, productName)
        self.register = register

    def func_xls2json(self, ModeTable, ChlGroupTable, OutputDataTable, options, xlrdsheet, index, register='PBSC'):
        json_dict = { "CfgFile" : "", "RecipeFile" : "", "BaseWaveFile" : ""}
        CfgFileString = ""
        RecipeFileString = ""
        # BaseWaveFileString = ""
        
        #read merged cells
        mergedCellsInf = xlrdsheet.merged_cells
        # print(mergedCellsInf)
        # print(len(mergedCellsInf))
        mergedCellsPos = []
        mergedCellsLen = []
        
        for mergedCell in mergedCellsInf:
            if ( mergedCell[2]==11 ):
                mergedCellsLen.append( mergedCell[1]-mergedCell[0] )
                mergedCellsPos.append( mergedCell[1]-3 )
        print(mergedCellsPos)
        print(mergedCellsLen)
    
    
        
        #Calculate the length of each row in ChlGroupTable
        #And load the data into _s_ChlGroupTable
        ChlGroupLength = []
        _s_ChlGroupTable = []
        for row in ChlGroupTable.index:
            tmp = ChlGroupTable.iloc[row, :]
            tmp = list(filter(None, tmp))
            _s_ChlGroupTable.extend(tmp)
            ChlGroupLength.append(len(tmp))
        # print(sum(ChlGroupLength))
        # print(ChlGroupLength)
        # print(_s_ChlGroupTable)
        
        # Write RecipeFile
        RecipeFile = []
        ParseType = '{:>02}'.format(options.loc[index, 'ParseType'])
        RecipeFile.append(ParseType)
        
        ModeCode = re.findall(".{2}", '{:>08}'.format(options.loc[index, 'RecipeID']))
        ModeCode.reverse()
        ModeCode = ' '.join(ModeCode)
        RecipeFile.append(ModeCode)
        
        ModeVersion = re.findall(".{2}", '{:>08}'.format(options.loc[index, 'ModeVersion']))
        ModeVersion.reverse()
        ModeVersion = ' '.join(ModeVersion)
        RecipeFile.append(ModeVersion)
        
        ChlGroupNumber = '{:>02}'.format(str(format(ChlGroupTable.shape[0], 'x')))
        RecipeFile.append(ChlGroupNumber)
        
        DmaOutputDataNumber = '{:>02}'.format(str(format(OutputDataTable.shape[0], 'x')))
        RecipeFile.append(DmaOutputDataNumber)
        
        _ChlGroupOffset = 0x11
        ChlGroupOffset = re.findall(".{2}", '{:>04}'.format(str(format(_ChlGroupOffset, 'x'))))
        ChlGroupOffset.reverse()
        ChlGroupOffset = ' '.join(ChlGroupOffset)
        RecipeFile.append(ChlGroupOffset)
        
        _DmaOutputDataOffset = _ChlGroupOffset + int(ChlGroupNumber, 16) + sum(ChlGroupLength)
        DmaOutputDataOffset = re.findall(".{2}", '{:>04}'.format(str(format(_DmaOutputDataOffset, 'x'))))
        DmaOutputDataOffset.reverse()
        DmaOutputDataOffset = ' '.join(DmaOutputDataOffset)
        RecipeFile.append(DmaOutputDataOffset)
        # print(DmaOutputDataOffset)
        
        if register == 'PBSC':
            _ModeDataOffset = _DmaOutputDataOffset + int(DmaOutputDataNumber, 16)*0x08
        elif register == 'POD':
            _ModeDataOffset = _DmaOutputDataOffset + int(DmaOutputDataNumber, 16)*0x02
        ModeDataOffset = re.findall(".{2}", '{:>04}'.format(str(format(_ModeDataOffset, 'x'))))
        ModeDataOffset.reverse()
        ModeDataOffset = ' '.join(ModeDataOffset)
        RecipeFile.append(ModeDataOffset)
        # print(ModeDataOffset)
        
    
        # Write ChlGroupTable
        ChlGroupTableSize = []
        for i in ChlGroupLength:
            ChlGroupTableSize.append('{:>02}'.format(str(format(i, 'x'))))
        ChlGroupTableSize = ' '.join(ChlGroupTableSize)
        RecipeFile.append(ChlGroupTableSize)
        
        s_ChlGroupTable = []
        ModeOutputData_t = list(OutputDataTable.loc[:, 'ModeOutputData_t'])
        for i in _s_ChlGroupTable:
            if i in ModeOutputData_t:
                # print(ModeOutputData_t.index(i))
                s_ChlGroupTable.append('{:>02}'.format(str(format(ModeOutputData_t.index(i), 'x'))))
            elif i == 'CHL_NULL':
                # print('ff')
                s_ChlGroupTable.append('ff')
                # s_ChlGroupTable.append('\r\n')
        s_ChlGroupTable = ' '.join(s_ChlGroupTable)
        RecipeFile.append(s_ChlGroupTable)
        # print(s_ChlGroupTable)
        
        
        # Write s_OutputDataTable
        s_OutputDataTable = []
        for row in OutputDataTable.index:
            if register == 'PBSC':
                tmp = re.findall(".{2}", '{:>08}'.format(OutputDataTable.loc[row, 'Pos']))
                tmp.reverse()
                tmp = ' '.join(tmp)
                s_OutputDataTable.append(tmp)
                tmp = re.findall(".{2}", '{:>08}'.format(OutputDataTable.loc[row, 'Neg']))
                tmp.reverse()
                tmp = ' '.join(tmp)
                s_OutputDataTable.append(tmp)
            elif register == 'POD':
                tmp = '{:>02}'.format(OutputDataTable.loc[row, 'Pos'])
                s_OutputDataTable.append(tmp)
                tmp = '{:>02}'.format(OutputDataTable.loc[row, 'Neg'])
                s_OutputDataTable.append(tmp)

            # s_OutputDataTable.append('\r\n')
        s_OutputDataTable = ' '.join(s_OutputDataTable)
        # print(s_OutputDataTable)
        RecipeFile.append(s_OutputDataTable)
        # print(s_OutputDataTable)
        
        #Write s_ModeTable
        s_ModeTable = []
        CHL_TABLE = []
        LastAddress = _ChlGroupOffset + int(ChlGroupNumber, 16)
        for x in ChlGroupLength:
            CHL_TABLE.append(LastAddress)
            LastAddress += x
        # print(CHL_TABLE)
        EmsType_t = {'EMS_SIGNLE':'00', 'EMS_BOTHWAY':'01', 'EMS_ALTERNATE':'02'}
        LFWaveType_t = {
                'LFWAVE_NULL':'00',
                'LFWAVE_SIN':'01',
                'LFWAVE_EXP':'02',
                'LFWAVE_LOG':'03',
                'LFWAVE_TRIA':'04',
                'LFWAVE_SAW':'05',
                'LFWAVE_SQUARE':'06',
                'LFWAVE_TRAPEZIA':'07',
                'LFWAVE_TRAPEZIA2':'08',
                'LFWAVE_TRAPEZIA3':'09',
                'LFWAVE_TRAPEZIA4':'0a',
                'LFWAVE_LOGSAW':'0b',
                'LFWAVE_EXPLOG':'0c',
                'LFWAVE_TRAPEZIASIN':'0d',
                'LFWAVE_ALLON':'0e',
                'LFWAVE_SAWLOG':'0f',
                'LFWAVE_LOGEXP':'10',
                'LFWAVE_TRAPEZI50':'11',
                'LFWAVE_SINH75_TRAPEZI25':'12',
                'LFWAVE_SINH60_TRAPEZI40':'13',
                'LFWAVE_SINH80_TRAPEZI20':'14',
                'LFWAVE_SINH25':'15',
                'LFWAVE_SINH100':'16',
                'LFWAVE_SINH50_EXPLOG50':'17',
                'LFWAVE_SINH25_EXPLOG':'18',
                'LFWAVE_LOG30_TRAP35_SINH35':'19',
                'LFWAVE_SINH25_TRAP75':'1a',
                'LFWAVE_CYCLEH25_TRAP75':'1b',
                'LFWAVE_SINH25_TRAP50_SINH25':'1c',
                'LFWAVE_CYCLEH25_TRAP50_CYCLEH25':'1d',
                'LFWAVE_SIN50_TRAP50':'1e',
                'LFWAVE_TRIAG25_SQURE25':'1f',
                'LFWAVE_TRIAG50_CYCLE50':'20',
                'LFWAVE_NUMBER':'21',
                'LFWAVE_REPEAT':'22'
        }
    
        # print(ModeTable)
        for row in ModeTable.index:
            # print(CHL_TABLE)
            CHL_TABLEX = re.findall(".{2}", '{:>08}'.format(str(format(CHL_TABLE[int(ModeTable.loc[row, 'CHL_TABLEX'])], 'x'))))
            CHL_TABLEX.reverse()
            CHL_TABLEX = ' '.join(CHL_TABLEX)
            s_ModeTable.append(CHL_TABLEX)
            # print(CHL_TABLEX)
        
            #EmsWaveInfo_t
            _Ems = ''
            _Ems += EmsType_t[ModeTable.iloc[row, 1]] + ' '
            _Ems += '{:>02}'.format(str(format(ModeTable.iloc[row, 2], 'x'))) + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.iloc[row, 3], 'x'))))
            tmp.reverse()
            _Ems += ' '.join(tmp)
            s_ModeTable.append(_Ems)
    
            #LFWaveInfo_t1
            _LFWave1 = ''
            _LFWave1 += LFWaveType_t[ModeTable.iloc[row, 4]] + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.iloc[row, 5], 'x'))))
            tmp.reverse()
            _LFWave1 += ' '.join(tmp) + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.iloc[row, 6], 'x'))))
            tmp.reverse()
            _LFWave1 += ' '.join(tmp)
            s_ModeTable.append(_LFWave1)
    
            #LFWaveInfo_t1
            _LFWave2 = ''
            _LFWave2 += LFWaveType_t[ModeTable.iloc[row, 7]] + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.iloc[row, 8], 'x'))))
            tmp.reverse()
            _LFWave2 += ' '.join(tmp) + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.iloc[row, 9], 'x'))))
            tmp.reverse()
            _LFWave2 += ' '.join(tmp)
            s_ModeTable.append(_LFWave2)
        
            #TotalTimes
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'TotalTimes'], 'x'))))
            tmp.reverse()
            _TotalTimes = ' '.join(tmp)
            s_ModeTable.append(_TotalTimes)
        
        
        #if has REPEAT
            if row in mergedCellsPos:
            # if ModeTable.loc[row, 'MODE_DELAY_XMS'] != '':
                REPEAT = []
                CHL_TABLEX = re.findall(".{2}", '{:>08}'.format(str(format(CHL_TABLE[0], 'x'))))
                CHL_TABLEX.reverse()
                CHL_TABLEX = ' '.join(CHL_TABLEX)
                REPEAT.append(CHL_TABLEX)
                REPEAT.append('01 fa 10 27')
                REPEAT.append('22')
                Lenpop = int(mergedCellsLen.pop(0))
                # print(type(Lenpop))
                # print(Lenpop)
                RepeatNum = re.findall(".{2}", '{:>04}'.format(str(format(Lenpop, 'x'))))
                RepeatNum.reverse()
                RepeatNum = ' '.join(RepeatNum)
                REPEAT.append(RepeatNum)
    
                print(row)
                print(row +1-Lenpop)
                print(ModeTable.iloc[row+1-Lenpop, 11])
                RepeatTimes = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.iloc[row+1-Lenpop, 11], 'x'))))
                RepeatTimes.reverse()
                RepeatTimes = ' '.join(RepeatTimes)
                REPEAT.append(RepeatTimes)
    
                REPEAT.append('00 0A 00 01 00 01 00')
                REPEAT = ' '.join(REPEAT)
                s_ModeTable.append(REPEAT)
                
        s_ModeTable = ' '.join(s_ModeTable)
        RecipeFile.append(s_ModeTable)
        
        RecipeFile = ' '.join(RecipeFile)
        RecipeFileString = RecipeFile
        # print('recipe file string: ')
        # print(RecipeFile)
        
        
        
        #Write CfgFile
        CfgFile = []
        CfgFileLength = '00 18'
        CfgFile.append(CfgFileLength)
        RecipeFileLength = re.findall(".{2}", '{:>08}'.format(str(format(len(RecipeFileString.split(' ')), 'x'))))
        RecipeFileLength = ' '.join(RecipeFileLength)
        CfgFile.append(RecipeFileLength)
        BaseWaveFileLength = '00 00 00 00'
        CfgFile.append(BaseWaveFileLength)
        
        CfgFileChecksum = ''
        RecipeFileChecksum = ''
        BaseWaveFileChecksum = '00 00 00 00'
        
        CfgChecksummedData = []
        
        RecipeFileComboID = '00 00 00 00'
        CfgChecksummedData.append(RecipeFileComboID)
        RecipeFileComboVersion = '00 01'
        CfgChecksummedData.append(RecipeFileComboVersion)
        RecipeFileNum = '01'
        CfgChecksummedData.append(RecipeFileNum)
        BaseWaveFileNum = '00'
        CfgChecksummedData.append(BaseWaveFileNum)
        
        RecipeID = re.findall(".{2}", '{:>08}'.format(options.loc[index, 'RecipeID']))
        # print(RecipeID)
        RecipeID = ' '.join(RecipeID)
        CfgChecksummedData.append(RecipeID)
        RecipeVersion = re.findall(".{2}", '{:>016}'.format(str(format(options.loc[index, 'RecipeVersion'], 'x'))))
        RecipeVersion = ' '.join(RecipeVersion)
        CfgChecksummedData.append(RecipeVersion)
        
        RecipeLength  = re.findall(".{2}", '{:>04}'.format(str(format(len(RecipeFileString.split(' ')), 'x'))))
        RecipeLength   = ' '.join(RecipeLength)
        CfgChecksummedData.append(RecipeLength)
        BaseWaveLength = '00 00'
        CfgChecksummedData.append(BaseWaveLength)
        
        CfgChecksummedData = ' '.join(CfgChecksummedData)
        # print(CfgChecksummedData)
        CheckList = CfgChecksummedData.split(' ')
        CheckList = [int(i, 16) for i in CheckList]
        CfgFileChecksum = re.findall(".{2}", '{:>08}'.format(str(format(sum(CheckList), 'x'))))
        CfgFileChecksum = ' '.join(CfgFileChecksum)
        CfgFile.append(CfgFileChecksum)
        # print(CfgFileChecksum)
        
        CheckList = RecipeFileString.split(' ')
        CheckList = [int(i, 16) for i in CheckList]
        RecipeFileChecksum = re.findall(".{2}", '{:>08}'.format(str(format(sum(CheckList), 'x'))))
        RecipeFileChecksum = ' '.join(RecipeFileChecksum)
        CfgFile.append(RecipeFileChecksum)
        CfgFile.append(BaseWaveFileChecksum )
        CfgFile.append(CfgChecksummedData)
        
        CfgFile = ' '.join(CfgFile)
        CfgFileString = CfgFile
        
        
        
        #putting data into the dictionary and converting it to the json file
        json_dict["CfgFile"] = CfgFileString
        json_dict["RecipeFile"] = RecipeFileString
        return json.dumps(json_dict, indent=4)
    
    def convert(self):
        self.cal_mode_number()
        self.cal_filename()
        for i in range(self.ModeNumber):
            # Read the excel file
            sheetModeTable = 's_ModeTable#' + str(i+1)
            ModeTable = pd.read_excel(self.xlsfile,sheet_name=sheetModeTable,header=1,keep_default_na=False,usecols="A:L")
            ChlGroupTable = pd.read_excel(self.xlsfile,sheet_name='s_ChlGroupTable',keep_default_na=False)
            OutputDataTable = pd.read_excel(self.xlsfile,sheet_name='s_OutputDataTable',keep_default_na=False,usecols="A:C")
            options = pd.read_excel(self.xlsfile,sheet_name='options',keep_default_na=False)
            workbook = xlrd.open_workbook(self.xlsfile,formatting_info=True)
            xlrdsheet = workbook.sheet_by_name(sheetModeTable)
            print(str(i+1)+'th mode')
            self.OutputData.append(self.func_xls2json(ModeTable,ChlGroupTable,OutputDataTable,options,xlrdsheet,i,register=self.register))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Usage: xls2c.py <xls_file> <product_name>')
        sys.exit(1)
    xlsfile = sys.argv[1]
    x2j = Xls2JsonMid(xlsfile,register='PBSC',productName=sys.argv[2])
    x2j.convert()
    output = x2j.get_data()

    for i in output:
        outputname = x2j.filename.pop(0) + '.json'
        print(outputname)
        with open(outputname, 'w') as f:
            f.write(i)
            f.close()
    sys.exit(0)




