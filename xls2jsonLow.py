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

from xls2 import Xls2


class Xls2JsonLow(Xls2):
    
    def func_xls2json(self, ModeTable, ChlGroupTable, PulseWidthTable, DZCarrierTable, OutputDataTable, options, index):
        json_dict = { "CfgFile" : "", "RecipeFile" : "", "BaseWaveFile" : ""}
        CfgFileString = ""
        RecipeFileString = ""
        # BaseWaveFileString = ""
        
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
        
        StdWidthRuleNumber = '{:>02}'.format(str(format(PulseWidthTable.shape[0], 'x')))
        RecipeFile.append(StdWidthRuleNumber)
        
        StdExtraParamNumber = '{:>02}'.format(str(format(DZCarrierTable.shape[0], 'x')))
        RecipeFile.append(StdExtraParamNumber)
        
        DmaOutputDataNumber = '{:>02}'.format(str(format(OutputDataTable.shape[0], 'x')))
        RecipeFile.append(DmaOutputDataNumber)
        
        _ChlGroupOffset = 0x17
        ChlGroupOffset = re.findall(".{2}", '{:>04}'.format(str(format(_ChlGroupOffset, 'x'))))
        ChlGroupOffset.reverse()
        ChlGroupOffset = ' '.join(ChlGroupOffset)
        RecipeFile.append(ChlGroupOffset)
        
        _StdWidthRuleOffset = _ChlGroupOffset + int(ChlGroupNumber, 16) + sum(ChlGroupLength)
        StdWidthRuleOffset = re.findall(".{2}", '{:>04}'.format(str(format(_StdWidthRuleOffset, 'x'))))
        StdWidthRuleOffset.reverse()
        StdWidthRuleOffset = ' '.join(StdWidthRuleOffset)
        RecipeFile.append(StdWidthRuleOffset)
        
        _StdExtraParamOffset = _StdWidthRuleOffset + int(StdWidthRuleNumber, 16)*0x05
        StdExtraParamOffset = re.findall(".{2}", '{:>04}'.format(str(format(_StdExtraParamOffset, 'x'))))
        StdExtraParamOffset.reverse()
        StdExtraParamOffset = ' '.join(StdExtraParamOffset)
        RecipeFile.append(StdExtraParamOffset)
        
        _DmaOutputDataOffset = _StdExtraParamOffset + int(StdExtraParamNumber, 16)*0x02
        DmaOutputDataOffset = re.findall(".{2}", '{:>04}'.format(str(format(_DmaOutputDataOffset, 'x'))))
        DmaOutputDataOffset.reverse()
        DmaOutputDataOffset = ' '.join(DmaOutputDataOffset)
        RecipeFile.append(DmaOutputDataOffset)
        # print(DmaOutputDataOffset)
        
        _ModeDataOffset = _DmaOutputDataOffset + int(DmaOutputDataNumber, 16)*0x08
        ModeDataOffset = re.findall(".{2}", '{:>04}'.format(str(format(_ModeDataOffset, 'x'))))
        ModeDataOffset.reverse()
        ModeDataOffset = ' '.join(ModeDataOffset)
        RecipeFile.append(ModeDataOffset)
        # RecipeFile.append('\r\n')
        # print(ModeDataOffset)
        
        
        # Write ChlGroupTable
        ChlGroupTableSize = []
        for i in ChlGroupLength:
            ChlGroupTableSize.append('{:>02}'.format(str(format(i, 'x'))))
        ChlGroupTableSize = ' '.join(ChlGroupTableSize)
        RecipeFile.append(ChlGroupTableSize)
        # RecipeFile.append('\r\n')
        
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
        
        # Write PulseWidthTable
        s_PulseWidthTable = []
        # print(PulseWidthTable)
        for row in PulseWidthTable.index:
            s_PulseWidthTable.append('{:>02}'.format(str(format(PulseWidthTable.loc[row, 'WidthMin'], 'x'))))
            s_PulseWidthTable.append('{:>02}'.format(str(format(PulseWidthTable.loc[row, 'WidthMax'], 'x'))))
            s_PulseWidthTable.append('{:>02}'.format(str(format(PulseWidthTable.loc[row, 'Constant'], 'x'))))
            s_PulseWidthTable.append('{:>02}'.format(str(format(PulseWidthTable.loc[row, 'Interval'], 'x'))))
            s_PulseWidthTable.append('{:>02}'.format(str(format(PulseWidthTable.loc[row,'StepValUs'], 'x'))))
            # s_PulseWidthTable.append('\r\n')
        s_PulseWidthTable = ' '.join(s_PulseWidthTable)
        RecipeFile.append(s_PulseWidthTable)
        # print(s_PulseWidthTable)
        
        # Write DZCarrierTable
        s_DZCarrierTable = []
        for row in DZCarrierTable.index:
            s_DZCarrierTable.append('{:>02}'.format(str(format(DZCarrierTable.loc[row, 'DeadZone'], 'x'))))
            s_DZCarrierTable.append('{:>02}'.format(str(format(DZCarrierTable.loc[row,  'Carrier'], 'x'))))
            # s_DZCarrierTable.append('\r\n')
        s_DZCarrierTable = ' '.join(s_DZCarrierTable)
        RecipeFile.append(s_DZCarrierTable)
        # print(s_DZCarrierTable)
        
        # Write s_OutputDataTable
        s_OutputDataTable = []
        for row in OutputDataTable.index:
            tmp = re.findall(".{2}", '{:>08}'.format(OutputDataTable.loc[row, 'Pos']))
            tmp.reverse()
            tmp = ' '.join(tmp)
            s_OutputDataTable.append(tmp)
            tmp = re.findall(".{2}", '{:>08}'.format(OutputDataTable.loc[row, 'Neg']))
            tmp.reverse()
            tmp = ' '.join(tmp)
            s_OutputDataTable.append(tmp)
            # s_OutputDataTable.append('\r\n')
        s_OutputDataTable = ' '.join(s_OutputDataTable)
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
        PulseWidth_t = list(PulseWidthTable.loc[:, 'PulseWidth_t'])
        DZCarrier_t = list(DZCarrierTable.loc[:, 'DZCarrier_t'])
        PulseType_t = {'PULSE_SINGLE':'00', 'PULSE_ALTERNATE':'01', 'PULSE_IN_TURN':'02', 'PULSE_ALTERNATE_REPEAT':'03'}
        BoostType_t = {'BOOST_CLOSE':'00', 'BOOST_AUTO':'01', 'BOOST_TRIA':'02', 'BOOST_SAWUP':'03',
                       'BOOST_SAWDOWN':'04', 'BOOST_UPHOLD':'05', 'BOOST_DOWNHOLD':'06'}
        AutoAdjust_t = {'AUTO_ADJUST_FRQ':'01', 'AUTO_ADJUST_VCYCLE':'02',
                        'AUTO_ADJUST_RUNTIME':'03', 'AUTO_ADJUST_IDLETIME':'04', 'AUTO_ADJUST_NONE':'00'}
        # print(ModeTable)
        for row in ModeTable.index:
            CHL_TABLEX = re.findall(".{2}", '{:>08}'.format(str(format(CHL_TABLE[int(ModeTable.loc[row, 'CHL_TABLEX'])], 'x'))))
            CHL_TABLEX.reverse()
            CHL_TABLEX = ' '.join(CHL_TABLEX)
            s_ModeTable.append(CHL_TABLEX)
            # print(CHL_TABLEX)
        
            _Pulse = ''
            _Pulse += PulseType_t[ModeTable.iloc[row, 1]] + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'FrqHZx10'], 'x'))))
            tmp.reverse()
            _Pulse += ' '.join(tmp) + ' '
            if ModeTable.loc[row, 'WidthIndex'] in PulseWidth_t:
                _Pulse += '{:>02}'.format(str(format(PulseWidth_t.index(ModeTable.loc[row, 'WidthIndex']), 'x'))) + ' '
                _Pulse += '{:>02}'.format(str(format(PulseWidth_t.index(ModeTable.loc[row, 'WidthIndex']), 'x'))) + ' '
            else:
                print(ModeTable.loc[row, 'WidthIndex'] + ' not in PulseWidthTable')
        
            if ModeTable.loc[row, 'ExtraParam'] in DZCarrier_t:
                _Pulse += '{:>02}'.format(str(format(DZCarrier_t.index(ModeTable.loc[row, 'ExtraParam']), 'x')))
            else:
                print(ModeTable.loc[row, 'ExtraParam'] + ' not in DZCarrierTable')
            s_ModeTable.append(_Pulse)
        
            _Boost = ''
            _Boost += BoostType_t[ModeTable.iloc[row, 5]] + ' '
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'Cycle'], 'x'))))
            tmp.reverse()
            _Boost += ' '.join(tmp)
            s_ModeTable.append(_Boost)
        
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'RunTimeMs'], 'x'))))
            tmp.reverse()
            _RunTimeMs = ' '.join(tmp)
            s_ModeTable.append(_RunTimeMs)
        
            tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'IdleMs'], 'x'))))
            tmp.reverse()
            _IdleMs = ' '.join(tmp)
            s_ModeTable.append(_IdleMs)
        
            _AutoAdjust = ''
            _AutoAdjust += AutoAdjust_t[ModeTable.iloc[row, 9]] + ' '
            if ModeTable.iloc[row, 9] == 'AUTO_ADJUST_NONE':
                _AutoAdjust += '00 00 00'
            else :
                _AutoAdjust += '{:>02}'.format(str(format(ModeTable.loc[row, 'Time'], 'x'))) + ' '
                tmp = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'Value'], 'x'))))
                tmp.reverse()
                _AutoAdjust += ' '.join(tmp)
            s_ModeTable.append(_AutoAdjust)
        
        #if has MODE_DELAY_XMS
            if ModeTable.loc[row, 'MODE_DELAY_XMS'] != '':
                DelayXms = []
                CHL_TABLEX = re.findall(".{2}", '{:>08}'.format(str(format(CHL_TABLE[0], 'x'))))
                CHL_TABLEX.reverse()
                CHL_TABLEX = ' '.join(CHL_TABLEX)
                DelayXms.append(CHL_TABLEX)
                DelayXms.append(PulseType_t['PULSE_SINGLE'])
                DelayXms.append('88 04 00 00 00 00 00 00 00 00')
                DelayValue = re.findall(".{2}", '{:>04}'.format(str(format(ModeTable.loc[row, 'MODE_DELAY_XMS'], 'x'))))
                DelayValue.reverse()
                DelayValue = ' '.join(DelayValue)
                DelayXms.append(DelayValue)
                DelayXms.append('00 00 00 00')
                # print(DelayXms)
                DelayXms = ' '.join(DelayXms)
                # s_ModeTable.append('\r\n')
                s_ModeTable.append(DelayXms)
                
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
        
        RecipeID = re.findall(".{2}", '{:>08}'.format(str(format(options.loc[index, 'RecipeID']))))
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
        self.getModeNumber()
        for i in range(self.ModeNumber):
            # Read the excel file
            sheetModeTable = 's_ModeTable#' + str(i+1)
            ModeTable = pd.read_excel(self.xlsfile,sheet_name=sheetModeTable,header=1,keep_default_na=False,usecols="A:M")
            ChlGroupTable = pd.read_excel(self.xlsfile,sheet_name='s_ChlGroupTable',keep_default_na=False)
            PulseWidthTable = pd.read_excel(self.xlsfile,sheet_name='s_PulseWidthTable',keep_default_na=False,usecols="A:F")
            DZCarrierTable = pd.read_excel(self.xlsfile,sheet_name='s_DZCarrierTable',keep_default_na=False,usecols="A:C")
            OutputDataTable = pd.read_excel(self.xlsfile,sheet_name='s_OutputDataTable',keep_default_na=False,usecols="A:C")
            options = pd.read_excel(self.xlsfile,sheet_name='options',keep_default_na=False)
            self.OutputData.append(self.func_xls2json(ModeTable, ChlGroupTable, PulseWidthTable, DZCarrierTable, OutputDataTable, options, i))


if __name__ == '__main__':
    xlsfile = sys.argv[1]
    x2j = Xls2JsonLow(xlsfile)
    x2j.convert()
    output = x2j.get_data()
    for i in output:
        outputname = 'outputLow' + str(output.index(i)) + '.json'
        with open(outputname, 'w') as f:
            f.write(i)
            f.close()
    sys.exit(0)




