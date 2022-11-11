import sys

import pandas as pd

from xls2 import Xls2


class Xls2CLow(Xls2):

    def func_xls2c_2(self, ModeTable, index):
        OutPut = []
        
        #define the output string
        ModeUnit = []

        #write ModeUnit header
        ModeUnit.append('static ModeUnit_t Mode'+str(index+1)+'[] = {\r', )
        #write ModeUnit body
        for row in ModeTable.index:
            block = '    {CHL_TABLEX('+'{:>02}'.format(str(ModeTable.loc[row, 'CHL_TABLEX']))+'),\r'
            block +='    {'+ModeTable.iloc[row, 1]+', '+str(ModeTable.loc[row, 'FrqHZx10'])+', WIDTH_SINGLE('+ModeTable.loc[row, 'WidthIndex']+'), '+ModeTable.loc[row, 'ExtraParam']+'},\r'
            block +='    BOOST_SET('+ModeTable.iloc[row, 5]+', '+str(ModeTable.loc[row, 'Cycle'])+'),\r'
            block +='    '+str(ModeTable.loc[row, 'RunTimeMs'])+',\r'
            block +='    '+str(ModeTable.loc[row, 'IdleMs'])+',\r'
            block +='    '+ModeTable.iloc[row, 9]+'},\r'
            if ModeTable.loc[row, 'MODE_DELAY_XMS'] != 0 and ModeTable.loc[row, 'MODE_DELAY_XMS'] != '':
                block +='    MODE_DELAY_XMS('+str(ModeTable.loc[row, 'MODE_DELAY_XMS'])+'),\r'
            ModeUnit.append(block+'\r')
        ModeUnit.append('};\r')
        ModeUnit = ''.join(ModeUnit)
        OutPut.append(ModeUnit)
        
        return OutPut
        # print(o_ModeUnit)
    
    def func_xls2c_1(self, ChlGroupTable, PulseWidthTable, DZCarrierTable):
        OutPut = []

        ChlGroup = []
        PulseWidth = []
        PulseWidthParam = []
        DZCarrier = []
        PulseExtraParam = []
    
        #Calculate the length of each row in ChlGroupTable
        #And load the data into _s_ChlGroupTable
        _ChlGroupTable = []
        for row in ChlGroupTable.index:
            tmp = ChlGroupTable.iloc[row, :]
            tmp = list(filter(None, tmp))
            _ChlGroupTable.append(tmp)
    
        #write ChlGroup
        for row in ChlGroupTable.index:
            #write ChlGroup header
            group = 'static const uint8_t s_ChlTable'+'{:>02}'.format(str(row))+'[] = {\r'
            #write ChlGroup body
            for chl in _ChlGroupTable[row]:
                group += '    '+chl+',\r'
            group += '};\r'
            ChlGroup.append(group)
        ChlGroup = ''.join(ChlGroup)
        OutPut.append(ChlGroup)
        # print(o_ChlGroup)
    
    
        #write PulseWidth header
        PulseWidth.append('static PulseWidth_t s_PulseWidthTable[WIDTH_RULE_NUMBER] = {\r')
        #write PulseWidth body
        for row in PulseWidthTable.index:
            block = '    {'+'{:>4}'.format(str(PulseWidthTable.loc[row, 'WidthMin']))
            block +=','    +'{:>4}'.format(str(PulseWidthTable.loc[row, 'WidthMax']))
            block +=','    +'{:>4}'.format(str(PulseWidthTable.loc[row, 'Constant']))
            block +=','    +'{:>4}'.format(str(PulseWidthTable.loc[row, 'Interval']))
            block +=','    +'{:>4}'.format(str(PulseWidthTable.loc[row, 'StepValUs']))+'},\r'
            PulseWidth.append(block)
        PulseWidth.append('};\r')
        PulseWidth = ''.join(PulseWidth)
        OutPut.append(PulseWidth)
        # print(o_PulseWidth)
    
        #write PulseWidthParam header
        PulseWidthParam.append('typedef enum\r{\r')
        #write PulseWidthParam body
        for row in PulseWidthTable.index:
            block = '    PULSE_{WidthMin}_{WidthMax}_{Constant}_{Interval}_{StepValUs},\r'.format(
                WidthMin=PulseWidthTable.loc[row, 'WidthMin'],
                WidthMax=PulseWidthTable.loc[row, 'WidthMax'],
                Constant=PulseWidthTable.loc[row, 'Constant'],
                Interval=PulseWidthTable.loc[row, 'Interval'],
                StepValUs=PulseWidthTable.loc[row, 'StepValUs'])
            PulseWidthParam.append(block)
        PulseWidthParam.append('    PULSE_TEST,\r')
        PulseWidthParam.append('    WIDTH_RULE_NUMBER,\r')
        PulseWidthParam.append('} PulseWidthParam_t;\r')
        PulseWidthParam = ''.join(PulseWidthParam)
        OutPut.append(PulseWidthParam)
        # print(o_PulseWidthParam)
    
    
        #write DZCarrier header
        DZCarrier.append('static DZCarrier_t s_DZCarrierTable[EXTRA_PARAM_NUMBER] = {\r')
        #write DZCarrier body
        for row in DZCarrierTable.index:
            block = '    {'+'{:>2}'.format(str(DZCarrierTable.loc[row, 'DeadZone']))
            block +=','    +'{:>2}'.format(str(DZCarrierTable.loc[row, 'Carrier']))+'},\r'
            DZCarrier.append(block)
        DZCarrier.append('};\r')
        DZCarrier = ''.join(DZCarrier)
        OutPut.append(DZCarrier)
        # print(o_DZCarrier)
    
        #write PulseExtraParam header
        PulseExtraParam.append('typedef enum\r{\r')
        #write PulseExtraParam body
        for row in DZCarrierTable.index:
            block = '    PULSE_D{DeadZone}_M{Carrier},\r'.format(
                DeadZone=DZCarrierTable.loc[row, 'DeadZone'],
                Carrier=DZCarrierTable.loc[row, 'Carrier'])
            PulseExtraParam.append(block)
        PulseExtraParam.append('    EXTRA_PARAM_NUMBER,\r')
        PulseExtraParam.append('} PulseExtraParam_t;\r')
        PulseExtraParam = ''.join(PulseExtraParam)
        OutPut.append(PulseExtraParam)
        # print(o_PulseExtraParam)
        return OutPut


    #get the data
    #  0:ChlGroup,
    #  1:PulseWidth,
    #  2:PulseWidthParam,
    #  3:DZCarrier,
    #  4:PulseExtraParam
    #  5~N:ModeUnit, 
    def convert(self):
        ChlGroupTable = pd.read_excel(self.xlsfile,sheet_name='s_ChlGroupTable',keep_default_na=False)
        PulseWidthTable = pd.read_excel(self.xlsfile,sheet_name='s_PulseWidthTable',keep_default_na=False,usecols="A:F")
        DZCarrierTable = pd.read_excel(self.xlsfile,sheet_name='s_DZCarrierTable',keep_default_na=False,usecols="A:C")
        self.OutputData.extend(self.func_xls2c_1(ChlGroupTable, PulseWidthTable, DZCarrierTable))
        self.cal_mode_number()
        for i in range(self.ModeNumber):
            sheetModeTable = 's_ModeTable#' + str(i+1)
            ModeTable = pd.read_excel(self.xlsfile,sheet_name=sheetModeTable,header=1,keep_default_na=False,usecols="A:M")
            self.OutputData.extend(self.func_xls2c_2(ModeTable,i))


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Usage: xls2c.py <xls_file>')
        sys.exit(1)
    x2c = Xls2CLow(sys.argv[1])
    x2c.convert()
    output = x2c.get_data()
    with open('outputLow.c', 'w') as f:
        print(''.join(output), file=f)
        f.close()
    sys.exit(0)

