
import pandas as pd


class Xls2:
    def __init__(self, xlsfile):
        self.xlsfile = xlsfile
        self.ModeNumber = 0
        self.OutputData = []
        
    def getModeNumber(self):
        options = pd.read_excel(self.xlsfile,sheet_name='options',keep_default_na=True,usecols='A:D')
        # print(options)
        options = options.dropna(axis=0,how='any')
        # print(options)
        self.ModeNumber = options.shape[0]
        return self.ModeNumber

    def convert(self):
        print('Converting...')
    
    def get_data(self):
        return self.OutputData
    
    