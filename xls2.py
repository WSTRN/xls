
import pandas as pd


class Xls2:
    def __init__(self, xlsfile, productName='Product'):
        self.xlsfile = xlsfile
        self.productName = productName
        self.ModeNumber = 0
        self.filename = []
        self.OutputData = []
        
    def getModeNumber(self):
        options = pd.read_excel(self.xlsfile,sheet_name='options',keep_default_na=True,usecols='A:D')
        # print(options)
        options = options.dropna(axis=0,how='any')
        # print(options)
        self.ModeNumber = options.shape[0]
        return self.ModeNumber
    
    def getfilename(self):
        options = pd.read_excel(self.xlsfile,sheet_name='options',keep_default_na=True,usecols='A:D',dtype=str)
        options = options.dropna(axis=0,how='any')
        for row in options.index:
            self.filename.append('Mode_'+self.productName+'_ID'+str(int(options.loc[row,'RecipeID'],16)))
        return self.filename

    def convert(self):
        print('Converting...')
    
    def get_data(self):
        return self.OutputData
    
    