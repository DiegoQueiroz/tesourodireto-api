from csv import writer as csvwriter

class QuickenCSV(object):
    '''
    Class for generating CSV files for Quicken.
    
    Quicken CSV files use the following format:
        SYMBOL , PRICE , DATE
    
        Field delimiter   = comma (,) or double space (  )
        Quote char        = double quotes (")
        Decimal separator = dot (.), see *1
        Date format       = see *2
        
        *1 - Decimal separator in Quicken IS DOT, regardless of system settings. 
        *2 - Date format in Quicken depends on system settings.
        
    More info at:
        http://quicken.intuit.com/support/help/backup--restore--file-issues/how-to-import-historical-security-data-into-quicken/GEN82637.html
    '''

    def __init__(self, quoteprefix, serie):
        '''
        Constructor
        '''
        self.fielddelimiter = ','
        self.quotechar      = '"'
        self.decimalsep     = '.'
        self.dateformat     = '%x' # use system locale
            
        self.quoteprefix    = quoteprefix
        self.__serie        = serie
        self.values         = {}
        
    def updateValues(self,values,clearData=False):
        if clearData:
            self.values = values
        else:
            self.values.update(values)

    def exportToFile(self,targetFile,clearFile=True):
        if len(self.values) == 0:
            return 0
        
        with open(targetFile,'wb' if clearFile else 'ab') as f:
            csvfile = csvwriter(f, delimiter=self.fielddelimiter,quotechar=self.quotechar)
            
            quotename = '{0}_{1}'.format(self.quoteprefix,self.__serie)
            data = [ (quotename, qdate.strftime(self.dateformat), qvalue)
                    for qdate, qvalue in sorted(self.values.items()) ]
            
            # write CSV
            csvfile.writerows(data)
            
            return len(data)
            
    def export(self,clearFile=True):
        quotename = '{0}_{1}'.format(self.quoteprefix,self.__serie)
        return self.exportToFile(quotename + '.csv',clearFile)
    
    