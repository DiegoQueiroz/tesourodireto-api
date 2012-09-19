import urllib
import os

def retrieveurl(url, filename=None):
    success = False
    
    fp = urllib.urlopen(url)
    try:
        headers = fp.info()
        code = fp.getcode()
        if code == 200:            
            if filename:
                tfp = open(filename, 'wb')
            else:
                import tempfile
                _, suffix = os.path.splitext(url)
                (fd, filename) = tempfile.mkstemp(suffix)
                tfp = os.fdopen(fd, 'wb')
            try:
                bs = 1024 * 8
                size = -1
                read = 0
                if "content-length" in headers:
                    size = int(headers["Content-Length"])
                
                while True:
                    block = fp.read(bs)
                    if block == "":
                        break
                    read += len(block)
                    tfp.write(block)
                    
                if size < 0 or read == size:
                    # retrieved file size is the same of the expected file size
                    success = True
            finally:
                tfp.close()
    finally:
        fp.close()
        
    return filename, success

if __name__ == '__main__':
    
    from datetime import date
    from csv import writer as csvwriter
    import xlrd
    import sys
    
    csv_delimiter  = ','
    quote_char     = '"'
    try:
        interval   = int(sys.argv[1]) # days
    except:
        interval   = 1 # in years
    
    baseurl = 'http://www.tesouro.fazenda.gov.br/tesouro_direto/download/historico/{0}/{1}_{0}.xls'
    docs = ['LFT', 'LTN', 'NTNC', 'NTNB', 'NTNB_Principal', 'NTNF', ]
    currentyear = date.today().year
    
    with open('FUNDOS_TESOURO_QUICKEN.csv','wb') as f:
        arqcsv = csvwriter(f, delimiter=csv_delimiter,quotechar=quote_char)
        
        for doc in docs:
            for year in range(currentyear,currentyear-interval,-1):
                if year == currentyear:
                    url = baseurl.format(year,doc)
                else:
                    # Files in history are prefixed with the word "historico"
                    # Also NTNB_Principal is spelled NTNBPrincipal (without underscore) in history
                    url = baseurl.format(year,'historico' + doc.replace('_',''))
                    
                ( xlsfile, success ) = retrieveurl(url)
                if success:
                    xls = xlrd.open_workbook(xlsfile)
                    
                    for sheet in xls.sheets():
                        securityname = 'TN_' + sheet.name.split(' ')[-1]
                        
                        for row in range(2,sheet.nrows):
                            
                            if sheet.ncols < 6:
                                # some old sheet (prior 2002) does not have "PU base" values
                                # skip them
                                continue
                            
                            quotedatecell = sheet.cell(row,0)
                            quotepricecell = sheet.cell(row,5)
                            
                            if quotepricecell.ctype == xlrd.XL_CELL_EMPTY:
                                # No quotes available
                                # skip line
                                continue
                            
                            if quotedatecell.ctype == xlrd.XL_CELL_DATE or quotedatecell.ctype == xlrd.XL_CELL_NUMBER:
                                quotedate = date(*xlrd.xldate_as_tuple(quotedatecell.value,xls.datemode)[:3]).strftime("%d/%m/%Y")
                            else:
                                quotedate = quotedatecell.value
                            
                            quoteprice = quotepricecell.value
                            
                            reg = [ securityname, quotedate, quoteprice ]
                            
                            arqcsv.writerow(reg)
                            
                    print('Downloaded %s prices from year %d.' % (doc,year))
    
        print('Finished!')
        
