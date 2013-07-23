from __future__ import print_function
from urllib2 import urlopen, Request, URLError
import os


def retrieveurl(url, filename=None):
    success = False

    # Tesouro website needs a stupid valid UserAgent. ;/
    useragent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.11' \
                ' (KHTML, like Gecko) Chrome/23.0.1271.64 Safari/537.11'

    try:
        fp = urlopen(Request(url, headers={'User-Agent': useragent}))
    except URLError:
        return None, success

    try:
        headers = fp.info()
        code = fp.getcode()
        if code == 200:  # HTTP CODE 200 OK
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
                if "Content-Length" in headers:
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

    csv_delimiter = ','
    quote_char = '"'
    # baseurl = 'http://www.tesouro.fazenda.gov.br/tesouro_direto/' \
    #           'download/historico/{0}/{1}_{0}.xls'
    baseurl_oldyears = 'http://www3.tesouro.gov.br/tesouro_direto/' \
                       'download/historico/{0}/historico{1}_{0}.xls'
    baseurl_2012 = 'http://www3.tesouro.gov.br/tesouro_direto/download/' \
                   'historico/{0}/{1}_{0}.xls'
    baseurl_currentyear = 'https://www.tesouro.fazenda.gov.br/' \
                          'images/arquivos/artigos/{1}_{0}.xls'

    try:
        interval = int(sys.argv[1])  # in years
    except IndexError:
        interval = 1  # in years

    try:
        docs = sys.argv[2]
    except IndexError:
        docs = ['LFT', 'LTN', 'NTN-C', 'NTN-B', 'NTN-B_Principal', 'NTN-F', ]

    currentyear = date.today().year

    with open('FUNDOS_TESOURO_QUICKEN.csv', 'wb') as f:
        arqcsv = csvwriter(f, delimiter=csv_delimiter, quotechar=quote_char)

        for doc in docs:
            for year in range(currentyear, currentyear - interval, -1):

                quotename = "%s year %d" % (doc, year)
                print('Downloading %-40s... ' % (quotename), end='')

                if year == currentyear:
                    url = baseurl_currentyear.format(year, doc)
                elif year == 2012:
                    # Different handling of 2012 year history
                    url = baseurl_2012.format(year, doc)
                else:
                    # Files in history are prefixed with the word "historico"
                    # Also NTN-B_Principal is spelled NTNBPrincipal
                    # (without underscore) in history
                    docrenamed = doc.replace('_', '').replace('-', '')
                    url = baseurl_oldyears.format(year, docrenamed)

                (xlsfile, success) = retrieveurl(url)
                if not success:
                    print('ERROR!')
                else:
                    xls = xlrd.open_workbook(xlsfile)

                    for sheet in xls.sheets():
                        securityname = 'TN_' + doc + '_' + \
                                       sheet.name.split(' ')[-1]

                        for row in range(2, sheet.nrows):

                            if sheet.ncols < 6:
                                # some old sheet (prior 2002) does
                                # not have "PU base" values
                                # skip them
                                continue

                            quotedatecell = sheet.cell(row, 0)
                            quotepricecell = sheet.cell(row, 5)

                            if quotepricecell.ctype == xlrd.XL_CELL_EMPTY:
                                # No quotes available
                                # skip line
                                continue

                            if (quotedatecell.ctype == xlrd.XL_CELL_DATE or
                                quotedatecell.ctype == xlrd.XL_CELL_NUMBER):
                                thedate = xlrd.xldate_as_tuple(
                                    quotedatecell.value, xls.datemode)[:3]
                                quotedate = date(*thedate).strftime("%d/%m/%Y")
                            else:
                                quotedate = quotedatecell.value

                            quoteprice = quotepricecell.value

                            reg = [securityname, quotedate, quoteprice]

                            arqcsv.writerow(reg)

                    print('success!')

        # Finished
