def xlsx_to_csv(xlsx_file, outfile='output.csv', sheet_index=0, row_start=1, row_end=None):
    """
    
    Converts the given xlsx file to csv format from a given range of row
    
    :param xlsx_file: The name of the xlsx file
    :param outfile: The name of the output file, Default: output.csv
    :param sheet_index: The index number of the sheet, Default: 0
    :param row_start: the starting row of the sheet, Default: 1
    :param row_end: the end of the row, Default: last row
    
    :returns: None
    
    Example Usage: xlsx_to_csv('./XLSX/Fiscal Deficits.xlsx', outfile='./XLSX/01_Fiscal_Deficits.csv', row_start=3, row_end=35)
    
    """
    import xlrd, csv
    
    wb = xlrd.open_workbook(xlsx_file)
    sheet = wb.sheet_by_index(sheet_index)
    output = open(outfile, 'w', encoding='utf8')
    wr = csv.writer(output, quoting=csv.QUOTE_ALL)
    
    if row_end is None:
        row_end = sheet.nrows
    
    for rowidx in range(row_start-1, row_end):
        row = sheet.row_values(rowidx)
        wr.writerow(row)
    wb.release_resources()
    output.close()
    del wb, sheet

def batch_convert(folder='./', outfolder='./CSV/', sheet_index=0, row_start=0, row_end=None):
    """
    Converts a batch of files *.xlsx to *.csv, mention the sheet_index, row_start and row_end
    
    Example Usage : batch_convert('./XLSX', row_start=3, row_end=35)
    """
    import glob, os, time
    start = time.time()
    
    files = glob.glob(folder+'/*.xlsx')
    if not os.path.exists(outfolder):
        os.makedirs(outfolder)
    outfilenames = [ outfolder + fname.replace(' ', '_') + '.csv' for fname in [os.path.splitext(os.path.basename(f))[0] for f in files] ]
    for (infile, outfile) in zip(files, outfilenames):
        xlsx_to_csv(xlsx_file=infile, outfile=outfile, sheet_index=sheet_index, row_start=row_start, row_end=row_end)
        
    end = time.time()
    
    print('DONE ! Processed {} files in {} seconds \nOutput Directory : {}'.format(len(files), end-start, os.path.abspath(outfolder)))
    
def clean_niti_data(csv_folder='./', outfolder='./PROCESSED/'):
    """
    Provide the CSV folder and this method will replace the missing values with NaN
    and the colnames will be cleaned
    
    Usage : clean_niti_data('./CSV/')
    """
    import re, functools, itertools, glob, os, time
    import pandas as pd
    
    start = time.time()
    
    files = glob.glob(csv_folder+'*.csv')
    
    if not os.path.exists(outfolder):
        os.makedirs(outfolder)
    
    for file in files:
        data = pd.read_csv(file)
        data.iloc[:, 1:] = data.iloc[:, 1:].apply( functools.partial(pd.to_numeric, errors='coerce'))
        newidx = [re.sub('[^0-9\-]+', '', colname) for colname in data.columns[1:]]
        newidx.insert(0, 'State')
        data.columns = pd.Index(newidx)
        data.to_csv(outfolder+os.path.splitext(os.path.basename(file))[0]+'.csv', index=False)
    
    end = time.time()
    
    print('Processed {} files in {} seconds'.format(len(files), end-start))