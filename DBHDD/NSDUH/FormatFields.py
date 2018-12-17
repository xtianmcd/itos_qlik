import csv
import os
import pandas as pd

def urlxlsx(url, root):
    tblurl_dict = pd.read_excel(url, sheet_name=None) #read all sheets
    del tblurl_dict[list(tblurl_dict.keys())[0]] #delete table of contents sheet
    # print((tblurl_dict.keys()))
    for k,v in tblurl_dict.items(): #save each sheet to separate csv file in rootpath
        if int(k.split(' ')[1])<10: tblurl_dict[k].to_csv(root+'2012_Tab0{}_#.csv'.format(k.split(' ')[1]), index=False)
        else: tblurl_dict[k].to_csv(root+'2012_Tab{}_#.csv'.format(k.split(' ')[1]), index=False)

def readcsv(readf):
    frows=[]
    rf = csv.reader(readf, delimiter=',')
    for line in rf:
        if line[0] == 'Order'or line[0][:4] == 'NOTE':
            #print(line)
            line = list([' '.join(i.split('\n')) for i in line])
            #print(line)
        frows.append(line)
    return frows

def writecsv(tbl, frws):
    with open(tbl, mode='w', newline='') as testwrite:
        tw = csv.writer(testwrite, delimiter=',')
        tw.writerows(frws)

if __name__=="__main__":

    rootpath = "G:\\appdev\\Projects - For Clients\\DBHDD\\Data\\DSW 1.0\\9_National Survey on Drug Use and Health (NSDUH), GA Data_ASCII\\Raw\\"
    xlsxurl = "https://www.samhsa.gov/data/sites/default/files/NSDUHsaeTotals2012.xlsx"

    urlxlsx(xlsxurl, rootpath)
    tables = list([rootpath+t for t in os.listdir(rootpath) if t.endswith('csv')])

    for table in tables:
        with open(table) as readfile:
            filerows = readcsv(readfile)
        writecsv(table, filerows)
