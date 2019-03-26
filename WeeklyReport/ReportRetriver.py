'''
Created on 2019年3月20日

@author: junli
'''

import os
import time
import collections
import csv
from urllib.request import urlopen
from  urllib.request import urlretrieve

from bs4 import BeautifulSoup


from tabula import read_pdf
import pandas

ReportIItem = collections.namedtuple('ReportIItem', ['name','url'])

WEB_SITE_URL="https://www.sge.com.cn/sjzx/hqzb"
MAX_PAGE=7

DOWNLOAD_FOLDER = r"files\downloads"
OUTPUT_FOLDER = r"files\weeklyreports"


def read_from_csv(file_name):
    items = []
    with open(file_name,'r') as file:
        reader = csv.reader(file)
        for row in reader:
            item = ReportIItem(row[0], row[1])
            items.append(item)
    return items


def write_to_file(file_name,report_items):
    with open(file_name,'w',newline='') as file:
        writer = csv.writer(file)
        for item in report_items:
            writer.writerow([item.name, item.url])
    
          
def analysis_content(page_url):
    html=urlopen(page_url)
    bs_obj = BeautifulSoup(html.read(),'html.parser')
    reports = []
    for link in bs_obj.find_all("a","txt fl"):
        url = link.get('href').strip()
        title =  link.find("span","fl").contents[0].strip()
        reports.append(ReportIItem(title, url))
    return reports


def get_all_report_items():
    all_reports = []
    for page in range(1,MAX_PAGE):
        url = "{}?p={}".format(WEB_SITE_URL, page)
        reports = analysis_content(url)
        all_reports.extend(reports)
    return all_reports

def download_reports(folder, report_items):
    for item in report_items:
        dest_file = os.path.join(folder, item.name+".pdf")
        if os.path.isfile(dest_file):
            print("Report {} already exists, skipping...".format(item.name))
        else:
            print("Try to download {} : url={}.".format(item.name, item.url))        
            urlretrieve(item.url, filename=dest_file)
            print("Download report {}, saved to {}".format(item.name, dest_file))
            time.sleep(2)

def remove_comma(obj):
    return str(obj).replace(',','').replace(' ','')

# Extract useful information from DataFrame, and compose a gold report
#  df[15]  ="黄金交易量前十名"
#  df[18] = "黄金代理成交量前十名 黄金自营成交量前十名"
def gen_gold_report(df):    
    df_gold_top10 =  df[15]
    
    # total buy and sell
    tbs =  df_gold_top10[0].str.split(' ',expand=True)[2:12][[1,2]]    
    tbs.columns = ['bank_name','total_buy_sell']
    tbs.set_index('bank_name', inplace=True)
    tbs['total_buy_sell'] = tbs['total_buy_sell'].apply(remove_comma).astype('float')

    # total buy
    tb =  df_gold_top10[1].str.split(' ',expand=True)[2:12][[1,2]]    
    tb.columns = ['bank_name','total_buy']
    tb.set_index('bank_name', inplace=True)
    tb['total_buy'] = tb['total_buy'].apply(remove_comma).astype('float')
     
    # total sell
    ts =df_gold_top10[2].str.split(' ',expand=True)[2:12][[1,2]]    
    ts.columns = ['bank_name','total_sell']
    ts.set_index('bank_name', inplace=True)
    ts['total_sell'] = ts['total_sell'].apply(remove_comma).astype('float')
    
    df_prop_broker = df[18]
    
    # brokage trading amount
    brokage =   df_prop_broker[1:11][[1,2]]
    brokage.columns =  ['bank_name','brokage']
    brokage.set_index('bank_name', inplace=True)
    brokage['brokage'] = brokage['brokage'].apply(remove_comma).astype('float')

    # proper trading amount
    prop = df_prop_broker[1:11][[4,5]]
    prop.columns =  ['bank_name','prop']
    prop.set_index('bank_name', inplace=True)
    prop['prop'] = prop['prop'].apply(remove_comma).astype('float')
    
    return tbs.join([tb,ts, prop, brokage], how='outer').eval('buy_sell_diff=total_buy-total_sell')

# Extract useful information from DataFrame, and compose a silver report
#   df[16] = "白银交易量前十名"
#   df[9] = "白银代理成交量前十名"
def gen_silver_report(df):
    df_silver_top10 =  df[16]
    tbs = df_silver_top10[4:14][[1,2]]
    tbs.columns = ['bank_name','total_buy_sell']
    tbs.set_index('bank_name', inplace=True)
    tbs['total_buy_sell'] = tbs['total_buy_sell'].apply(remove_comma).astype('float')
    
    tb = df_silver_top10[4:14][[4,5]]
    tb.columns = ['bank_name','total_buy']
    tb.set_index('bank_name', inplace=True)
    tb['total_buy'] = tb['total_buy'].apply(remove_comma).astype('float')
    
    ts = df_silver_top10[4:14][[7,8]]
    ts.columns = ['bank_name','total_sell']
    ts.set_index('bank_name', inplace=True)
    ts['total_sell'] = ts['total_sell'].apply(remove_comma).astype('float')
    
    df_silver_broker = df[19]
    brokage =   df_silver_broker[1:11][[1,2]]
    brokage.columns =  ['bank_name','brokage']
    brokage.set_index('bank_name', inplace=True)
    brokage['brokage'] = brokage['brokage'].apply(remove_comma).astype('float')
    
    return tbs.join([tb,ts,  brokage], how='outer').eval('buy_sell_diff=total_buy-total_sell')


def pdf_to_excel(pdf_file, exel_output):
    
    df = read_pdf(pdf_file,pages="all", multiple_tables=True)
    gold_rpt_df = gen_gold_report(df)    
    silver_rpt_df = gen_silver_report(df)
    
    with pandas.ExcelWriter(exel_output) as writer:
        gold_rpt_df.to_excel(writer, sheet_name="gold")
        silver_rpt_df.to_excel(writer, sheet_name="silver")

def gen_reports_batch():
    reports = ["20190311-20190315周报",
               "20190304-20190308周报",
               "20190225-20190301周报",
               "20190218-20190222周报",
               "20190211-20190215周报",
               "20190128-20190201周报",
               "20190121-20190125周报",
               "20190114-20190118周报",
               "20190107-20190111周报",
               "20190102-20190104周报"]

    for report_name in reports:
        print("Processing {}".format(report_name))
        pdf_file = os.path.join(DOWNLOAD_FOLDER, report_name + ".pdf")
        output_file = os.path.join(OUTPUT_FOLDER, report_name + ".xlsx")
        pdf_to_excel(pdf_file, output_file)
        print("Done: {}".format(report_name).center(50, '-'))

def download_and_convert_latest(n):
    reports = analysis_content(WEB_SITE_URL)
    if n> len(reports):
        print("{} reports in first page, set n {}-->{}".format(len(reports), n, len(reports)))
        n = len(reports)
    reports_to_do = reports[:n]
    download_reports(DOWNLOAD_FOLDER, reports_to_do )
    
    
    for latest_report in reports_to_do:    
        print("Processing '{}' ...".format(latest_report.name))
        pdf_file = os.path.join(DOWNLOAD_FOLDER, latest_report.name + ".pdf")
        output_file = os.path.join(OUTPUT_FOLDER, latest_report.name + ".xlsx")
        if os.path.isfile(output_file):
            print("Target file '{}' exists, skipping...".format(output_file))
        else:
            pdf_to_excel(pdf_file, output_file)            
            print("Saved to file: '{}'".format(output_file))
        
        print("Done: {}".format(latest_report.name).center(80, '-'))

if __name__ == '__main__':
    #default: analysis first page, and generate report for the first one
    download_and_convert_latest(1)
