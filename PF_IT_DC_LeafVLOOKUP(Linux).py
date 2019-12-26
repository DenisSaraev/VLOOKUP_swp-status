#!/usr/bin/env python3.7

from datetime import datetime
start_time = datetime.now()

print ('The script is running\nPlease wait\n...')

import os
import glob
import pandas as pd
import textfsm

# path to nexus configs (switch port)
path=r'/home/denis.saraev/my_python/reports/PF_DC_IT-Leafs_switchport/'
# path to final file (switch port)
path2=r'my_python/scripts/PF_DC_IT_VLOOKUP/Result/'

# path to nexus configs (interface status)
path3 = r'/home/denis.saraev/my_python/reports/PF_DC_IT-Leafs_status/'
# path to final file (interface status)
path4=r'my_python/scripts/PF_DC_IT_VLOOKUP/Result/'

# path to vlookup table
path5=r'my_python/scripts/PF_DC_IT_VLOOKUP/Result/'

# Opening xlsx file for finall data (interface status)
writer = pd.ExcelWriter(path4 + 'StatusReportBrief.xlsx', engine='xlsxwriter')

# this list will collect names of pages in 'StatusReportBrief.xlsx'. we will need it later for VLOOKUP
page_names=[]

# Processing with every nexus 'sh int status' file
for every_file in glob.glob(os.path.join(path3, '*.csv')):
    try:
        # we need hostname for each list's name in xlsx
        hostname = os.path.basename(every_file)
        # delete filename extension (.csv)
        hostname = os.path.splitext(hostname)[0]

        page_names.append(hostname)

        # making dataframe
        conf = pd.read_csv(every_file)
        df = pd.DataFrame(conf)
        # delete excess rows
        df = df.drop([0, 1, 2, 3, 5])
        df = df.drop(df.index[55:])
        # turn the correct row into a header
        df.columns = df.iloc[0]
        df = df.drop([4])
        # delete not correct index
        df = df.reset_index()
        df = df.drop(['index'], axis=1)

        # making temp csv file
        df.to_csv(path4 + 'StatusTempFile.csv', index=False)

        # read temp csv as file with 'columns with a fixed width'
        df_new = pd.read_fwf(path4 + 'StatusTempFile.csv')

        # transform this file into xlsx
        df_new.to_excel(writer, sheet_name=hostname)
        workbook = writer.book
        # each config will have own list in xlsx
        worksheet = writer.sheets[hostname]

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1})

        text_format = workbook.add_format({
            'bold': False,
            'text_wrap': False,
            'valign': 'top',
            'border': 1})

        # width of columns
        worksheet.set_column('A:A', 0, text_format)
        worksheet.set_column('B:B', 8.57, text_format)
        worksheet.set_column('C:C', 19.71, text_format)
        worksheet.set_column('D:D', 10.29, text_format)
        worksheet.set_column('E:E', 8.86, text_format)
        worksheet.set_column('F:F', 11.14, text_format)
        worksheet.set_column('G:G', 10.43, text_format)
        worksheet.set_column('H:H', 10.57, text_format)

        # colour for header
        for col_num, value in enumerate(df_new.columns.values):
            worksheet.write(0, col_num + 1, value, header_format)
        # filter in the header
        worksheet.autofilter('A1:H1')

    # this 'except' skipping conf file with empty data. for example - switch unreachable or collector cant login on it
    except KeyError:
        continue

writer.save()
# delete temp file
os.remove(path4 + 'StatusTempFile.csv')

print ('30% complete\n...')

# Opening xlsx file for finall data (switchport)
writer = pd.ExcelWriter(path2 + 'SwitchportReportBrief.xlsx', engine='xlsxwriter')

# Processing with every nexus 'sh int switchport' file
for every_file in glob.glob(os.path.join(path, '*.csv')):
    try:
        # reading each file
        with open(every_file,'r') as file:
            nexus_conf = file.read()

        # we need hostname for each list's name in xlsx
        hostname=os.path.basename(every_file)
        # delete filename extension (.csv)
        hostname=os.path.splitext(hostname)[0]

        # open template for textfsm 'sh int switchport'
        with open(r'/home/denis.saraev/my_python/packages/templates_textfsm/cisco_nxos_show_interfaces_switchport.textfsm') as template:
            fsm = textfsm.TextFSM(template)
        # textfsm parsing in config
        result=fsm.ParseText(nexus_conf)

        # making dataframe
        df = pd.DataFrame(result)
        # create temp csv file
        df.to_csv(path2+'SwitchportTempFile.csv', index=False, sep='|', header=fsm.header)
        # read temp csv as file with 'columns with a fixed width'
        df_new = pd.read_csv(path2+'SwitchportTempFile.csv', sep='|')

        # we need the same format of interface's name for VLOOKUP.
        # in 'sh int status' its 'Eth1/1', but in 'sh sw port' its 'Ethernet1/1'
        # the next step makes them the same
        df_new['INTERFACE'].replace(regex=True, inplace=True, to_replace=r'Ethernet', value=r'Eth')
        # also we need the same names of columns in both tables
        df_new.rename(columns={'INTERFACE':'Port'}, inplace=True)

        # transform this file into xlsx
        df_new.to_excel(writer, sheet_name=hostname)
        workbook = writer.book
        # each config will have own list in xlsx
        worksheet = writer.sheets[hostname]

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': False,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1})

        text_format = workbook.add_format({
            'bold': False,
            'text_wrap': False,
            'valign': 'top',
            'border': 1})

        # width of columns

        worksheet.set_column('A:A', 0, text_format)
        worksheet.set_column('B:B', 12.14, text_format)
        worksheet.set_column('C:C', 14.29, text_format)
        worksheet.set_column('D:D', 24.57, text_format)
        worksheet.set_column('E:E', 8.14, text_format)
        worksheet.set_column('F:F', 15.14, text_format)
        worksheet.set_column('G:G', 15.14, text_format)
        worksheet.set_column('H:H', 19.29, text_format)
        worksheet.set_column('I:I', 13.86, text_format)

        # colour for header
        for col_num, value in enumerate(df_new.columns.values):
            worksheet.write(0, col_num + 1, value, header_format)
        # filter in the header
        worksheet.autofilter('A1:I1')


    # this 'except' skipping conf file with empty data. for example - switch unreachable or collector cant login on it
    except KeyError:
        continue

writer.save()
# delete temp file
os.remove(path2 + 'SwitchportTempFile.csv')

print ('60% complete\n...')

# Opening xlsx file for finall data (VLOOKUP)
writer = pd.ExcelWriter(path5 + 'InterfaceVLookupBrief.xlsx', engine='xlsxwriter')

# VLOOKUP for both tables
for each_page in page_names:
    # index_col=0 for dont import index column
    df_status=pd.read_excel(path4+'StatusReportBrief.xlsx',sheet_name=each_page, index_col=0)
    df_switchport=pd.read_excel(path2+'SwitchportReportBrief.xlsx',sheet_name=each_page, index_col=0)
    df_sum=df_status.merge(df_switchport, on='Port', how='left')
    # transform this file into xlsx
    df_sum.to_excel(writer, sheet_name=each_page)
    workbook = writer.book
    # each config will have own list in xlsx
    worksheet = writer.sheets[each_page]

    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': False,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1})

    text_format = workbook.add_format({
        'bold': False,
        'text_wrap': False,
        'valign': 'top',
        'border': 1})

    worksheet.set_column('A:A', 0, text_format)
    worksheet.set_column('B:B', 8.57, text_format)
    worksheet.set_column('C:C', 19.71, text_format)
    worksheet.set_column('D:D', 10.29, text_format)
    worksheet.set_column('E:E', 8.86, text_format)
    worksheet.set_column('F:F', 11.14, text_format)
    worksheet.set_column('G:G', 10.43, text_format)
    worksheet.set_column('H:H', 10.57, text_format)

    worksheet.set_column('I:I', 14.29, text_format)
    worksheet.set_column('J:J', 24.57, text_format)
    worksheet.set_column('K:K', 8.14, text_format)
    worksheet.set_column('L:L', 15.14, text_format)
    worksheet.set_column('M:M', 15.14, text_format)
    worksheet.set_column('N:N', 19.29, text_format)
    worksheet.set_column('O:O', 13.86, text_format)

    # colour for header
    for col_num, value in enumerate(df_sum.columns.values):
        worksheet.write(0, col_num + 1, value, header_format)
    # filter in the header
    worksheet.autofilter('A1:O1')

writer.save()

end_time = datetime.now()
print('Finished\nDuration: {}'.format(end_time - start_time))
print ('Your files\n/home/denis.saraev/my_python/scripts/PF_DC_IT_VLOOKUP/Result/InterfaceVLookupBrief.xlsx')
