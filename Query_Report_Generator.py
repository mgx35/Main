import datetime
import tkinter as tk
from tkinter import ttk, END
from tkinter import filedialog as fd
from tkinter.filedialog import askopenfilename
import os
import numpy as np
import pandas as pd
from pandas import Series, DataFrame
import xlrd
from tkinter import StringVar
from tkinter import *
import time
import calendar
from datetime import date
from datetime import datetime

# creating the date object of today's date
todays_date = date.today()

# Create a GUI app
app = tk.Tk()

# Specify the title and dimensions to app
app.title('Query Report Generator')
app.geometry('615x510')

# Create a textfield for instructions to add query file
query_instr = ttk.Label(app, text="Select query file:")
query_instr.grid(column=0, row=0)

# Create a textfield for file name
q_file_name = tk.Text(app, height=1, width=50)
q_file_name.grid(column=2, row=0)

# Create an open file button
open_q_file_button = ttk.Button(app, text='Browse', command=lambda: open_query_file())
open_q_file_button.grid(column=1, row=0)

# Create a textfield for instructions to add pvc file
pvc_instr = ttk.Label(app, text="Select Prev v Curr file:")
pvc_instr.grid(column=0, row=1)

# Create a textfield for file name
pvc_file_name = tk.Text(app, height=1, width=50)
pvc_file_name.grid(column=2, row=1)

# Create an open file button
open_pvc_button = ttk.Button(app, text='Browse',
                         command=lambda: open_pvc_file())
open_pvc_button.grid(column=1, row=1)
#Pre project query frame
pre_proj_q_frame= tk.LabelFrame(app,text="Pre-Project Queries")
pre_proj_q_frame.grid(column=0, row=3, columnspan=3,sticky=tk.W, pady=5)

# Create button to generate a Data-Priority Report
gen_data_priority_rep = ttk.Button(pre_proj_q_frame, text='Generate Data Priority Report', width=31,
                         command=lambda: sizes_vs_preproject())
gen_data_priority_rep.grid(column=0, row=0, pady=3, columnspan=1)

# Create button to generate a Query-Priority Report
gen_q_priority_rep = ttk.Button(pre_proj_q_frame, text='Generate Query Priority Report', width=31,
                         command=lambda: pre_proj_q_vs_sizes())
gen_q_priority_rep.grid(column=0, row=1, pady=3, columnspan=1)

#PPC button
PPC_snapshot_button = ttk.Button(app, text='Generate Data Snapshot',
                         command=lambda: generate_PPC_snaphot())
PPC_snapshot_button.grid(column=2, row=2, sticky= "nw")

#blanking of the status notes
#success_note_data_priority = ttk.Label(pre_proj_q_frame, text="", width=10)
###success_note_query_priority.grid(column=1, row=1, padx=3)

#Reason for change query frame
rch_q_frame = tk.LabelFrame(app,text="Reasons for Change")
rch_q_frame.grid(column=0, row=6, columnspan=3,sticky=tk.W, pady=5)
#last year to consider DD menu
l_y_to_cons_label=ttk.Label(rch_q_frame,text="Last year to consider changes:")
l_y_to_cons_label.grid(column=0, row=0,pady=3,padx=2)
#yr_opt = [2019,2020,2021,2023,2024]
yr_opt_variable = IntVar()
yr_opt = []
yo=-5
while yo<=5:
    yr_opt.append(todays_date.year+yo)
    yo+=1

yr_opt_variable.set(yr_opt[1]) # default value
yr_lim = ttk.OptionMenu(rch_q_frame, yr_opt_variable,yr_opt_variable.get(),*yr_opt)
yr_lim.grid(column=1, row=0)
#significance threshold sel
sig_q_t=DoubleVar()
sig_q_label=ttk.Label(rch_q_frame,text="Enter Query Significance Limit:")
sig_q_label.grid(column=0, row=1,pady=3,padx=2)
sig_q_t = tk.Text(rch_q_frame, height=1,width=5)
sig_q_t.insert('1.0',"10")
sig_q_t.grid(column=1, row=1)

#Reasons for change report button
gen_changes_vs_rch_rep = ttk.Button(rch_q_frame, text='Generate Reasons for Change Report', width=50,
                         command=lambda:  generate_sig_changes_v_rch())
gen_changes_vs_rch_rep.grid(column=0, row=4, pady=3, columnspan=4)

#Post project query frame
post_proj_q_frame= tk.LabelFrame(app,text="Post-Project Queries")
post_proj_q_frame.grid(column=0, row=10, columnspan=3,sticky=tk.W, pady=5)
#txt field for snapshot sel
query_instr = ttk.Label(post_proj_q_frame, text="Select snapshot file:")
query_instr.grid(column=0, row=0)

#date selection frame
date_frame = tk.LabelFrame(app,text="Query cutoff date")
date_frame.grid(column=0, row=2, columnspan=3,sticky=tk.W, pady=5,padx=5,ipady=5,ipadx=5)

#year
cod_yr_lim_label=ttk.Label(date_frame,text="Year")
cod_yr_lim_label.grid(column=0, row=0, pady=3,padx=2)
cod_yr_opt_variable = StringVar()
cod_yr_opt = []
cod_yo=-3
while cod_yo<=0:
    cod_yr_opt.append(todays_date.year+cod_yo)
    cod_yo+=1
cod_yr_opt_variable.set(cod_yr_opt[2]) # default value
cod_yr_lim = ttk.OptionMenu(date_frame, cod_yr_opt_variable, cod_yr_opt_variable.get(), *cod_yr_opt)
cod_yr_lim.grid(column=0, row=1)

#month
cod_mo_lim_label=ttk.Label(date_frame,text="Month")
cod_mo_lim_label.grid(column=1, row=0, pady=3,padx=2)

cod_mo_opt = ['01','02','03','04','05','06','07','08','09','10','11','12']
cod_mo_opt_variable = StringVar()

cod_mo_lim = ttk.OptionMenu(date_frame, cod_mo_opt_variable, str(todays_date.month+1), *cod_mo_opt)
cod_mo_lim.grid(column=1, row=1)
#Date
cod_da_lim_label=ttk.Label(date_frame,text="Date")
cod_da_lim_label.grid(column=2, row=0, pady=3,padx=2)
cod_da_opt = ['01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31']
cod_da_opt_variable = StringVar()

cod_da_lim = ttk.OptionMenu(date_frame, cod_da_opt_variable, str(todays_date.day), *cod_da_opt)
cod_da_lim.grid(column=2, row=1)

# Create a textfield for file name
snap_file_name = tk.Text(post_proj_q_frame, height=1, width=50)
snap_file_name.grid(column=2, row=0)

# Create button to generate a Data-Priority Report
snap_data_priority_rep = ttk.Button(post_proj_q_frame, text='Generate Data Priority Report', width=31,
                         command=lambda: snap_vs_size_data_priority())
snap_data_priority_rep.grid(column=0, row=1, pady=3, columnspan=2)

# Create button to generate a Query-Priority Report
snap_q_priority_rep = ttk.Button(post_proj_q_frame, text='Generate Query Priority Report', width=31,
                         command=lambda: snap_vs_size_q_priority())
snap_q_priority_rep.grid(column=0, row=2, pady=3, columnspan=2)
#open file button
open_snap_file_button = ttk.Button(post_proj_q_frame, text='Browse', command=lambda: open_snap_file())
open_snap_file_button.grid(column=1, row=0)

def open_query_file():
    # Specify the file types
    filetypes = [("Excel files", "*.xlsx; *.xls")]

    # Show the open file dialog by specifying path
    qf = fd.askopenfile(filetypes=filetypes,
                       initialdir="D:/Downloads")
    # clear prev entry
    q_file_name.delete("1.0", "end")
    #get the query filepath
    query_filepath = os.path.abspath(qf.name)
    # Insert the text extracted from file in a textfield
    q_file_name.insert('0.0', query_filepath)

def open_pvc_file():
    # Specify the file types

    filetypes = [("Excel files", "*.xlsx; *.xls")]
    #clear prev entry
    pvc_file_name.delete("1.0","end")
    # Show the open file dialog by specifying path
    pvc_f = fd.askopenfile(filetypes=filetypes,
                       initialdir="D:/Downloads")
    #get the query filepath
    query_filepath = os.path.abspath(pvc_f.name)
    # Insert the text extracted from file in a textfield
    pvc_file_name.insert('1.0', query_filepath)

def sizes_vs_preproject():
    #success_note_data_priority = ttk.Label(pre_proj_q_frame, text="Working...", width=10)
    #success_note_data_priority.grid(column=1, row=0, padx=3)
    # read and rename Geography to Country
    # q_0= pd.read_excel(r'C:\Users\Golov\Downloads\CFS2023 queries.xlsx')
    pathq = q_file_name.get("0.0", "end")
    pathq = str(pathq[0:len(pathq) - 1])
    q_0 = pd.read_excel(pathq)
    q_0.rename(columns={'Geography': 'Country'}, inplace=True)
    # clean useless data from df

    q_1 = q_0.set_index(['Country', 'Product'])
    q_1.drop(index=np.NaN, columns=['Unnamed: 0', 'Unnamed: 1', 'Resolved / RA approved', 'Client', 'Sub-project'],
             inplace=True)
    # remove all entries from prior editions
    str_date = cod_yr_opt_variable.get() + '-' + cod_mo_opt_variable.get() + '-' + cod_da_opt_variable.get()
    start_date = str_date
    mask = (q_1['Date'] > start_date)
    q_2 = q_1.loc[mask]
    # filter for pre-check queries
    q_pre_chk = q_2.loc[q_2['Notes'].str.contains("pre-checks", case=False, na=False)]
    q_pre_chk_cln = q_pre_chk.loc[:, ['Dataset', 'Query', 'Response']]

    #filter for sizes queries
    mask2=(q_pre_chk_cln['Dataset']=='Size')
    q_3=q_pre_chk_cln.loc[mask2]
    pre_q_sizes=q_3.loc[:,['Query','Response']]
    pre_q_sizes.rename(columns={'Query':'Size Query','Response':'Size Query Response'},inplace=True)
    # filter for share queries
    mask2 = (q_pre_chk_cln['Dataset'] == 'Market Shares/Ranks')
    q_5 = q_pre_chk_cln.loc[mask2]
    pre_q_shares = q_5.loc[:, ['Query', 'Response']]
    pre_q_shares.rename(columns={'Query': 'Share Query', 'Response': 'Share Query Response'}, inplace=True)
    # filter for measure queries
    mask2 = (q_pre_chk_cln['Dataset'] == 'Measure')
    q_6 = q_pre_chk_cln.loc[mask2]
    pre_q_measures = q_6.loc[:, ['Query', 'Response']]
    pre_q_measures.rename(columns={'Query': 'Measure Query', 'Response': 'Measure Query Response'}, inplace=True)
    pre_check_q_column_sorted_interim = pre_q_sizes.merge(pre_q_shares, how='left', on=['Country', 'Product'])
    pre_check_q_column_sorted = pre_check_q_column_sorted_interim.merge(pre_q_measures, how='left', on=['Country', 'Product'])

    pathpvc= pvc_file_name.get("0.0","end")
    pathpvc = str(pathpvc[0:len(pathpvc) - 1])
    pvc0 = pd.read_excel(pathpvc)
    # set index to country and product
    new_col = np.array(pvc0.head(1))
    pvc0.columns = new_col[0]
    pvc1 = pvc0.drop(0)
    pvc2 = pvc1.set_index(['Country', 'Product'])
    # remove useless columns
    pvc31 = pvc2.drop(columns=['Sub-project', 'Sector'])
    # select columns for comparison with pre check
    imp_col_f_pchk = ['Data type', 'Unit']
    col = pvc31.columns
    for i in col:
        if i.find('curr') >= 0 or i.find('diff') >= 0:
            imp_col_f_pchk.append(i)
        else:
            imp_col_f_pchk = imp_col_f_pchk
    pvc3 = pvc31.loc[:, imp_col_f_pchk]

    pvc_precheck = pvc3.loc[:, imp_col_f_pchk]
    # merging of the two files
    pre_check_vs_data = pvc_precheck.merge(pre_check_q_column_sorted, how='left', on=['Country', 'Product'])
    pre_check_vs_data_clean = pre_check_vs_data.loc[:,
                              ['Data type', 'Unit', 'curr2015', 'diff2015', 'curr2016', 'diff2016',
                               'curr2017', 'diff2017', 'curr2018', 'diff2018', 'curr2019', 'diff2019',
                               'curr2020', 'diff2020', 'curr2021', 'diff2021', 'curr2022', 'diff2022',
                               'curr2023', 'diff2023', 'curr2024', 'diff2024', 'curr2025', 'diff2025',
                               'curr2026', 'diff2026', 'Size Query', 'Size Query Response',
                               'Share Query', 'Share Query Response', 'Measure Query',
                               'Measure Query Response']]
    pre_check_vs_data_clean.reset_index(inplace=True)
    pre_check_vs_data_clean.to_excel('Data_vs_Precheck.xlsx', index=False)
    #success_note_data_priority = ttk.Label(pre_proj_q_frame, text="Success!", width=10)
    #success_note_data_priority.grid(column=1, row=0,padx=3)

def pre_proj_q_vs_sizes():
    #success_note_query_priority= ttk.Label(pre_proj_q_frame, text="Working...",width=10)
    #success_note_query_priority.grid(column=1,row=1, padx=3)
    #read and rename Geography to Country
    #q_0= pd.read_excel(r'C:\Users\Golov\Downloads\CFS2023 queries.xlsx')
    pathq= q_file_name.get("0.0","end")
    pathq = str(pathq[0:len(pathq) - 1])
    q_0 = pd.read_excel(pathq)
    q_0.rename(columns={'Geography':'Country'},inplace=True)
    #clean useless data from df
    q_1=q_0.set_index(['Country','Product'])
    q_1.drop(index=np.NaN,columns=['Unnamed: 0','Unnamed: 1','Resolved / RA approved','Client','Sub-project'],inplace=True)
    #remove all entries from prior editions
    str_date = cod_yr_opt_variable.get() + '-' + cod_mo_opt_variable.get() + '-' + cod_da_opt_variable.get()
    start_date = str_date
    mask = (q_1['Date'] > start_date)
    q_2=q_1.loc[mask]
    #filter for pre-check queries
    q_pre_chk = q_2.loc[q_2['Notes'].str.contains("pre-checks", case=False, na=False)]

    q_pre_chk_cln=q_pre_chk.loc[:,['Dataset', 'Query','Response']]

    #read pvc file
    #pvc0=pd.read_excel(r'C:\Users\Golov\Downloads\Sizes_prev_vs_curr_gbl_excel.xls')
    pathpvc= pvc_file_name.get("0.0","end")
    pathpvc = str(pathpvc[0:len(pathpvc) - 1])
    pvc0 = pd.read_excel(pathpvc)
    #set index to country and product
    new_col= np.array(pvc0.head(1))
    pvc0.columns=new_col[0]
    pvc1=pvc0.drop(0)
    pvc2=pvc1.set_index(['Country','Product'])
    # remove useless columns
    pvc31 = pvc2.drop(columns=['Sub-project', 'Sector'])
    # select columns for comparison with pre check
    imp_col_f_pchk = ['Data type', 'Unit']
    col = pvc31.columns
    for i in col:
        if i.find('curr') >= 0 or i.find('diff') >= 0:
            imp_col_f_pchk.append(i)
        else:
            imp_col_f_pchk = imp_col_f_pchk
    pvc3 = pvc31.loc[:, imp_col_f_pchk]
    pvc_precheck=pvc3.loc[: , imp_col_f_pchk]
    #combining dataframes
    pre_check_vs_data = q_pre_chk_cln.merge(pvc_precheck, how = 'left', on = ['Country', 'Product'])
    #reorder to put queries at the end
    pre_check_vs_data_corr_ord=pre_check_vs_data[['Data type', 'Unit', 'curr2015',
           'diff2015', 'curr2016', 'diff2016', 'curr2017', 'diff2017', 'curr2018',
           'diff2018', 'curr2019', 'diff2019', 'curr2020', 'diff2020', 'curr2021',
           'diff2021', 'curr2022', 'diff2022', 'curr2023', 'diff2023', 'curr2024',
           'diff2024', 'curr2025', 'diff2025', 'curr2026', 'diff2026','Dataset', 'Query', 'Response']]
    pre_check_vs_data_corr_ord.reset_index(inplace=True)
    pre_check_vs_data_corr_ord.to_excel('Pre_proj_vs_Data.xlsx',index=False)
    #success_note_query_priority= ttk.Label(pre_proj_q_frame, text="Success!", width=10)
    #success_note_query_priority.grid(column=1,row=1,padx=3)

def generate_PPC_snaphot():
    # read excel file
    #pvc0 = pd.read_excel(r'C:\Users\Golov\Downloads\Sizes_prev_vs_curr_gbl_excel.xls')
    pathpvc = pvc_file_name.get("0.0", "end")
    pathpvc = str(pathpvc[0:len(pathpvc) - 1])
    pvc0 = pd.read_excel(pathpvc)
    new_col = np.array(pvc0.head(1))
    pvc0.columns = new_col[0]
    pvc1 = pvc0.drop(0)
    pvc2 = pvc1.set_index(['Country', 'Data type', 'Sector', 'Product',
                           'Unit'])
    col = pvc2.columns
    # select only curr rows
    curr_col = []
    for i in col:
        if i.find('curr') >= 0:
            curr_col.append(i)
        else:
            curr_col = curr_col
    pvc3 = pvc2.loc[:, curr_col]
    #file naming and export to excel
    pvc3.reset_index(inplace=True)
    PPC_snapshot_name= str(datetime.now().timestamp()) + '_Curr_Data_Snapshot.xlsx'
    pvc3.to_excel(PPC_snapshot_name, index=False)

def generate_sig_changes_v_rch():
    # read and rename Geography to Country
    #q_0 = pd.read_excel(r'C:\Users\Golov\Downloads\CFS2023 queries.xlsx')
    pathq= q_file_name.get("0.0","end")
    pathq = str(pathq[0:len(pathq) - 1])
    q_0 = pd.read_excel(pathq)
    q_0.rename(columns={'Geography': 'Country'}, inplace=True)
    # clean useless data from df
    q_1 = q_0.set_index(['Country', 'Product'])
    q_11 = DataFrame(q_1.loc[:,['Type','Dataset','Query','Response','Date','Notes']])
    # remove all entries from prior editions
    #start_date = '2022-03-03'
    str_date = cod_yr_opt_variable.get() + '-' + cod_mo_opt_variable.get() + '-' + cod_da_opt_variable.get()
    start_date=datetime.strptime(str_date, '%Y-%m-%d')
    mask = (q_11['Date'] > start_date)
    q_2 = q_11.loc[mask]
    #filter for Type for reason for change queries for sizes
    mask3=(q_2['Type']=='Reason for change')
    q_all_rch=q_2.loc[mask3]
    mask4=(q_all_rch['Dataset']=='Size')
    q_size_rch=q_all_rch.loc[mask4]
    q_size_rch_cln=DataFrame(q_size_rch.loc[:,'Response'])
    q_size_rch_cln.rename(columns={"Response":"Reasons for Change: Sizes"},inplace=True)
    mask5 = (q_all_rch['Dataset'] == 'Measure')
    q_rch_measure = q_all_rch.loc[mask5]
    q_rch_measure_cln = DataFrame(q_rch_measure.loc[:, 'Response'])
    q_rch_measure_cln.rename(columns={'Response': 'Reasons for Change: Measures'}, inplace=True)
    q_rch_comb = q_size_rch_cln.merge(q_rch_measure_cln, how='left', on=['Country', 'Product'])
    #read pvc file
    #pvc0 = pd.read_excel(r'C:\Users\Golov\Downloads\Sizes_prev_vs_curr_gbl_excel.xls')
    pathpvc = pvc_file_name.get("0.0", "end")
    pathpvc = str(pathpvc[0:len(pathpvc) - 1])
    pvc0 = pd.read_excel(pathpvc)
    new_col = np.array(pvc0.head(1))
    pvc0.columns = new_col[0]
    pvc1 = pvc0.drop(0)
    # generate list of researched mkts
    all_mkts = pvc1['Country']
    all_mkts_uniq = all_mkts.unique()
    res_mkts = np.delete(all_mkts_uniq, np.where(
        all_mkts_uniq == [['World'], ['Western Europe'], ['Eastern Europe'], ['North America'], ['Latin America'],
                          ['Asia Pacific'], ['Australasia'], ['Middle East and Africa']]))
    all_cats = pvc1['Product']
    all_cats_uniq = all_cats.unique()
    excl_cat_list = ['Consumer Foodservice', 'Consumer Foodservice by Type',
                     'Chained Consumer Foodservice (duplicate)',
                     'Independent Consumer Foodservice (duplicate)', 'Cafés/Bars', 'Bars/Pubs', 'Cafés',
                     'Juice/Smoothie Bars',
                     'Specialist Coffee and Tea Shops',
                     'Full-Service Restaurants', 'Chained Full-Service Restaurants',
                     'Independent Full-Service Restaurants',
                     'Full-Service Restaurants by Type',
                     'Asian Full-Service Restaurants',
                     'European Full-Service Restaurants',
                     'Latin American Full-Service Restaurants',
                     'Middle Eastern Full-Service Restaurants',
                     'North American Full-Service Restaurants',
                     'Pizza Full-Service Restaurants',
                     'Other Full-Service Restaurants',
                     'Limited-Service Restaurants',
                     'Limited-Service Restaurants by Type',
                     'Asian Limited-Service Restaurants',
                     'Bakery Products Limited-Service Restaurants',
                     'Burger Limited-Service Restaurants',
                     'Chicken Limited-Service Restaurants',
                     'Convenience Stores Limited-Service Restaurants',
                     'Fish Limited-Service Restaurants',
                     'Ice Cream Limited-Service Restaurants',
                     'Latin American Limited-Service Restaurants',
                     'Middle Eastern Limited-Service Restaurants',
                     'Pizza Limited-Service Restaurants',
                     'Other Limited-Service Restaurants',
                     'Self-Service Cafeterias', 'Street Stalls/Kiosks',
                     'Consumer Foodservice by Chained/Independent',
                     'Chained Consumer Foodservice', 'Chained Cafés/Bars (duplicate)',
                     'Chained Bars/Pubs (duplicate)', 'Chained Cafés (duplicate)',
                     'Chained Juice/Smoothie Bars (duplicate)',
                     'Chained Specialist Coffee and Tea Shops (duplicate)',
                     'Chained Limited-Service Restaurants (duplicate)',
                     'Chained Asian Limited-Service Restaurants (duplicate)',
                     'Chained Bakery Products Limited-Service Restaurants (duplicate)',
                     'Chained Burger Limited-Service Restaurants (duplicate)',
                     'Chained Chicken Limited-Service Restaurants (duplicate)',
                     'Chained Convenience Stores Limited-Service Restaurants (duplicate)',
                     'Chained Fish Limited-Service Restaurants (duplicate)',
                     'Chained Ice Cream Limited-Service Restaurants (duplicate)',
                     'Chained Latin American Limited-Service Restaurants (duplicate)',
                     'Chained Middle Eastern Limited-Service Restaurants (duplicate)',
                     'Chained Pizza Limited-Service Restaurants (duplicate)',
                     'Chained Other Limited-Service Restaurants (duplicate)',
                     'Chained Full-Service Restaurants (duplicate)',
                     'Chained Asian Full-Service Restaurants (duplicate)',
                     'Chained European Full-Service Restaurants (duplicate)',
                     'Chained Latin American Full-Service Restaurants (duplicate)',
                     'Chained Middle Eastern Full-Service Restaurants (duplicate)',
                     'Chained North American Full-Service Restaurants (duplicate)',
                     'Chained Pizza Full-Service Restaurants (duplicate)',
                     'Chained Other Full-Service Restaurants (duplicate)',
                     'Chained Self-Service Cafeterias (duplicate)',
                     'Chained Street Stalls/Kiosks (duplicate)',
                     'Independent Consumer Foodservice',
                     'Independent Cafés/Bars (duplicate)',
                     'Independent Bars/Pubs (duplicate)',
                     'Independent Cafés (duplicate)',
                     'Independent Juice/Smoothie Bars (duplicate)',
                     'Independent Specialist Coffee and Tea Shops (duplicate)',
                     'Independent Limited-Service Restaurants (duplicate)',
                     'Independent Asian Limited-Service Restaurants (duplicate)',
                     'Independent Bakery Products Limited-Service Restaurants (duplicate)',
                     'Independent Burger Limited-Service Restaurants (duplicate)',
                     'Independent Chicken Limited-Service Restaurants (duplicate)',
                     'Independent Convenience Stores Limited-Service Restaurants (duplicate)',
                     'Independent Ice Cream Limited-Service Restaurants (duplicate)',
                     'Independent Fish Limited-Service Restaurants (duplicate)',
                     'Independent Latin American Limited-Service Restaurants (duplicate)',
                     'Independent Middle Eastern Limited-Service Restaurants (duplicate)',
                     'Independent Pizza Limited-Service Restaurants (duplicate)',
                     'Independent Other Limited-Service Restaurants (duplicate)',
                     'Independent Full-Service Restaurants (duplicate)',
                     'Independent Asian Full-Service Restaurants (duplicate)',
                     'Independent European Full-Service Restaurants (duplicate)',
                     'Independent Latin American Full-Service Restaurants (duplicate)',
                     'Independent Middle Eastern Full-Service Restaurants (duplicate)',
                     'Independent North American Full-Service Restaurants (duplicate)',
                     'Independent Pizza Full-Service Restaurants (duplicate)',
                     'Independent Other Full-Service Restaurants (duplicate)',
                     'Independent Self-Service Cafeterias (duplicate)',
                     'Independent Street Stalls/Kiosks (duplicate)',
                     'Consumer Foodservice by Location',
                     'Consumer Foodservice Through Standalone',
                     'Consumer Foodservice Through Leisure',
                     'Consumer Foodservice Through Retail',
                     'Consumer Foodservice Through Lodging',
                     'Consumer Foodservice Through Travel',
                     'Consumer Foodservice Eat-In/Takeaway',
                     'Consumer Foodservice Eat-In', 'Consumer Foodservice Takeaway',
                     'Consumer Foodservice Home Delivery',
                     'Consumer Foodservice Drive-Through',
                     'Consumer Foodservice Online/Offline Ordering',
                     'Consumer Foodservice Online Ordering',
                     'Consumer Foodservice Offline Ordering',
                     'Retailing convenience stores and forecourt retailers',
                     'Convenience Stores', 'Forecourt Retailers',
                     'Consumer Foodservice Food Sales',
                     'Consumer Foodservice Drink Sales']
    res_cats_uniq_mask = np.isin(all_cats_uniq, excl_cat_list, invert=True)
    res_cats_uniq = all_cats_uniq[res_cats_uniq_mask]
    # remove all non research cats and mkts
    pvc2 = pvc1.set_index(['Country', 'Product'])
    pvc_rchmkt = pvc2.loc[res_mkts, :]
    pvc_fin = pvc_rchmkt.loc[slice(None), res_cats_uniq, :]
    # separate column types
    pvc_all_col = pvc_fin.columns
    diff_col = ['Data type', 'Sector', 'Unit']
    actd_col = ['Data type', 'Sector', 'Unit']
    prev_col = ['Data type', 'Sector', 'Unit']
    curr_col = ['Data type', 'Sector', 'Unit']
    for i in pvc_all_col:
        if str(i).find("diff") >= 0:
            diff_col.append(i)
        elif str(i).find("2022") == 0 and str(i).find("act") < 0:
            diff_col.append(i)
        elif str(i).find("actDiff") >= 0:
            actd_col.append(i)
        elif str(i).find("prev") >= 0:
            prev_col.append(i)
        elif str(i).find("curr") >= 0:
            curr_col.append(i)
        else:
            continue
    diff_col_arr = np.array(diff_col)
    # remove all data after target year
    ty = yr_opt_variable.get()
    i = 0
    yr_excl_list = []
    while i <= 10:
        ty += 1
        yr = 'diff' + str(ty)
        yr_excl_list.append(yr)
        i += 1
    yr_excl_list
    hist_diff_col_mask = np.isin(diff_col_arr, yr_excl_list, invert=True)
    hist_diff_col_mask
    hist_diff_col = diff_col_arr[hist_diff_col_mask]
    pvc_fin_hist = pvc_fin.loc[:, hist_diff_col]
    hist_diff_col_only_data = np.delete(hist_diff_col, [0, 1, 2])
    # filter out everything under the change limiter
    diff_lim = float(sig_q_t.get('1.0','1.9'))
    sig_hist_diff = pvc_fin_hist[(abs(pvc_fin_hist.loc[:, hist_diff_col_only_data]) >= diff_lim).any(1)]
    # merging of the two df
    pvc_sig_hist_vs_q_sorted = sig_hist_diff.merge(q_rch_comb, how='left', on=['Country', 'Product'])
    pvc_sig_hist_vs_q_sorted.reset_index(inplace=True)
    # to excel
    pvc_sig_hist_vs_q_sorted.to_excel('Significant_Changes_vs_Reasons_for_Change.xlsx', index=False)

def open_snap_file():
    # Specify the file types
    filetypes = [("Excel files", "*.xlsx; *.xls")]

    # Show the open file dialog by specifying path
    snf = fd.askopenfile(filetypes=filetypes,initialdir=r"C:\Users\Golov\PycharmProjects\CFSquery analysis")
    # clear prev entry
    snap_file_name.delete("1.0", "end")
    #get the query filepath
    snf_filepath = os.path.abspath(snf.name)
    # Insert the text extracted from file in a textfield
    snap_file_name.insert('0.0', snf_filepath)

def snap_vs_size_data_priority():
    # read query file and rename Geography to Country
    #q_0_post = pd.read_excel(r'C:\Users\Golov\Downloads\CFS2023 queries.xlsx')
    pathq = q_file_name.get("0.0", "end")
    pathq = str(pathq[0:len(pathq) - 1])
    q_0_post = pd.read_excel(pathq)
    q_0_post.rename(columns={'Geography': 'Country'}, inplace=True)
    # clean useless data from df
    q_1_post = q_0_post.set_index(['Country', 'Product'])
    q_1_post.drop(index=np.NaN, columns=['Unnamed: 0', 'Unnamed: 1', 'Resolved / RA approved', 'Client', 'Sub-project'],
                  inplace=True)
    # remove all entries from prior editions
    str_date = cod_yr_opt_variable.get() + '-' + cod_mo_opt_variable.get() + '-' + cod_da_opt_variable.get()
    start_date = str_date
    mask = (q_1_post['Date'] > start_date)
    q_2_post = q_1_post.loc[mask]
    # filter for pre-check queries
    q_post_chk = q_2_post.loc[q_2_post['Notes'].str.contains("post-checks", case=False, na=False)]

    # filter for sizes queries
    mask2 = (q_post_chk['Dataset'] == 'Size')
    q_3_post = q_post_chk.loc[mask2]
    post_q_sizes = q_3_post.loc[:, ['Query', 'Response']]
    post_q_sizes.rename(columns={'Query': 'Size Query', 'Response': 'Size Query Response'}, inplace=True)
    # filter for share queries
    mask2 = (q_post_chk['Dataset'] == 'Market Shares/Ranks')
    q_5_post = q_post_chk.loc[mask2]
    post_q_shares = q_5_post.loc[:, ['Query', 'Response']]
    post_q_shares.rename(columns={'Query': 'Share Query', 'Response': 'Share Query Response'}, inplace=True)
    # filter for measure queries
    mask2 = (q_post_chk['Dataset'] == 'Measure')
    q_6_post = q_post_chk.loc[mask2]
    post_q_measures = q_6_post.loc[:, ['Query', 'Response']]
    post_q_measures.rename(columns={'Query': 'Measure Query', 'Response': 'Measure Query Response'}, inplace=True)
    post_check_q_column_sorted_interim = post_q_sizes.merge(post_q_shares, how='left', on=['Country', 'Product'])
    post_check_q_column_sorted = post_check_q_column_sorted_interim.merge(post_q_measures, how='left',
                                                                          on=['Country', 'Product'])
    # read snapshot
    pathsnap = snap_file_name.get("0.0","end")
    pathsnap = str(pathsnap[0:len(pathsnap) - 1])
    snap0 = pd.read_excel(pathsnap)
    # set index
    snap1 = snap0.set_index(['Country', 'Sector', 'Product', 'Data type', 'Unit'])
    #pvc0 = pd.read_excel(r'C:\Users\Golov\Downloads\Sizes_prev_vs_curr_gbl_excel.xls')
    pathpvc= pvc_file_name.get("0.0","end")
    pathpvc = str(pathpvc[0:len(pathpvc) - 1])
    pvc0 = pd.read_excel(pathpvc)
    # set index to country and product
    new_col = np.array(pvc0.head(1))
    pvc0.columns = new_col[0]
    pvc1 = pvc0.drop(0)
    pvc2 = pvc1.set_index(['Country', 'Sector', 'Product', 'Data type', 'Unit'])
    # select only curr rows
    col=pvc2.columns
    curr_col = []
    for i in col:
        if i.find('curr') >= 0:
            curr_col.append(i)
        else:
            curr_col = curr_col
    pvc3 = pvc2.loc[:, curr_col]
    act_diff_pvc_snap = snap1 - pvc3
    pct_diff_pvc_snap = act_diff_pvc_snap / snap1 * 100
    pct_diff_pvc_snap1 = pct_diff_pvc_snap.reset_index()
    pct_diff_pvc_snap2 = pct_diff_pvc_snap1.set_index(['Country', 'Product'])
    # rename cols
    diff_pvc_snap_cols = list(pct_diff_pvc_snap2.columns)
    diff_pvc_snap_cols_rev = list(np.char.replace(diff_pvc_snap_cols, 'curr', 'diff'))
    snap_diff_renaming_dict = {diff_pvc_snap_cols[i]: diff_pvc_snap_cols_rev[i] for i in range(len(diff_pvc_snap_cols))}
    pct_diff_pvc_snap2.rename(columns=snap_diff_renaming_dict, inplace=True)
    # merge files
    post_check_vs_data = pct_diff_pvc_snap2.merge(post_check_q_column_sorted, how='left', on=['Country', 'Product'])
    post_check_vs_data2 = post_check_vs_data.reset_index()
    post_check_vs_data2.to_excel('Data Changes vs Post Checks.xlsx', index=False)

def snap_vs_size_q_priority():
    pathq = q_file_name.get("0.0", "end")
    pathq = str(pathq[0:len(pathq) - 1])
    q_0_post = pd.read_excel(pathq)
    q_0_post.rename(columns={'Geography': 'Country'}, inplace=True)
    # clean useless data from df
    q_1_post = q_0_post.set_index(['Country', 'Product'])
    q_1_post.drop(index=np.NaN, columns=['Unnamed: 0', 'Unnamed: 1', 'Resolved / RA approved', 'Client', 'Sub-project'],
                  inplace=True)
    str_date = cod_yr_opt_variable.get() + '-' + cod_mo_opt_variable.get() + '-' + cod_da_opt_variable.get()
    start_date = str_date
    mask = (q_1_post['Date'] > start_date)
    q_2_post = q_1_post.loc[mask]
    # filter for post-check queries
    q_post_chk = q_2_post.loc[q_2_post['Notes'].str.contains("post-checks", case=False, na=False)]

    # read snapshot
    pathsnap = snap_file_name.get("0.0", "end")
    pathsnap = str(pathsnap[0:len(pathsnap) - 1])
    snap0 = pd.read_excel(pathsnap)
    # set index
    snap1 = snap0.set_index(['Country', 'Sector', 'Product', 'Data type', 'Unit'])
    # pvc0 = pd.read_excel(r'C:\Users\Golov\Downloads\Sizes_prev_vs_curr_gbl_excel.xls')
    pathpvc = pvc_file_name.get("0.0", "end")
    pathpvc = str(pathpvc[0:len(pathpvc) - 1])
    pvc0 = pd.read_excel(pathpvc)
    # set index to country and product
    new_col = np.array(pvc0.head(1))
    pvc0.columns = new_col[0]
    pvc1 = pvc0.drop(0)
    pvc2 = pvc1.set_index(['Country', 'Sector', 'Product', 'Data type', 'Unit'])
    # select only curr rows
    col = pvc2.columns
    curr_col = []
    for i in col:
        if i.find('curr') >= 0:
            curr_col.append(i)
        else:
            curr_col = curr_col
    pvc3 = pvc2.loc[:, curr_col]
    act_diff_pvc_snap = snap1 - pvc3
    pct_diff_pvc_snap = act_diff_pvc_snap / snap1 * 100
    pct_diff_pvc_snap1 = pct_diff_pvc_snap.reset_index()
    pct_diff_pvc_snap2 = pct_diff_pvc_snap1.set_index(['Country', 'Product'])
    # rename cols
    diff_pvc_snap_cols = list(pct_diff_pvc_snap2.columns)
    diff_pvc_snap_cols_rev = list(np.char.replace(diff_pvc_snap_cols, 'curr', 'diff'))
    snap_diff_renaming_dict = {diff_pvc_snap_cols[i]: diff_pvc_snap_cols_rev[i] for i in range(len(diff_pvc_snap_cols))}
    pct_diff_pvc_snap2.rename(columns=snap_diff_renaming_dict, inplace=True)
    # merge files
    post_check_vs_data = q_post_chk.merge(pct_diff_pvc_snap2, how='left', on=['Country', 'Product'])
    post_check_vs_data2 = post_check_vs_data.reset_index()
    post_check_vs_data2.to_excel('Post Checks vs Data Changes.xlsx', index=False)

# Button for closing
exit_button = Button(app, text="Exit", width=30, command=app.destroy)
exit_button.grid(column=1, row=15,padx=10,pady=20,columnspan=2)

app.mainloop()