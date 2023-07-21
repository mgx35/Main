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
from keras.models import Sequential
from keras.layers import LSTM, Dense
from sklearn.linear_model import LinearRegression
from itertools import product
from xlsxwriter import Workbook

global scale_tr_start
global scale_tr_end
global scale_pred_start
global training_data_selections
global frame_yr
global first_year
global last_year
global lrt_output

# Create a GUI app
app = tk.Tk()

# Specify the title and dimensions to app
app.title('Unit Price Estimator')
app.geometry('800x600')

# Create a textfield for instructions to add pvc file
sz_instr = ttk.Label(app, text="Select Sizes Report file:")
sz_instr.grid(column=0, row=1)

# Create a textfield for file name
sz_file_name = tk.Text(app, height=1, width=50)
sz_file_name.grid(column=2, row=1)

# Create an open file button
open_sz_button = ttk.Button(app, text='Browse',
                         command=lambda: open_sizes_file())
open_sz_button.grid(column=1, row=1)

# Create a load file button
load_button = ttk.Button(app, text='Load',
                         command=lambda: load_sizes_file())
load_button.grid(column=1, row=2)

def open_sizes_file():
    # Specify the file types
    filetypes = [("Excel files", "*.xlsx; *.xls")]

    # Show the open file dialog by specifying path
    sz = fd.askopenfile(filetypes=filetypes,
                       initialdir="D:/Downloads")
    # clear prev entry
    sz_file_name.delete("1.0", "end")
    #get the query filepath
    sz_filepath = os.path.abspath(sz.name)
    # Insert the text extracted from file in a textfield
    sz_file_name.insert('0.0', sz_filepath)

def load_sizes_file():
    path_sz = sz_file_name.get("0.0", "end")
    path_sz = str(path_sz[0:len(path_sz) - 1])
    size_0 = pd.read_excel(path_sz)

    # read file and set cols to be correct
    #size_0 = pd.read_excel(r'C:\Users\Golov\Downloads\sizes_review_cty_excel.xls')
    new_col = np.array(size_0.head(1))
    size_0.columns = new_col[0]
    size_0 = size_0.drop(0)
    # remove non-res prod cat
    all_cats = size_0['Product']
    all_cats_uniq = all_cats.unique()
    excl_cat_list = ['Consumer Foodservice', 'Consumer Foodservice by Type', 'Takeaway',
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
                     'Consumer Foodservice Through Leisure', 'Drive-Through'
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
                     'Convenience Stores', 'Forecourt Retailers', 'Consumer Foodservice Drink Sales',
                     'Consumer Foodservice Food Sales', 'Eat-In', 'Delivery', 'Drive-Through'
                                                                              'Consumer Foodservice Drink Sales',
                     'Travel', 'Inbound Arrivals', 'Arrivals by Country of Origin', 'Arrivals from Asia Pacific',
                     'Arrivals from Afghanistan', 'Arrivals from American Samoa', 'Arrivals from Armenia',
                     'Arrivals from Azerbaijan', 'Arrivals from Bangladesh', 'Arrivals from Bhutan',
                     'Arrivals from Brunei Darussalam', 'Arrivals from Cambodia', 'Arrivals from China',
                     'Arrivals From Fiji', 'Arrivals from French Polynesia', 'Arrivals from Guam',
                     'Arrivals from Hong Kong, China', 'Arrivals from India', 'Arrivals from Indonesia',
                     'Arrivals from Japan', 'Arrivals from Kazakhstan', 'Arrivals from Kiribati',
                     'Arrivals from Kyrgyzstan', 'Arrivals from Laos', 'Arrivals from Macau, China',
                     'Arrivals from Malaysia', 'Arrivals from Maldives', 'Arrivals from Mongolia',
                     'Arrivals from Myanmar', 'Arrivals from Nauru', 'Arrivals from Nepal',
                     'Arrivals from New Caledonia', 'Arrivals from North Korea', 'Arrivals from Pakistan',
                     'Arrivals from Papua New Guinea', 'Arrivals from Philippines', 'Arrivals from Samoa',
                     'Arrivals from Singapore', 'Arrivals from Solomon Islands', 'Arrivals from South Korea',
                     'Arrivals from Sri Lanka', 'Arrivals from Taiwan', 'Arrivals from Tajikistan',
                     'Arrivals from Thailand', 'Arrivals from Tonga', 'Arrivals from Turkmenistan',
                     'Arrivals from Tuvalu', 'Arrivals from Uzbekistan', 'Arrivals from Vanuatu',
                     'Arrivals from Vietnam', 'Arrivals from Australasia', 'Arrivals from Australia',
                     'Arrivals from New Zealand', 'Arrivals from Eastern Europe', 'Arrivals from Albania',
                     'Arrivals from Belarus', 'Arrivals from Bosnia and Herzegovina', 'Arrivals from Bulgaria',
                     'Arrivals from Croatia', 'Arrivals from Czech Republic', 'Arrivals from Estonia',
                     'Arrivals from Georgia', 'Arrivals from Hungary', 'Arrivals from Kosovo', 'Arrivals from Latvia',
                     'Arrivals from Lithuania', 'Arrivals from North Macedonia', 'Arrivals from Moldova',
                     'Arrivals from Montenegro', 'Arrivals from Poland', 'Arrivals from Romania',
                     'Arrivals from Russia', 'Arrivals from Serbia', 'Arrivals from Slovakia', 'Arrivals from Slovenia',
                     'Arrivals from Ukraine', 'Arrivals from Latin America', 'Arrivals from Anguilla',
                     'Arrivals from Antigua and Barbuda', 'Arrivals from Argentina', 'Arrivals from Aruba',
                     'Arrivals from Bahamas', 'Arrivals from Barbados', 'Arrivals from Belize', 'Arrivals from Bermuda',
                     'Arrivals from Bolivia', 'Arrivals from Brazil', 'Arrivals from British Virgin Islands',
                     'Arrivals from Cayman Islands', 'Arrivals from Chile', 'Arrivals from Colombia',
                     'Arrivals from Costa Rica', 'Arrivals from Cuba', 'Arrivals from Curaçao',
                     'Arrivals from Dominica', 'Arrivals from Dominican Republic', 'Arrivals from Ecuador',
                     'Arrivals from El Salvador', 'Arrivals from French Guiana', 'Arrivals from Grenada',
                     'Arrivals from Guadeloupe', 'Arrivals from Guatemala', 'Arrivals from Guyana',
                     'Arrivals from Haiti', 'Arrivals from Honduras', 'Arrivals from Jamaica',
                     'Arrivals from Martinique', 'Arrivals from Mexico', 'Arrivals from Nicaragua',
                     'Arrivals from Panama', 'Arrivals from Paraguay', 'Arrivals from Peru',
                     'Arrivals from Puerto Rico', 'Arrivals from Sint Maarten', 'Arrivals from St Kitts and Nevis',
                     'Arrivals from St Lucia', 'Arrivals from St Vincent and the Grenadines', 'Arrivals from Suriname',
                     'Arrivals from Trinidad and Tobago', 'Arrivals from Uruguay', 'Arrivals from US Virgin Islands',
                     'Arrivals from Venezuela', 'Arrivals from Middle East and Africa', 'Arrivals from Algeria',
                     'Arrivals from Angola', 'Arrivals from Bahrain', 'Arrivals from Benin', 'Arrivals from Botswana',
                     'Arrivals from Burkina Faso', 'Arrivals from Burundi', 'Arrivals from Cameroon',
                     'Arrivals from Cabo Verde', 'Arrivals from Central African Republic', 'Arrivals from Chad',
                     'Arrivals from Comoros', 'Arrivals from Congo, Democratic Republic',
                     'Arrivals from Congo-Brazzaville', 'Arrivals from Djibouti', 'Arrivals from Egypt',
                     'Arrivals from Equatorial Guinea', 'Arrivals from Eritrea', 'Arrivals from Ethiopia',
                     'Arrivals from Gabon', 'Arrivals from Gambia', 'Arrivals from Ghana', 'Arrivals from Guinea',
                     'Arrivals from Guinea-Bissau', 'Arrivals from Iran', 'Arrivals from Iraq', 'Arrivals from Israel',
                     'Arrivals from Jordan', 'Arrivals from Kenya', 'Arrivals from Kuwait', 'Arrivals from Lebanon',
                     'Arrivals from Lesotho', 'Arrivals from Liberia', 'Arrivals from Libya',
                     'Arrivals from Madagascar', 'Arrivals from Malawi', 'Arrivals from Mali',
                     'Arrivals from Mauritania', 'Arrivals from Mauritius', 'Arrivals from Morocco',
                     'Arrivals from Mozambique', 'Arrivals from Namibia', 'Arrivals from Niger',
                     'Arrivals from Nigeria', 'Arrivals from Oman', 'Arrivals from Qatar', 'Arrivals from Réunion',
                     'Arrivals from Rwanda', 'Arrivals from Sao Tomé e Príncipe', 'Arrivals from Saudi Arabia',
                     'Arrivals from Senegal', 'Arrivals from Seychelles', 'Arrivals from Sierra Leone',
                     'Arrivals from Somalia', 'Arrivals from South Africa', 'Arrivals from South Sudan',
                     'Arrivals from Sudan', 'Arrivals from Eswatini', 'Arrivals from Syria', 'Arrivals from Tanzania',
                     'Arrivals from Togo', 'Arrivals from Tunisia', 'Arrivals from Uganda',
                     'Arrivals from United Arab Emirates', 'Arrivals from Yemen', 'Arrivals from Zambia',
                     'Arrivals from Zimbabwe', 'Arrivals from North America', 'Arrivals from Canada',
                     'Arrivals from US', 'Arrivals from Western Europe', 'Arrivals from Andorra',
                     'Arrivals from Austria', 'Arrivals from Belgium', 'Arrivals from Cyprus', 'Arrivals from Denmark',
                     'Arrivals from Finland', 'Arrivals from France', 'Arrivals from Germany',
                     'Arrivals from Gibraltar', 'Arrivals from Greece', 'Arrivals from Iceland',
                     'Arrivals from Ireland', 'Arrivals from Italy', 'Arrivals from Liechtenstein',
                     'Arrivals from Luxembourg', 'Arrivals from Malta', 'Arrivals from Monaco',
                     'Arrivals from Netherlands', 'Arrivals from Norway', 'Arrivals from Portugal',
                     'Arrivals from Spain', 'Arrivals from Sweden', 'Arrivals from Switzerland', 'Arrivals from Turkey',
                     'Arrivals from United Kingdom', 'Arrivals from Other Countries', 'Inbound Length of Stay',
                     'Inbound Tourism by Mode of Transport', 'Air Arrivals', 'Business Air Arrivals',
                     'Leisure Air Arrivals', 'Land Arrivals', 'Business Land Arrivals', 'Leisure Land Arrivals',
                     'Rail Arrivals', 'Business Rail Arrivals', 'Leisure Rail Arrivals', 'Water Arrivals',
                     'Business Water Arrivals', 'Leisure Water Arrivals', 'Inbound Tourism by Purpose of Visit',
                     'Business Inbound', 'MICE Inbound', 'Other Business Inbound', 'Leisure Inbound', 'VFR Inbound',
                     'Other Leisure Inbound', 'Inbound Tourism Spending', 'Inbound Business Tourism Spending',
                     'Inbound Leisure Tourism Spending', 'Inbound Spending on Lodging',
                     'Inbound Spending Excluding Lodging', 'Inbound Spending on Activities', 'Inbound Spending on Food',
                     'Inbound Spending on Shopping', 'Inbound Spending on Retail Shopping',
                     'Inbound Spending on Duty-free Shopping', 'Inbound Spending on Travel Modes',
                     'Inbound Spending on Other', 'Inbound City Arrivals', 'Vienna', 'Salzburg', 'Sölden',
                     'Saalbach-Hinterglemm', 'Mayrhofen', 'Ischgl', 'Sankt Anton am Arlberg', 'Lech', 'Innsbruck',
                     'Linz', 'Domestic Tourism', 'Domestic Tourism By Destination', 'Steiermark', 'Kärnten',
                     'Oberösterreich', 'Niederösterreich', 'Tirol', 'Burgenland', 'Wien', 'Vorarlberg',
                     'Domestic Tourism Destination Subtype 10', 'Domestic Tourism Destination Leading Cities', 'Graz',
                     'Domestic Tourism Destination Leading City Subtype 6',
                     'Domestic Tourism Destination Leading City Subtype 7',
                     'Domestic Tourism Destination Leading City Subtype 8',
                     'Domestic Tourism Destination Leading City Subtype 9',
                     'Domestic Tourism Destination Leading City Subtype 10', 'Domestic Tourism by Mode of Transport',
                     'Domestic Tourism by Air', 'Domestic Business Tourism By Air', 'Domestic Leisure Tourism By Air',
                     'Domestic Tourism by Land', 'Domestic Business Tourism By Land',
                     'Domestic Leisure Tourism By Land', 'Domestic Tourism by Rail',
                     'Domestic Business Tourism By Rail', 'Domestic Leisure Tourism By Rail',
                     'Domestic Tourism by Water', 'Domestic Business Tourism By Water',
                     'Domestic Leisure Tourism By Water', 'Domestic Tourism by Purpose of Visit', 'Domestic Business',
                     'Domestic Other Business', 'Domestic Leisure', 'Domestic Other Leisure', 'Domestic Spending',
                     'Domestic Spending on Lodging', 'Domestic Spending on Shopping', 'Domestic Spending on Other',
                     'Outbound Departures', 'Outbound Departures Source Markets', 'Outbound Departures to Asia Pacific',
                     'Outbound Departures to Afghanistan', 'Outbound Departures to American Samoa',
                     'Outbound Departures to Armenia', 'Outbound Departures to Azerbaijan',
                     'Outbound Departures to Bangladesh', 'Outbound Departures to Bhutan',
                     'Outbound Departures to Brunei Darussalam', 'Outbound Departures to Cambodia',
                     'Outbound Departures to China', 'Outbound Departures to Fiji',
                     'Outbound Departures to French Polynesia', 'Outbound Departures to Guam',
                     'Outbound Departures to Hong Kong, China', 'Outbound Departures to India',
                     'Outbound Departures to Indonesia', 'Outbound Departures to Japan',
                     'Outbound Departures to Kazakhstan', 'Outbound Departures to Kiribati',
                     'Outbound Departures to Kyrgyzstan', 'Outbound Departures to Laos',
                     'Outbound Departures to Macau, China', 'Outbound Departures to Malaysia',
                     'Outbound Departures to Maldives', 'Outbound Departures to Mongolia',
                     'Outbound Departures to Myanmar', 'Outbound Departures to Nauru', 'Outbound Departures to Nepal',
                     'Outbound Departures to New Caledonia', 'Outbound Departures to North Korea',
                     'Outbound Departures to Pakistan', 'Outbound Departures to Papua New Guinea',
                     'Outbound Departures to Philippines', 'Outbound Departures to Samoa',
                     'Outbound Departures to Singapore', 'Outbound Departures to Solomon Islands',
                     'Outbound Departures to South Korea', 'Outbound Departures to Sri Lanka',
                     'Outbound Departures to Taiwan', 'Outbound Departures to Tajikistan',
                     'Outbound Departures to Thailand', 'Outbound Departures to Tonga',
                     'Outbound Departures to Turkmenistan', 'Outbound Departures to Tuvalu',
                     'Outbound Departures to Uzbekistan', 'Outbound Departures to Vanuatu',
                     'Outbound Departures to Vietnam', 'Outbound Departures to Australasia',
                     'Outbound Departures to Australia', 'Outbound Departures to New Zealand',
                     'Outbound Departures to Eastern Europe', 'Outbound Departures to Albania',
                     'Outbound Departures to Belarus', 'Outbound Departures to Bosnia and Herzegovina',
                     'Outbound Departures to Bulgaria', 'Outbound Departures to Croatia',
                     'Outbound Departures to Czech Republic', 'Outbound Departures to Estonia',
                     'Outbound Departures to Georgia', 'Outbound Departures to Hungary',
                     'Outbound Departures to Kosovo', 'Outbound Departures to Latvia',
                     'Outbound Departures to Lithuania', 'Outbound Departures to North Macedonia',
                     'Outbound Departures to Moldova', 'Outbound Departures to Montenegro',
                     'Outbound Departures to Poland', 'Outbound Departures to Romania', 'Outbound Departures to Russia',
                     'Outbound Departures to Serbia', 'Outbound Departures to Slovakia',
                     'Outbound Departures to Slovenia', 'Outbound Departures to Ukraine',
                     'Outbound Departures to Latin America', 'Outbound Departures to Anguilla',
                     'Outbound Departures to Antigua and Barbuda', 'Outbound Departures to Argentina',
                     'Outbound Departures to Aruba', 'Outbound Departures to Bahamas',
                     'Outbound Departures to Barbados', 'Outbound Departures to Belize',
                     'Outbound Departures to Bermuda', 'Outbound Departures to Bolivia',
                     'Outbound Departures to Brazil', 'Outbound Departures to British Virgin Islands',
                     'Outbound Departures to Cayman Islands', 'Outbound Departures to Chile',
                     'Outbound Departures to Colombia', 'Outbound Departures to Costa Rica',
                     'Outbound Departures to Cuba', 'Outbound Departures to Curaçao', 'Outbound Departures to Dominica',
                     'Outbound Departures to Dominican Republic', 'Outbound Departures to Ecuador',
                     'Outbound Departures to El Salvador', 'Outbound Departures to French Guiana',
                     'Outbound Departures to Grenada', 'Outbound Departures to Guadeloupe',
                     'Outbound Departures to Guatemala', 'Outbound Departures to Guyana',
                     'Outbound Departures to Haiti', 'Outbound Departures to Honduras',
                     'Outbound Departures to Jamaica', 'Outbound Departures to Martinique',
                     'Outbound Departures to Mexico', 'Outbound Departures to Nicaragua',
                     'Outbound Departures to Panama', 'Outbound Departures to Paraguay', 'Outbound Departures to Peru',
                     'Outbound Departures to Puerto Rico', 'Outbound Departures to Sint Maarten',
                     'Outbound Departures to St Kitts and Nevis', 'Outbound Departures to St Lucia',
                     'Outbound Departures to St Vincent and the Grenadines', 'Outbound Departures to Suriname',
                     'Outbound Departures to Trinidad and Tobago', 'Outbound Departures to Uruguay',
                     'Outbound Departures to US Virgin Islands', 'Outbound Departures to Venezuela',
                     'Outbound Departures to Middle East and Africa', 'Outbound Departures to Algeria',
                     'Outbound Departures to Angola', 'Outbound Departures to Bahrain', 'Outbound Departures to Benin',
                     'Outbound Departures to Botswana', 'Outbound Departures to Burkina Faso',
                     'Outbound Departures to Burundi', 'Outbound Departures to Cameroon',
                     'Outbound Departures to Cabo Verde', 'Outbound Departures to Central African Republic',
                     'Outbound Departures to Chad', 'Outbound Departures to Comoros',
                     'Outbound Departures to Congo, Democratic Republic', 'Outbound Departures to Congo-Brazzaville',
                     'Outbound Departures to Djibouti', 'Outbound Departures to Egypt',
                     'Outbound Departures to Equatorial Guinea', 'Outbound Departures to Eritrea',
                     'Outbound Departures to Ethiopia', 'Outbound Departures to Gabon', 'Outbound Departures to Gambia',
                     'Outbound Departures to Ghana', 'Outbound Departures to Guinea',
                     'Outbound Departures to Guinea-Bissau', 'Outbound Departures to Iran',
                     'Outbound Departures to Iraq', 'Outbound Departures to Israel', 'Outbound Departures to Jordan',
                     'Outbound Departures to Kenya', 'Outbound Departures to Kuwait', 'Outbound Departures to Lebanon',
                     'Outbound Departures to Lesotho', 'Outbound Departures to Liberia', 'Outbound Departures to Libya',
                     'Outbound Departures to Madagascar', 'Outbound Departures to Malawi',
                     'Outbound Departures to Mali', 'Outbound Departures to Mauritania',
                     'Outbound Departures to Mauritius', 'Outbound Departures to Morocco',
                     'Outbound Departures to Mozambique', 'Outbound Departures to Namibia',
                     'Outbound Departures to Niger', 'Outbound Departures to Nigeria', 'Outbound Departures to Oman',
                     'Outbound Departures to Qatar', 'Outbound Departures to Réunion', 'Outbound Departures to Rwanda',
                     'Outbound Departures to Sao Tomé e Príncipe', 'Outbound Departures to Saudi Arabia',
                     'Outbound Departures to Senegal', 'Outbound Departures to Seychelles',
                     'Outbound Departures to Sierra Leone', 'Outbound Departures to Somalia',
                     'Outbound Departures to South Africa', 'Outbound Departures to South Sudan',
                     'Outbound Departures to Sudan', 'Outbound Departures to Eswatini', 'Outbound Departures to Syria',
                     'Outbound Departures to Tanzania', 'Outbound Departures to Togo', 'Outbound Departures to Tunisia',
                     'Outbound Departures to Uganda', 'Outbound Departures to United Arab Emirates',
                     'Outbound Departures to Yemen', 'Outbound Departures to Zambia', 'Outbound Departures to Zimbabwe',
                     'Outbound Departures to North America', 'Outbound Departures to Canada',
                     'Outbound Departures to US', 'Outbound Departures to Western Europe',
                     'Outbound Departures to Andorra', 'Outbound Departures to Austria',
                     'Outbound Departures to Belgium', 'Outbound Departures to Cyprus',
                     'Outbound Departures to Denmark', 'Outbound Departures to Finland',
                     'Outbound Departures to France', 'Outbound Departures to Germany',
                     'Outbound Departures to Gibraltar', 'Outbound Departures to Greece',
                     'Outbound Departures to Iceland', 'Outbound Departures to Ireland', 'Outbound Departures to Italy',
                     'Outbound Departures to Liechtenstein', 'Outbound Departures to Luxembourg',
                     'Outbound Departures to Malta', 'Outbound Departures to Monaco',
                     'Outbound Departures to Netherlands', 'Outbound Departures to Norway',
                     'Outbound Departures to Portugal', 'Outbound Departures to Spain', 'Outbound Departures to Sweden',
                     'Outbound Departures to Switzerland', 'Outbound Departures to Turkey',
                     'Outbound Departures to United Kingdom', 'Outbound Departures to Other Destinations',
                     'Outbound Length of Stay', 'Outbound Tourism by Mode of Transport', 'Air Outbound',
                     'Business Air Outbound', 'Leisure Air Outbound', 'Land Outbound', 'Business Land Outbound',
                     'Leisure Land Outbound', 'Rail Outbound', 'Business Rail Outbound', 'Leisure Rail Outbound',
                     'Water Outbound', 'Business Water Outbound', 'Leisure Water Outbound',
                     'Outbound Tourism by Purpose of Visit', 'Business Outbound', 'Leisure Outbound',
                     'Outbound Tourism Spending', 'Outbound Business Spending', 'Outbound Leisure Spending',
                     'Outbound Spending on Lodging', 'Outbound Spending on Activities', 'Outbound Spending on Food',
                     'Outbound Spending on Shopping', 'Outbound Spending on Retail Shopping',
                     'Outbound Spending on Duty-free Shopping', 'Outbound Spending on Travel Modes',
                     'Outbound Spending on Other', 'Travel Modes', 'Airlines', 'Airlines by Category',
                     'Scheduled Airlines', 'Ancillary Revenue', 'International Airlines', 'Airlines by Channel',
                     'Airlines Online', 'Airlines Online via Direct', 'Airlines Online via Intermediaries',
                     'Air through Package Holidays', 'Air Online Sales less Air through Package Holidays',
                     'Airlines Offline', 'Airlines Offline via Direct', 'Airlines Offline via Intermediaries',
                     'Surface Travel Modes', 'Surface Travel Modes by Category', 'Surface Travel Modes by Channel',
                     'Surface Travel Modes Online', 'Surface Travel Modes Online via Direct',
                     'Surface Travel Modes Online via Intermediaries', 'Surface Travel Modes Offline',
                     'Surface Travel Modes Offline via Direct', 'Surface Travel Modes Offline via Intermediaries',
                     'Lodging (Destination)', 'Hotels by Category', 'Hotels by Channel', 'Hotels Online',
                     'Hotels Online via Direct', 'Hotels Online via Intermediaries', 'Hotels Offline',
                     'Hotels Offline via Direct', 'Hotels Offline via Intermediaries', 'Short-Term Rentals Online',
                     'Short-Term rentals Online via Direct', 'Short-term Rentals Online via Intermediaries',
                     'Short-Term Rentals Offline', 'Short-term Rentals Offline via Direct',
                     'Short-term Rentals Offline via Intermediaries', 'Other Lodging', 'Other Lodging by Category',
                     'Other Lodging by Channel', 'Other Lodging Online', 'Other Lodging Online Direct',
                     'Other Lodging Online Intermediaries', 'Other Lodging Offline', 'Other Lodging Offline via Direct',
                     'Other Lodging Offline via Intermediaries', 'Lodging (Destination) by Channel',
                     'Lodging (Destination) Online', 'Lodging (Destination) Online via Direct',
                     'Lodging (Destination) Online via Intermediaries', 'Lodging (Destination) Offline',
                     'Lodging (Destination) Offline via Direct', 'Lodging (Destination) Offline via Intermediaries',
                     'In-Destination Spending', 'Attractions', 'Experiences', 'Shopping', 'Retail Shopping',
                     'Duty-Free Shopping', 'Wellness', 'Other In-Destination Spending',
                     'In-Destination Spending by Channel', 'In-Destination Spending Online',
                     'In-Destination Spending Online Direct', 'In-Destination Spending Online Intermediaries',
                     'In-Destination Spending Offline', 'In-Destination Spending Offline Direct',
                     'In-Destination Spending Offline Intermediaries', 'Booking', 'Booking Offline', 'Booking Online',
                     'Mobile Travel', 'Leisure Travel', 'Leisure Air Travel Online',
                     'Leisure Air Travel Online via Direct', 'Leisure Air Travel Online via Intermediaries',
                     'Leisure Air Travel Offline', 'Leisure Air Travel Offline via Direct',
                     'Leisure Air Travel Offline via Intermediaries', 'Leisure Car Rental Online',
                     'Leisure Car Rental Online via Direct', 'Leisure Car Rental Online via Intermediaries',
                     'Leisure Car Rental Offline', 'Leisure Car Rental Offline via Direct',
                     'Leisure Car Rental Offline via Intermediaries', 'Leisure Cruise Online',
                     'Leisure Cruise Online via Direct', 'Leisure Cruise Online via Intermediaries',
                     'Leisure Cruise Offline', 'Leisure Cruise Offline via Direct',
                     'Leisure Cruise Offline via Intermediaries', 'Leisure Experiences and Attractions Online',
                     'Leisure Experiences and Attractions Online via Direct',
                     'Leisure Experiences and Attractions Online via Intermediaries',
                     'Leisure Experiences and Attractions Offline',
                     'Leisure Experiences and Attractions Offline via Direct',
                     'Leisure Experiences and Attractions Offline via Intermediaries',
                     'Leisure Lodging (Source) Online', 'Leisure Lodging (Source) Online via Direct',
                     'Leisure Lodging (Source) Online via Intermediaries', 'Leisure Lodging (Source) Offline',
                     'Leisure Lodging (Source) Offline via Direct',
                     'Leisure Lodging (Source) Offline via Intermediaries', 'Leisure Packages Online',
                     'Leisure Packages Online via Intermediaries', 'Leisure Packages Offline',
                     'Leisure Packages Offline via Intermediaries', 'Leisure Surface Travel Online',
                     'Leisure Surface Travel Online via Direct', 'Leisure Surface Travel Online via Intermediaries',
                     'Leisure Surface Travel Offline', 'Leisure Surface Travel Offline via Direct',
                     'Leisure Surface Travel Offline via Intermediaries', 'Leisure Other Travel Products Online',
                     'Leisure Other Travel Products Online via Direct',
                     'Leisure Other Travel Products Online via Intermediaries', 'Leisure Other Travel Products Offline',
                     'Leisure Other Travel Products Offline via Direct',
                     'Leisure Other Travel Products Offline via Intermediaries', 'Business Travel',
                     'Business Air Travel Online', 'Business Air Travel Online via Direct',
                     'Business Air Travel Online via Intermediaries', 'Business Air Travel Offline',
                     'Business Air Travel Offline via Direct', 'Business Air Travel Offline via Intermediaries',
                     'Business Car Rental Online', 'Business Car Rental Online via Direct',
                     'Business Car Rental Online via Intermediaries', 'Business Car Rental Offline',
                     'Business Car Rental Offline via Direct', 'Business Car Rental Offline via Intermediaries',
                     'Business Lodging Online', 'Business Lodging Online via Direct',
                     'Business Lodging Online via Intermediaries', 'Business Lodging Offline',
                     'Business Lodging Offline via Direct', 'Business Lodging Offline via Intermediaries',
                     'Business Other Online', 'Business Other Online via Direct',
                     'Business Other Online via Intermediaries', 'Business Other Offline',
                     'Business Other Offline via Direct', 'Business Other Offline via Intermediaries',
                     'Travel Intermediaries', 'Travel Intermediaries Online', 'Travel Intermediaries Offline',
                     'Direct Suppliers', 'Direct Suppliers Online', 'Direct Suppliers Offline', 'Leading Airports',
                     'Vienna International Airport', 'Graz Airport', 'Innsbruck Airport',
                     'Kärntern Airport (Klagenfurt)', 'Blue Danube Airport Linz', 'Salzburg Airport',
                     'Leading Airports Subtype 7', 'Leading Airports Subtype 8', 'Leading Airports Subtype 9',
                     'Leading Airports Subtype 10',
                     'Sum of Cards by Function', 'ATM Cards', 'Commercial Charge Cards', 'Personal Charge Cards',
                     'Commercial Credit Cards', 'Personal Credit Cards', 'Commercial Debit Cards',
                     'Personal Debit Cards', 'Pre-Paid Cards', 'Closed Loop Pre-Paid Cards', 'Open Loop Pre-Paid Cards',
                     'Transactions', 'Total Card Transactions', 'Card Payment Transactions', 'Charge Card Transactions',
                     'Credit Card Transactions', 'Debit Card Transactions', 'Pre-Paid Card Transactions',
                     'Commercial Payment Transactions', 'Commercial Card Payment Transactions',
                     'Commercial Charge Card Transactions (duplicate)',
                     'Commercial Credit Card Transactions (duplicate)',
                     'Commercial Debit Card Transactions (duplicate)',
                     'Commercial Electronic Direct/ACH Transactions (duplicate)',
                     'Commercial Paper Payment Transactions (duplicate)', 'Commercial Cash Transactions (duplicate)',
                     'Commercial Other Paper Transactions (duplicate)', 'Personal Payment Transactions',
                     'Personal Card Payment Transactions', 'Personal Charge Card Transactions (duplicate 2)',
                     'Personal Credit Card Transactions (duplicate 2)', 'Personal Debit Card Transactions (duplicate)',
                     'Pre-Paid Card Transactions (duplicate 2)', 'Closed Loop Pre-Paid Card Transactions (duplicate 2)',
                     'Open Loop Pre-Paid Card Transactions (duplicate 2)', 'Store Card Transactions (duplicate 2)',
                     'Personal Electronic Direct/ACH Transactions (duplicate)',
                     'Personal Paper Payment Transactions (duplicate)', 'Personal Cash Transactions (duplicate)',
                     'Personal Other Paper Transactions (duplicate)', 'Total Non-Card Transactions',
                     'Electronic Direct/ACH Transactions', 'Paper Payment Transactions',
                     'Commercial Paper Payment Transactions', 'Personal Paper Payment Transactions', 'Consumer Lending',
                     'Consumer Credit', 'Average Personal Credit Card Balance', 'Non-Card Lending', 'Auto Lending',
                     'Durables Lending', 'Education Lending', 'Home Lending', 'Other Personal Lending',
                     'Mortgages/Housing', 'Mobile Payments', 'Mobile E-Commerce Payments', 'Data Checks',
                     'Alternative Financial Service Providers', 'Payday', 'Remote Payments',
                     'Sum (Card Holder not Present)', 'Charge Card Transactions (Card Holder not Present)',
                     'Credit Card Transactions (Card Holder not Present)',
                     'Debit Card Transactions (Card Holder not Present)',
                     'Open Loop Pre-Paid Card Transactions (Card Holder not Present)']
    filtered_cats = [cat for cat in all_cats if cat not in excl_cat_list]
    size = size_0[size_0['Product'].isin(filtered_cats)].copy()
    size = size.drop(['Sub-project', 'Region', 'Sector', 'Unit'], axis=1)
    size = size.fillna(0)
    last_year=size.columns[-1]
    first_year=last_year-20

    #year frame set up
    frame_yr = tk.Frame(app, borderwidth=1, relief=tk.RAISED)
    frame_yr.grid(column=0, row=3,columnspan=3, sticky=tk.W)
    frame_yr.columnconfigure(1, minsize=300)
    frame_yr.columnconfigure(2, minsize=200)
    #year scales
    fist_train_yr_instr = ttk.Label(frame_yr, text="First year of train data:")
    fist_train_yr_instr.grid(column=0, row=3, sticky=tk.E)
    scale_tr_start = tk.Scale(frame_yr, from_=first_year, to=last_year, orient=tk.HORIZONTAL)
    scale_tr_start.set(last_year-20)
    scale_tr_start.grid(column=1, row=3, sticky=tk.W+tk.E)

    last_train_yr_instr = ttk.Label(frame_yr, text="Last year of train data:")
    last_train_yr_instr.grid(column=0, row=4, sticky=tk.E)
    scale_tr_end = tk.Scale(frame_yr, from_=first_year, to=last_year, orient=tk.HORIZONTAL)
    scale_tr_end.set(last_year-7)
    scale_tr_end.grid(column=1, row=4, sticky=tk.W+tk.E)

    first_yr_pred_instr = ttk.Label(frame_yr, text="First year to model:")
    first_yr_pred_instr.grid(column=0, row=5, sticky=tk.E)
    scale_pred_start = tk.Scale(frame_yr, from_=first_year, to=last_year, orient=tk.HORIZONTAL)
    scale_pred_start.set(last_year-6)
    scale_pred_start.grid(column=1, row=5, sticky=tk.W+tk.E)

    #dtype frame set up
    frame_dt = tk.Frame(app, borderwidth=1, relief=tk.RAISED)
    frame_dt.grid(column=0, row=4,columnspan=3, sticky=tk.W)

    # Get unique values from the 'Data type' column
    unique_values = size['Data type'].unique()
    checkbox_var_o = tk.IntVar()
    checkbox_o = tk.Checkbutton(frame_dt, text='Outlets', variable=checkbox_var_o, onvalue=1, offvalue=0)
    # Check if 'Outlets' is present in the unique values
    if 'Outlets' in unique_values:
        checkbox_var_o.set(1)
        checkbox_o.select()
    else:
        checkbox_var_o.set(0)
        checkbox_o.deselect()
        checkbox_o.config(state=tk.DISABLED)

    checkbox_o.grid(column=0, row=0, sticky=tk.W)

    sel_rsp_type = tk.StringVar()

    radio_button1 = tk.Radiobutton(frame_dt, text='Foodservice rsp (con/con YrCurr, local)',
                                   value='Foodservice rsp (con/con YrCurr, local)', variable=sel_rsp_type)
    radio_button1.grid(column=1, row=0, sticky=tk.W)
    radio_button1.configure(state='normal' if radio_button1['value'] in unique_values else 'disabled')

    radio_button2 = tk.Radiobutton(frame_dt, text='Foodservice rsp (curr/curr, local)',
                                   value='Foodservice rsp (curr/curr, local)', variable=sel_rsp_type)
    radio_button2.grid(column=1, row=1, sticky=tk.W)
    radio_button2.configure(state='normal' if radio_button2['value'] in unique_values else 'disabled')

    radio_button3 = tk.Radiobutton(frame_dt, text='Foodservice rsp US$ (con/con YrCurr, fixed-exg-YrCur)',
                                   value='Foodservice rsp US$ (con/con YrCurr, fixed-exg-YrCur)', variable=sel_rsp_type)
    radio_button3.grid(column=1, row=2, sticky=tk.W)
    radio_button3.configure(state='normal' if radio_button3['value'] in unique_values else 'disabled')

    radio_button4 = tk.Radiobutton(frame_dt, text='Foodservice rsp euro (con/con YrCurr, fixed-exg-YrCur)',
                                   value='Foodservice rsp euro (con/con YrCurr, fixed-exg-YrCur)',
                                   variable=sel_rsp_type)
    radio_button4.grid(column=1, row=3, sticky=tk.W)
    radio_button4.configure(state='normal' if radio_button4['value'] in unique_values else 'disabled')

    radio_button5 = tk.Radiobutton(frame_dt, text='None',
                                   value='None',
                                   variable=sel_rsp_type)
    radio_button5.grid(column=1, row=4, sticky=tk.W)
    radio_button5.configure(state='normal')

    # Select the first enabled radio button
    enabled_buttons = [radio_button1, radio_button2, radio_button3, radio_button4, radio_button5]
    for button in enabled_buttons:
        if button['state'] == 'normal':
            sel_rsp_type.set(button['value'])
        break

     # Create a regression model button
    run_reg_button = Button(app, text='Run Regression Model',
                             command=lambda: run_reg_mod(scale_tr_start, scale_tr_end, scale_pred_start,checkbox_var_o,sel_rsp_type,frame_yr))
    run_reg_button.grid(column=2, row=4, padx=50, ipadx=15, ipady=15, sticky=tk.E)

def run_reg_mod(scale_tr_start, scale_tr_end, scale_pred_start,checkbox_var_o,sel_rsp_type,frame_yr):
    first_yr_of_train = int(scale_tr_start.get())
    last_yr_of_train = int(scale_tr_end.get())
    first_yr_pred = int(scale_pred_start.get())
    selected_types_train = []
    # Check if the 'Outlets' checkbox is selected
    if checkbox_var_o.get() == 1:
        selected_types_train.append('Outlets')
    if sel_rsp_type=='None':
        selected_types_train=selected_types_train
    else:
        selected_types_train.append(sel_rsp_type.get())

    selected_types_target = ['Transactions']
    predicted_data_type = [selected_types_target[0] + ' (Modelled)']

    # read file and set cols to be correct
    path_sz = sz_file_name.get("0.0", "end")
    path_sz = str(path_sz[0:len(path_sz) - 1])
    size_0 = pd.read_excel(path_sz)
    #size_0 = pd.read_excel(r'C:\Users\Golov\Downloads\sizes_review_cty_excel.xls')
    new_col = np.array(size_0.head(1))
    size_0.columns = new_col[0]
    size_0 = size_0.drop(0)
    # remove non-res prod cat
    all_cats = size_0['Product']
    excl_cat_list = ['Consumer Foodservice', 'Consumer Foodservice by Type', 'Takeaway',
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
                     'Consumer Foodservice Through Leisure', 'Drive-Through'
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
                     'Convenience Stores', 'Forecourt Retailers', 'Consumer Foodservice Drink Sales',
                     'Consumer Foodservice Food Sales', 'Eat-In', 'Delivery', 'Drive-Through'
                                                                              'Consumer Foodservice Drink Sales',
                     'Travel', 'Inbound Arrivals', 'Arrivals by Country of Origin', 'Arrivals from Asia Pacific',
                     'Arrivals from Afghanistan', 'Arrivals from American Samoa', 'Arrivals from Armenia',
                     'Arrivals from Azerbaijan', 'Arrivals from Bangladesh', 'Arrivals from Bhutan',
                     'Arrivals from Brunei Darussalam', 'Arrivals from Cambodia', 'Arrivals from China',
                     'Arrivals From Fiji', 'Arrivals from French Polynesia', 'Arrivals from Guam',
                     'Arrivals from Hong Kong, China', 'Arrivals from India', 'Arrivals from Indonesia',
                     'Arrivals from Japan', 'Arrivals from Kazakhstan', 'Arrivals from Kiribati',
                     'Arrivals from Kyrgyzstan', 'Arrivals from Laos', 'Arrivals from Macau, China',
                     'Arrivals from Malaysia', 'Arrivals from Maldives', 'Arrivals from Mongolia',
                     'Arrivals from Myanmar', 'Arrivals from Nauru', 'Arrivals from Nepal',
                     'Arrivals from New Caledonia', 'Arrivals from North Korea', 'Arrivals from Pakistan',
                     'Arrivals from Papua New Guinea', 'Arrivals from Philippines', 'Arrivals from Samoa',
                     'Arrivals from Singapore', 'Arrivals from Solomon Islands', 'Arrivals from South Korea',
                     'Arrivals from Sri Lanka', 'Arrivals from Taiwan', 'Arrivals from Tajikistan',
                     'Arrivals from Thailand', 'Arrivals from Tonga', 'Arrivals from Turkmenistan',
                     'Arrivals from Tuvalu', 'Arrivals from Uzbekistan', 'Arrivals from Vanuatu',
                     'Arrivals from Vietnam', 'Arrivals from Australasia', 'Arrivals from Australia',
                     'Arrivals from New Zealand', 'Arrivals from Eastern Europe', 'Arrivals from Albania',
                     'Arrivals from Belarus', 'Arrivals from Bosnia and Herzegovina', 'Arrivals from Bulgaria',
                     'Arrivals from Croatia', 'Arrivals from Czech Republic', 'Arrivals from Estonia',
                     'Arrivals from Georgia', 'Arrivals from Hungary', 'Arrivals from Kosovo', 'Arrivals from Latvia',
                     'Arrivals from Lithuania', 'Arrivals from North Macedonia', 'Arrivals from Moldova',
                     'Arrivals from Montenegro', 'Arrivals from Poland', 'Arrivals from Romania',
                     'Arrivals from Russia', 'Arrivals from Serbia', 'Arrivals from Slovakia', 'Arrivals from Slovenia',
                     'Arrivals from Ukraine', 'Arrivals from Latin America', 'Arrivals from Anguilla',
                     'Arrivals from Antigua and Barbuda', 'Arrivals from Argentina', 'Arrivals from Aruba',
                     'Arrivals from Bahamas', 'Arrivals from Barbados', 'Arrivals from Belize', 'Arrivals from Bermuda',
                     'Arrivals from Bolivia', 'Arrivals from Brazil', 'Arrivals from British Virgin Islands',
                     'Arrivals from Cayman Islands', 'Arrivals from Chile', 'Arrivals from Colombia',
                     'Arrivals from Costa Rica', 'Arrivals from Cuba', 'Arrivals from Curaçao',
                     'Arrivals from Dominica', 'Arrivals from Dominican Republic', 'Arrivals from Ecuador',
                     'Arrivals from El Salvador', 'Arrivals from French Guiana', 'Arrivals from Grenada',
                     'Arrivals from Guadeloupe', 'Arrivals from Guatemala', 'Arrivals from Guyana',
                     'Arrivals from Haiti', 'Arrivals from Honduras', 'Arrivals from Jamaica',
                     'Arrivals from Martinique', 'Arrivals from Mexico', 'Arrivals from Nicaragua',
                     'Arrivals from Panama', 'Arrivals from Paraguay', 'Arrivals from Peru',
                     'Arrivals from Puerto Rico', 'Arrivals from Sint Maarten', 'Arrivals from St Kitts and Nevis',
                     'Arrivals from St Lucia', 'Arrivals from St Vincent and the Grenadines', 'Arrivals from Suriname',
                     'Arrivals from Trinidad and Tobago', 'Arrivals from Uruguay', 'Arrivals from US Virgin Islands',
                     'Arrivals from Venezuela', 'Arrivals from Middle East and Africa', 'Arrivals from Algeria',
                     'Arrivals from Angola', 'Arrivals from Bahrain', 'Arrivals from Benin', 'Arrivals from Botswana',
                     'Arrivals from Burkina Faso', 'Arrivals from Burundi', 'Arrivals from Cameroon',
                     'Arrivals from Cabo Verde', 'Arrivals from Central African Republic', 'Arrivals from Chad',
                     'Arrivals from Comoros', 'Arrivals from Congo, Democratic Republic',
                     'Arrivals from Congo-Brazzaville', 'Arrivals from Djibouti', 'Arrivals from Egypt',
                     'Arrivals from Equatorial Guinea', 'Arrivals from Eritrea', 'Arrivals from Ethiopia',
                     'Arrivals from Gabon', 'Arrivals from Gambia', 'Arrivals from Ghana', 'Arrivals from Guinea',
                     'Arrivals from Guinea-Bissau', 'Arrivals from Iran', 'Arrivals from Iraq', 'Arrivals from Israel',
                     'Arrivals from Jordan', 'Arrivals from Kenya', 'Arrivals from Kuwait', 'Arrivals from Lebanon',
                     'Arrivals from Lesotho', 'Arrivals from Liberia', 'Arrivals from Libya',
                     'Arrivals from Madagascar', 'Arrivals from Malawi', 'Arrivals from Mali',
                     'Arrivals from Mauritania', 'Arrivals from Mauritius', 'Arrivals from Morocco',
                     'Arrivals from Mozambique', 'Arrivals from Namibia', 'Arrivals from Niger',
                     'Arrivals from Nigeria', 'Arrivals from Oman', 'Arrivals from Qatar', 'Arrivals from Réunion',
                     'Arrivals from Rwanda', 'Arrivals from Sao Tomé e Príncipe', 'Arrivals from Saudi Arabia',
                     'Arrivals from Senegal', 'Arrivals from Seychelles', 'Arrivals from Sierra Leone',
                     'Arrivals from Somalia', 'Arrivals from South Africa', 'Arrivals from South Sudan',
                     'Arrivals from Sudan', 'Arrivals from Eswatini', 'Arrivals from Syria', 'Arrivals from Tanzania',
                     'Arrivals from Togo', 'Arrivals from Tunisia', 'Arrivals from Uganda',
                     'Arrivals from United Arab Emirates', 'Arrivals from Yemen', 'Arrivals from Zambia',
                     'Arrivals from Zimbabwe', 'Arrivals from North America', 'Arrivals from Canada',
                     'Arrivals from US', 'Arrivals from Western Europe', 'Arrivals from Andorra',
                     'Arrivals from Austria', 'Arrivals from Belgium', 'Arrivals from Cyprus', 'Arrivals from Denmark',
                     'Arrivals from Finland', 'Arrivals from France', 'Arrivals from Germany',
                     'Arrivals from Gibraltar', 'Arrivals from Greece', 'Arrivals from Iceland',
                     'Arrivals from Ireland', 'Arrivals from Italy', 'Arrivals from Liechtenstein',
                     'Arrivals from Luxembourg', 'Arrivals from Malta', 'Arrivals from Monaco',
                     'Arrivals from Netherlands', 'Arrivals from Norway', 'Arrivals from Portugal',
                     'Arrivals from Spain', 'Arrivals from Sweden', 'Arrivals from Switzerland', 'Arrivals from Turkey',
                     'Arrivals from United Kingdom', 'Arrivals from Other Countries', 'Inbound Length of Stay',
                     'Inbound Tourism by Mode of Transport', 'Air Arrivals', 'Business Air Arrivals',
                     'Leisure Air Arrivals', 'Land Arrivals', 'Business Land Arrivals', 'Leisure Land Arrivals',
                     'Rail Arrivals', 'Business Rail Arrivals', 'Leisure Rail Arrivals', 'Water Arrivals',
                     'Business Water Arrivals', 'Leisure Water Arrivals', 'Inbound Tourism by Purpose of Visit',
                     'Business Inbound', 'MICE Inbound', 'Other Business Inbound', 'Leisure Inbound', 'VFR Inbound',
                     'Other Leisure Inbound', 'Inbound Tourism Spending', 'Inbound Business Tourism Spending',
                     'Inbound Leisure Tourism Spending', 'Inbound Spending on Lodging',
                     'Inbound Spending Excluding Lodging', 'Inbound Spending on Activities', 'Inbound Spending on Food',
                     'Inbound Spending on Shopping', 'Inbound Spending on Retail Shopping',
                     'Inbound Spending on Duty-free Shopping', 'Inbound Spending on Travel Modes',
                     'Inbound Spending on Other', 'Inbound City Arrivals', 'Vienna', 'Salzburg', 'Sölden',
                     'Saalbach-Hinterglemm', 'Mayrhofen', 'Ischgl', 'Sankt Anton am Arlberg', 'Lech', 'Innsbruck',
                     'Linz', 'Domestic Tourism', 'Domestic Tourism By Destination', 'Steiermark', 'Kärnten',
                     'Oberösterreich', 'Niederösterreich', 'Tirol', 'Burgenland', 'Wien', 'Vorarlberg',
                     'Domestic Tourism Destination Subtype 10', 'Domestic Tourism Destination Leading Cities', 'Graz',
                     'Domestic Tourism Destination Leading City Subtype 6',
                     'Domestic Tourism Destination Leading City Subtype 7',
                     'Domestic Tourism Destination Leading City Subtype 8',
                     'Domestic Tourism Destination Leading City Subtype 9',
                     'Domestic Tourism Destination Leading City Subtype 10', 'Domestic Tourism by Mode of Transport',
                     'Domestic Tourism by Air', 'Domestic Business Tourism By Air', 'Domestic Leisure Tourism By Air',
                     'Domestic Tourism by Land', 'Domestic Business Tourism By Land',
                     'Domestic Leisure Tourism By Land', 'Domestic Tourism by Rail',
                     'Domestic Business Tourism By Rail', 'Domestic Leisure Tourism By Rail',
                     'Domestic Tourism by Water', 'Domestic Business Tourism By Water',
                     'Domestic Leisure Tourism By Water', 'Domestic Tourism by Purpose of Visit', 'Domestic Business',
                     'Domestic Other Business', 'Domestic Leisure', 'Domestic Other Leisure', 'Domestic Spending',
                     'Domestic Spending on Lodging', 'Domestic Spending on Shopping', 'Domestic Spending on Other',
                     'Outbound Departures', 'Outbound Departures Source Markets', 'Outbound Departures to Asia Pacific',
                     'Outbound Departures to Afghanistan', 'Outbound Departures to American Samoa',
                     'Outbound Departures to Armenia', 'Outbound Departures to Azerbaijan',
                     'Outbound Departures to Bangladesh', 'Outbound Departures to Bhutan',
                     'Outbound Departures to Brunei Darussalam', 'Outbound Departures to Cambodia',
                     'Outbound Departures to China', 'Outbound Departures to Fiji',
                     'Outbound Departures to French Polynesia', 'Outbound Departures to Guam',
                     'Outbound Departures to Hong Kong, China', 'Outbound Departures to India',
                     'Outbound Departures to Indonesia', 'Outbound Departures to Japan',
                     'Outbound Departures to Kazakhstan', 'Outbound Departures to Kiribati',
                     'Outbound Departures to Kyrgyzstan', 'Outbound Departures to Laos',
                     'Outbound Departures to Macau, China', 'Outbound Departures to Malaysia',
                     'Outbound Departures to Maldives', 'Outbound Departures to Mongolia',
                     'Outbound Departures to Myanmar', 'Outbound Departures to Nauru', 'Outbound Departures to Nepal',
                     'Outbound Departures to New Caledonia', 'Outbound Departures to North Korea',
                     'Outbound Departures to Pakistan', 'Outbound Departures to Papua New Guinea',
                     'Outbound Departures to Philippines', 'Outbound Departures to Samoa',
                     'Outbound Departures to Singapore', 'Outbound Departures to Solomon Islands',
                     'Outbound Departures to South Korea', 'Outbound Departures to Sri Lanka',
                     'Outbound Departures to Taiwan', 'Outbound Departures to Tajikistan',
                     'Outbound Departures to Thailand', 'Outbound Departures to Tonga',
                     'Outbound Departures to Turkmenistan', 'Outbound Departures to Tuvalu',
                     'Outbound Departures to Uzbekistan', 'Outbound Departures to Vanuatu',
                     'Outbound Departures to Vietnam', 'Outbound Departures to Australasia',
                     'Outbound Departures to Australia', 'Outbound Departures to New Zealand',
                     'Outbound Departures to Eastern Europe', 'Outbound Departures to Albania',
                     'Outbound Departures to Belarus', 'Outbound Departures to Bosnia and Herzegovina',
                     'Outbound Departures to Bulgaria', 'Outbound Departures to Croatia',
                     'Outbound Departures to Czech Republic', 'Outbound Departures to Estonia',
                     'Outbound Departures to Georgia', 'Outbound Departures to Hungary',
                     'Outbound Departures to Kosovo', 'Outbound Departures to Latvia',
                     'Outbound Departures to Lithuania', 'Outbound Departures to North Macedonia',
                     'Outbound Departures to Moldova', 'Outbound Departures to Montenegro',
                     'Outbound Departures to Poland', 'Outbound Departures to Romania', 'Outbound Departures to Russia',
                     'Outbound Departures to Serbia', 'Outbound Departures to Slovakia',
                     'Outbound Departures to Slovenia', 'Outbound Departures to Ukraine',
                     'Outbound Departures to Latin America', 'Outbound Departures to Anguilla',
                     'Outbound Departures to Antigua and Barbuda', 'Outbound Departures to Argentina',
                     'Outbound Departures to Aruba', 'Outbound Departures to Bahamas',
                     'Outbound Departures to Barbados', 'Outbound Departures to Belize',
                     'Outbound Departures to Bermuda', 'Outbound Departures to Bolivia',
                     'Outbound Departures to Brazil', 'Outbound Departures to British Virgin Islands',
                     'Outbound Departures to Cayman Islands', 'Outbound Departures to Chile',
                     'Outbound Departures to Colombia', 'Outbound Departures to Costa Rica',
                     'Outbound Departures to Cuba', 'Outbound Departures to Curaçao', 'Outbound Departures to Dominica',
                     'Outbound Departures to Dominican Republic', 'Outbound Departures to Ecuador',
                     'Outbound Departures to El Salvador', 'Outbound Departures to French Guiana',
                     'Outbound Departures to Grenada', 'Outbound Departures to Guadeloupe',
                     'Outbound Departures to Guatemala', 'Outbound Departures to Guyana',
                     'Outbound Departures to Haiti', 'Outbound Departures to Honduras',
                     'Outbound Departures to Jamaica', 'Outbound Departures to Martinique',
                     'Outbound Departures to Mexico', 'Outbound Departures to Nicaragua',
                     'Outbound Departures to Panama', 'Outbound Departures to Paraguay', 'Outbound Departures to Peru',
                     'Outbound Departures to Puerto Rico', 'Outbound Departures to Sint Maarten',
                     'Outbound Departures to St Kitts and Nevis', 'Outbound Departures to St Lucia',
                     'Outbound Departures to St Vincent and the Grenadines', 'Outbound Departures to Suriname',
                     'Outbound Departures to Trinidad and Tobago', 'Outbound Departures to Uruguay',
                     'Outbound Departures to US Virgin Islands', 'Outbound Departures to Venezuela',
                     'Outbound Departures to Middle East and Africa', 'Outbound Departures to Algeria',
                     'Outbound Departures to Angola', 'Outbound Departures to Bahrain', 'Outbound Departures to Benin',
                     'Outbound Departures to Botswana', 'Outbound Departures to Burkina Faso',
                     'Outbound Departures to Burundi', 'Outbound Departures to Cameroon',
                     'Outbound Departures to Cabo Verde', 'Outbound Departures to Central African Republic',
                     'Outbound Departures to Chad', 'Outbound Departures to Comoros',
                     'Outbound Departures to Congo, Democratic Republic', 'Outbound Departures to Congo-Brazzaville',
                     'Outbound Departures to Djibouti', 'Outbound Departures to Egypt',
                     'Outbound Departures to Equatorial Guinea', 'Outbound Departures to Eritrea',
                     'Outbound Departures to Ethiopia', 'Outbound Departures to Gabon', 'Outbound Departures to Gambia',
                     'Outbound Departures to Ghana', 'Outbound Departures to Guinea',
                     'Outbound Departures to Guinea-Bissau', 'Outbound Departures to Iran',
                     'Outbound Departures to Iraq', 'Outbound Departures to Israel', 'Outbound Departures to Jordan',
                     'Outbound Departures to Kenya', 'Outbound Departures to Kuwait', 'Outbound Departures to Lebanon',
                     'Outbound Departures to Lesotho', 'Outbound Departures to Liberia', 'Outbound Departures to Libya',
                     'Outbound Departures to Madagascar', 'Outbound Departures to Malawi',
                     'Outbound Departures to Mali', 'Outbound Departures to Mauritania',
                     'Outbound Departures to Mauritius', 'Outbound Departures to Morocco',
                     'Outbound Departures to Mozambique', 'Outbound Departures to Namibia',
                     'Outbound Departures to Niger', 'Outbound Departures to Nigeria', 'Outbound Departures to Oman',
                     'Outbound Departures to Qatar', 'Outbound Departures to Réunion', 'Outbound Departures to Rwanda',
                     'Outbound Departures to Sao Tomé e Príncipe', 'Outbound Departures to Saudi Arabia',
                     'Outbound Departures to Senegal', 'Outbound Departures to Seychelles',
                     'Outbound Departures to Sierra Leone', 'Outbound Departures to Somalia',
                     'Outbound Departures to South Africa', 'Outbound Departures to South Sudan',
                     'Outbound Departures to Sudan', 'Outbound Departures to Eswatini', 'Outbound Departures to Syria',
                     'Outbound Departures to Tanzania', 'Outbound Departures to Togo', 'Outbound Departures to Tunisia',
                     'Outbound Departures to Uganda', 'Outbound Departures to United Arab Emirates',
                     'Outbound Departures to Yemen', 'Outbound Departures to Zambia', 'Outbound Departures to Zimbabwe',
                     'Outbound Departures to North America', 'Outbound Departures to Canada',
                     'Outbound Departures to US', 'Outbound Departures to Western Europe',
                     'Outbound Departures to Andorra', 'Outbound Departures to Austria',
                     'Outbound Departures to Belgium', 'Outbound Departures to Cyprus',
                     'Outbound Departures to Denmark', 'Outbound Departures to Finland',
                     'Outbound Departures to France', 'Outbound Departures to Germany',
                     'Outbound Departures to Gibraltar', 'Outbound Departures to Greece',
                     'Outbound Departures to Iceland', 'Outbound Departures to Ireland', 'Outbound Departures to Italy',
                     'Outbound Departures to Liechtenstein', 'Outbound Departures to Luxembourg',
                     'Outbound Departures to Malta', 'Outbound Departures to Monaco',
                     'Outbound Departures to Netherlands', 'Outbound Departures to Norway',
                     'Outbound Departures to Portugal', 'Outbound Departures to Spain', 'Outbound Departures to Sweden',
                     'Outbound Departures to Switzerland', 'Outbound Departures to Turkey',
                     'Outbound Departures to United Kingdom', 'Outbound Departures to Other Destinations',
                     'Outbound Length of Stay', 'Outbound Tourism by Mode of Transport', 'Air Outbound',
                     'Business Air Outbound', 'Leisure Air Outbound', 'Land Outbound', 'Business Land Outbound',
                     'Leisure Land Outbound', 'Rail Outbound', 'Business Rail Outbound', 'Leisure Rail Outbound',
                     'Water Outbound', 'Business Water Outbound', 'Leisure Water Outbound',
                     'Outbound Tourism by Purpose of Visit', 'Business Outbound', 'Leisure Outbound',
                     'Outbound Tourism Spending', 'Outbound Business Spending', 'Outbound Leisure Spending',
                     'Outbound Spending on Lodging', 'Outbound Spending on Activities', 'Outbound Spending on Food',
                     'Outbound Spending on Shopping', 'Outbound Spending on Retail Shopping',
                     'Outbound Spending on Duty-free Shopping', 'Outbound Spending on Travel Modes',
                     'Outbound Spending on Other', 'Travel Modes', 'Airlines', 'Airlines by Category',
                     'Scheduled Airlines', 'Ancillary Revenue', 'International Airlines', 'Airlines by Channel',
                     'Airlines Online', 'Airlines Online via Direct', 'Airlines Online via Intermediaries',
                     'Air through Package Holidays', 'Air Online Sales less Air through Package Holidays',
                     'Airlines Offline', 'Airlines Offline via Direct', 'Airlines Offline via Intermediaries',
                     'Surface Travel Modes', 'Surface Travel Modes by Category', 'Surface Travel Modes by Channel',
                     'Surface Travel Modes Online', 'Surface Travel Modes Online via Direct',
                     'Surface Travel Modes Online via Intermediaries', 'Surface Travel Modes Offline',
                     'Surface Travel Modes Offline via Direct', 'Surface Travel Modes Offline via Intermediaries',
                     'Lodging (Destination)', 'Hotels by Category', 'Hotels by Channel', 'Hotels Online',
                     'Hotels Online via Direct', 'Hotels Online via Intermediaries', 'Hotels Offline',
                     'Hotels Offline via Direct', 'Hotels Offline via Intermediaries', 'Short-Term Rentals Online',
                     'Short-Term rentals Online via Direct', 'Short-term Rentals Online via Intermediaries',
                     'Short-Term Rentals Offline', 'Short-term Rentals Offline via Direct',
                     'Short-term Rentals Offline via Intermediaries', 'Other Lodging', 'Other Lodging by Category',
                     'Other Lodging by Channel', 'Other Lodging Online', 'Other Lodging Online Direct',
                     'Other Lodging Online Intermediaries', 'Other Lodging Offline', 'Other Lodging Offline via Direct',
                     'Other Lodging Offline via Intermediaries', 'Lodging (Destination) by Channel',
                     'Lodging (Destination) Online', 'Lodging (Destination) Online via Direct',
                     'Lodging (Destination) Online via Intermediaries', 'Lodging (Destination) Offline',
                     'Lodging (Destination) Offline via Direct', 'Lodging (Destination) Offline via Intermediaries',
                     'In-Destination Spending', 'Attractions', 'Experiences', 'Shopping', 'Retail Shopping',
                     'Duty-Free Shopping', 'Wellness', 'Other In-Destination Spending',
                     'In-Destination Spending by Channel', 'In-Destination Spending Online',
                     'In-Destination Spending Online Direct', 'In-Destination Spending Online Intermediaries',
                     'In-Destination Spending Offline', 'In-Destination Spending Offline Direct',
                     'In-Destination Spending Offline Intermediaries', 'Booking', 'Booking Offline', 'Booking Online',
                     'Mobile Travel', 'Leisure Travel', 'Leisure Air Travel Online',
                     'Leisure Air Travel Online via Direct', 'Leisure Air Travel Online via Intermediaries',
                     'Leisure Air Travel Offline', 'Leisure Air Travel Offline via Direct',
                     'Leisure Air Travel Offline via Intermediaries', 'Leisure Car Rental Online',
                     'Leisure Car Rental Online via Direct', 'Leisure Car Rental Online via Intermediaries',
                     'Leisure Car Rental Offline', 'Leisure Car Rental Offline via Direct',
                     'Leisure Car Rental Offline via Intermediaries', 'Leisure Cruise Online',
                     'Leisure Cruise Online via Direct', 'Leisure Cruise Online via Intermediaries',
                     'Leisure Cruise Offline', 'Leisure Cruise Offline via Direct',
                     'Leisure Cruise Offline via Intermediaries', 'Leisure Experiences and Attractions Online',
                     'Leisure Experiences and Attractions Online via Direct',
                     'Leisure Experiences and Attractions Online via Intermediaries',
                     'Leisure Experiences and Attractions Offline',
                     'Leisure Experiences and Attractions Offline via Direct',
                     'Leisure Experiences and Attractions Offline via Intermediaries',
                     'Leisure Lodging (Source) Online', 'Leisure Lodging (Source) Online via Direct',
                     'Leisure Lodging (Source) Online via Intermediaries', 'Leisure Lodging (Source) Offline',
                     'Leisure Lodging (Source) Offline via Direct',
                     'Leisure Lodging (Source) Offline via Intermediaries', 'Leisure Packages Online',
                     'Leisure Packages Online via Intermediaries', 'Leisure Packages Offline',
                     'Leisure Packages Offline via Intermediaries', 'Leisure Surface Travel Online',
                     'Leisure Surface Travel Online via Direct', 'Leisure Surface Travel Online via Intermediaries',
                     'Leisure Surface Travel Offline', 'Leisure Surface Travel Offline via Direct',
                     'Leisure Surface Travel Offline via Intermediaries', 'Leisure Other Travel Products Online',
                     'Leisure Other Travel Products Online via Direct',
                     'Leisure Other Travel Products Online via Intermediaries', 'Leisure Other Travel Products Offline',
                     'Leisure Other Travel Products Offline via Direct',
                     'Leisure Other Travel Products Offline via Intermediaries', 'Business Travel',
                     'Business Air Travel Online', 'Business Air Travel Online via Direct',
                     'Business Air Travel Online via Intermediaries', 'Business Air Travel Offline',
                     'Business Air Travel Offline via Direct', 'Business Air Travel Offline via Intermediaries',
                     'Business Car Rental Online', 'Business Car Rental Online via Direct',
                     'Business Car Rental Online via Intermediaries', 'Business Car Rental Offline',
                     'Business Car Rental Offline via Direct', 'Business Car Rental Offline via Intermediaries',
                     'Business Lodging Online', 'Business Lodging Online via Direct',
                     'Business Lodging Online via Intermediaries', 'Business Lodging Offline',
                     'Business Lodging Offline via Direct', 'Business Lodging Offline via Intermediaries',
                     'Business Other Online', 'Business Other Online via Direct',
                     'Business Other Online via Intermediaries', 'Business Other Offline',
                     'Business Other Offline via Direct', 'Business Other Offline via Intermediaries',
                     'Travel Intermediaries', 'Travel Intermediaries Online', 'Travel Intermediaries Offline',
                     'Direct Suppliers', 'Direct Suppliers Online', 'Direct Suppliers Offline', 'Leading Airports',
                     'Vienna International Airport', 'Graz Airport', 'Innsbruck Airport',
                     'Kärntern Airport (Klagenfurt)', 'Blue Danube Airport Linz', 'Salzburg Airport',
                     'Leading Airports Subtype 7', 'Leading Airports Subtype 8', 'Leading Airports Subtype 9',
                     'Leading Airports Subtype 10',
                     'Sum of Cards by Function', 'ATM Cards', 'Commercial Charge Cards', 'Personal Charge Cards',
                     'Commercial Credit Cards', 'Personal Credit Cards', 'Commercial Debit Cards',
                     'Personal Debit Cards', 'Pre-Paid Cards', 'Closed Loop Pre-Paid Cards', 'Open Loop Pre-Paid Cards',
                     'Transactions', 'Total Card Transactions', 'Card Payment Transactions', 'Charge Card Transactions',
                     'Credit Card Transactions', 'Debit Card Transactions', 'Pre-Paid Card Transactions',
                     'Commercial Payment Transactions', 'Commercial Card Payment Transactions',
                     'Commercial Charge Card Transactions (duplicate)',
                     'Commercial Credit Card Transactions (duplicate)',
                     'Commercial Debit Card Transactions (duplicate)',
                     'Commercial Electronic Direct/ACH Transactions (duplicate)',
                     'Commercial Paper Payment Transactions (duplicate)', 'Commercial Cash Transactions (duplicate)',
                     'Commercial Other Paper Transactions (duplicate)', 'Personal Payment Transactions',
                     'Personal Card Payment Transactions', 'Personal Charge Card Transactions (duplicate 2)',
                     'Personal Credit Card Transactions (duplicate 2)', 'Personal Debit Card Transactions (duplicate)',
                     'Pre-Paid Card Transactions (duplicate 2)', 'Closed Loop Pre-Paid Card Transactions (duplicate 2)',
                     'Open Loop Pre-Paid Card Transactions (duplicate 2)', 'Store Card Transactions (duplicate 2)',
                     'Personal Electronic Direct/ACH Transactions (duplicate)',
                     'Personal Paper Payment Transactions (duplicate)', 'Personal Cash Transactions (duplicate)',
                     'Personal Other Paper Transactions (duplicate)', 'Total Non-Card Transactions',
                     'Electronic Direct/ACH Transactions', 'Paper Payment Transactions',
                     'Commercial Paper Payment Transactions', 'Personal Paper Payment Transactions', 'Consumer Lending',
                     'Consumer Credit', 'Average Personal Credit Card Balance', 'Non-Card Lending', 'Auto Lending',
                     'Durables Lending', 'Education Lending', 'Home Lending', 'Other Personal Lending',
                     'Mortgages/Housing', 'Mobile Payments', 'Mobile E-Commerce Payments', 'Data Checks',
                     'Alternative Financial Service Providers', 'Payday', 'Remote Payments',
                     'Sum (Card Holder not Present)', 'Charge Card Transactions (Card Holder not Present)',
                     'Credit Card Transactions (Card Holder not Present)',
                     'Debit Card Transactions (Card Holder not Present)',
                     'Open Loop Pre-Paid Card Transactions (Card Holder not Present)']
    filtered_cats = [cat for cat in all_cats if cat not in excl_cat_list]
    size = size_0[size_0['Product'].isin(filtered_cats)].copy()
    size = size.drop(['Sub-project', 'Region', 'Sector', 'Unit'], axis=1)
    size = size.fillna(0)
    last_year = size.columns[-1]
    first_year = last_year - 20
    init_yr_train = int(first_yr_of_train - int(first_year))
    fin_yr_train = int(last_yr_of_train - int(first_year))
    init_yr_fcst = int(first_yr_pred - int(first_year)-1)
    df = size
    df_o = df
    dfc = df_o.columns
    all_countries = df['Country'].unique()
    all_prod = df['Product'].unique()
    lrt_output = pd.DataFrame()
    # year error messages
    text_box_fyot_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
    text_box_fyot_mssg.grid(column=2, row=3)
    text_box_fyot_mssg.delete(0, "end")
    text_box_lyot_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
    text_box_lyot_mssg.grid(column=2, row=4)
    text_box_lyot_mssg.delete(0, "end")
    text_box_fyofcst_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
    text_box_fyofcst_mssg.grid(column=2, row=5)
    text_box_fyofcst_mssg.delete(0, "end")
    if first_yr_of_train>=last_yr_of_train:
        fyot_mssg = 'This has to be smaller than both of the below'
        lyot_mssg = 'This has to be larger than the above'
        text_box_fyot_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
        text_box_fyot_mssg.grid(column=2, row=3)
        text_box_fyot_mssg.delete(0, "end")
        text_box_fyot_mssg.insert(0, fyot_mssg)

        text_box_lyot_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
        text_box_lyot_mssg.grid(column=2, row=4)
        text_box_lyot_mssg.delete(0, "end")
        text_box_lyot_mssg.insert(0, lyot_mssg)
    elif first_yr_of_train>=first_yr_pred:
        fyot_mssg = 'This has to be smaller than both of the below'
        fyofcst_mssg = 'This has to be larger than both of the above'
        text_box_fyot_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
        text_box_fyot_mssg.grid(column=2, row=3)
        text_box_fyot_mssg.delete(0, "end")
        text_box_fyot_mssg.insert(0, fyot_mssg)

        text_box_fyofcst_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
        text_box_fyofcst_mssg.grid(column=2, row=5)
        text_box_fyofcst_mssg.delete(0, "end")
        text_box_fyofcst_mssg.insert(0, fyofcst_mssg)
    elif last_yr_of_train>=first_yr_pred:
        lyot_mssg = 'This has to be smaller than the below'
        text_box_lyot_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
        text_box_lyot_mssg.grid(column=2, row=4)
        text_box_lyot_mssg.delete(0, "end")
        text_box_lyot_mssg.insert(0, lyot_mssg)
        fyofcst_mssg = 'This has to be larger than both of the above'
        text_box_fyofcst_mssg = tk.Entry(frame_yr, width=50, relief=tk.FLAT, bg="SystemButtonFace")
        text_box_fyofcst_mssg.grid(column=2, row=5)
        text_box_fyofcst_mssg.delete(0, "end")
        text_box_fyofcst_mssg.insert(0, fyofcst_mssg)

    for country, prod in [(country, prod) for country in all_countries for prod in all_prod]:
        try:
            df_cou_prod = df[(df['Country'] == country) & (df['Product'] == prod)].drop(['Country', 'Product'], axis=1)
            df_cou_prod_train = df_cou_prod[df_cou_prod['Data type'].isin(selected_types_train)]
            df_cou_prod_target = df_cou_prod[df_cou_prod['Data type'].isin(selected_types_target)]
            df_cou_prod_train.set_index('Data type', inplace=True)
            df_cou_prod_target.set_index('Data type', inplace=True)
            df_cou_prod_target = df_cou_prod_target.transpose()

            df_cou_prod_train = df_cou_prod_train.transpose()
            df_cou_prod_train_np = df_cou_prod_train.to_numpy()
            df_cou_prod_target_np = df_cou_prod_target.to_numpy()
            x_train = df_cou_prod_train_np[init_yr_train:fin_yr_train, :]
            y_train = df_cou_prod_target_np[init_yr_train:fin_yr_train, :]
            x_fcst = df_cou_prod_train_np[init_yr_fcst:, :]
            sklearn_model = LinearRegression().fit(x_train, y_train)
            sklearn_y_predictions = sklearn_model.predict(x_fcst)
            res = np.append(country, predicted_data_type)
            res = np.append(res, prod)
            t=int(first_yr_pred)
            t_pred_yrs = [t]
            while t<last_year:
                t+=1
                t_pred_yrs.append(t)

            res = np.append(res, sklearn_y_predictions)
            res = pd.DataFrame(res)
            res = res.transpose()

            lrt_output = pd.concat([lrt_output, res])
        except Exception as e:
            continue
    lrt_output_columns = ['Country', 'Data type', 'Product'] + t_pred_yrs
    lrt_output.columns=lrt_output_columns
    #writer=pd.ExcelWriter(lrt_output, engine='xlsxwriter',engine_kwargs={'options': {'strings_to_numbers': True}})
    #writer = pd.ExcelWriter(lrt_output, engine='xlsxwriter')
    #lrt_output.to_excel(writer,'Regression Model Output.xlsx', index=False)
    # Create a Pandas Excel writer using the xlsxwriter engine
    writer = pd.ExcelWriter('Regression Model Output.xlsx', engine='xlsxwriter')
    lrt_output.to_excel(writer, 'Regression Model Output', index=False)

    # Open the Excel workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Regression Model Output']

    # Create a format to display numbers as numbers (not text) in the Excel sheet
    num_format = workbook.add_format({'num_format': '#,##0.00'})

    # Get the dimensions of the DataFrame (number of rows and columns)
    num_rows, num_cols = lrt_output.shape

    # Loop through each column and apply the number format to the entire column, starting from row 1
    #for col_num in range(num_cols):
     #   column_letter = chr(65 + col_num)  # Convert column number to Excel column letter (A, B, C, ...)
      #  column_range = f'{column_letter}2:{column_letter}{num_rows + 1}'  # +1 to include header row
       # worksheet.set_column(column_range, None, num_format)

    # Save the Excel file
    writer = pd.ExcelWriter('Regression Model Output.xlsx', engine='xlsxwriter')

    # Save the DataFrame to Excel with the 'float_format' parameter
    lrt_output.to_excel(writer, 'Regression Model Output', index=False, float_format='%.2f')

    writer.close()

# Button for closing
exit_button = Button(app, text="Exit", width=30, command=app.destroy)
exit_button.grid(column=1, row=20,padx=10,pady=20,columnspan=2)

app.mainloop()
