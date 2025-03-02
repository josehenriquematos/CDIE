''' 
How to use "CDIE.py":

It's pretty easy to use this code. First you need to set the reference date that you want to view the Copom market expectations to the "date"
variable. After that, you should check if the Selic over rate is correct in the "selic_over_rate" variable (Note: the effective Selic rate is 
approximately 10 bps below Selic target rate which is set by Banco Central do Brasil. For example, if Selic rate announced by BCB is 13.25%,
the effective Selic rate or Selic over rate will be 13.15%). 

'''
import requests
import pandas as pd
from pandas.tseries.offsets import BMonthBegin
from workalendar.america import Brazil
from pandas.tseries.offsets import BDay
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt

# DI curve's date:
date = "28/02/2025"

# B3 data scrap:
url = f"https://www2.bmf.com.br/pages/portal/bmfbovespa/lumis/lum-ajustes-do-pregao-ptBR.asp?dData1={date}"
DI_data = pd.read_html(requests.get(url).text)
table = DI_data[0]

# Extracting only the DI's adjustment rows:
start = table[table['Mercadoria'] == "DI1 - DI de 1 dia"].index[0]
end = table[table['Mercadoria'] == "DOL - Dólar comercial"].index[0]
table.loc[start:end, 'Mercadoria'] = table.loc[start:end, 'Mercadoria'].fillna("DI1 - DI de 1 dia")
table_2 = table[table['Mercadoria'] == "DI1 - DI de 1 dia"]
table_3 = table_2.drop(columns = ['Mercadoria', 'Preço de ajuste anterior', 'Variação', 'Valor do ajuste por contrato (R$)'])

# Scraping saved holidays:
with open("feriados_jhenriquematos.txt", "r") as archive:
    holidays_list = [row.strip() for row in archive]

# Convertion of holidays list to datetime type
converted_holidays = [datetime.strptime(date, "%Y-%m-%d %H:%M:%S") for date in holidays_list]

# Defining is_working_day() function to remove holidays in working days. Reason to create this is to help me get
# the first working day of each DI settlement month:
def is_working_day(date, holidays=None):
    if isinstance(date, str):
        date = datetime.strptime(date, "%Y-%m-%d")
    
    # Verify if it's weekend
    if date.weekday() >= 5:  # 5 = Saturday, 6 = Sunday
        return False
    
    # Verify if it's holiday given a holiday list
    if holidays and date in holidays:
        return False
    
    return True
    
# Dataframes with DI contract dates
def DI_dates(value):
    # Mapping the first letter to month's number
    months_aux = {
    'F': 1,
    'G': 2,
    'H': 3, 
    'J': 4,
    'K': 5, 
    'M': 6, 
    'N': 7, 
    'Q': 8, 
    'U': 9, 
    'V': 10, 
    'X': 11, 
    'Z': 12
    }

    years_aux = {
    '23': 2023, 
    '24': 2024, 
    '25': 2025, 
    '26': 2026, 
    '27': 2027, 
    '28': 2028, 
    '29': 2029, 
    '30': 2030, 
    '31': 2031, 
    '32': 2032, 
    '33': 2033, 
    '34': 2034, 
    '35': 2035, 
    '36': 2036, 
    '37': 2037, 
    '38': 2038, 
    '39': 2039, 
    '40': 2040, 
    '41': 2041, 
    '42': 2042, 
    '43': 2043, 
    '44': 2044, 
    '45': 2045
    }

    first_letter = value[0].upper()
    last_numbers = value[1:]

    # Getting corresponding month and year
    month = months_aux.get(first_letter)
    year = years_aux.get(last_numbers)

    if month and year:
        # Calculate the first working day of the month
        first_working_day = pd.Timestamp(f'{year}-{month:02d}-01')
        # Verify if the initial date is working day and isn't a holiday
        while not is_working_day(first_working_day, holidays=converted_holidays):
            first_working_day += BDay(1)  # Go to the next working day
        return first_working_day.date()
    else:
        return None  # Return None if something is wrong

table_3 = table_3.rename(columns={"Vencimento": "Events", "Preço de ajuste Atual": "Adjustment Price"})
table_3["Dates"] = table_3['Events'].apply(DI_dates)

# New order: Dates first, follow by other columns
highlighted_column = "Dates"
other_columns = [col for col in table_3.columns if col != highlighted_column]
new_order = [highlighted_column] + other_columns
table_3 = table_3[new_order]

#Defining Copom dates:
copom_dates = pd.DataFrame({
    "Dates": pd.to_datetime([
        "29-01-2025",
        "19-03-2025",
        "07-05-2025",
        "18-06-2025",
        "30-07-2025",
        "17-09-2025",
        "05-11-2025",
        "10-12-2025"
    ], dayfirst=True)
})

# Putting Copom dates in Dataframe
table_4 = pd.concat([table_3, copom_dates], ignore_index=True)
table_4["Dates"] = pd.to_datetime(table_4["Dates"], dayfirst=True)
table_5 = table_4.fillna({"Events": "Copom"})

# Calculating NDU:
fix_date = pd.to_datetime(date, dayfirst=True)

# Defining get_working_days_delta and calculate_working_days that will effectivelly calculate the NDU:
def get_working_days_delta(start_date, end_date, holidays=None):
    start_date = pd.to_datetime(start_date)
    end_date = pd.to_datetime(end_date)
    all_dates = pd.date_range(start=start_date, end=end_date)
    working_days = [date for date in all_dates if date.weekday() < 5]
    if holidays:
        holidays = pd.to_datetime(holidays)
        working_days = [day for day in working_days if day not in holidays]
    return len(working_days)

def calculate_working_days(fix_date, reference_date):
    # Convert dates to the right format
    fix_date = pd.to_datetime(fix_date, dayfirst=True)
    reference_date = pd.to_datetime(reference_date, dayfirst=True)
    # List of working days between two dates
    working_days = get_working_days_delta(fix_date, reference_date, holidays=converted_holidays)-1
    return working_days

# Apply the function in DataFrame
table_5['NDU'] = table_5['Dates'].apply(lambda reference_date: calculate_working_days(fix_date, reference_date))

# Remove rows where the NDU column is negative. Reason: it's possible that you may want to check Copom market expectations in reference 
# dates that some Copom meetings have already occured.
removed_rows = table_5.index[table_5["NDU"].lt(0)].tolist()
table_5 = table_5.drop(removed_rows)
table_5 = table_5.sort_values(by="Dates").reset_index(drop=True)

# Defining closement rates based on closement adjusment:
table_5['Adjustment Price'] = table_5['Adjustment Price'].str.replace('.', '', regex=True)  # Remove points
table_5['Adjustment Price'] = table_5['Adjustment Price'].str.replace(',', '.', regex=True)  # Change commas to points
table_5['Adjustment Price'] = pd.to_numeric(table_5['Adjustment Price'])
table_5["Closement rates"] = table_5.apply(lambda row: (pow((100000/row["Adjustment Price"]),(252/row["NDU"])) - 1)*100, axis = 1)
table_5 = table_5.drop(columns=["Adjustment Price"])

# Defining Selic Over Rate and it's insertion in curve:
selic_over_rate = 13.15
table_5.loc[len(table_5)] = [fix_date, "Selic Over Today", 1, selic_over_rate]
table_5 = table_5.reset_index(drop=False)
table_5 = table_5.drop(columns=["index"])
table_5.sort_values(by="Dates", ascending = True, inplace = True)
table_5 = table_5.reset_index(drop=False)
table_5 = table_5.drop(columns=["index"])

# Calculating a interpolated rate on Copom days:
table_5["Closement rates"] = table_5["Closement rates"].fillna(0)
interpolation_indexes = table_5.index[table_5["Closement rates"].eq(0)].tolist()

# Function to calculate interpolated rates on Copom reunion days. The interpolation used was Flat Forward 252
def interpolated_rates(dataframe, indexes, column):
    for interpolated_indexes in indexes:
        rate = (pow(pow((1 + table_5.loc[interpolated_indexes - 1, "Closement rates"]/100),(table_5.loc[interpolated_indexes - 1, "NDU"]/252))*pow(pow((1+table_5.loc[interpolated_indexes + 1, "Closement rates"]/100),(table_5.loc[interpolated_indexes + 1, "NDU"]/252))/pow((1+table_5.loc[interpolated_indexes - 1, "Closement rates"]/100),(table_5.loc[interpolated_indexes - 1, "NDU"]/252)),((table_5.loc[interpolated_indexes, "NDU"]-table_5.loc[interpolated_indexes - 1, "NDU"])/(table_5.loc[interpolated_indexes + 1, "NDU"]-table_5.loc[interpolated_indexes - 1, "NDU"]))),(252/table_5.loc[interpolated_indexes, "NDU"]))-1)*100
        dataframe.loc[interpolated_indexes, column] = rate
    return dataframe

table_5 = interpolated_rates(table_5, interpolation_indexes, "Closement rates")
table_5 = table_5.sort_values(by="Dates").reset_index(drop=True)

# Removing the DI contract with the lowest maturity one day before its settlement. Reason: within two days before the maturity the contract lose
# its usage in the market as a hedge or speculation, for example. Additionally, it creates a wrong forward rate calculation in this code
removed_DI_contract = table_5.index[table_5["NDU"].eq(1)].tolist()
table_5 = table_5.drop(max(removed_DI_contract))
table_5 = table_5.sort_values(by="Dates").reset_index(drop=True)

# Definition of forward rates:
def forward_rate(row):
    if row.name == 0:
        return selic_over_rate
    return (pow(pow((1+table_5.loc[row.name, "Closement rates"]/100),(table_5.loc[row.name, "NDU"]/252))/pow((1+table_5.loc[row.name - 1, "Closement rates"]/100),(table_5.loc[row.name - 1, "NDU"]/252)),(252/(table_5.loc[row.name, "NDU"] - table_5.loc[row.name - 1, "NDU"]))) - 1)*100

table_5["Forward rate"] = table_5.apply(forward_rate, axis=1)

# Calculating the forward rate variation in each term:
def forward_variation(row):
    if row.name == 0:
        return 0
    return round((table_5.loc[row.name, "Forward rate"] - table_5.loc[row.name - 1, "Forward rate"])*100,2)

table_5["Forward variation"] = table_5.apply(forward_variation, axis=1)

# Creation of consolidated Copom pricing per reunion:

# Create groups based on "Copom lines"
table_5['Group'] = (table_5['Events'] == 'Copom').cumsum()

# Sum only the lines between each group
table_5['Copom Pricing'] = table_5.groupby('Group')['Forward variation'].transform('sum')

# CDIE Dashboard
bps_copom = table_5.loc[table_5['Events'] == "Copom", "Copom Pricing"].tolist()
bps_copom.pop(-1)
last_copom_index = table_5[table_5["Events"] == "Copom"].index[-1]
last_copom = table_5.loc[last_copom_index, "Forward variation"].tolist()
bps_copom.extend([last_copom])
dates_copom_list = table_5.loc[table_5['Events'] == "Copom", "Dates"].tolist()
dates_copom_list_formatted = [date.strftime("%b-%Y") for date in dates_copom_list]

# Creating bar chart:
fig, ax = plt.subplots()

x = np.arange(len(dates_copom_list_formatted))
ax.bar(x, bps_copom, width=0.8, align='center', color='black')    
plt.xticks(x, dates_copom_list_formatted, fontsize=7.5)
plt.yticks(np.arange(0, max(bps_copom) + 25, 25))
ax = plt.gca()
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
# Add main title
plt.suptitle('CDIE', fontsize=16, fontweight='light', fontfamily='rockwell', x=0.135, y=0.9)

# Add sub title
plt.title(f'COPOM Market Expectations, bps - Last: {fix_date:%d-%m-%Y}', fontsize=12, fontweight='light', fontfamily='rockwell', x=0.36, y=1.05)

# Adjusting layout to avoid overlap
plt.tight_layout(rect=[0, 0, 1, 0.95])
ax.legend()

plt.show()