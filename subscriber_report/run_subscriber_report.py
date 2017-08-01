import subscriber_report as sr
import sys
import pandas as pd

filepath = sys.argv[1]

agg_types = ['week', 'month', 'quarter']

all_dfs = {}

rep = sr.subscriber_report(filepath, 'week')
markets = rep.get_markets()

# Create dataframe for each Market and Agg type combination. Assign to dictionary to call later on in script
for market in markets:
    for agg_type in agg_types:

        rep = sr.subscriber_report(filepath, agg_type, market)
        data = rep.build_dataset()

        key = market+agg_type

        all_dfs[key] = {'market':market, 'agg_type':agg_type, 'data':data}


for agg_type in agg_types:

    rep = sr.subscriber_report(filepath, agg_type)
    data = rep.build_dataset()

    key = 'aggregate'+agg_type

    all_dfs[key] = {'market':'aggregate', 'agg_type':agg_type, 'data':data}


num_reports = len(all_dfs)
dict_keys = all_dfs.keys()

week_dfs = {}
month_dfs = {}
quarter_dfs = {}

writer = pd.ExcelWriter('subscriber_report.xlsx', engine='xlsxwriter')

### Get aggregate dataframes with key of Aggregate
for key in dict_keys:
    market = all_dfs[key]['market']
    agg_type = all_dfs[key]['agg_type']
    data = all_dfs[key]['data']

    if agg_type == 'week' and market == 'aggregate':
        week_dfs['Aggregate'] = data
    if agg_type == 'month' and market == 'aggregate':
        month_dfs['Aggregate'] = (data)
    if agg_type == 'quarter' and market == 'aggregate':
        quarter_dfs['Aggregate'] = (data)

### Then the rest
for key in dict_keys:
    market = all_dfs[key]['market']
    agg_type = all_dfs[key]['agg_type']
    data = all_dfs[key]['data']

    if agg_type == 'week' and market != 'aggregate':
        week_dfs[market] = data
    if agg_type == 'month' and market != 'aggregate':
        month_dfs[market] = data
    if agg_type == 'quarter' and market != 'aggregate':
        quarter_dfs[market] = data

# Loop through each df to create report
week_key_list = []
month_key_list = []
quarter_key_list = []

# Get aggregates first
week_key_list.append('Aggregate')
month_key_list.append('Aggregate')
quarter_key_list.append('Aggregate')

# Make list for other keys and sort alphabetically
other_week_key_list = sorted(key for key in week_dfs.keys() if key != 'Aggregate')
other_month_key_list = sorted(key for key in month_dfs.keys() if key != 'Aggregate')
other_quarter_key_list = sorted(key for key in quarter_dfs.keys() if key != 'Aggregate')

# Alphabetically sorted list of keys with aggregate first
final_week_key_list = week_key_list + other_week_key_list
final_month_key_list = month_key_list + other_month_key_list
final_quarter_key_list = quarter_key_list + other_quarter_key_list

# Build each sheet
def write_sheet(key_list, agg_dfs, agg):
    """
    With each Agg Type as its own sheet, create the sheets and format
    
    Args:
        key_list (Python List): List containing the sorted keys 
        agg_dfs (Python Dictionary): Dictionary containing the Market as the key and the DataFrame as the value 
        agg (string): Agg Type, accepts 'Week', 'Month', or 'Quarter' 
    """

    row = 0
    workbook=writer.book
    for key in key_list:
        sheetname = '{agg}ly Report'.format(agg=agg.capitalize())

        df = agg_dfs[key]
        start_row = row + 2
        df.to_excel(writer, sheet_name=sheetname, startrow=start_row, startcol = 0)

        worksheet = writer.sheets[sheetname]
        worksheet.write_string(row=row, col=0, string = key)

        col_format = workbook.add_format({'italic':True, 'bold':True, 'font_size':16})
        worksheet.set_column('A:A', 20, col_format)

        row = start_row + len(df.index) + 5




write_sheet(final_week_key_list, week_dfs, 'Week')
write_sheet(final_month_key_list, month_dfs, 'Month')
write_sheet(final_quarter_key_list, quarter_dfs, 'Quarter')



writer.save()