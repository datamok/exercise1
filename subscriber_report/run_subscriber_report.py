import subcriber_report as sr
import sys
import pandas as pd

filepath = sys.argv[1]

agg_types = ['week', 'month', 'quarter']

all_dfs = {}

rep = sr.subscriber_report(filepath, 'week')
markets = rep.get_markets()

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

week_dfs = []
month_dfs = []
quarter_dfs = []

writer = pd.ExcelWriter('subscriber_report.xlsx', engine='xlsxwriter')

# for key in dict_keys:
#     market = all_dfs[key]['market']
#     agg_type = all_dfs[key]['agg_type']
#     data = all_dfs[key]['data']
#
#     sheetname = market.capitalize() + ' Aggregated by '+ agg_type.capitalize()
#
#     if agg_type == 'week' and market != 'aggregate':
#         data.to_excel(writer, sheet_name=sheetname)
#     if agg_type == 'week' and market == 'aggregate':
#         data.to_excel(writer,sheet_name='Total Aggregates by '+agg_type.capitalize())
#
#     if agg_type == 'month' and market != 'aggregate':
#         data.to_excel(writer, sheet_name=sheetname)
#     if agg_type == 'month' and market == 'aggregate':
#         data.to_excel(writer,sheet_name='Total Aggregates by '+agg_type.capitalize())
#
#     if agg_type == 'quarter' and market != 'aggregate':
#         data.to_excel(writer, sheet_name=sheetname)
#     if agg_type == 'quarter' and market == 'aggregate':
#         data.to_excel(writer,sheet_name='Total Aggregates by '+agg_type.capitalize())
#
#

### Hack to get aggregates first in every list
for key in dict_keys:
    market = all_dfs[key]['market']
    agg_type = all_dfs[key]['agg_type']
    data = all_dfs[key]['data']

    if agg_type == 'week' and market == 'aggregate':
        week_dfs.append(data)
    if agg_type == 'month' and market == 'aggregate':
        month_dfs.append(data)
    if agg_type == 'quarter' and market == 'aggregate':
        quarter_dfs.append(data)

### Then the rest
for key in dict_keys:
    market = all_dfs[key]['market']
    agg_type = all_dfs[key]['agg_type']
    data = all_dfs[key]['data']
# hack to get aggregates first
    if agg_type == 'week' and market != 'aggregate':
        week_dfs.append(data)
    if agg_type == 'month' and market != 'aggregate':
        month_dfs.append(data)
    if agg_type == 'quarter' and market != 'aggregate':
        quarter_dfs.append(data)

# Loop through each df to create report
row = 0
for df in week_dfs:
    df.to_excel(writer, sheet_name = 'Weekly Report', startrow=row, startcol=0)
    row = row + len(df.index) + 5

row = 0
for df in month_dfs:
    df.to_excel(writer, sheet_name = 'Monthly Report', startrow=row, startcol=0)
    row = row + len(df.index) + 5

row = 0
for df in quarter_dfs:
    df.to_excel(writer, sheet_name = 'Quarterly Report', startrow=row, startcol=0)
    row = row + len(df.index) + 5


writer.save()