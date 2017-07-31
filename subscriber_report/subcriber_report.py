import pandas as pd
import numpy as np
import datetime
import calendar

class subscriber_report(object):

    def __init__(self, filepath, agg_type, market=None, year_filter=None):

        self.filepath = filepath
        self.agg_type = agg_type
        self.year_filter = year_filter
        self.market = market

    def read_data(self, filepath, agg_type):

        data = pd.read_excel(filepath)
        data['date'] = data.activity_date.apply(lambda x: x.date())
        data['year'] = data.date.apply(lambda x: x.year)
        data['month'] = data.date.apply(lambda x: x.month)
        data['week'] = data.date.apply(lambda x: x.strftime('%U'))
        data['dom'] = data.date.apply(lambda x: x.day)
        data['day'] = data.date.apply(lambda x: x.isoweekday())
        data['last_dom'] = data.date.apply(lambda x: calendar.monthrange(x.year, x.month)[1])
        data['is_last_dom'] = np.where(data.dom == data.last_dom, 1, 0)
        data['quarter'] = data.month.apply(lambda x: (x-1)  // 3 + 1)
        data['is_last_day_of_quarter'] = np.where(((data.month == 3) & (data.dom == 31)) | ((data.month == 6) & (data.dom == 30)) |
                      ((data.month == 9) & (data.dom == 30)) | ((data.month == 12) & (data.dom == 31)), 1, 0)
        #data['aggyear'] = data.year.apply(lambda x: agg_type+str(x))
        return data

    def get_ending_subscribers(self, data, agg_type, market):

        if agg_type == 'week':
            df = data[data.day == 7]
        elif agg_type == 'month':
            df = data[data.is_last_dom == 1]
        elif agg_type == 'quarter':
            df = data[data.is_last_day_of_quarter == 1]

        if market:
            total_subscribers = df[df.market == market][['total_subscribers', agg_type, 'year']]
        else:
            total_subscribers = df.groupby([agg_type, 'year']).agg({'total_subscribers':sum}).reset_index()

        return total_subscribers

    def build_agg_data(self, data, total_subscribers, agg_type, market=None, year_filter=None):

        if year_filter:
            df = data[data.year == year_filter]
        else:
            df = data

        if market:
            df = df[df.market == market]

        agg_data = df.groupby([agg_type, 'year']).agg(
            {'new_subscriptions': sum,
             'self_install': sum,
             'professional_install': sum,
             'disconnects': sum,
             'post_install_returns': sum,
             'total_disconnects': sum})

        agg_data['net_gain'] = agg_data.new_subscriptions - agg_data.total_disconnects

        merged_agg_data = agg_data.reset_index().merge(total_subscribers, on=[agg_type, 'year'])
        merged_agg_data['beginning_subs'] = merged_agg_data.total_subscribers.shift(1)

        end_df = merged_agg_data.set_index(['year', agg_type]).sort_index().transpose().reindex(['beginning_subs', 'new_subscriptions', 'self_install', 'professional_install',\
                          'total_disconnects', 'post_install_returns', 'disconnects',\
                          'net_gain', 'total_subscribers']).rename(index={
                                                                'beginning_subs':'Beginning Subscribers',
                                                                'new_subscriptions':'Total Connects',
                                                                'self_install':'Self Installs',
                                                                'professional_install':'Pro Installs',
                                                                'total_disconnects':'Total Disconnects',
                                                                'post_install_returns':'Post Install Returns',
                                                                'disconnects':'Disconnects',
                                                                'net_gain':'Net Gain',
                                                                'total_subscribers':'Ending Subs'
                          })

        return end_df

    def build_dataset(self):
        filepath = self.filepath
        agg_type = self.agg_type
        year_filter = self.year_filter
        market = self.market

        data = self.read_data(filepath, agg_type)
        total_subscribers = self.get_ending_subscribers(data, agg_type, market)
        report = self.build_agg_data(data, total_subscribers, agg_type, year_filter)

        return report

    def get_markets(self):
        filepath = self.filepath
        data = pd.read_excel(filepath)
        markets = data.market.unique()

        return markets









