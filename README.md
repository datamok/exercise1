# Exercise 1 : Automate Excel Report

### Required Packages:

* Pandas 
* Numpy
* Datetime
* Calendar
* XlsxWriter
* Sys

### How to Run:

1. Clone the Github Repo to your local machine
2. Navigate to `exercise1/subscriber_report/`
3. From the `subscriber_report` directory, run `python run_subscriber_report.py 'path/to/data/file.xslx'`
4. The report will be dropped in the current working directory `exercise1/subscriber_report` as `subscriber_report.xlsx`

### How it Works

After examining the report, it became clear to me that the raw data is being aggregated by different slices. Primarily by `market` and some measure of time, otherwise known as `agg_type`. Each section in the example report is a combination of a `market` and `agg_type`. That being said, the code was designed to pass in parameters about a certain market-agg_type combination and return a dataframe with the desired metrics. 

There are three main pieces to the class. Read the data, find ending subscribers, and build the aggregated data. I will touch on these in more detail.

**Read the Data**

The first piece is to read in the data from the raw excel file. Since the aggregations would be sliced by different time periods, I added fields to identify time attributes such as day of week, day of month, month of year, last day of the current month, etc. because I would need to use these as flags for filtering downstream. 

**Get Ending Subscribers**

One of the more difficult pieces was knowing how many subscribers the `agg_type` period ended with. For example, if aggregating by `week`, I had to find the number of total_subscribers on the 7th day of the prior week. If `month`. I needed to identify the last day of each month, and grab the total_subscribers from the month prior where the day was equal to the last day of the month.

I used the flags and time attributes in the `read_data` function to identify the right rows of data that contained the ending subscibers in each time period and output these into a dataframe, indexed by the agg_type and year.

**Build Aggregated Data**

Now that I had the ending subscribers piece, I then filter the data by any filers provided (year or market) and aggregate. `Net gain` is a field that is calculated after the fact, and then I performed a lag function on the `total_subscribers` field to get the `beginning_subscribers` field. The dataframe is then transposed and organized and returned in the output. 

### Limitations

The code could be more dry. I end up repeating myself sometimes and I probably don't always use the most efficient or pythonic means of getting what I need.

Also, the format of the output doesn't exactly match the one provided because of a limitation of the package I used, `xlsxwriter`. Apparently, indeces (in report: bolded and outlined) aren't able to be reformatted because they already have a default format and it can't be overridden. This means I wasn't able to match the example report exactly. 


### Additional Info / Discussion

Before I started writing the modules and orchestrator, I did a lot of work up front to step through the process on a specific example. In other words, I tried to create the Atlanta 2017 Week over Week report from the data. Once I was able to get the results in a table, I started to generalize and work my way outward from there. 

The work I did for this is in the `Notebooks` folder in the repo. However, these are a `.ipynb` filetype, which is only accessible through `jupyter`. Jupyter Notebooks allow me to run one line of code at a time and saves objects to memory. This helps me greatly in looking at and understanding data and how certain functions work. The work I did isn't as neat and tidy as the code that I wrote, but it is there if you want to get a glimpse into my thought process. 

If you want to look at these notebooks, I strongly suggest installing the `anaconda` distribution.  Once installed, you can hop into the `Notebooks` directory and run `jupyter notebook` in your terminal. A jupyter hub will appear in your browser, at which point you can click on the notebook and view the code.

---

Please let me know if you have any questions or have trouble getting the code to run. Thanks! 

-Kyle 