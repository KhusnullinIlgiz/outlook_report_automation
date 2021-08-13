import os
import re
from datetime import datetime
from textwrap import wrap

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd
import xlsxwriter


def gen_report(path, subject):
    """
    This function is used to read csv file from given directory, make an analise and create report
    with plot and data in excel file saved in the output directory.
    :param path: directory to csv file attached to the Email
    :param subject: subject of the Email
    :return:None
    """
    # read csv file to pandas df
    df = pd.read_csv(path, index_col=None)

    # Time column transformation from unix timestamp format
    try:
        df.Time = pd.to_datetime(df.Time, unit='ms')
    except:
        df.Time = pd.to_datetime(df.Time)

    # creating output directory for saving report
    dir_path = os.path.join('C:\Projects\outputs',
                            subject.lower() + "_" + str(datetime.now().strftime("%Y_%m_%d %H_%M_%S")))
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)

    # creating xlsx file in output dir
    workbook = xlsxwriter.Workbook(os.path.join(dir_path, subject + '.xlsx'))

    # csv file can consist several columns. First column is always Time in ('%Y-%m-%d %H:%M:%S') format which is general
    # for all lines of the graph. Each of the rest columns represent values of line to be analysed. Each line will be
    # added to xlsx file as a new worksheet with max,min and graph
    for i in range(1, len(df.columns)):

        # create new df of Time column and i'th line of the graph
        # dropin 'undefined'/'None' values
        new = (df[df.columns[0]], df[df.columns[i]])
        labels = ("time", df.columns[i])
        new_df = pd.DataFrame.from_dict(dict(zip(labels, new)))
        new_df = new_df[new_df[new_df.columns[1]] != "undefined"]

        # rounding values of column
        if "cpu" in path.lower():
            new_df[new_df.columns[1]] = new_df[new_df.columns[1]].apply(lambda x: x * 100).astype(float).round(3)
        else:
            new_df[new_df.columns[1]] = new_df[new_df.columns[1]].astype(float).round(3)

        # extracting max value and time point of the line
        max_val = new_df[new_df.columns[1]].max()
        idx_max = new_df[new_df[new_df.columns[1]] == new_df[new_df.columns[1]].max()].index.values
        time_max_val = new_df.time.iloc[idx_max[0]]

        # extracting min value and time point of the line
        min_val = new_df[new_df.columns[1]].min()
        idx_min = new_df[new_df[new_df.columns[1]] == new_df[new_df.columns[1]].min()].index.values
        time_min_val = new_df.time.iloc[idx_min[0]]

        avg_val = new_df[new_df.columns[1]].mean().round(3)

        # creating plot of the line
        # setting size, title, lables, legend
        # plotting line values through time, adding max/min points on the plot, setting text of max/min points
        # setting x axis time in ('%Y-%m-%d %H:%M:%S') format
        plot_title = new_df.columns[1]
        fig = plt.figure(figsize=(8, 8))
        ax = fig.add_subplot(111)
        ax.set_title("\n".join(wrap(plot_title, 80)), fontsize=9)
        ax.plot(new_df.time, new_df[new_df.columns[1]], label="series data")
        ax.plot(time_max_val, max_val, 'o', color='r', label="max value")
        ax.plot(time_min_val, min_val, 'o', color='g', label="min value")
        plt.legend()
        plt.text(time_max_val, max_val, str(max_val))
        plt.text(time_min_val, min_val, str(min_val))
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H:%M:%S'))

        # fitting x axis labels
        # saving plot as .png format in output dir together with xlsx file
        fig.autofmt_xdate()
        img = os.path.join(dir_path, re.sub('[^0-9a-zA-Z]+', '_', plot_title) + str(i) + ".png")
        plt.savefig(img)

        # creating worksheet of the xlsx file
        worksheet = workbook.add_worksheet(re.sub('[^0-9a-zA-Z]+', '_', plot_title)[0:10] + str(i))
        worksheet.set_column('A:A', 30)

        # writing results to the worksheets
        # adding .png plot to the worksheet
        worksheet.write('A2', "max: " + str(max_val))
        worksheet.write('A3', "min: " + str(min_val))
        worksheet.write('A4', "avg: " + str(avg_val))
        worksheet.write('A5', "time_max: " + str(time_max_val))
        worksheet.write('A6', "time_min: " + str(time_min_val))
        worksheet.insert_image('B2', img)
    # after all columns were analysed -> close xlsx file
    workbook.close()
