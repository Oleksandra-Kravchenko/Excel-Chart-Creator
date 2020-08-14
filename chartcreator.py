"""
This code creates bar charts based on a pivot table that uses data from a given workbook. 
It saves every worksheet from the workbook as separate excel workbook that includes a table and a chart.
""" 
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from openpyxl import Workbook

# the workbook the contains the data for the plot
FILE = 'data.xlsx'

def chart_creator():
    # open the workbook and create a list of sheets in it 
    table = pd.ExcelFile(FILE)
    sheets = table.sheet_names
    
    data = [pd.read_excel(FILE, sheet_name = sheet) for sheet in sheets]
    # create a separate workbook for each sheet from the main main workbook
    for s in range(len(sheets)):
        book = Workbook()
        book.save(f'{sheets[s]}.xlsx')
    # create a chart for every workbook     
    for s in range(len(sheets)):
        df = pd.DataFrame(data[s])
        # create a pivot table
        df_table = pd.pivot_table(df, index='Year', values='Sales', columns='Quarter',
                                  margins=True, margins_name='Total',
                                  aggfunc='sum', fill_value = 'N/A')
        df_table1 = df_table.iloc[:-1]
        # create a chart that shows sales amount in every quarter and total sales
        #grouped by year
        chart = df_table1.plot(kind='bar', width=0.7, title=sheets[s])
        chart_img = BytesIO()
        plt.figure()
        chart.figure.savefig(chart_img)
        # wrtie to workbooks 
        writer = pd.ExcelWriter(f'{sheets[s]}.xlsx', 
                                engine= 'xlsxwriter')
        # save new workbooks
        df_table.to_excel(writer,sheet_name=sheets[s])
        worksheet = writer.sheets[f'{sheets[s]}']
        worksheet.insert_image('G1', '', {'image_data':chart_img})
        writer.save()
# run chart creator        
chart_creator()
