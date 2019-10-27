import matplotlib.pyplot as plt
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Color
from openpyxl.styles import colors
from PIL import Image
import time
import datetime

"""
A program to create iv curves for data of socal cells in an excel file.
Default start row is 1.
Assuming that if some data measured in the same day is of the same cell,
only the most recent data is correct.
"""

#NUM_COL is a constant. Will not change.
#df.shape returns a tuple (num_col, num_row)
NUM_COLS = 12

#Voltage is row number 1. Number of row in dataframe = num_row in excel - 2
start_row_num = 1
chart_title = "IV curve - "

def create_graph(df, start_row_num = 1):
    """
    Create a scatter plot of the data for one cell, incuding Forward and 
    Reverse data. Data is stored in dataframe (df).
    """
    
    x_forward = []
    y_forward = []
    
    x_reverse = []
    y_reverse = []
    
    v_row_forward = start_row_num
    i_row_forward = start_row_num + 1
    
    v_row_reverse = start_row_num + NUM_COLS + 1 - 2 
    #to get the date
    i_row_reverse = start_row_num + NUM_COLS + 2
    
    forward_data = df.iloc[list(range(v_row_forward - 1, i_row_forward + 1))]
    reverse_data = df.iloc[list(range(v_row_reverse, i_row_reverse + 1))]
    #-1 because we want to include the row with the name of the cell
    #only in forward data
    
    for index, row in forward_data.iterrows():
        #Get forward data
        l_row = list(row)
        if "Comment:" in l_row:
            name = l_row[1]
        elif "Voltage (V)" in l_row:
            x_forward = l_row[1:]
        elif "Current (A)" in l_row:
            y_forward = l_row[1:]
    
    for index,row in reverse_data.iterrows():
        #Get reverse data
        l_row = list(row)
        if "Voltage (V)" in l_row:
            x_reverse = l_row[1:]
            date_index = index - 2
            #Take only the date, not the time
            date = str(df.iloc[[date_index]].values.tolist()[0][0])[:10]
            #date = reverse_data.iloc[index, 0]
            
        if "Current (A)" in l_row:
            y_reverse = l_row[1:]   
    
    #Create the graph
    plt.scatter(x_forward, y_forward, s = 2, label = "Forward", color = 'b', \
                marker =',')
    plt.scatter(x_reverse, y_reverse, s = 2, label = "Reverse", color = 'r')
    plt.title(chart_title + name + " measured " + date, fontweight = 'bold')
    plt.xlabel('Voltage (V)')
    plt.ylabel('Current (A)')
    plt.legend()
    
    #Add the axes through the origin
    plt.axvline(x=0, color = 'k', linewidth = 0.5, alpha = 0.8)
    plt.axhline(y=0, color = 'k', linewidth = 0.5, alpha = 0.8)
    
    fig_file_name = name + " measured " + date + '.png'
    plt.savefig(fig_file_name)
    plt.close()
    
    return fig_file_name
    

def iv_graphs():
    """
    Create all iv curves from data in the excel file. The file can be updated
    by fillign in the start row. The start_row of a new set of data is the row
    of the date of the forward data.
    
    Data needs to be consistent: Forward - Reverse, Forward - Reverse for the
    program to work smoothly
    
    """
    source_name = input("What is the name of the source excel file? ") + ".xlsx"
    start_row = int(input("What is the number of the start row? "))
    start_time = time.time()
    
    #graph_pos is row number to place the graph. The column is determined 
    #as column D.
    
    #row_num is the row we are currently in. It serves to keep track of the
    #number of charts need to be created.
    row_num = start_row
    graph_pos = start_row
    
    #Create dataframe
    df = pd.read_excel(source_name)
    wb = openpyxl.load_workbook(source_name)
    #Using the default sheet name: Sheet1
    sheet = wb.worksheets[0]   
    
    while row_num < df.shape[0]:
        #df.shape[0] is the number of rows
        
        img_file = create_graph(df, start_row)
        img = openpyxl.drawing.image.Image(img_file)
        sheet.add_image(img, 'D'+str(graph_pos))
        wb.save(source_name)        
        
        graph_pos += NUM_COLS*2 + 2
        start_row += NUM_COLS*2 + 2
        row_num = start_row + 22
    
    #Comment this if don't want to change Efficiency text color
    eff_index_list = []
    for index,row in df.iterrows():
        l_row = list(row)
        if "Comment:" in l_row:
            if "Reverse" not in l_row[1] and "d" not in l_row[1]:
                #name_index = index
                eff_index = index + 3
                #Efficiency is 3 rows away from name row
                eff_index_list.append(eff_index)
    
    
    #Indexing in excel is different 
    xl_eff_list = [str(i+2) for i in eff_index_list]
    
    #Making efficiency red
    for row_num in xl_eff_list:
        sheet['A' + row_num].font = Font(color = colors.RED)
        sheet['B' + row_num].font = Font(color = colors.RED)               
    #end comment    
    wb.save(source_name) 
    
    ext_time = time.time() - start_time
    
    print("Plotting completed. Plotting took " + \
          str(datetime.timedelta(seconds=round(ext_time))) + ".")
                
#Run the program automatically when this Python file is opened
iv_graphs()
