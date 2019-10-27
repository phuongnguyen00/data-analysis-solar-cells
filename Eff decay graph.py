import matplotlib.pyplot as plt
import pandas as pd
import openpyxl

"""
A program to graph the decay of efficiency of one particular cell over time.
(Forward data)
Assumming the current format of data.
"""

source_file = input("What is the source file name? ") +".xlsx"

df = pd.read_excel(source_file)
wb = openpyxl.load_workbook(source_file)
sheet = wb.worksheets[0]

#Name of cell. The whole file contains data about one cell.
name = sheet['B2'].value
graph_title = "Efficiency Decay of Cell " + name

#List of row numbers of all dates:
date_list = list(range(sheet.min_row, sheet.max_row, 26))
days_x_axis = [0]

for i in range(len(date_list) - 1):
    former_row = str(date_list[i+1])
    later_row = str(date_list[i])
    interval = (sheet['A'+ former_row].value - \
        sheet['A'+ later_row].value).days
    days_x_axis.append(interval)

eff_index_list = []

#Find the column number of the Efficiencies. Only take the forward direction.
for index, row in df.iterrows():
    l_row = list(row)
    if "Comment:" in l_row:
        if "Reverse" not in l_row[1] and "d" not in l_row[1]:
            #name_index = index
            eff_index = index + 3
            eff_index_list.append(eff_index)
            
#Indexing in excel is different 
xl_eff_list = [str(i+2) for i in eff_index_list]

eff_y_axis = []

for row_num in xl_eff_list:
    eff_y_axis.append(sheet['B' + row_num].value)

#Plot the graph with the x-axis as Time and the y-axis as Efficiency
plt.plot(days_x_axis, eff_y_axis, marker = 'o')
plt.title(graph_title)
plt.xlabel("Time (Days)")
plt.ylabel(sheet['A5'].value)
plt.show()
plt.savefig(graph_title)


    
    




            
            
            