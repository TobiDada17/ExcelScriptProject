
import tkinter as tk
from tkinter import filedialog, Text, Entry
import pandas as pd
from openpyxl import Workbook
from tkinter import messagebox

import os

root = tk.Tk()

files = {'mapping': "", 'data': "", 'save path' : ""}

v = tk.IntVar()

def selectMapping():
    filename = filedialog.askopenfilename(initialdir= "/", title = "Select File",
                                          filetypes=(("csv","*.csv"), ("all files", "*.*") ))
    print(filename)
    files['mapping'] = filename
    updateList()

def selectData():
    filename = filedialog.askopenfilename(initialdir= "/", title = "Select File",
                                          filetypes=(("csv","*.csv"), ("all files", "*.*") ))
    print(filename)
    files['data'] = filename
    updateList()

def updateSavePath():
    filename = filedialog.askdirectory(initialdir= "/", title = "Select File")
    print(filename)
    files['save path'] = filename
    updateList()


def updateList():
    for widget in frame.winfo_children():
        widget.destroy()

    for key, file in files.items():
        label = tk.Label(frame, text=(key + ': ' + file), bg="gray")
        label.pack()

def yearFormat(year):
    switcher = {
        '2010/11': '2010',
        '2011/12': '2011',
        '2012/13': '2012',
        '2013/14': '2013',
        '2014/15': '2014',
        '2015/16': '2015',
        '2016/17': '2016',
        '2017/18': '2017',
        '2018/19': '2018',
        '2019/20': '2019',
        '2020/21': '2020'
    }
    return switcher.get(year, "Invalid Year")


def runScript():
    df = pd.read_csv(files['mapping'], skip_blank_lines=True)
    df2 = pd.read_csv(files['data'], na_values=['x'], skip_blank_lines=True)
    metricName = metricEntry.get()
    newDataFrame = {'Area Code': [],
                    'Area Names': [],
                    'Year': [],
                    metricName: []
                    }
    yearHeaders = []
    for year in df2.columns:
        if not 'Area' in year:
            yearHeaders.append(year)

    for (columnName, columnData) in df.iteritems():
        compareName = 'Area code' if v == 1 else 'Area'
        if (columnName == compareName ):
            idTable = columnData.values

            for id in idTable:

                if (v == 1):
                    row = df2.loc[df2['Area Codes'] == id]
                else:
                    row = df2.loc[df2['Area Names'] == id]
                    areaCode = df.loc[df['Area'] == id]
                if not row.empty:
                    for year in yearHeaders:
                        if(v == 1):
                            newDataFrame['Area Code'].append(id)
                            newDataFrame['Area Names'].append(str(row['Area Names'].values[0]))
                        else:
                            newDataFrame['Area Code'].append(str(areaCode['Area code'].values[0]))
                            newDataFrame['Area Names'].append(id)

                        newDataFrame['Year'].append(yearFormat(year))
                        newDataFrame[metricName].append(str(row[year].values[0]))
    df3 = pd.DataFrame(newDataFrame)
    print(files['save path'])
    df3.to_excel(files['save path'] + '/newExcelFile.xlsx', sheet_name=metricName, index=False)
    messagebox.showinfo("Success", "A new Excel File has been create and saved to your desktop")


canvas = tk.Canvas(root, height=700, width=700, bg="#263D42")
canvas.pack()

frame =tk.Frame(root, bg="white")
frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

selectMappings = tk.Button(root, text="Select Mappings", padx=10,
                           pady=5, fg="red", bg="#263D42", command=selectMapping)
selectMappings.pack()

selectMetricData = tk.Button(root, text="Select Data", padx=10,
                           pady=5, fg="red", bg="#263D42", command=selectData)
selectMetricData.pack()

selectMetricData = tk.Button(root, text="Save Excel Path", padx=10,
                           pady=5, fg="red", bg="#263D42", command=updateSavePath)
selectMetricData.pack()

metricLabel = tk.Label(root, text=('Metric Name'), fg="red")
metricLabel.pack()

metricEntry = Entry(root)
metricEntry.pack()


tk.Label(root,
        text="""Is Area Code Present""",
        justify = tk.LEFT,
        padx = 20).pack()
tk.Radiobutton(root,
              text="Yes",
              padx = 20,
              variable=v,
              value=1).pack(anchor=tk.W)
tk.Radiobutton(root,
              text="No",
              padx = 20,
              variable=v,
              value=2).pack(anchor=tk.W)

runSelectScript = tk.Button(root, text="Run Script", padx=10,
                           pady=5, fg="red", bg="#263D42", command=runScript)
runSelectScript.pack()


root.mainloop();