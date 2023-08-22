import pandas as pd
import os
from tkinter import *
from tkinter import filedialog

def get_path():
    global pl
    pl = []
    def clickOK():
        root.destroy()
    def clickOpen():
        root.filename = filedialog.askopenfilename(initialdir='/', title='Select An Excel File',
                                                   filetypes=(('excel files', '*.xlsx'),),multiple=True)
        file_list = root.tk.splitlist(root.filename)
        for i in file_list:
            pl.append(i)
        myEnt.insert(0, root.filename)

    root = Tk()
    root.title('Design Summary')

    myLabel = Label(root, text='Please find the excel file that you want to read: ', font=15)
    myLabel.pack()

    myEnt = Entry(root, width=45)
    myEnt.pack()

    myButton1 = Button(root, text='...', command=clickOpen)
    myButton2 = Button(root, text='OK', command=clickOK, width=8)
    myButton1.pack()
    myButton2.pack()

    root.mainloop()
    # print(l)
    return pl

def extract_row_data(get_path):
    filepath =[]
    for i in get_path:
        filepath.append(i)
        #print(filepath)
    result = []
    for i in filepath:
        tdn_sum = pd.read_excel(i,
                               sheet_name='Intro',
                               skiprows=7,
                               usecols='B,E,G',
                               nrows=17,
                               header=None)
        tdn_num = tdn_sum.iloc[[0, 1, 4, 5, 8, 9, 12, 13, 14], [0]]
        tdn_list = tdn_num[1].tolist()
        concrete_quant = round(tdn_sum.iat[16, 2],2)
        total_steel_quant = round(tdn_sum.iat[15,2],2)
        seg_sheet = pd.read_excel(i,
                                  sheet_name='SW',
                                  skiprows=29,
                                  usecols='D:I',
                                  nrows=4,
                                  header=None)

        segmentation = seg_sheet.values[0].tolist()
        segmentation_str = ','.join(map(str,segmentation))
        height = seg_sheet.iat[2,1].tolist()
        breadth = seg_sheet.iat[3,1].tolist()
        row_data = [height,breadth,segmentation_str]+tdn_list+[concrete_quant,total_steel_quant]
        result.append(row_data)
    print(result)
    return result




def create_new_frame(extract_row_data):
    filename =[]
    for i in pl:
        #print(i)
        filename.append(os.path.basename(i))
    pier_numbers = []
    for i in filename:
        pier_numbers.append(os.path.splitext(i)[0])
    #print(pier_numbers)

    col_names = ['H(m)', 'B(m)', 'Segmentation',
                 'T1a', 'T1b', 'T2a', 'T2b', 'T3a', 'T3b', 'B1', 'B2', 'B3',
                 'Concrete Quantity (m3)', 'Total Steel (Ton)']

    df_new = pd.DataFrame(extract_row_data,
                      columns=col_names,
                      index=pier_numbers)
    output = df_new.to_excel('summary.xlsx')
    return output

create_new_frame(extract_row_data(get_path()))
#extract_row_data(get_path())


