import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import xlsxwriter
from datetime import datetime

# creating excel sheet to export:
workbook = xlsxwriter.Workbook('Current_File.xlsx')

#getting today's date to name the file:
now = datetime.now()
date_time = now.strftime("%b-%d-%Y")



#function to upload the files:
def UploadAction_1(event=None):
    filename1 = filedialog.askopenfilename()
    global Current_File
    Current_File = filename1
    print('Selected:', filename1)
def UploadAction_2(event=None):
    filename2 = filedialog.askopenfilename()
    global Prev_File
    Prev_File = filename2
    print('Selected:', filename2)
#function for Sheet FS_Bench_Supply
def Compare():
    FS_Bench_Supply()
    FS_JAVA_Bench()
    FS_Cross_IG()


#funxtion of sheet with checkbox
def FS_Bench_Supply():
    if chkValue1.get()==True:
        global to_df_bs
        yester_df_bs = pd.read_excel(Prev_File, sheet_name='FS_Bench Supply')
        to_df_bs = pd.read_excel(Current_File, sheet_name='FS_Bench Supply')
        to_only_df_bs = to_df_bs[~to_df_bs['Personnel No'].isin(yester_df_bs['Personnel No'])]
        to_df_bs['New_Comer'] = to_df_bs['Personnel No'].isin(to_only_df_bs['Personnel No'])
        writer = ExcelWriter(workbook)
        to_df_bs.to_excel('FS_Bench_Supply_'+ date_time +'.xlsx', index=False)
        print("Done!")
    else:
        print ("Nothing to do")
def FS_JAVA_Bench():
    if chkValue2.get()==True:
        global to_df_js
        yester_df_js = pd.read_excel(Prev_File, sheet_name=5)
        to_df_js = pd.read_excel(Current_File, sheet_name=5)
        to_only_df_js = to_df_js[~to_df_js['Personnel No'].isin(yester_df_js['Personnel No'])]
        to_df_js['New_Comer'] = to_df_js['Personnel No'].isin(to_only_df_js['Personnel No'])
        writer = ExcelWriter(workbook)
        to_df_js.to_excel('FS_JAVA_Bench_'+ date_time +'.xlsx', index=False)
        print("Done!")
    else:
        print ("Nothing to do")

def FS_Cross_IG():
    if chkValue3.get()==True:
        global to_df_IG
        yester_df_IG = pd.read_excel(Prev_File, sheet_name='Cross IG Supply > 2 Weeks',header=0, skiprows=2)
        to_df_IG = pd.read_excel(Current_File, sheet_name='Cross IG Supply > 2 Weeks', header=0, skiprows=2)
        to_only_df_IG = to_df_IG[~to_df_IG['Personnel No'].isin(yester_df_IG['Personnel No'])]
        to_df_IG['New_Comer'] = to_df_IG['Personnel No'].isin(to_only_df_IG['Personnel No'])
        writer = ExcelWriter(workbook)
        to_df_IG.to_excel('FS_Cross_IG_'+ date_time +'.xlsx', index=False)
        print("Done!")
    else:
        print ("Nothing to do")
#function for pop-up(FITER) window
def FS_Bench_Filter_Popup():
    toplevel = tk.Toplevel()
    label1 = tk.Label(toplevel, text="FS BENCH FILTER ", height=0, width=100)
    label1.pack()
    label2 = tk.Label(toplevel, text="UNDER CONSTRUCTION", height=0, width=100)
    label2.pack()
def FS_JAVA_Bench_Filter_Popup():
    toplevel = tk.Toplevel()
    label1 = tk.Label(toplevel, text="FS JAVA BENCH FILTER ", height=0, width=100)
    label1.pack()
    label2 = tk.Label(toplevel, text="UNDER CONSTRUCTION", height=0, width=100)
    label2.pack()
def CROSS_IG_Bench_Filter_Popup():
    toplevel = tk.Toplevel()
    label1 = tk.Label(toplevel, text="CROSS IG BENCH FILTER ", height=0, width=100)
    label1.pack()
    label2 = tk.Label(toplevel, text="UNDER CONSTRUCTION", height=0, width=100)
    label2.pack()

#function with drop drop list:


#UI Setup
root = tk.Tk()
root.title("Staffing Tool   |   DU 10   |   Kitty Hawk")
#checkbox
chkValue1 = tk.BooleanVar()
chkValue1.set(True)
chkValue2 = tk.BooleanVar()
chkValue2.set(True)
chkValue3 = tk.BooleanVar()
chkValue3.set(True)

# base UI
#basecanvas
canvas = tk.Canvas(root, height='650', width='650', bg='#1D3557' )
canvas.pack()

#title frame
Main_frame = tk.Frame(canvas, bg='#A8DADC')
Main_frame.place(x = 10,y = 10,relx=0.1, rely=0.1, relheight=0.1, relwidth=0.70)

Title_label = tk.Label(Main_frame, text='Staffing Tool', bg='#FFE66D', font=("Courier", 44),justify='center' )
Title_label.grid(row=1, column=1)

#Label and upload Button
label_1 = tk.Label(canvas, text="FS_Bench_Supply :", bg='#E63946', fg='#F1FAEE', font=("Courier", 14),justify='left')
label_1.place(x = 30,y = 200)

Prev_F_b2 = tk.Button(canvas, text='Previous File', command=UploadAction_2, width=13,height=1, bg='#F1FAEE', fg='#1D3557', font=("Courier", 18))
Prev_F_b2.place(x = 30,y = 250)

Curr_F_b1 = tk.Button(canvas, text='Current File', command=UploadAction_1,width=13,height=1, bg='#F1FAEE', fg='#1D3557', font=("Courier", 18))
Curr_F_b1.place(x = 30,y = 335)

#Label and checkbox
label_2 = tk.Label(canvas, text="Sheets Need to Compared :", bg='#E63946', fg='#F1FAEE', font=("Courier", 14),justify='left')
label_2.place(x = 350,y = 200)
ck_b1= tk.Checkbutton(canvas, text="FS Bench Supply ", variable=chkValue1,width=22,height=1, bg='#F1FAEE', fg='#1D3557', font=("Courier", 14),justify='left' )
ck_b1.place(x = 350,y = 250)
ck_b2 =tk.Checkbutton(canvas, text="FS  Java  Bench ", variable=chkValue2,width=22,height=1, bg='#F1FAEE', fg='#1D3557', font=("Courier", 14),justify='left')
ck_b2.place(x = 350,y = 300)
ck_b3 = tk.Checkbutton(canvas,text="Cross IG  Bench ", variable=chkValue3,width=22,height=1, bg='#F1FAEE', fg='#1D3557', font=("Courier", 14),justify='left')
ck_b3.place(x = 350 ,y = 350)

#Get Sheet Button
FS_Bench_Supply_b3 = tk.Button(canvas, text= 'Get Sheets', command = Compare,  bg='#E63946', fg='#F1FAEE', font=("Courier", 18))
FS_Bench_Supply_b3.place(x = 205,y = 425)

#filter label:
label_3 = tk.Label(canvas, text="Sheets Need to be filter with Skill :", bg='#E63946', fg='#F1FAEE', font=("Courier", 14),justify='center')
label_3.place(x = 30,y = 500)

#filter buttons
Filter_FS_b4 = tk.Button(canvas, text= 'FS Bench', command = FS_Bench_Filter_Popup ,  bg='#F1FAEE', fg='#1D3557', font=("Courier", 18))
Filter_FS_b4.place(x = 30 ,y = 550)

Filter_JA_b5  = tk.Button(canvas, text= 'FS   JAVA', command = FS_JAVA_Bench_Filter_Popup ,  bg='#F1FAEE', fg='#1D3557', font=("Courier", 18))
Filter_JA_b5.place(x = 230,y = 550)

Filter_IG_b6  = tk.Button(canvas, text= 'Cross IG', command = CROSS_IG_Bench_Filter_Popup ,  bg='#F1FAEE', fg='#1D3557', font=("Courier", 18))
Filter_IG_b6.place(x = 460,y = 550)

#base root mainloop
root.mainloop()
