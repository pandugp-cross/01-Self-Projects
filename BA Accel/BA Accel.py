# -*- coding: utf-8 -*-
"""
Created on Wed Jun 21 22:01:22 2023

@author: Pariz Nanap Pandu Gumelar Pratama
"""
import tkinter as tk
import os
import pandas as pd
import shutil
root=tk.Tk()
root.geometry("300x200")
pathinput="Input/Input BAUT.xlsx"
options_list=["Create zip From List","Create Notepad From PPM Link List"]
value_inside = tk.StringVar(root)
value_inside.set("Create zip From List")
question_menu = tk.OptionMenu(root, value_inside, *options_list)
question_menu.pack()
checkinput=os.path.exists(pathinput)
fileprocessed=0
def zipping(x):
    shutil.make_archive("Output/"+str(x), 'zip',root_dir="Output/"+str(x))
    shutil.rmtree("Output/"+str(x))
def createfolder(x):
    checkf=os.path.exists("Output/"+str(x)+"/"+str(x))
    if checkf==False:
        os.mkdir("Output/"+str(x))
        os.mkdir("Output/"+str(x)+"/"+str(x))
def matrixpo(ag,src,dst):
    files=os.listdir("Reference/Matrix PO")
    for file in files:
        if file.startswith(str(ag)+"."):
            shutil.copy("Reference/Matrix PO/"+str(file), dst)
def customfile(c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,dst):
    files=os.listdir("Reference/Custom File")
    for file in files:
        if file.startswith(str(c1)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c2)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c3)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c4)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c5)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c6)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c7)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c8)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c9)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
    for file in files:
        if file.startswith(str(c10)+"."):
            shutil.copy("Reference/Custom File/"+str(file), dst)
def BASSAC(x,y,z,aa,ab,ac,ad,ae,af,ag,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10):
    global fileprocessed
    src=""
    dst=""
    createfolder(x)
    dst="Output/"+str(x)+"/"+str(x)
    baseline=os.path.exists("Reference/Baseline Approved/"+str(y))
    if baseline==True:
        src="Reference/Baseline Approved/"+str(y)
        baselinef=os.listdir("Reference/Baseline Approved/"+str(y))
        shutil.copytree(src,dst,dirs_exist_ok = True)
    Capture=os.path.exists("Reference/Capture Term of Payment")
    if Capture==True:
        src="Reference/Capture Term of Payment"
        capturef=os.listdir("Reference/Capture Term of Payment")
        shutil.copytree(src,dst,dirs_exist_ok = True)
    Capture=os.path.isfile("Reference/Endorsed/"+str(af)+".msg")
    if Capture==True:
        src="Reference/Endorsed/"+str(af)+".msg"
        shutil.copy(src,dst)
    matrixpo(ag, src, dst)
    customfile(c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, dst)
    zipping(x)
    fileprocessed=fileprocessed+1
def BAPAC(x,y,z,aa,ab,ac,ad,ae,af,ag,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10):
    global fileprocessed
    src=""
    dst=""
    createfolder(x)
    dst="Output/"+str(x)+"/"+str(x)
    baseline=os.path.exists("Reference/Baseline Approved/"+str(y))
    if baseline==True:
        src="Reference/Baseline Approved/"+str(y)
        baselinef=os.listdir("Reference/Baseline Approved/"+str(y))
        shutil.copytree(src,dst,dirs_exist_ok = True)
    Capture=os.path.exists("Reference/Capture Term of Payment")
    if Capture==True:
        src="Reference/Capture Term of Payment"
        capturef=os.listdir("Reference/Capture Term of Payment")
        shutil.copytree(src,dst,dirs_exist_ok = True)
    Capture=os.path.exists("Reference/CR Approved/"+str(aa))
    if Capture==True:
        src="Reference/CR Approved/"+str(aa)
        capturef=os.listdir("Reference/CR Approved/"+str(aa))
        subf=dst+"/"+str(aa)
        shutil.copytree(src,subf,dirs_exist_ok = True)
    Capture=os.path.isfile("Reference/KK Approved/"+str(ab)+".pdf")
    if Capture==True:
        src="Reference/KK Approved/"+str(ab)+".pdf"
        shutil.copy(src,dst)
    Capture=os.path.isfile("Reference/CA Integration/"+str(ac)+".pdf")
    if Capture==True:
        src="Reference/CA Integration/"+str(ac)+".pdf"
        shutil.copy(src,dst)
    Capture=os.path.isfile("Reference/CA Optim/"+str(ad)+".pdf")
    if Capture==True:
        src="Reference/CA Optim/"+str(ad)+".pdf"
        shutil.copy(src,dst)
    Capture=os.path.isfile("Reference/Endorsed/"+str(af)+".msg")
    if Capture==True:
        src="Reference/Endorsed/"+str(af)+".msg"
        shutil.copy(src,dst)
    matrixpo(ag, src, dst)
    customfile(c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, dst)
    zipping(x)
    fileprocessed=fileprocessed+1
def BASG(x,y,z,aa,ab,ac,ad,ae,af,ag,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10):
    global fileprocessed
    createfolder(x)
    src=""
    dst=""
    dst="Output/"+str(x)+"/"+str(x)
    Capture=os.path.exists("Reference/Capture Term of Payment")
    if Capture==True:
        src="Reference/Capture Term of Payment"
        capturef=os.listdir("Reference/Capture Term of Payment")
        shutil.copytree(src,dst,dirs_exist_ok = True)
    Capture=os.path.isfile("Reference/CA Integration/"+str(ac)+".pdf")
    if Capture==True:
        src="Reference/CA Integration/"+str(ac)+".pdf"
        shutil.copy(src,dst)
    Capture=os.path.isfile("Reference/CA Optim/"+str(ad)+".pdf")
    if Capture==True:
        src="Reference/CA Optim/"+str(ad)+".pdf"
        shutil.copy(src,dst)
    Capture=os.path.isfile("Reference/Punchlist Clearance/"+str(ae)+".msg")
    if Capture==True:
        src="Reference/Punchlist Clearance/"+str(ae)+".msg"
        shutil.copy(src,dst)
    Capture=os.path.isfile("Reference/Endorsed/"+str(af)+".msg")
    if Capture==True:
        src="Reference/Endorsed/"+str(af)+".msg"
        shutil.copy(src,dst)
    matrixpo(ag, src, dst)
    customfile(c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, dst)
    zipping(x)
    fileprocessed=fileprocessed+1
def NotF(x,y,z,aa,ab,ac,ad,ae,af,ag,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10):
    global fileprocessed
    fileprocessed=fileprocessed
def modes():
    code="{}".format(value_inside.get())
    if code=="Create zip From List":
        createrar()
    elif code=="Create Notepad From PPM Link List":
        createnote()
def notepads(x,y):
    global fileprocessed
    if y !="" and pd.isna(y) is False:
        checkf=os.path.exists("Output/Notepad")
        if checkf==False:
            os.mkdir("Output/Notepad")
        with open("Output/Notepad/"+str(x)+'.txt', 'w') as f:
            f.write(str(y))
            fileprocessed=fileprocessed+1
def createnote():
    global fileprocessed
    fileprocessed=0
    BAInput = pd.read_excel(pathinput,sheet_name="Data",skiprows=3)
    index=(len(BAInput.index))
    for x in range (index):
        notepads(BAInput.iloc[x,6],BAInput.iloc[x,24])
    message = tk.Label(root, text="Finished Processing to Notepad, "+str(fileprocessed)+" out of "+str(index)+" Are Processed")
    message.pack()
    root.mainloop()
def createrar():
    global fileprocessed
    fileprocessed=0
    BAInput = pd.read_excel(pathinput,sheet_name="Data",skiprows=3)
    index=(len(BAInput.index))
    for x in range (index):
        option(BAInput.iloc[x,6], BAInput.iloc[x,2], BAInput.iloc[x,4],BAInput.iloc[x,10],BAInput.iloc[x,7],BAInput.iloc[x,9],BAInput.iloc[x,11],BAInput.iloc[x,12],BAInput.iloc[x,8],BAInput.iloc[x,13]
               ,BAInput.iloc[x,14],BAInput.iloc[x,15],BAInput.iloc[x,16],BAInput.iloc[x,17],BAInput.iloc[x,18]
               ,BAInput.iloc[x,19],BAInput.iloc[x,20],BAInput.iloc[x,21],BAInput.iloc[x,22],BAInput.iloc[x,23])
    message = tk.Label(root, text="Finished Processing to RAR, "+str(fileprocessed)+" out of "+str(index)+" Are Processed")
    message.pack()
    root.mainloop()
switch={
        "BA Installation/SSAC":BASSAC,
        "BAUT/BAMS PAC":BAPAC,
        "BAMS/BAMG":BASG
        }
def option(x,y,z,aa,ab,ac,ad,ae,af,ag,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10):
    return switch.get(z,NotF)(x,y,z,aa,ab,ac,ad,ae,af,ag,c1,c2,c3,c4,c5,c6,c7,c8,c9,c10)
if checkinput==False:
    message = tk.Label(root, text="Folder Not Exist, Pastikan Folder input ada dan File Input BAUT.xlsx ada didalam")
    message.pack()
    root.mainloop()
else:
    submit_mode=tk.Button(root, text='Submit', command=modes)
    submit_mode.pack()
    root.mainloop()
