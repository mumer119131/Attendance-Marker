import openpyxl
import numpy as np
import getpass
import PySimpleGUI as sg
from datetime import date
import time


agList=["2019-ag-6051",
"2019-ag-6052",
"2019-ag-6053",
"2019-ag-6054",
"2019-ag-6055",
"2019-ag-6056",
"2019-ag-6057",
"2019-ag-6058",
"2019-ag-6059",
"2019-ag-6060",
"2019-ag-6061",
"2019-ag-6062",
"2019-ag-6063",
"2019-ag-6065",
"2019-ag-6066",
"2019-ag-6067",
"2019-ag-6068",
"2019-ag-6069",
"2019-ag-6070",
"2019-ag-6071",
"2019-ag-6072",
"2019-ag-6073",
"2019-ag-6074",
"2019-ag-6075",
"2019-ag-6076",
"2019-ag-6077",
"2019-ag-6078",
"2019-ag-6079",
"2019-ag-6080",
"2019-ag-6081",
"2019-ag-6082",
"2019-ag-6084",
"2019-ag-6085",
"2019-ag-6086",
"2019-ag-6087",
"2019-ag-6088",
"2019-ag-6089",
"2019-ag-6090",
"2019-ag-6091"]
currentDate=date.today()
currentTime=time.localtime()
current_time = time.strftime("%H-%M-%S", currentTime)
username = getpass.getuser()
with open("C:/Users/"+username+"/Desktop/"+str(currentDate)+"_"+str(current_time)+".xlsx","w"):
    pass

wb = openpyxl.load_workbook("C:/Users/"+username+"/Desktop/"+str(currentDate)+"_"+str(current_time)+".xlsx")
ws=wb.active

dataFromFile=[]
chatListDir=''


sg.theme("DarkTeal2")
layout = [[sg.T("")], [sg.Text("Choose a file: "), sg.Input(), sg.FileBrowse(key="-IN-")],[sg.Button("Submit")]]

###Building Window
window = sg.Window('My File Browser', layout, size=(600,150))
    
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event=="Exit":
        break
    elif event == "Submit":
        chatListDir=values["-IN-"]
        break

n=1
with open (chatListDir) as f:
     dataFromFile=np.loadtxt(f,dtype=str,delimiter="\n").tolist()
# for i in agList:
    
#     ws["A"+str(n)]=i
#     n+=1
print(ws["A1"].value)
for i in range(1,40):
    for j in dataFromFile:
           
             if(ws['A'+str(i)].value==j):
                 ws['B'+str(i)]="P"
                 break
             else:
                 ws['B'+str(i)]="A"
                     
wb.save('C:/Users/Muhammad Umer/Desktop/pytry.xlsx')

