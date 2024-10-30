#create a python program can read excel file that contains information of studants and generate a csv file to add user to google admin
import os

import pandas as pd
#add folder path that  contains excel file
ir = "path"


l=os.listdir(ir)

from glob import glob

# get the absolute paths of all Excel files 
all_excel_files = glob(ir+"*.xlsx")
#prepare csv column names
"""Prof Mail """
listFinlMail = pd.DataFrame(columns=['First Name','Last Name','Email Address','Password','Password Hash Function','Org Unit Path','New Primary Email','Recovery Email','Home Secondary Email','Work Secondary Email','Recovery Phone','Work Phone','Home Phone','Mobile Phone','Work Address','Home Address','Employee ID','Employee Type','Employee Title','Manager Email','Department','Cost Center','Building ID','Floor Name','Floor Section','Change Password at Next Sign-In','New Status','Advanced Protection Program enrollment'])
#read excel  files one by one and get data 
for i in all_excel_files:
    filei = pd.read_excel(i)
    flname = i.split('/')[-1]
    listSp = pd.DataFrame(columns=['us','Last Name','Email Address','Password','Password Hash Function','Org Unit Path','New Primary Email','Recovery Email','Home Secondary Email','Work Secondary Email','Recovery Phone','Work Phone','Home Phone','Mobile Phone','Work Address','Home Address','Employee ID','Employee Type','Employee Title','Manager Email','Department','Cost Center','Building ID','Floor Name','Floor Section','Change Password at Next Sign-In','New Status','Advanced Protection Program enrollment'])
    sub = filei[["Date de naiss.","Mat. Etudiant","Prénom","Nom"]]    
    listSp['First Name'] = sub['Prénom'].str.capitalize() 
    listSp['Last Name'] = sub['Nom'].str.upper()
    sp=flname[1:-5].lower()
    listSp['Email Address'] = sub['Nom'].str.lower().str.replace(' ', '')+'.'+sub['Prénom'].str.lower().str.replace(' ', '')+'.'+sp+'@univ-ghardaia.dz'
    listSp['Org Unit Path'] = "/"
    listSp['Department'] = filei['Libellé filière']
    listSp['Password'] = 'Gh'+sub['Date de naiss.']
    listFinlMail=pd.concat([listFinlMail,listSp])
#save output csv file     
listFinlMail.to_csv(ir+"allMail.csv",index=False,sep=';',encoding='utf-8')

