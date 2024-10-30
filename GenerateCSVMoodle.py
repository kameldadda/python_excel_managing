#create a python program can read excel file that contains information of studants and generate a csv file to add it to moodle
import os

import pandas as pd

# add folder path that  contains excel file
diractory="path"

# get the absolute paths of all Excel files 
from glob import glob
all_excel_files = glob(diractory+"*.xlsx")

#prepare csv column names
'''Moodle'''
listFinMoodle=pd.DataFrame(columns=['username','password','firstname','lastname','email','cohort1'])
#read excel  files one by one and get data 
for i in all_excel_files:
    filei = pd.read_excel(i)
    sp = i.split('/')[-1][1:-5].lower()
    listElearnSp = pd.DataFrame(columns=['username','password','firstname','lastname','email','cohort1'])
    sub = filei[["Date de naiss.","Mat. Etudiant","Prénom","Nom"]]
    listElearnSp['username']=sub["Mat. Etudiant"]
    listElearnSp['password'] = 'Gh'+sub['Date de naiss.']
    listElearnSp['firstname']=sub["Prénom"].str.capitalize() 
    listElearnSp['lastname']=sub["Nom"].str.upper()    
    listElearnSp['email'] = sub['Nom'].str.lower().str.replace(' ', '')+'.'+sub['Prénom'].str.lower().str.replace(' ', '')+'@univ-ghardaia.dz'
    listElearnSp['cohort1'] = filei['Code niveau']+'_'+filei['Libellé filière']+'_2425'
    listFinMoodle=pd.concat([listFinMoodle,listElearnSp])

#save output csv file 
listFinMoodle.to_csv(diractory+"MoodleALL1st.csv",index=False,sep=';',encoding='utf-8')
