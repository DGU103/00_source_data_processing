# -*- coding: utf-8 -*-
"""
Created on Mon Aug  5 11:13:58 2024

@author: ara112
"""

""

import glob
import pandas as pd




# Ele Regex----------------------------------------------------------------
           
#path = "D://E3D//Digital-review//*ELE*MCT*.csv"
path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*ELE*-Naming.csv"

          
for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column1 = []  
           new_column2 = []  
           new_column3 = []  
           new_column4 = []  
           new_column5 = []  
           new_column6 = []  
     
           for index, row in df.iterrows():
       
               Tag_number = row['Name']            
               Tag_class = row['Type']
               Tag_Desc = row['Description']
               Tag_Status = ''
               Tag_Action = ''
               E3D_ID = row['Name']
               
               
               
       
               new_column1.append(Tag_number)
               new_column2.append(Tag_class)
               new_column3.append(Tag_Desc)
               new_column4.append(Tag_Status)
               new_column5.append(Tag_Action)
               new_column6.append(E3D_ID)
               
               
               
               
           df['Tag_number'] = new_column1
           df['Tag_number'] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           #df['Tag_number] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           df['Tag_class'] = new_column2
           df['Tag_description'] = new_column3
           df['Tag_description'] = df['Tag_description'].apply(lambda x: x.replace('unset', '',1)[:100])
           df['Action'] = new_column4
           df['Status'] = new_column5
           df['E3D_ID'] = new_column6

           
           fname_validate = fname.replace('.csv', '')
        
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_E3D_Non-Tags.csv', columns=['Tag_number','Tag_class','Tag_description','Action','Status','E3D_ID'],index=False)   
           


# Ins Regex----------------------------------------------------------------
           
#path = "D://E3D//Digital-review//*ELE*MCT*.csv"
path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*INS*-Naming.csv"

          
for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column1 = []  
           new_column2 = []  
           new_column3 = []  
           new_column4 = []  
           new_column5 = []  
           new_column6 = []  
     
           for index, row in df.iterrows():
       
               Tag_number = row['Name']            
               Tag_class = row['Type']
               Tag_Desc = row['Description']
               Tag_Status = ''
               Tag_Action = ''
               E3D_ID = row['Name']
               
               
               
       
               new_column1.append(Tag_number)
               new_column2.append(Tag_class)
               new_column3.append(Tag_Desc)
               new_column4.append(Tag_Status)
               new_column5.append(Tag_Action)
               new_column6.append(E3D_ID)
               
               
               
               
           df['Tag_number'] = new_column1
           df['Tag_number'] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           #df['Tag_number] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           df['Tag_class'] = new_column2
           df['Tag_description'] = new_column3
           df['Tag_description'] = df['Tag_description'].apply(lambda x: x.replace('unset', '',1)[:100])
           df['Action'] = new_column4
           df['Status'] = new_column5
           df['E3D_ID'] = new_column6

           
           fname_validate = fname.replace('.csv', '')
        
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_E3D_Non-Tags.csv', columns=['Tag_number','Tag_class','Tag_description','Action','Status','E3D_ID'],index=False)   
           

# Tel Regex----------------------------------------------------------------
           
#path = "D://E3D//Digital-review//*ELE*MCT*.csv"
path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*TEL*-Naming.csv"

          
for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column1 = []  
           new_column2 = []  
           new_column3 = []  
           new_column4 = []  
           new_column5 = []  
           new_column6 = []  
     
           for index, row in df.iterrows():
       
               Tag_number = row['Name']            
               Tag_class = row['Type']
               Tag_Desc = row['Description']
               Tag_Status = ''
               Tag_Action = ''
               E3D_ID = row['Name']
               
               
               
       
               new_column1.append(Tag_number)
               new_column2.append(Tag_class)
               new_column3.append(Tag_Desc)
               new_column4.append(Tag_Status)
               new_column5.append(Tag_Action)
               new_column6.append(E3D_ID)
               
               
               
               
           df['Tag_number'] = new_column1
           df['Tag_number'] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           #df['Tag_number] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           df['Tag_class'] = new_column2
           df['Tag_description'] = new_column3
           df['Tag_description'] = df['Tag_description'].apply(lambda x: x.replace('unset', '',1)[:100])
           df['Action'] = new_column4
           df['Status'] = new_column5
           df['E3D_ID'] = new_column6

           
           fname_validate = fname.replace('.csv', '')
        
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_E3D_Non-Tags.csv', columns=['Tag_number','Tag_class','Tag_description','Action','Status','E3D_ID'],index=False)   
           


# Equi Trim Regex----------------------------------------------------------------
           
#path = "D://E3D//Digital-review//*ELE*MCT*.csv"
path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*EQUI*Trim*-Naming.csv"

          
for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column1 = []  
           new_column2 = []  
           new_column3 = []  
           new_column4 = []  
           new_column5 = []  
           new_column6 = []  
     
           for index, row in df.iterrows():
       
               Tag_number = row['Name']            
               Tag_class = row['Type']
               Tag_Desc = row['Description']
               Tag_Status = ''
               Tag_Action = ''
               E3D_ID = row['Name']
               
               
               
       
               new_column1.append(Tag_number)
               new_column2.append(Tag_class)
               new_column3.append(Tag_Desc)
               new_column4.append(Tag_Status)
               new_column5.append(Tag_Action)
               new_column6.append(E3D_ID)
               
               
               
               
           df['Tag_number'] = new_column1
           df['Tag_number'] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           #df['Tag_number] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           df['Tag_class'] = new_column2
           df['Tag_description'] = new_column3
           df['Tag_description'] = df['Tag_description'].apply(lambda x: x.replace('unset', '',1)[:100])
           df['Action'] = new_column4
           df['Status'] = new_column5
           df['E3D_ID'] = new_column6

           
           fname_validate = fname.replace('.csv', '')
        
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_E3D_Non-Tags.csv', columns=['Tag_number','Tag_class','Tag_description','Action','Status','E3D_ID'],index=False)   
           

# Pipe Supp Regex----------------------------------------------------------------
           
#path = "D://E3D//Digital-review//*ELE*MCT*.csv"
path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*Pipe*Supp*-Naming.csv"

          
for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column1 = []  
           new_column2 = []  
           new_column3 = []  
           new_column4 = []  
           new_column5 = []  
           new_column6 = []  
     
           for index, row in df.iterrows():
       
               Tag_number = row['Name']            
               Tag_class = row['Type']
               Tag_Desc = row['Description']
               Tag_Status = ''
               Tag_Action = ''
               E3D_ID = row['Name']
               
               
               
       
               new_column1.append(Tag_number)
               new_column2.append(Tag_class)
               new_column3.append(Tag_Desc)
               new_column4.append(Tag_Status)
               new_column5.append(Tag_Action)
               new_column6.append(E3D_ID)
               
               
               
               
           df['Tag_number'] = new_column1
           df['Tag_number'] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           #df['Tag_number] = df['Tag_number'].apply(lambda x: x.replace('/', '',1)[:100])
           df['Tag_class'] = new_column2
           df['Tag_description'] = new_column3
           df['Tag_description'] = df['Tag_description'].apply(lambda x: x.replace('unset', '',1)[:100])
           df['Action'] = new_column4
           df['Status'] = new_column5
           df['E3D_ID'] = new_column6

           
           fname_validate = fname.replace('.csv', '')
        
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_E3D_Non-Tags.csv', columns=['Tag_number','Tag_class','Tag_description','Action','Status','E3D_ID'],index=False)   
           




print("All files are validated successfully.!!!")
