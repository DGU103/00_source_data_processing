# -*- coding: utf-8 -*-
"""
Created on Sun Jul 21 11:42:01 2024

@author: ara112
"""

# -*- coding: utf-8 -*-
"""
Created on Sun Jul 21 07:53:42 2024

@author: ara112
"""

import glob
import pandas as pd
import re
import os
import zipfile


#zip all files to archive----------------------------------------------

from datetime import datetime
zdate = datetime.today().strftime('%Y-%m-%d')

#print(zdate)


PATH = '//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/Archive'

if not os.path.exists(PATH):
    os.makedirs(PATH)
    


pathfilename_type = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*.csv"
filename=[]

for fname in glob.glob(pathfilename_type):
    filename.append(fname)  
    
    
#print(filename)    

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/Archive/" + zdate + ".zip"

with zipfile.ZipFile("//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/Archive/" + zdate + ".zip", mode="w") as archive:    
     for filen in filename:
         archive.write(filen)
         

# Delete all Validation files------------------------------------------
path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*Validation*.csv"
for fname in glob.glob(path):
    #print(fname)
    os.remove(fname)
    #print('file deleted')

# DF for reading regex config data

config = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Configs/Non_Tagged_Items\Config.csv"

df_config=pd.read_csv(config, sep=";")
##print(df_config)






# Ele MCT Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*ELE*MCT*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ele_MCT_rege = re.match(regex_template,'Ele_MCT_regex')
    if Ele_MCT_rege is None:
        pass
           
    else:
        #print(Ele_MCT_rege)
        Ele_MCT_regex = regex_value
        #print(Ele_MCT_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ele_MCT_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


    
    
# Ins MCT Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*INS*MCT*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ins_MCT_rege = re.match(regex_template,'Ins_MCT_regex')
    if Ins_MCT_rege is None:
        pass
           
    else:
        #print(Ins_MCT_rege)
        Ins_MCT_regex = regex_value
        #print(Ins_MCT_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ins_MCT_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Tel MCT Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*TEL*MCT*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Tel_MCT_rege = re.match(regex_template,'Tel_MCT_regex')
    if Tel_MCT_rege is None:
        pass
           
    else:
        #print(Tel_MCT_rege)
        Tel_MCT_regex = regex_value
        #print(Tel_MCT_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Tel_MCT_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)

  

# Ele_Tray Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*ELE*TRAY*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ele_Tray_rege = re.match(regex_template,'Ele_Tray_regex')
    if Ele_Tray_rege is None:
        pass
           
    else:
        #print(Ele_Tray_rege)
        Ele_Tray_regex = regex_value
        #print(Ele_Tray_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ele_Tray_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Ele_Supp Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*ELE*SUPP*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ele_Supp_rege = re.match(regex_template,'Ele_Supp_regex')
    if Ele_Supp_rege is None:
        pass
           
    else:
        #print(Ele_Supp_rege)
        Ele_Supp_regex = regex_value
        #print(Ele_Supp_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ele_Supp_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Tel_Supp Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*TEL*SUPP*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Tel_Supp_rege = re.match(regex_template,'Tel_Supp_regex')
    if Tel_Supp_rege is None:
        pass
           
    else:
        #print(Tel_Supp_rege)
        Tel_Supp_regex = regex_value
        #print(Tel_Supp_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Tel_Supp_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Ins_Supp Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*INS*SUPP*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ins_Supp_rege = re.match(regex_template,'Ins_Supp_regex')
    if Ins_Supp_rege is None:
        pass
           
    else:
        #print(Ins_Supp_rege)
        Ins_Supp_regex = regex_value
        #print(Ins_Supp_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ins_Supp_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)



# Ele_Earth Bar Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*ELE*Earth*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ele_Earth_rege = re.match(regex_template,'Ele_Earth_regex')
    if Ele_Earth_rege is None:
        pass
           
    else:
        #print(Ele_Earth_rege)
        Ele_Earth_regex = regex_value
        #print(Ele_Earth_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ele_Earth_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Ins_Tray Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*INS*TRAY*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Ins_Tray_rege = re.match(regex_template,'Ins_Tray_regex')
    if Ins_Tray_rege is None:
        pass
           
    else:
        #print(Ins_Tray_rege)
        Ins_Tray_regex = regex_value
        #print(Ins_Tray_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Ins_Tray_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Tel_Tray Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*TEL*TRAY*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Tel_Tray_rege = re.match(regex_template,'Tel_Tray_regex')
    if Tel_Tray_rege is None:
        pass
           
    else:
        #print(Tel_Tray_rege)
        Tel_Tray_regex = regex_value
        #print(Tel_Tray_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Tel_Tray_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Pipe_Supp Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*PIPE*SUPP*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    Pipe_Supp_rege = re.match(regex_template,'Pipe_Supp_regex')
    if Pipe_Supp_rege is None:
        pass
           
    else:
        #print(Pipe_Supp_rege)
        Pipe_Supp_regex = regex_value
        #print(Pipe_Supp_regex)
        
        for fname in glob.glob(path):
           #print(fname)
           
           df=pd.read_csv(fname, sep=";")
           #print(df)
           
           new_column = []  
       
           for index, row in df.iterrows():
       
               Tag_No = row['Name']
       
               Tag_No_Validate = re.fullmatch(Pipe_Supp_regex, Tag_No)
               if Tag_No_Validate is None:
                   new_column.append('Invalid')
               else:
                   new_column.append('Valid')     
           df['Name Validation'] = new_column

           
           fname_validate = fname.replace('.csv', '')
           df.to_csv(fname_validate + '_validation_report.csv', index=False)


# Trim Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*TRIM*Naming.csv"


for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']

    Trim_line_rege = re.match(regex_template,'Trim_Line_regex')

    if Trim_line_rege is None:
        pass
           
    else:
        Trim_line_regex = regex_value
        
        for index, row in df_config.iterrows():
            
            regex_template = row['TemplateID']
            regex_value = row['regex']
            
            
            
            Pipe_Tag_rege = re.match(regex_template,'Pipe_Tag_regex')
            
            if Pipe_Tag_rege is None:
                pass
                   
            else:
                #print(Trim_Line_rege)
                Pipe_Tag_regex = regex_value
                #print(Trim_Line_regex)
                
                for fname in glob.glob(path):
                   #print(fname)
                   
                   df=pd.read_csv(fname, sep=";")
                   #print(df)
                   
                   new_column = []  
               
                   for index, row in df.iterrows():
               
                       Tag_No = row['Name']
               
                       Tag_No_Validate = re.fullmatch(Pipe_Tag_regex, Tag_No)
                       if Tag_No_Validate is None:
                           Tag_No_Validate2 = re.fullmatch(Trim_line_regex, Tag_No)
                           if Tag_No_Validate2 is None:
                               new_column.append('Invalid')
                           else:
                               new_column.append('Valid')
                       else:
                           new_column.append('Valid Pipe Naming')     
                   df['Name Validation'] = new_column
                   df = df.drop(df[df['Name Validation'] == 'Valid Pipe Naming'].index)

                   fname_validate = fname.replace('.csv', '')
                   df.to_csv(fname_validate + '_validation_report.csv', index=False)
                   


# Catalog Regex----------------------------------------------------------------

path = "//als.local/NOC/Data/Appli/DigitalAsset/MP/RUYA_data/Source/E3D/Non-Tagged/*CATALOG*Naming.csv"

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    SPCO_rege = re.match(regex_template,'SPCO_regex')
	
    if SPCO_rege is None:
        pass
           
    else:
        #print(SPCO_rege)
        SPCO_regex = regex_value
        #print(SPCO_regex)
        	
	
for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    CATREF_rege = re.match(regex_template,'CATREF_regex')
	
    if CATREF_rege is None:
        pass
           
    else:
        #print(CATREF_rege)
        CATREF_regex = regex_value
        #print(CATREF_regex)
 	
for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    CATE_rege = re.match(regex_template,'CATE_regex')
	
    if CATE_rege is None:
        pass
           
    else:
        #print(CATE_rege)
        CATE_regex = regex_value
        #print(CATE_regex)
	
	
for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    SDTE_rege = re.match(regex_template,'SDTE_regex')
	
    if SDTE_rege is None:
        pass
           
    else:
        #print(SDTE_rege)
        SDTE_regex = regex_value
        #print(SDTE_regex)


for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    MATXT_rege = re.match(regex_template,'MATXT_regex')
	
    if MATXT_rege is None:
        pass
           
    else:
        #print(MATXT_rege)
        MATXT_regex = regex_value
        #print(MATXT_regex)

for index, row in df_config.iterrows():
    
    regex_template = row['TemplateID']
    regex_value = row['regex']
    CMPREF_rege = re.match(regex_template,'CMPREF_regex')
	
    if CMPREF_rege is None:
        pass
           
    else:
        #print(CMPREF_rege)
        CMPREF_regex = regex_value
        #print(CMPREF_regex)

for fname in glob.glob(path):
    #print(fname)
    
    df=pd.read_csv(fname, sep=";")
    ##print(df)
    
    SPCO_name_column = []  
    CATREF_name_column = []  
    CATE_name_column = []  
    SDTE_name_column = []  
    MATXT_name_column = []  
    CMPREF_name_column = []  


    for index, row in df.iterrows():

        SPCO_name = row['SPCO']
        CATREF_name = row['CATREF']
        CATE_name = row['CATE']
        SDTE_name = row['SDTE']
        MATXT_name = row['MATXT']
        CMPREF_name = row['CMPREF']

        #print (SPCO_name)
        #print (type(SPCO_name))
        
        if type(SPCO_name) == str:         
            SPCO_name_Validate = re.fullmatch(SPCO_regex, SPCO_name)
            if SPCO_name_Validate is None:
                SPCO_name_column.append('Invalid')
            else:
                SPCO_name_column.append('Valid') 
        else: SPCO_name_column.append('Invalid')

        if type(CATREF_name) == str:         
            CATREF_name_Validate = re.fullmatch(CATREF_regex, CATREF_name)
            if CATREF_name_Validate is None:
                CATREF_name_column.append('Invalid')
            else:
                CATREF_name_column.append('Valid') 
        else: CATREF_name_column.append('Invalid')

        if type(CATE_name) == str:         
            CATE_name_Validate = re.fullmatch(CATE_regex, CATE_name)
            if CATE_name_Validate is None:
                CATE_name_column.append('Invalid')
            else:
                CATE_name_column.append('Valid') 
        else: CATE_name_column.append('Invalid')

        if type(SDTE_name) == str:         
            SDTE_name_Validate = re.fullmatch(SDTE_regex, SDTE_name)
            if SDTE_name_Validate is None:
                SDTE_name_column.append('Invalid')
            else:
                SDTE_name_column.append('Valid') 
        else: SDTE_name_column.append('Invalid')

        if type(MATXT_name) == str:         
            MATXT_name_Validate = re.fullmatch(MATXT_regex, MATXT_name)
            if MATXT_name_Validate is None:
                MATXT_name_column.append('Invalid')
            else:
                MATXT_name_column.append('Valid') 
        else: MATXT_name_column.append('Invalid')

        if type(CMPREF_name) == str:         
            CMPREF_name_Validate = re.fullmatch(CMPREF_regex, CMPREF_name)
            if CMPREF_name_Validate is None:
                CMPREF_name_column.append('Invalid')
            else:
                CMPREF_name_column.append('Valid') 
        else: CMPREF_name_column.append('Invalid')
                
 
    
    df['SPCO Validation'] = SPCO_name_column
    df['CATREF Validation'] = CATREF_name_column
    df['CATE Validation'] = CATE_name_column
    df['SDTE Validation'] = SDTE_name_column
    df['MATXT Validation'] = MATXT_name_column
    df['CMPREF Validation'] = CMPREF_name_column
    #print (df)
    
    
    
    
    fname_validate = fname.replace('.csv', '')
    df.to_csv(fname_validate + '_validation_report.csv', index=False)

print("All files are validated successfully.!!!")