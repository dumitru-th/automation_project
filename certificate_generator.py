# -*- coding: utf-8 -*-
"""
Created on Sat Sep 30 00:18:42 2023

@author: Dumitru
"""
print("hi")

def belt_generator(kup):
    if kup==10:
        return "Alba"
    elif kup==9:
        return "Alba cu Galben"
    elif kup==8:
        return "Galbena"
    elif kup==7:
        return "Galbena cu Verde"
    elif kup==6:
        return "Verde"
    elif kup==5:
        return "Verde cu Albastru"
    elif kup==4:
        return "Albastra"
    elif kup==3:
        return "Albastra cu Rosu"
    elif kup==2:
        return "Rosie"
    elif kup==1:
        return "Rosie cu Negru"
    else:
        return ""
    

import re, os
database={}
with open("exam_data.txt") as file:
    file= file.readlines()
    pattern=re.compile(r"^ROS\d\d\d\d")
    i=0
    for e in file :
        matches=re.match(pattern,e)    
        if matches: 
            athlete_id=e.strip()
            try:
                database[athlete_id]={}
                i+=1
                database[athlete_id]["Last Name"]=file[i].strip().split()[0]
                database[athlete_id]["First Name"]=" ".join(file[i].strip().split()[1:])

                i+=1
                database[athlete_id]["Certificate ID"]=file[i].strip()
                i+=1
                database[athlete_id]["Old Kup Grade"]=file[i].strip()
                i+=1
                database[athlete_id]["New Kup Grade"]=file[i].strip().split()[1]
                    
                database[athlete_id]["Belt"]=belt_generator(int(file[i].strip().split()[1]))
            
                i+=1
            except:  
                pass
      

        
           
from docxtpl import DocxTemplate
import os
def first_name_chk(fn):
    name_to_write =""
    x=len(fn)
    if x>15:
        f_names=fn.split()
        for name in f_names:
           if  len(name_to_write+name)<=15:
                name_to_write+=name+" "
        
    else:
        name_to_write=fn
    return name_to_write    
                
                
def get_exam_code():
    return database[athlete_id]["Certificate ID"][:6]

doc=DocxTemplate("template.docx")

for key in database.keys():
    
        context={
                'n': database[f"{key}"]["New Kup Grade"],
                'last_name1':database[f"{key}"]["Last Name"].upper(),
                'first_name1':first_name_chk(database[f"{key}"]["First Name"].upper()),
                'last_name2':database[f"{key}"]["Last Name"].upper(),
                'first_name2':database[f"{key}"]["First Name"].upper(),
                'belt':database[f"{key}"]["Belt"], 'certificate_code':database[f"{key}"]["Certificate ID"]
                }
        doc.render(context)
        

        # Define the directory where you want to save the generated docx files
        output_dir = f"certificate_create\\{get_exam_code()}"

        # Ensure the output directory exists, if not, create it
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Define the filename for the generated docx file
        filename = f"{database[f'{key}']['Last Name']} {database[f'{key}']['First Name']} kup {database[f'{key}']['New Kup Grade']}.docx"
        # Join the output directory and filename to create the full path
        output_file = os.path.join(output_dir, filename)

        doc.save(output_file)
       
         
