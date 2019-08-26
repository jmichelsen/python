#This script prints out a csv export of the x, y and z axis coordinates tag data
#based on the dxf files in the user's targeted directory
#The values are stored in a dictionary with the filename as the key, and then ultimately migrates it to a dataframe so it
#can be exported to an excel spreadsheet
import os
import easygui as eg
import pandas as pd
import dxfgrabber as dx

dir_dxf = eg.diropenbox(msg='Please open Directory that stores DXF files', title='DXF files')

tag_dict = {}
for file in os.listdir(dir_dxf):
    filepath = os.path.join(dir_dxf, file)
    dxf = dx.readfile(filepath)

    #The tag metadata lives in the entities section of the DXF file
    #and we only care about the type of entity that has 'TEXT' and then 
    #the data that is needed lives in the align_point  field
    for entity in dxf.entities:
        if entity.dxftype == 'TEXT':
            tag_text = entity.text
            coordinates = list(entity.align_point)

            #some of these coordinates have deimal points that
            #enter into the millionth, so rounding is preferred
            for idx,c in enumerate(coordinates):
                coordinates[idx] = round(c)
            
            #removes the .dxf extension from filename since it is a preferred
            #dictionary key
            file_no_ext = os.path.splitext(file)[0]
            tag_dict[file_no_ext] = coordinates

columns = ['Graphic', 'X', 'Y', 'Rotation']
df= pd.DataFrame(columns = columns)

for key, value in tag_dict.items():
  df = df.append({'Graphic':key, 'X' : value[0], 'Y' : value[1],
                        'Rotation': value[2]}, ignore_index = True)

out_dir = eg.diropenbox(msg='Please open Directory that will store output file')
df.to_excel(os.path.join(out_dir,'Tag Data.xlsx'), sheet_name='Sheet1', index = False)
