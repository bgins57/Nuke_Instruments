
import glob2
import numpy as np
import pandas as pd
from pandas import DataFrame
filenames = glob2.glob("*.txt") #only looks for files ending in txt, in the current directory
systems = ["C51", "B21", "B32"] 
null_values = ["1.000-2","2.000-2","3.000-2","0.000+0", "7.000-2", "6.000-2", "4.000-2", "5.000-2", "8.000-2", "9.000-2", "RESET"]
final_data = DataFrame() #contains everything from a given file
final_data_OPRM = DataFrame() #contains the counts and relative signal data, by looking for "OPRM" in the header
final_data_FLUX = DataFrame() #contains the APRM and LPRM data by looking for FLUX in the header

# now read through the files and create the dataframe
for fname in filenames:

    with open(fname) as fobj:
        lines = fobj.readlines() # reads each line in the file as a string, all strings in one list named "lines"

    data_points = [] #this will get the header information out of the input file, first column will be time
    for line in lines:
        for sys in systems:
            if sys in line:
                data_points.append(line[4:-1]) # data_points now contains a list of header information

    lines_split = []
    for line in lines:
        lines_split.append(line.split()) #this creates a list with each data line an entry in the list
        
# this will delete the lines up to the start of the data, now that we have pulled out the header information
# some files have 10 points, some have 6, so cover both caes

    if ['DATE', 'TIME', '1','2','3','4','5','6', '7', '8', '9','10'] in lines_split:
        first = lines_split.index(['DATE', 'TIME', '1','2','3','4','5','6', '7', '8', '9','10'])
        del lines_split[0:first + 3] #using 3 also cuts out the first line of data that has "DATE" in it
    elif ['DATE', 'TIME', '1','2','3','4','5','6'] in lines_split:
        first = lines_split.index(['DATE', 'TIME', '1','2','3','4','5','6'])
        del lines_split[0:first + 3] #using 3 also cuts out the first line of data that has "DATE" in it which is extraneous.
        
#The next section will delete all the page breaks by looking for the next blank line which occurs in the page breaks
 
    next = 0
    while next < len(lines_split):
        next = lines_split.index([]) #looks for a blank line
        del lines_split[next-1:next+13] #deletes the blanks and header lines between pages

    lines_split_arr = np.array(lines_split) # has all the data

    data_points_arr = np.array(data_points) # has the data points' header information

    time_arr = lines_split_arr[:, 0]# has the time steps since they are in column 0
    data_arr = lines_split_arr[:, 1:len(data_points) + 1] # has the actual data
    
    final_table = DataFrame(data_arr, index = time_arr, columns = data_points_arr)
    
#the following replaces the PPC inserted values for bad sensors with a string zero to avoid errors in the conversion to numeric

    for nu in null_values:
        final_table.replace(nu, "0", inplace=True)
    
# converts the numerical data to float rather than string
    final_table_float = final_table.apply(pd.to_numeric) #Convrts the data to float rather than string
    
#now split the final_table_float into 2 pieces, one for APRM/LPRM and one for OPRMs. Do this by adding columns
#based on the column name
# cycle through the columns names in final_table_float and look for OPRM or FLUX

    for col in final_table_float.columns:
        if "OPRM" in col:
            final_data_OPRM[col] = final_table_float[col]
        elif "FLUX" in col:
            final_data_FLUX[col] = final_table_float[col] 
 
# at this point have 2 dataframes, one for OPRM, one for FLUX
final_data_OPRM.sort_index(axis=1, inplace=True)
final_data_FLUX.sort_index(axis=1, inplace=True)
            


# FLUX noise calculation which is 2 times the normalized standard deviation
    
averages = final_data_FLUX.mean(0) # gets the mean of each column
FLUX_normal = (final_data_FLUX/averages) #the normalized dataframe
FLUX_noise = 100*2*(FLUX_normal.std(0)) # a series with the noise values
    
# the following three lines create a new line in the index that has the calculated values
final_data_FLUX_trans = final_data_FLUX.T
final_data_FLUX_trans["%_noise"]= FLUX_noise.values
final_FLUX = final_data_FLUX_trans.T
    
#now move the noise row to the top by reindexing
val = ["%_noise"]
idx = val + final_FLUX.index.drop(val).tolist()
final_FLUX = final_FLUX.reindex(idx)
  

# Do the same for the OPRMs but only need the max of each column

maxes = final_data_OPRM.max(0) #gets the max of each column
final_data_OPRM_trans = final_data_OPRM.T
final_data_OPRM_trans["Maximum"] = maxes.values
final_OPRM = final_data_OPRM_trans.T
    
val_O = ["Maximum"]
idx_O = val_O +final_OPRM.index.drop(val_O).tolist()
final_OPRM = final_OPRM.reindex(idx_O)
 

#save the files into an EXCEL workbook, sheet 1 is the FLUX and sheet 2 is the OPRMs
fname_result = "TEST 99A and 99C.xlsx"

writer = pd.ExcelWriter(fname_result, engine = 'xlsxwriter')
final_FLUX.to_excel(writer, sheet_name = 'A-LPRM', startrow = 1, header =False)
final_OPRM.to_excel(writer, sheet_name = 'OPRM', startrow = 1, header = False)
   
# format the header row so it wraps text etc
wb = writer.book
ws1 = writer.sheets['A-LPRM']
ws2 = writer.sheets['OPRM']

header_format = wb.add_format({'bold':True, 'text_wrap':True, 'border': 1})
for col_num, value in enumerate(final_FLUX.columns.values):
    ws1.write(0, col_num +1, value, header_format)
for col_num, value in enumerate(final_OPRM.columns.values):
    ws2.write(0, col_num +1, value, header_format)

writer.save()

