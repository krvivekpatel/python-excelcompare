import pandas as pd 
  
#Reading two Excel Sheets 

file1= 'C:/Vivek/pre.xlsx' 
file2= 'C:/Vivek/post.xlsx'
excel1 = pd.read_excel(file1,sheet_name=0, names=['Primary', 'Date1'], skiprows=0) 
excel2 = pd.read_excel(file2,sheet_name=0,  names=['Primary', 'Date1'], skiprows=0)
# print (excel1)
# print (excel2)
#DID = pd.read_excel(file1, sheet_name=0, header=None, usecols=[0, 1, 6], names=['Primary', 'Date1'], skiprows=0)
# chest n cold
# For example...
# usecols => read only specific col indexes
# dtype => specifying the data types
# skiprows => skip number of rows from the top.

# df = pd.DataFrame(excel1, index = idx)
# excel1.sort_index(by=["Primary"])
# excel1=excel1.reindex()
# excel2.sort_index(by=["Primary"])
# excel2=excel2.reindex()  
# Iterating the Columns Names of both Sheets 
for i,j in zip(excel1,excel2): 
     
    # Creating empty lists to append the columns values     
    a,b =[],[] 
  
    # Iterating the columns values 
    for m, n in zip(excel1[i],excel2[j]): 
  
        # Appending values in lists 
        a.append(m) 
        b.append(n) 
  
    # Sorting the lists 
    #a.sort() 
    #b.sort() 
  
    # Iterating the list's values and comparing them 
    for m, n in zip(range(len(a)), range(len(b))): 
        if a[m] != b[n]: 
            print('Column name : \'{}\' and Row Number : {}'.format(i,m)) 