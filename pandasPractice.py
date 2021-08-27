import pandas as pd
from pandas import read_excel
from pandas import ExcelFile, ExcelWriter

df = pd.read_excel('workshop.xlsx', sheet_name='Power Bi Workshop')
#print(df)

# Printing top and botton rows
# print(df.head(10))
# print(df.tail())

# Display the index columns
# print(df.index)

# getting quick statistics about the data
# print(df.describe())

#Transpose of the table
# print(df.T)

#Slicing th edf
#slice = df[:3] #like python list list
#slice = df[10:150] #using index columns

#selection by lables
#slice = df.loc[4,['First Name'] ]
#print(slice)


