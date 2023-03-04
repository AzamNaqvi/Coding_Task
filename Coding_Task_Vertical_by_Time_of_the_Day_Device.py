import os
import pandas as pd

print(os.getcwd())

os.chdir(r"C:\Users\mohsi\OneDrive\Desktop\News_Corp\Datasets")

#Loading multiple sheets in one dataframe
Verticals_by_Day_Device = pd.concat(pd.read_excel('NCA_Verticals_by_Time_of_Day_Device.xlsx', sheet_name=None,skiprows=range(1,10)), ignore_index=True)

#Setting Column header
Verticals_by_Day_Device.columns = Verticals_by_Day_Device.iloc[0]

#Dropping first row
Verticals_by_Day_Device=Verticals_by_Day_Device.drop(0)

#Removing duplicate column header by adding suffix
cols = pd.Series(Verticals_by_Day_Device.columns)
dup_count = cols.value_counts()
for dup in cols[cols.duplicated()].unique():
    cols[cols[cols == dup].index.values.tolist()] = [dup + str(i) for i in range(1, dup_count[dup]+1)]

Verticals_by_Day_Device.columns = cols

#selecting required columns
Verticals_by_Day_Device=Verticals_by_Day_Device[["Item1","Item2","Total1","Desktop/Tablet1","Mobile1","Total2","Desktop/Tablet2","Mobile2"]]

#Removing text rows and NA
Verticals_by_Day_Device=Verticals_by_Day_Device[~Verticals_by_Day_Device['Item1'].str.contains("Item")]
Verticals_by_Day_Device=Verticals_by_Day_Device[~Verticals_by_Day_Device['Item2'].str.contains("Total",na=False)]

# Renaming Columns

Verticals_by_Day_Device.rename(columns = {'Item1':'Brand','Item2':'Hour of the Day'}, inplace = True)

#separate dataframe for page views and unique views
Verticals_by_Day_Device_unique_view=Verticals_by_Day_Device[['Brand','Hour of the Day','Total2','Desktop/Tablet2','Mobile2']]
Verticals_by_Day_Device_page_view=Verticals_by_Day_Device[['Brand','Hour of the Day','Total1','Desktop/Tablet1','Mobile1']]

#Transposing the dataframe for Device Type Unique Visitors
Verticals_by_Day_Device_unique_view=Verticals_by_Day_Device_unique_view.melt(id_vars=['Brand','Hour of the Day'],var_name='Device Type',value_name='Unique Visitors')
#Transposing the dataframe for Device Type Page Views
Verticals_by_Day_Device_page_view=Verticals_by_Day_Device_page_view.melt(id_vars=['Brand','Hour of the Day'],var_name='Device Type',value_name='Page Views')


#Joining two dataframes together
df_final_Day_Device=Verticals_by_Day_Device_unique_view.join(Verticals_by_Day_Device_page_view['Page Views'])

#Renaming Brand
df_final_Day_Device=df_final_Day_Device.replace({'Brand':{'AA: p2 Brand - ':'', 'Section - ':'News.com.au-'}}, regex=True)



#Final Dataframe
df_final_Day_Device['Device Type'] = df_final_Day_Device['Device Type'].str.replace('2', '')







