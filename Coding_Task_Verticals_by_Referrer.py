import os
import pandas as pd

print(os.getcwd())

os.chdir(r"C:\Users\mohsi\OneDrive\Desktop\News_Corp\Datasets")


#Reading Excel File by skipping rows
Verticals_by_Referrer=pd.read_excel('NCA_Verticals_by_Referrer_Device.xlsx'
,skiprows=range(1,10))

#Setting Column header
Verticals_by_Referrer.columns = Verticals_by_Referrer.iloc[0]

#Dropping first row
Verticals_by_Referrer=Verticals_by_Referrer.drop(0)


#renaming duplicate column names

cols = pd.Series(Verticals_by_Referrer.columns)
dup_count = cols.value_counts()
for dup in cols[cols.duplicated()].unique():
    cols[cols[cols == dup].index.values.tolist()] = [dup + str(i) for i in range(1, dup_count[dup]+1)]

Verticals_by_Referrer.columns = cols


Verticals_by_Referrer=Verticals_by_Referrer[["Item1","Item2","Total","Desktop/Tablet","Mobile"]]

#drop NA, delete item from item1 and total from item2

Verticals_by_Referrer=Verticals_by_Referrer.dropna()
Verticals_by_Referrer=Verticals_by_Referrer[~Verticals_by_Referrer['Item1'].str.contains("Item")]
Verticals_by_Referrer=Verticals_by_Referrer[~Verticals_by_Referrer['Item2'].str.contains("Total",na=False)]


# Renaming Columns

Verticals_by_Referrer.rename(columns = {'Item1':'Brand','Item2':'Referrer Type'}, inplace = True)

#Transposing by Device Type for referrer Instances and Final Dataframe


Final_df_Referrer=Verticals_by_Referrer.melt(id_vars=['Brand','Referrer Type'],var_name='Device Type',value_name='Referrer Instances')

#Renaming Brands
Final_df_Referrer=Final_df_Referrer.replace({'Brand':{'AA: p2 Brand - ':'', 'Section - ':'News.com.au-'}}, regex=True)

#Writing Final Table to csv
Final_df_Referrer.to_csv('NCA_Vertical_by_Referrer_Device_Final',encoding='utf-8',index=False)

