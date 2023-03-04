import os
import pandas as pd
from pandas import datetime as dt
import matplotlib.pyplot as plt


print(os.getcwd())

os.chdir(r"C:\Users\mohsi\OneDrive\Desktop\News_Corp\Datasets")


#Reading Excel File by skipping rows, there are lot of rows with unstructured data
Verticals_by_Month=pd.read_excel('NCA_Verticals_by_Month_Device.xlsx'
,skiprows=range(1,10))

#Setting Column header
Verticals_by_Month.columns = Verticals_by_Month.iloc[0]

#Dropping first row
Verticals_by_Month=Verticals_by_Month.drop(0)


#renaming duplicate column names
#1 for Page View and 2 for Unique Visitors
cols = pd.Series(Verticals_by_Month.columns)
dup_count = cols.value_counts()
for dup in cols[cols.duplicated()].unique():
    cols[cols[cols == dup].index.values.tolist()] = [dup + str(i) for i in range(1, dup_count[dup]+1)]

Verticals_by_Month.columns = cols

#Selecting only the required Columns
Verticals_by_Month=Verticals_by_Month[["Item1","Item2","Total1","Desktop/Tablet1","Mobile1","Total2","Desktop/Tablet2","Mobile2"]]



#cleaning
#----------drop NA, delete item from item1 and Total from item2

Verticals_by_Month=Verticals_by_Month.dropna()
Verticals_by_Month=Verticals_by_Month[~Verticals_by_Month['Item1'].str.contains("Item")]
Verticals_by_Month=Verticals_by_Month[~Verticals_by_Month['Item2'].str.contains("Total",na=False)]


# Renaming Columns

Verticals_by_Month.rename(columns = {'Item1':'Brand','Item2':'Months'}, inplace = True)


#separate dataframe for page views and unique views
Verticals_by_Month_unique_view=Verticals_by_Month[['Brand','Months','Total2','Desktop/Tablet2','Mobile2']]
Verticals_by_Month_page_view=Verticals_by_Month[['Brand','Months','Total1','Desktop/Tablet1','Mobile1']]

#Transposing the dataframe for Device Type Unique Visitors
Verticals_by_Month_unique_view=Verticals_by_Month_unique_view.melt(id_vars=['Brand','Months'],var_name='Device Type',value_name='Unique Visitors')
#Transposing the dataframe for Device Type Page Views
Verticals_by_Month_page_view=Verticals_by_Month_page_view.melt(id_vars=['Brand','Months'],var_name='Device Type',value_name='Page Views')

#Joining both dataframes together
df_final_Verticals_by_Month=Verticals_by_Month_unique_view.join(Verticals_by_Month_page_view['Page Views'])

#Removing un necessary texts from rows and Creating Final Dataframe
df_final_Verticals_by_Month['Device Type'] = df_final_Verticals_by_Month['Device Type'].str.replace('2', '')


#Renaming Brand names
df_final_Verticals_by_Month = df_final_Verticals_by_Month.replace({'Brand':{'AA: p2 Brand - ':'', 'Section - ':'News.com.au-'}}, regex=True)

#Writing Final Table to csv
df_final_Verticals_by_Month.to_csv('NCA_Vertical_by_Month_Final',encoding='utf-8',index=False)


#----------Visualization monthly trend by Brand

df_plot=df_final_Verticals_by_Month


df_plot=df_plot[df_plot['Brand'].str.contains("Newscomau")]

plt.ticklabel_format(style='plain')

plt.ticklabel_format(useOffset=False)

(df_plot.pivot_table(index='Months', columns='Brand', values='Page Views')
   .resample('M').sum().plot())

(df_plot.pivot_table(index='Months', columns='Brand', values='Unique Visitors')
   .resample('M').sum().plot())


