
# %% [markdown]
# #### Example sales report summary

# %%
import datetime
import pandas as pd
import os

path = r"D:\My Documents\Desktop\python_office_11_OCT_2023\src\data\sales_data"
today = datetime.date.today()
year = today.year-1

xlxs_file_lists = []

for root,dirs,files in os.walk(path):
      for name in files:
        file_path = os.path.join(root,name)
        if file_path.split("\\")[-2] == str(year): #change filter
            xlxs_file_lists.append(file_path)
xlxs_file_lists

# %%
df_lists = []

for f in xlxs_file_lists:
    df = pd.read_excel(f)
    df_lists.append(df)
df_lists

# %%
df_summary = pd.concat(df_lists)
df_summary

# %%
pivot = pd.pivot_table(df_summary,index="transaction_date",columns="store",values="amount",aggfunc="sum")
pivot

# %%
summary_monthly = pivot.resample("M").sum()
summary_monthly

# %%
import matplotlib
fig = summary_monthly.plot(kind="bar",figsize=(20,12),fontsize=26,title="monthly sale summary").get_figure()

# %%
import xlwings as xw

import datetime
now = datetime.datetime.now()
date_file_name = f'{str(now.date())}_{str(now.time()).split(".")[0].replace(":","_")}'


template = xw.Book(r"D:\My Documents\Desktop\python_office_11_OCT_2023\src\data\sale_template.xlsx")

app = xw.apps.active
sheet = template.sheets["summary"]
sheet["A1"].value = summary_monthly

pivot_page = template.sheets["pivot"]
pivot_page["A1"].value = pivot

#add picture
sheet_report = template.sheets["report"]
sheet_report["A1"].value = "Summary by month"
sheet_report['A1'].font.size = 24
sheet_report["A1"].api.Font.Bold = True
plot= sheet_report.pictures.add(fig,top=sheet["A3"].top,left=sheet["A3"].left)
plot.width = plot.width*0.8
plot.height = plot.height*0.8

template.save(f"D:/My Documents/Desktop/python_office_11_OCT_2023/src/export/summary_sale_report_{date_file_name}.xlsx")
template.close()
app.kill()