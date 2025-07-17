import pandas as pd
import os

excel_conversion_folder = "D:/SLDC Gujarat Web Scraping + Excel Conversion/excel_conversion"
combined_df = []

for file in os.listdir(excel_conversion_folder):
    if file.endswith(".xlsx"):
        print("Loading File {0}...".format(file))
        combined_df.append(pd.read_excel(os.path.join(excel_conversion_folder,file),sheet_name="Wind Energy"))
        combined_df.sort()

master_df = pd.concat(combined_df,axis=0).to_excel("masterfile.xlsx",index=False)