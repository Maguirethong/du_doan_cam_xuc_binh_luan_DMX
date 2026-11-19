import pandas as pd
import os

# Lấy danh sách tất cả file Excel
file_list = [r"Comment_cleaned\comments_{}_cleaned.xlsx".format(i) for i in range(1,6)]

# Đọc và gộp các file
df_list = [pd.read_excel(file) for file in file_list]

df_merged = pd.concat(df_list, ignore_index=True)
df_merged["STT"] = df_merged.index + 1
# Xuất thành một file Excel
df_merged.to_excel("comments_merged.xlsx", index=False)
