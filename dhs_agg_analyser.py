import pandas as pd
import os
#import xlrd


def month(row):
	a = str(row['PostDate'])
	m = a[:-6]
	return m

filepath = f"C:/Users/Varna/Documents/datasets/DHS Data/"

file_dir = os.listdir(filepath)

print(file_dir)

df_data_aggregate = pd.DataFrame()

for i in file_dir:
	
	filename = f"{filepath}{i}"
	print(filename)
	file = pd.ExcelFile(filename)
	print(file.sheet_names)
	for j in file.sheet_names:
		temp_df = pd.DataFrame()
		temp_df = pd.read_excel(filename,sheet_name=j)
		temp_df['dept'] = j
		temp_df['month'] = temp_df.apply(lambda row:month(row),axis=1)
		print(temp_df.shape)
		df_data_aggregate = df_data_aggregate.append(temp_df)
		print(df_data_aggregate.shape)


print(df_data_aggregate.info())

#df_aggregate = df_data_aggregate.groupby(['dept','Merchant','MCC Description','Merchant State/Province','month']).agg({'TransactionID':'count','Transaction Amount':'sum'}).reset_index()

df_aggregate = df_data_aggregate.groupby(['dept','Merchant','MCC Description']).agg({'TransactionID':'count','Transaction Amount':'sum'}).reset_index()


print(df_aggregate.info())

print(df_aggregate.head(20))


output_filepath = f"{filepath}aggregated_consoldiated.xlsx"
print(output_filepath)

writer = pd.ExcelWriter(output_filepath)

df_aggregate.to_excel(writer,"aggregated_numbers", index=False)

writer.save()