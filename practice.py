import pandas as pd
import os
import pytz
from tqdm import tqdm
from datetime import datetime,timedelta
import numpy as np
main_folder = "C:/My workspace/Python/IMD/Data"
df_output = pd.DataFrame()
lt = []
start_date = datetime(2024,4,1,5,30)
end_date = datetime(2025,3,31,5,15)
step = timedelta(minutes = 15)
while start_date <= end_date:
    lt.append(start_date)
    start_date += step


for subfolder in os.listdir(main_folder):
   # print(subfolder)
    subfolder_path = os.path.join(main_folder, subfolder)
    if os.path.isdir(subfolder_path):
        excel_files = [
            f for f in os.listdir(subfolder_path)
            if f.lower().endswith(('.xlsx', '.csv'))
        ]
        dataf = pd.read_csv(os.path.join(subfolder_path, excel_files[0]))
        date_time = [col for col in dataf.columns if "DateTime" in col]
        dataf[date_time[0]] = pd.to_datetime(dataf[date_time[0]])
        formatted_date = datetime.strptime(subfolder, "%Y%m%d").strftime("%Y-%m-%d")
        filtered_df = dataf[dataf[date_time[0]].dt.date == pd.to_datetime(formatted_date).date()]
        filtered_df[date_time[0]] = filtered_df[date_time[0]].dt.tz_localize('GMT').dt.tz_convert('Asia/Kolkata')
        filtered_df[date_time[0]] = filtered_df[date_time[0]].dt.tz_localize(None)
        date_time_col = [col for col in filtered_df.columns if "DateTime" in col]
        date_time_values = filtered_df[date_time_col[0]]
        dataframe = pd.DataFrame({"Date and Time":date_time_values})
        for file in excel_files:
            station = file.split("_", 4)[-1].replace(".csv", "")
            file_path = os.path.join(subfolder_path, file)
            df = pd.read_csv(file_path)
            d_time = [col for col in df.columns if "DateTime" in col]
            df[d_time[0]] = pd.to_datetime(df[d_time[0]])
            filtered_df = df[df[d_time[0]].dt.date == pd.to_datetime(formatted_date).date()]
            for col in filtered_df.columns:
                if "T2" in col:
                    dataframe[f"Temperature_{station}"] = filtered_df[col]
                elif "SWDOWN" in col:
                    dataframe[f"GHI_{station}"] = filtered_df[col]

        df_output = pd.concat([df_output,dataframe],ignore_index=True)

available_dates = sorted(df_output['Date and Time'].to_list())
available_dates_set = set(available_dates)
lt_index = {timestamp: idx for idx, timestamp in enumerate(lt)}


def average(a,b,c):
    return (a+b+c)/3

output_columns = df_output.columns.to_list()
new_stations = output_columns[output_columns.index("Temperature_AAPL_BKN2"):]
df_output.drop(columns = new_stations,inplace = True)


df_output.set_index("Date and Time", inplace=True)

for i in tqdm(lt,desc = "Filling missing timestamps"):
    if i not in df_output.index:
        row = {}
        for col in df_output.columns:
            try:
                val1 = df_output.at[i - timedelta(minutes=96 * 15), col]
                val2 = df_output.at[i - timedelta(minutes=192 * 15), col]
                val3 = df_output.at[i - timedelta(minutes=288 * 15), col]
                row[col] = average(val1, val2, val3)
            except KeyError:
                row[col] = None

        df_output.loc[i] = row

df_output = df_output.sort_index().reset_index()

df_output.set_index('Date and Time', inplace=True)

df_output = df_output.sort_index()
columns = df_output.columns

# Loop over each column
start_time = datetime.strptime("05:30:00", "%H:%M:%S").time()
end_time = datetime.strptime("19:45:00", "%H:%M:%S").time()

for col in tqdm(df_output.columns,desc="Filling columns"):
    series = df_output[col]

    for i in range(len(series)):
        val = series.iloc[i]
        timestamp = series.index[i]
        # If value is 0 or NaN
        if start_time <= timestamp.time() <= end_time and (val == 0 or val < 0 or pd.isna(val)):
            # 1. Try previous 3 non-zero values
            prev_vals = []
            j = i - 1
            while j >= 0 and len(prev_vals) < 3:
                if series.iloc[j] != 0 and series.iloc[j] > 0 and pd.notna(series.iloc[j]):
                    prev_vals.append(series.iloc[j])
                j -= 1

            if len(prev_vals) == 3:
                series.iloc[i] = sum(prev_vals) / 3
                continue  # Skip next check

            # 2. Try next 3 non-zero values
            next_vals = []
            j = i + 1
            while j < len(series) and len(next_vals) < 3:
                if series.iloc[j] != 0 and series.iloc[j] > 0 and pd.notna(series.iloc[j]):
                    next_vals.append(series.iloc[j])
                j += 1

            if len(next_vals) == 3:
                series.iloc[i] = sum(next_vals) / 3

    # Replace modified column back
    df_output[col] = series

df_output = df_output.sort_index().reset_index()

df_output.set_index('Date and Time', inplace=True)

for parameter in tqdm(df_output.columns,desc = "Handling outliers"):
    data = df_output[parameter]
    q1 = np.percentile(data.dropna(),25)
    q3 = np.percentile(data.dropna(),75)
    iqr = q3 - q1
    lower_fence = q1 - (1.5*iqr)
    higher_fence = q3 + (1.5*iqr)
    for i in range(len(data)):
        value = data.iloc[i]
        if lower_fence <= value <= higher_fence:
            continue
        else:
            if value > higher_fence:
                data.iloc[i] = higher_fence
            elif value < lower_fence:
                data.iloc[i] = lower_fence
    df_output[parameter] = data

df_output = df_output.sort_index().reset_index()
start_new = pd.Timestamp("2024-04-01 05:30:00")
end_new = pd.Timestamp("2025-03-30 23:45:00")
df_output = df_output[(df_output['Date and Time'] >= start_new) & (df_output['Date and Time'] <= end_new)]

xls_meter_data = pd.ExcelFile("Compiled_Actual data(till 18 23 MARCH 25)1 WEEKS MISSING (1).xlsx")
meter_data = pd.read_excel(xls_meter_data,sheet_name="NR")
meter_data['Datetime'] = meter_data['date'].dt.date.astype(str) + ' ' + meter_data['Date & Time Block'].str.split('-').str[0] + ':00'
meter_data['Datetime'] = pd.to_datetime(meter_data['Datetime'])

start = pd.Timestamp("2024-04-01 05:30:00")
end = pd.Timestamp("2025-03-30 23:45:00")
filter_df1 = meter_data[(meter_data['Datetime'] >= start) & (meter_data['Datetime'] <= end)]
filter_df1.drop(columns = ['date','Date & Time Block'],inplace = True)

wr_data = pd.read_excel(xls_meter_data,sheet_name="WR")
wr_data['Datetime'] = wr_data['Date '].dt.date.astype(str) + ' ' + wr_data['Time block'].str.split('-').str[0] + ':00'
wr_data['Datetime'] = pd.to_datetime(wr_data['Datetime'])
start = pd.Timestamp("2024-04-01 05:30:00")
end = pd.Timestamp("2025-03-30 23:45:00")
filter_df2 = wr_data[(wr_data['Datetime'] >= start) & (wr_data['Datetime'] <= end)]
filter_df2.drop(columns = ['Date ','Time block'],inplace = True)
filter_df2.set_index('Datetime', inplace=True)

for i in tqdm(lt,desc = "Filling missing timestamps again for meter data"):
    if i not in filter_df2.index:
        row = {}
        for col in filter_df2.columns:
            try:
                val1 = filter_df2.at[i - timedelta(minutes=96 * 15), col]
                val2 = filter_df2.at[i - timedelta(minutes=192 * 15), col]
                val3 = filter_df2.at[i - timedelta(minutes=288 * 15), col]
                row[col] = average(val1, val2, val3)
            except KeyError:
                row[col] = 0

        filter_df2.loc[i] = row

filter_df2 = filter_df2.sort_index().reset_index()


filter_df2.drop(columns = ['Datetime'],inplace = True)

filter_df = pd.concat([filter_df1,filter_df2], axis=1)
starting = pd.Timestamp("2024-04-01 05:30:00")
ending = pd.Timestamp("2025-03-30 23:45:00")
filter_df = filter_df[(filter_df['Datetime'] >= starting) & (filter_df['Datetime'] <= ending)]

#MAPPING
mapped = pd.ExcelFile("mapping_interns.xlsx")
mapped_df = pd.read_excel(mapped,sheet_name="Sheet3")

mapping = {}
for i in range(len(mapped_df)):
    key = mapped_df.loc[i,"ISRO Plants"]
    value = mapped_df.loc[i,"Meter Data"]
    mapping[key] = value

suffixes_to_remove = ['Temperature_','GHI_']
remove_col = []
df_output_cols = set([col.replace(suffix, '') for col in df_output.columns for suffix in suffixes_to_remove if col.startswith(suffix)])
for i in df_output_cols:
    if i not in mapping:
        remove_col.append(f"Temperature_{i}")
        remove_col.append(f"GHI_{i}")
    elif mapping[i] not in filter_df.columns.to_list():
        remove_col.append(f"Temperature_{i}")
        remove_col.append(f"GHI_{i}")

df_output.drop(columns=remove_col,inplace=True)


length = len(df_output.columns) - 1
destination = int(1.5*length) + 1
for i in tqdm(range(1,destination,3),desc="Inserting MW data"):
    s = df_output.columns[i]
    index = s.index('_')
    s_new = s[index+1:]
    values = filter_df[mapping[s_new]].to_list()
    df_output.insert(i,f"MW_{s_new}",values)

df_output.set_index('Date and Time', inplace=True)

for col in tqdm(df_output.columns,desc="Filling MW columns"):
    if col[0:2] == "MW":
        series = df_output[col]
        for i in range(len(series)):
            val = series.iloc[i]
            timestamp = series.index[i]
            # If value is 0 or NaN
            if start_time <= timestamp.time() <= end_time and (val == 0 or val < 0 or pd.isna(val)):
                # 1. Try previous 3 non-zero values
                prev_vals = []
                j = i - 1
                while j >= 0 and len(prev_vals) < 3:
                    if series.iloc[j] != 0 and series.iloc[j] > 0 and pd.notna(series.iloc[j]):
                        prev_vals.append(series.iloc[j])
                    j -= 1

                if len(prev_vals) == 3:
                    series.iloc[i] = sum(prev_vals) / 3
                    continue  # Skip next check

                # 2. Try next 3 non-zero values
                next_vals = []
                j = i + 1
                while j < len(series) and len(next_vals) < 3:
                    if series.iloc[j] != 0 and series.iloc[j] > 0 and pd.notna(series.iloc[j]):
                        next_vals.append(series.iloc[j])
                    j += 1

                if len(next_vals) == 3:
                    series.iloc[i] = sum(next_vals) / 3

        # Replace modified column back
        df_output[col] = series
df_output = df_output.sort_index().reset_index()

df_output.set_index('Date and Time', inplace=True)

for parameter in tqdm(df_output.columns,desc = "Handling outliers"):
    if parameter[0:2] == 'MW':
        data = df_output[parameter]
        q1 = np.percentile(data.dropna(),25)
        q3 = np.percentile(data.dropna(),75)
        iqr = q3 - q1
        lower_fence = q1 - (1.5*iqr)
        higher_fence = q3 + (1.5*iqr)
        for i in range(len(data)):
            value = data.iloc[i]
            if lower_fence <= value <= higher_fence:
                continue
            else:
                if value > higher_fence:
                    data.iloc[i] = higher_fence
                elif value < lower_fence:
                    data.iloc[i] = lower_fence
        df_output[parameter] = data

df_output = df_output.sort_index().reset_index()

df_output.to_excel("output1.xlsx",index=False)






