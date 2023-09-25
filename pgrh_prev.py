import pandas as pd
import os

dir_path = os.path.dirname(os.path.realpath(__file__))
data_dir = os.path.join(dir_path, 'Input')
data_dir_output = os.path.join(dir_path, 'Output')

# Specify the name of the Excel file
file_name1= "PG_RH_data.xlsx"
file_name3= "Mains_Failure (OSS).xlsx"
file_name4= "Ext. Alarm Huawei (EMS Alarm).xlsx"
file_name6= "DC_Low (OSS).xlsx"
file_name8= "Grid Fail Alarm(RMS).xlsx"
file_name9= "Site _down (OSS).xlsx"
file_name10= "RMS_DC Low Alarm.xlsx"

file_path1 = os.path.join(data_dir, file_name1)
file_path3 = os.path.join(data_dir, file_name3)
file_path4 = os.path.join(data_dir, file_name4)
file_path6 = os.path.join(data_dir, file_name6)
file_path8 = os.path.join(data_dir, file_name8)
file_path9 = os.path.join(data_dir, file_name9)
file_path10 = os.path.join(data_dir, file_name10)

df1 = pd.read_excel(file_path1)
df3 = pd.read_excel(file_path3)
df4 = pd.read_excel(file_path4)
df6 = pd.read_excel(file_path6)
df8 = pd.read_excel(file_path8)
df9 = pd.read_excel(file_path9)
df11 = pd.read_excel(file_path10)

#Data preprocessing
df10 = df9.copy()
df1.rename(columns={'Site_Code': 'SiteCode'}, inplace=True)

df3.rename(columns={'FirstOccurrence': 'OSS MF_Start'}, inplace=True)
df3.rename(columns={'ClearTimestamp': 'OSS MF_Clear'}, inplace=True)

df6.rename(columns={'FirstOccurrence': 'OSS DC_Start'}, inplace=True)
df6.rename(columns={'ClearTimestamp': 'OSS DC_Clear'}, inplace=True)

df8.rename(columns={'FirstOccurrence': 'RMS GF_Start'}, inplace=True)
df8.rename(columns={'ClearTimestamp': 'RMS GF_Clear'}, inplace=True)

df10.rename(columns={'FIRSTOCCURRENCE': 'SD_Start_(During PG Run)'}, inplace=True)
df10.rename(columns={'CLEARTIMESTAMP': 'SD_End_(During PG Run)'}, inplace=True)
df10.rename(columns={'SITECODE': 'SiteCode'}, inplace=True)

df9.rename( columns={'FIRSTOCCURRENCE': 'SD_Start_(Prior PG Run)'}, inplace=True)
df9.rename(columns={'CLEARTIMESTAMP': 'SD_End_(Prior PG Run)'}, inplace=True)
df9.rename(columns={'SITECODE': 'SiteCode'}, inplace=True)

df11.rename( columns={'Generated At': 'RMS DC_Start'}, inplace=True)
df11.rename(columns={'Rectified At': 'RMS DC_Clear'}, inplace=True)
df11.rename(columns={'SITE ID ': 'SiteCode'}, inplace=True)


df3.drop_duplicates()
# Filter the rows with "AC Mains Failure" in the "Name" column
df5 = df4[df4['Name'] == 'AC Mains Failure'].copy()
# Reset the index of the new dataframe if needed
df5.reset_index(drop=True, inplace=True)
df5.rename(columns={'Alarm Source': 'SiteCode'}, inplace=True)
df5.rename(columns={'Occurred On (NT)': 'EMS MF_Start'}, inplace=True)
df5.rename(columns={'Cleared On (NT)': 'EMS MF_Clear'}, inplace=True)


# Filter the rows with "AC Mains Failure" in the "Name" column
df7 = df4[df4['Name'].isin(['DC Low Alarm', 'DC Low Voltage'])].copy()
# Reset the index of the new dataframe if needed
df7.reset_index(drop=True, inplace=True)
df7.rename(columns={'Alarm Source': 'SiteCode'}, inplace=True)
df7.rename(columns={'Occurred On (NT)': 'EMS DC_Start'}, inplace=True)
df7.rename(columns={'Cleared On (NT)': 'EMS DC_Clear'}, inplace=True)


# Convert columns to datetime if they're not already
df1['PG_Start_Time'] = pd.to_datetime(df1['PG_Start_Time'], errors='coerce')
df1['PG_End_Time'] = pd.to_datetime(df1['PG_End_Time'], errors='coerce')

# Convert the OSS MF_Start and OSS MF_Clear columns in df3 to the desired format
df3['OSS MF_Start'] = pd.to_datetime(df3['OSS MF_Start'], format='%d/%m/%Y %H:%M:%S')
df3['OSS MF_Clear'] = pd.to_datetime(df3['OSS MF_Clear'], format='%d/%m/%Y %H:%M:%S')
df5['EMS MF_Start'] = pd.to_datetime(df5['EMS MF_Start'], errors='coerce')
df5['EMS MF_Clear'] = pd.to_datetime(df5['EMS MF_Clear'], errors='coerce')
df6['OSS DC_Start'] = pd.to_datetime(df6['OSS DC_Start'], errors='coerce')
df6['OSS DC_Clear'] = pd.to_datetime(df6['OSS DC_Clear'], errors='coerce')
df7['EMS DC_Start'] = pd.to_datetime(df7['EMS DC_Start'], errors='coerce')
df7['EMS DC_Clear'] = pd.to_datetime(df7['EMS DC_Clear'], errors='coerce')
df8['RMS GF_Start'] = pd.to_datetime(df8['RMS GF_Start'], format="%d-%m-%Y %H:%M:%S", errors='coerce')
df8['RMS GF_Clear'] = pd.to_datetime(df8['RMS GF_Clear'], format="%d-%m-%Y %H:%M:%S", errors='coerce')
df9['SD_Start_(Prior PG Run)'] = pd.to_datetime(df9['SD_Start_(Prior PG Run)'], errors='coerce')
df9['SD_End_(Prior PG Run)'] = pd.to_datetime(df9['SD_End_(Prior PG Run)'], errors='coerce')
df10['SD_Start_(During PG Run)'] = pd.to_datetime(df10['SD_Start_(During PG Run)'], errors='coerce')
df10['SD_End_(During PG Run)'] = pd.to_datetime(df10['SD_End_(During PG Run)'], errors='coerce')
df11['RMS DC_Start'] = pd.to_datetime(df11['RMS DC_Start'], format="%d-%m-%Y %H:%M:%S", errors='coerce')
df11['RMS DC_Clear'] = pd.to_datetime(df11['RMS DC_Clear'], format="%d-%m-%Y %H:%M:%S", errors='coerce')
# Initialize result dataframe
result_df = df1.copy()

# Create dictionaries for quick lookups
df3_dict = {}
for _, row in df3.iterrows():
    site_code = row['SiteCode']
    oss_start = row['OSS MF_Start']
    if site_code not in df3_dict:
        df3_dict[site_code] = []
    df3_dict[site_code].append(
        {'start': oss_start, 'clear': row['OSS MF_Clear']})

df5_dict = {}
for _, row in df5.iterrows():
    site_code = row['SiteCode']
    ems_start = row['EMS MF_Start']
    if site_code not in df5_dict:
        df5_dict[site_code] = []
    df5_dict[site_code].append(
        {'start': ems_start, 'clear': row['EMS MF_Clear']})

# Create dictionaries for quick lookups
df6_dict = {}
for _, row in df6.iterrows():
    site_code = row['SiteCode']
    oss_dc_start = row['OSS DC_Start']
    if site_code not in df6_dict:
        df6_dict[site_code] = []
    df6_dict[site_code].append(
        {'start': oss_dc_start, 'clear': row['OSS DC_Clear']})

# Create dictionaries for quick lookups
df7_dict = {}
for _, row in df7.iterrows():
    site_code = row['SiteCode']
    ems_dc_start = row['EMS DC_Start']
    if site_code not in df7_dict:
        df7_dict[site_code] = []
    df7_dict[site_code].append(
        {'start': ems_dc_start, 'clear': row['EMS DC_Clear']})

# Create dictionaries for quick lookups
df11_dict = {}
for _, row in df11.iterrows():
    site_code = row['SiteCode']
    rms_dc_start = row['RMS DC_Start']
    if site_code not in df11_dict:
        df11_dict[site_code] = []
    df11_dict[site_code].append(
        {'start': rms_dc_start, 'clear': row['RMS DC_Clear']})

# Create dictionaries for quick lookups
df8_dict = {}
for _, row in df8.iterrows():
    site_code = row['SiteCode']
    rms_start = row['RMS GF_Start']
    if site_code not in df8_dict:
        df8_dict[site_code] = []
    df8_dict[site_code].append(
        {'start': rms_start, 'clear': row['RMS GF_Clear']})

# Create dictionaries for quick lookups
df10_dict = {}
for _, row in df10.iterrows():
    site_code = row['SiteCode']
    sd_during_start = row['SD_Start_(During PG Run)']
    if site_code not in df10_dict:
        df10_dict[site_code] = []
    df10_dict[site_code].append(
        {'start': sd_during_start, 'clear': row['SD_End_(During PG Run)']})


# Create dictionaries for quick lookups
df9_dict = {}
for _, row in df9.iterrows():
    site_code = row['SiteCode']
    sd_prior_start = row['SD_Start_(Prior PG Run)']
    if site_code not in df9_dict:
        df9_dict[site_code] = []
    df9_dict[site_code].append(
        {'start': sd_prior_start, 'clear': row['SD_End_(Prior PG Run)']})

for index, row in df1.iterrows():
    site_code = row['SiteCode']
    pg_start_time = row['PG_Start_Time']
    pg_end_time = row['PG_End_Time']

    # Process OSS data
    if site_code in df3_dict:
        relevant_oss_entries = [entry for entry in df3_dict[site_code] if entry['start'] <= pg_start_time]
        if relevant_oss_entries:
            latest_oss_entry = max(relevant_oss_entries, key=lambda x: x['start'])
            oss_mf_clear = latest_oss_entry['clear']
            gap_timedelta = pg_start_time - oss_mf_clear
            result_df.at[index, 'OSS MF_Start'] = latest_oss_entry['start']
            result_df.at[index, 'OSS MF_Clear'] = oss_mf_clear

            if gap_timedelta.total_seconds() <= 0:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_OSS'] = 'MF_OSS persisted/clear after PG'
            elif gap_timedelta.total_seconds() <= 600:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_OSS'] = 'MF_OSS ≤ 10 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 1800:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_OSS'] = 'MF_OSS clear ≤ 30 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 86400:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_OSS'] = 'MF_OSS found but clear before PG'
            else:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_OSS'] = 'No MF_OSS found'
        else:
            result_df.at[index, 'Comments_MF_OSS'] = 'No MF_OSS alarm found before PG run'
    else:
         result_df.at[index, 'Comments_MF_OSS'] = 'No MF_OSS alarm found before PG run'
     
    # Process EMS data
    if site_code in df5_dict:
        relevant_ems_entries = [entry for entry in df5_dict[site_code] if entry['start'] <= pg_start_time]
        if relevant_ems_entries:
            latest_ems_entry = max(relevant_ems_entries, key=lambda x: x['start'])
            ems_mf_clear = latest_ems_entry['clear']
            gap_timedelta = pg_start_time - ems_mf_clear
            result_df.at[index, 'EMS MF_Start'] = latest_ems_entry['start']
            result_df.at[index, 'EMS MF_Clear'] = ems_mf_clear

            if gap_timedelta.total_seconds() <= 0:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_EMS'] = 'MF_EMS persisted/clear after PG'
            elif gap_timedelta.total_seconds() <= 600:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_EMS'] = 'MF_EMS ≤ 10 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 1800:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_EMS'] = 'MF_EMS clear ≤ 30 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 86400:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_EMS'] = 'MF_EMS found but clear before PG'
            else:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS'] = gap_timedelta
                result_df.at[index, 'Comments_MF_EMS'] = 'No MF_EMS found'
        else:
            result_df.at[index, 'Comments_MF_EMS'] = 'No MF_EMS alarm found before PG run'
    else:
         result_df.at[index, 'Comments_MF_EMS'] = 'No MF_EMS alarm found before PG run'
            
    # Process OSS_dc data
    if site_code in df6_dict:
        relevant_oss_dc_entries = [entry for entry in df6_dict[site_code] if entry['start'] <= pg_start_time]
        if relevant_oss_dc_entries:
            latest_oss_dc_entry = max(relevant_oss_dc_entries, key=lambda x: x['start'])
            oss_dc_clear = latest_oss_dc_entry['clear']
            gap_timedelta = pg_start_time - oss_dc_clear
            result_df.at[index, 'OSS DC_Start'] = latest_oss_dc_entry['start']
            result_df.at[index, 'OSS DC_Clear'] = oss_dc_clear
            
            if gap_timedelta.total_seconds() <= 0:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_OSS_DC'] = 'OSS_DC persisted/clear after PG'
            elif gap_timedelta.total_seconds() <= 600:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_OSS_DC'] = 'OSS_DC ≤ 10 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 1800:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_OSS_DC'] = 'OSS_DC clear ≤ 30 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 86400:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_OSS_DC'] = 'OSS_DC found but clear before PG'
            else:
                result_df.at[index, 'Gap(PG Start - MF_clear)_OSS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_OSS_DC'] = 'No OSS_DC found'
        else:
            result_df.at[index, 'Comments_OSS_DC'] = 'No OSS_DC alarm found before PG run'
    else:
        result_df.at[index, 'Comments_OSS_DC'] = 'No OSS_DC alarm found before PG run'
            
    #Process EMS_DC data       
    if site_code in df7_dict:
        relevant_ems_dc_entries = [entry for entry in df7_dict[site_code] if entry['start'] <= pg_start_time]
        if relevant_ems_dc_entries:
            latest_ems_dc_entry = max(relevant_ems_dc_entries, key=lambda x: x['start'])
            ems_dc_clear = latest_ems_dc_entry['clear']
            gap_timedelta = pg_start_time - ems_dc_clear
            result_df.at[index, 'EMS DC_Start'] = latest_ems_dc_entry['start']
            result_df.at[index, 'EMS DC_Clear'] = ems_dc_clear

            if gap_timedelta.total_seconds() <= 0:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_EMS_DC'] = 'EMS_DC persisted/clear after PG'
            elif gap_timedelta.total_seconds() <= 600:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_EMS_DC'] = 'EMS_DC ≤ 10 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 1800:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_EMS_DC'] = 'EMS_DC clear ≤ 30 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 86400:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_EMS_DC'] = 'EMS_DC found but clear before PG'
            else:
                result_df.at[index, 'Gap(PG Start - MF_clear)_EMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_EMS_DC'] = 'No EMS_DC found'

        else:
           result_df.at[index, 'Comments_EMS_DC'] = 'No EMS_DC alarm found before PG run'
    else:
       result_df.at[index, 'Comments_EMS_DC'] = 'No EMS_DC alarm found before PG run'
           
    # Process EMS_dc data
    if site_code in df11_dict:
        relevant_rms_dc_entries = [entry for entry in df11_dict[site_code] if entry['start'] <= pg_start_time]
        if relevant_rms_dc_entries:
            latest_rms_dc_entry = max(relevant_rms_dc_entries, key=lambda x: x['start'])
            rms_dc_clear = latest_rms_dc_entry['clear']
            gap_timedelta = pg_start_time - rms_dc_clear
            result_df.at[index, 'RMS DC_Start'] = latest_rms_dc_entry['start']
            result_df.at[index, 'RMS DC_Clear'] = rms_dc_clear
        
            if gap_timedelta.total_seconds() <= 0:
                result_df.at[index, 'Gap(PG Start - MF_clear)_RMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_RMS_DC'] = 'RMS_DC persisted/clear after PG'
            elif gap_timedelta.total_seconds() <= 600:
                result_df.at[index, 'Gap(PG Start - MF_clear)_RMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_RMS_DC'] = 'RMS_DC ≤ 10 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 1800:
                result_df.at[index, 'Gap(PG Start - MF_clear)_RMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_RMS_DC'] = 'RMS_DC clear ≤ 30 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 86400:
                result_df.at[index, 'Gap(PG Start - MF_clear)_RMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_RMS_DC'] = 'RMS_DC found but clear before PG'
            else:
                result_df.at[index, 'Gap(PG Start - MF_clear)_RMS_DC'] = gap_timedelta
                result_df.at[index, 'Comments_RMS_DC'] = 'No RMS_DC found'
        else:
           result_df.at[index, 'Comments_RMS_DC'] = 'No RMS_DC alarm found before PG run'
    else:
        result_df.at[index, 'Comments_RMS_DC'] = 'No RMS_DC alarm found before PG run'


    # Process RMS_MF data
    if site_code in df8_dict:
        relevant_rms_entries = [entry for entry in df8_dict[site_code] if entry['start'] <= pg_start_time]
        if relevant_rms_entries:
            latest_rms_entry = max(relevant_rms_entries, key=lambda x: x['start'])
            rms_mf_clear = latest_rms_entry['clear']
            gap_timedelta = pg_start_time - rms_mf_clear
            result_df.at[index, 'RMS GF_Start'] = latest_rms_entry['start']
            result_df.at[index, 'RMS GF_Clear'] = rms_mf_clear
        
            if gap_timedelta.total_seconds() <= 0:
                result_df.at[index, 'Gap(PG Start - GF_clear)_RMS'] = gap_timedelta
                result_df.at[index, 'Comments_GF_RMS'] = 'GF_RMS persisted/clear after PG'
            elif gap_timedelta.total_seconds() <= 600:
                result_df.at[index, 'Gap(PG Start - GF_clear)_RMS'] = gap_timedelta
                result_df.at[index, 'Comments_GF_RMS'] = 'GF_RMS ≤ 10 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 1800:
                result_df.at[index, 'Gap(PG Start - GF_clear)_RMS'] = gap_timedelta
                result_df.at[index, 'Comments_GF_RMS'] = 'GF_RMS clear ≤ 30 minutes (within before)'
            elif gap_timedelta.total_seconds() <= 86400:
                result_df.at[index, 'Gap(PG Start - GF_clear)_RMS'] = gap_timedelta
                result_df.at[index, 'Comments_GF_RMS'] = 'GF_RMS found but clear before PG'
            else:
                result_df.at[index, 'Gap(PG Start - GF_clear)_RMS'] = gap_timedelta
                result_df.at[index, 'Comments_GF_RMS'] = 'No GF_RMS found'
        
        else:
           result_df.at[index, 'Comments_GF_RMS'] = 'No GF_RMS alarm found before PG run'
    else:
       result_df.at[index, 'Comments_GF_RMS'] = 'No GF_RMS alarm found before PG run'
    
    #SD_Prior Data Processing       
    if site_code in df9_dict:
        relevant_sd_prior_entries = [entry for entry in df9_dict[site_code] if entry['start'] <= pg_start_time]

        if relevant_sd_prior_entries:
            latest_sd_prior_entry = max(relevant_sd_prior_entries, key=lambda x: x['start'])
            sd_prior_clear = latest_sd_prior_entry['clear']
            gap_timedelta = pg_start_time - sd_prior_clear
            result_df.at[index, 'SD_Start_(Prior PG Run)'] =latest_sd_prior_entry['start']
            result_df.at[index, 'SD_End_(Prior PG Run)'] = sd_prior_clear
            if gap_timedelta.total_seconds() <= -600:# Clear before 10 minutes
                result_df.at[index, 'Gap(PG Start - MF_clear)_SD_Prior'] = gap_timedelta
                result_df.at[index, 'Comments_SD_Prior'] = 'Clear before 10 minutes'
            elif gap_timedelta.total_seconds() >= 600:  # Clear after 10 minutes
                result_df.at[index, 'Gap(PG Start - MF_clear)_SD_Prior'] = gap_timedelta
                result_df.at[index, 'Comments_SD_Prior'] = 'Clear after 10 minutes'
            else:
                result_df.at[index, 'Gap(PG Start - MF_clear)_SD_Prior'] = gap_timedelta
                result_df.at[index, 'Comments_SD_Prior'] = 'No SD_Prior alarm found'
        else:
            result_df.at[index, 'Comments_SD_Prior'] = 'No SD_Prior alarm found'
    else:
        result_df.at[index, 'Comments_SD_Prior'] = 'No SD_Prior alarm found'

    # SD_During Data Processing
    if site_code in df10_dict:
        relevant_sd_during_entries = [
        entry for entry in df10_dict[site_code] if entry['start'] >= pg_start_time and entry['clear'] <= pg_end_time
    ]
        if relevant_sd_during_entries:
            latest_sd_during_entry = max(
            relevant_sd_during_entries, key=lambda x: x['start'])
            sd_during_start = latest_sd_during_entry['start']
            sd_during_clear = latest_sd_during_entry['clear']

            sd_during_start = pd.to_datetime(sd_during_start)
            sd_during_clear = pd.to_datetime(sd_during_clear)
            
            # Calculate the gap as a float number (in days)
            gap = (sd_during_clear - sd_during_start).total_seconds() / (60 * 60 * 24)

            result_df.at[index, 'Gap(PG Start - MF_clear)_SD_During'] = gap

            if gap == 0:
               result_df.at[index, 'Comments_SD_During'] = 'No Overlap'
            elif gap == row['Total PG RH']:
               result_df.at[index, 'Comments_SD_During'] = 'Full Overlap'
            else:
               result_df.at[index, 'Comments_SD_During'] = 'Partial Overlap'

               result_df.at[index, 'SD_Start_(During PG Run)'] = sd_during_start
               result_df.at[index, 'SD_End_(During PG Run)'] = sd_during_clear
        else:
            result_df.at[index, 'Comments_SD_During'] = 'No Overlap'
    else:
       result_df.at[index, 'Comments_SD_During'] = 'No Overlap'



# Reorder columns as per the desired sequence
desired_columns = [
    'TT_Issue_ Date','eQuip TT_NO','SiteCode','Zone','PG_Owner','PG_Run_Type','PG_Controller_ID','PG RH  Category','PG_Start_Time','PG_End_Time','Total PG RH',	'Justification for DG site PG RH',	'Justification for Non Controller PG RH','Remarks (If any)','Operator Name','Service Vendor',
    'OSS MF_Start','OSS MF_Clear','Gap(PG Start - MF_clear)_OSS','Comments_MF_OSS','EMS MF_Start','EMS MF_Clear','Gap(PG Start - MF_clear)_EMS','Comments_MF_EMS','OSS DC_Start','OSS DC_Clear','Gap(PG Start - MF_clear)_OSS_DC','Comments_OSS_DC','EMS DC_Start','EMS DC_Clear','Gap(PG Start - MF_clear)_EMS_DC','Comments_EMS_DC','RMS DC_Start','RMS DC_Clear','Gap(PG Start - MF_clear)_RMS_DC','Comments_RMS_DC','RMS GF_Start','RMS GF_Clear','Gap(PG Start - GF_clear)_RMS','Comments_GF_RMS','SD_Start_(Prior PG Run)','SD_End_(Prior PG Run)','Gap(PG Start - MF_clear)_SD_Prior','Comments_SD_Prior','SD_Start_(During PG Run)','SD_End_(During PG Run)','Gap(PG Start - MF_clear)_SD_During','Comments_SD_During']
#,'Gap(PG Start - MF_clear)_SD_Prior','Comments_SD_Prior'
result_df = result_df[desired_columns]
# Save the result to Excel
result_df.to_excel(data_dir_output+ "\\" + "Result.xlsx", index=None)
