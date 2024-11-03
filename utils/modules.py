import pandas as pd
import os
from datetime import timedelta

class IPA_PLC_Inspector:
    def __init__(self, target, plc_time_sync_min,plc_time_sync_sec, excel_file, wav_dir, out_dir):
        self.excel_file = excel_file
        self.wav_dir = wav_dir
        self.out_dir = out_dir
        self.target = target
        self.plc_time_sync_min = plc_time_sync_min
        self.plc_time_sync_sec = plc_time_sync_sec
        self.df = self.load_data()
        self.grouped_data = self.process_data()

    def load_data(self):
        df = pd.read_excel(self.excel_file)
        df['timestamp'] = pd.to_datetime(df[self.target], format='%Y %m %d %H:%M:%S')
        df['time_diff'] = df['timestamp'].diff().fillna(pd.Timedelta(seconds=0))
        return df

    def process_data(self):
        self.df['group'] = (self.df['time_diff'] > timedelta(seconds=64)).cumsum()
        grouped = self.df.groupby('group')['timestamp'].agg(['first', 'last']).reset_index(drop=True)
        grouped['[PLC]start'] = grouped['first'].dt.strftime('%Y%m%d%H%M%S')
        grouped['[PLC]end'] = grouped['last'].dt.strftime('%Y%m%d%H%M%S')
        
        grouped['[Sync]start'] = (grouped['first'] - timedelta(minutes=self.plc_time_sync_min,seconds=self.plc_time_sync_sec)).dt.strftime('%Y%m%d%H%M%S')
        grouped['[Sync]end'] = (grouped['last'] - timedelta(minutes=self.plc_time_sync_min,seconds=self.plc_time_sync_sec)).dt.strftime('%Y%m%d%H%M%S')

        grouped['duration (min(s))'] = (grouped['last'] - grouped['first']).dt.total_seconds() / 60
        grouped['case#'] = range(1, len(grouped) + 1)
        grouped['# of wav files'] = grouped.apply(lambda row: self.count_wav_files(row['[PLC]start'], row['[PLC]end']), axis=1)
        return grouped[['case#', '[PLC]start', '[PLC]end', '[Sync]start', '[Sync]end', 'duration (min(s))', '# of wav files']]

    def count_wav_files(self, start, end):
        start_time = pd.to_datetime(start, format='%Y%m%d%H%M%S')
        end_time = pd.to_datetime(end, format='%Y%m%d%H%M%S')
        count = 0

        for hour in range(start_time.hour, end_time.hour + 1):
            hour_dir = os.path.join(self.wav_dir, start_time.strftime('%Y%m%d') + f'{hour:02}')
            if os.path.exists(hour_dir):
                for filename in os.listdir(hour_dir):
                    if filename.endswith('.wav'):
                        timestamp_str = filename[:-4]
                        file_time = pd.to_datetime(timestamp_str, format='%Y%m%d%H%M%S', errors='coerce')

                        if file_time is not None and start_time <= file_time <= end_time:
                            count += 1

        return count

    def save_output(self):
        self.grouped_data.to_excel(self.out_dir, index=False)

    def print_summary(self):
        for _, row in self.grouped_data.iterrows():
            print(f"Case#: {row['case#']}, [PLC] Start: {row['[PLC]start']}, [PLC] End: {row['[PLC]end']}, [Sync] Start: {row['[Sync]start']}, [Sync] End: {row['[Sync]end']}, Duration: {row['duration (min(s))']} minutes, # of wav files: {row['# of wav files']}")

