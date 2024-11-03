from utils.modules import * 


analyzer = IPA_PLC_Inspector(excel_file='aaa.xlsx',target = '시각', plc_time_sync_min = 42, plc_time_sync_sec = 30,
                                wav_dir='/home/mg/True_NAS/PoC_IPA_01/1_motor-50kt/E0001S01', out_dir='test.xlsx')
analyzer.save_output()
analyzer.print_summary()
