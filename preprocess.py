import pandas as pd
import numpy as np


def getBPcomb(BP_df, all_df):
    minIdx = BP_df['收縮壓 (mmHg)'].idxmin()
    try:
        BPcomb = str(int(all_df['收縮壓 (mmHg)'].iloc[minIdx])) + '/' + str(int(all_df['舒張壓 (mmHg)'].iloc[minIdx]))
    except:
        BPcomb = None
    
    return BPcomb


class CAPDExcel():
    def __init__(self, uploadfile, downloadfile, start_date, end_date) -> None:
        self.uploadfile = uploadfile
        self.downloadfile = downloadfile
        self.start_date = start_date
        self.end_date = end_date

        self.workBook = pd.ExcelWriter(self.downloadfile, engine='xlsxwriter').book
        self.workSheet = self.workBook.add_worksheet('透析紀錄')

        # Additional infomation format
        addinfo_dict = {
            'font_size': '12',
            'font_name': 'Regular'
        }
        addinfo_YM_dict = addinfo_dict.copy()
        addinfo_YM_dict['align'] = 'right'
        self.addinfo_UO_format = self.workBook.add_format(addinfo_dict)
        self.addinfo_YM_format = self.workBook.add_format(addinfo_YM_dict)

        # Detail format
        detail_based_dict = {
            'align': 'center',
            'valign': 'vcenter',
            'font_size': '9',
            'font_name': 'Arial Unicode MS',
            'border': 1
        }
        self.detail_based_format = self.workBook.add_format(detail_based_dict)

        detail_date_dict = detail_based_dict.copy()
        detail_date_dict['num_format'] = 'yyyy/mm/dd'
        self.detail_date_format = self.workBook.add_format(detail_date_dict)

        detail_DP_dict = detail_based_dict.copy()
        detail_DP_dict['font_name'] = 'Arial'
        self.detail_DP_format = self.workBook.add_format(detail_DP_dict)

        detail_UF_dict = detail_based_dict.copy()
        detail_UF_dict['font_color'] = 'red'
        self.detail_UF_format = self.workBook.add_format(detail_UF_dict)

        colname_dict = detail_based_dict.copy()
        colname_dict['font_size'] = '10'
        colname_dict['font_name'] = 'Light Regular'
        self.colname_format = self.workBook.add_format(colname_dict)

    
    def intervalSet(self, dfRow, tp, r, c): # (dfRow, timePeriod, rowId, colId)
        self.workSheet.merge_range(f'{c}{r}:{c}{r+1}', dfRow[f'{tp}(濃度)'], self.detail_DP_format)
        self.workSheet.write(f'{chr(ord(c)+1)}{r}', 2000 if dfRow[f'{tp}(排出)'] else None, self.detail_based_format)
        self.workSheet.write(f'{chr(ord(c)+1)}{r+1}', dfRow[f'{tp}(排出)'], self.detail_UF_format)
        if(dfRow['時段頻率'] < 0.15 and dfRow[f'{tp}(排出)']):
            self.workSheet.write_comment(f'{chr(ord(c)+1)}{r}', '透析時間: ' + dfRow[f'{tp}(時段)'].strftime("%m-%d %H:%M"))
        self.workSheet.set_column(f'{c}:{c}', 3)
        self.workSheet.set_column(f'{chr(ord(c)+1)}:{chr(ord(c)+1)}', 7.88)


    def getDataframe(self):
        df = pd.read_excel(self.uploadfile, engine='openpyxl').sort_values(by=['時間戳記'])
        df = df[df['時間戳記'].diff() > pd.Timedelta(hours=1)]
        df['日期'] = df['時間戳記'].dt.date
        df['時段'] = df.groupby('日期').cumcount() + 1
        df['脫水量 (cc)'] = df['脫水量 (cc)'].shift(-1)
        df = df[(df['日期'] >= self.start_date) & (df['日期'] < self.end_date)]
        df = df.reset_index(drop=True)

        weight_min = df[['日期', '體重 (kg)']].groupby('日期').min().reset_index()

        BP_onSBPmin = df[['日期', '收縮壓 (mmHg)', '舒張壓 (mmHg)']].groupby('日期').apply(getBPcomb, all_df=df).rename('血壓 (mm/Hg)').reset_index()

        time_max = df[['日期', '時段']].groupby('日期').max().reset_index()
        time_freq = (time_max['時段'].value_counts()/len(time_max)).sort_index().to_dict()
        time_max_freq = time_max.replace(time_freq).rename(columns={'時段': '時段頻率'})

        DP_list = df[['日期', '透析液濃度 (%)']].groupby('日期')['透析液濃度 (%)'].agg(lambda x: list(x) + [np.nan] * (5-len(x))).reset_index()
        DP_list[['1(濃度)', '2(濃度)', '3(濃度)', '4(濃度)', '5(濃度)']] = pd.DataFrame(DP_list['透析液濃度 (%)'].tolist())
        DP_list = DP_list.drop(columns=['透析液濃度 (%)'])

        UF_list = df[['日期', '脫水量 (cc)']].groupby('日期')['脫水量 (cc)'].agg(lambda x: list(x) + [np.nan] * (5-len(x))).reset_index()
        UF_list[['1(排出)', '2(排出)', '3(排出)', '4(排出)', '5(排出)']] = pd.DataFrame(UF_list['脫水量 (cc)'].tolist())
        UF_list = UF_list.drop(columns=['脫水量 (cc)'])

        time_list = df[['日期', '時間戳記']].groupby('日期')['時間戳記'].agg(lambda x: list(x) + [np.nan] * (5-len(x))).reset_index()
        time_list[['1(時段)', '2(時段)', '3(時段)', '4(時段)', '5(時段)']] = pd.DataFrame(time_list['時間戳記'].tolist())
        time_list = time_list.drop(columns=['時間戳記'])

        merge_dfs = [BP_onSBPmin, weight_min, UF_list, DP_list, time_max_freq, time_list]
        CAPD = merge_dfs[0]
        for merge_df in merge_dfs[1:]:
            CAPD = pd.merge(CAPD, merge_df, on='日期')
        CAPD = CAPD.replace({np.nan: None})

        return CAPD


    def getExcel(self):
        self.workSheet.write('B1', '尿量: cc', self.addinfo_UO_format)
        self.workSheet.write('O1', f'{self.start_date.year-1911}年{self.start_date.month}月', self.addinfo_YM_format)
        self.workSheet.merge_range('A2:A3', '日期', self.colname_format)
        self.workSheet.set_column('A:A', 10)
        self.workSheet.merge_range('B2:B3', '天', self.colname_format)
        self.workSheet.set_column('B:B', 3)
        self.workSheet.merge_range('C2:C3', '血壓', self.colname_format)
        self.workSheet.merge_range('D2:D3', '體重', self.colname_format)
        self.workSheet.merge_range('E2:N2', '透析液濃度、注入量、排出量', self.colname_format)
        self.workSheet.merge_range('E3:F3', '1', self.colname_format)
        self.workSheet.merge_range('G3:H3', '2', self.colname_format)
        self.workSheet.merge_range('I3:J3', '3', self.colname_format)
        self.workSheet.merge_range('K3:L3', '4', self.colname_format)
        self.workSheet.merge_range('M3:N3', '5', self.colname_format)
        self.workSheet.merge_range('O2:O3', '全日脫水量', self.colname_format)
        self.workSheet.set_column('O:O', 9.5)

        sidx = 4
        CAPD_df = self.getDataframe()
        tp_col = ['E', 'G', 'I', 'K', 'M']
        for idx, row in CAPD_df.iterrows():
            self.workSheet.merge_range(f'A{sidx}:A{sidx+1}', row['日期'], self.detail_date_format)
            self.workSheet.merge_range(f'B{sidx}:B{sidx+1}', f'=DAY(A{sidx})', self.detail_based_format)
            self.workSheet.merge_range(f'C{sidx}:C{sidx+1}', row['血壓 (mm/Hg)'], self.detail_based_format)
            self.workSheet.merge_range(f'D{sidx}:D{sidx+1}', row['體重 (kg)'], self.detail_based_format)
            for i in range(5):
                tp = i + 1
                self.intervalSet(row, tp, sidx, tp_col[i])
            

            self.workSheet.merge_range(f'O{sidx}:O{sidx+1}', f'=F{sidx+1}+H{sidx+1}+J{sidx+1}+L{sidx+1}+N{sidx+1}'+ \
                                                             f'-F{sidx}-H{sidx}-J{sidx}-L{sidx}-N{sidx}', self.detail_based_format)
            self.workSheet.set_row(sidx, 12)
            self.workSheet.set_row(sidx+1, 12)

            sidx += 2

        self.workSheet.print_area(f'B1:O{sidx-1}')
        self.workBook.close()
        


