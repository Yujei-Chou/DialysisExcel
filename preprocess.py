import pandas as pd
import numpy as np

def timePeriodTrans(hr):
    if(5 <= hr and hr <= 7):
        return 1
    elif(10 <= hr and hr <= 13):
        return 2
    elif(16 <= hr and hr <= 19):
        return 3
    elif(21 <= hr and hr <=23):
        return 4
    else:
        pass

def getBPcomb(BP_df, all_df):
    minIdx = BP_df['收縮壓 (mmHg)'].idxmin()
    try:
        BPcomb = str(int(all_df['收縮壓 (mmHg)'].iloc[minIdx])) + '/' + str(int(all_df['舒張壓 (mmHg)'].iloc[minIdx]))
    except:
        BPcomb = None
    
    return BPcomb


def getDataframe(file, start_date, end_date):
    org_df = pd.read_excel(file, engine='openpyxl')
    org_df['日期'] = org_df['時間戳記'].dt.date
    org_df['時段'] = org_df['時間戳記'].dt.hour.apply(timePeriodTrans)
    org_df['脫水量 (cc)'] = org_df['脫水量 (cc)'].shift(-1)
    org_df = org_df[(org_df['日期'] >= start_date) & (org_df['日期'] < end_date)]
    org_df = org_df.drop(columns=['時間戳記'])

    pad_df = pd.DataFrame({'時間戳記': pd.date_range(start=start_date, end=end_date, freq='6H')})[:-1]
    pad_df['日期'] = pad_df['時間戳記'].dt.date
    pad_df['時段'] = pad_df['時間戳記'].dt.hour // 6 + 1
    pad_df = pad_df.drop(columns=['時間戳記'])

    df = pd.merge(org_df, pad_df, on=['日期' , '時段'], how='outer').sort_values(by=['日期', '時段']).reset_index(drop=True)

    weight_min = df[['日期', '體重 (kg)']].groupby('日期').min().reset_index()

    BP_onSBPmin = df[['日期', '收縮壓 (mmHg)', '舒張壓 (mmHg)']].groupby('日期').apply(getBPcomb, all_df=df).rename('血壓 (mm/Hg)').reset_index()

    DP_list = df[['日期', '透析液濃度 (%)']].groupby('日期')['透析液濃度 (%)'].apply(list).reset_index()
    DP_list[['早上(濃度)', '中午(濃度)', '下午(濃度)', '晚上(濃度)']] = pd.DataFrame(DP_list['透析液濃度 (%)'].tolist())
    DP_list = DP_list.drop(columns=['透析液濃度 (%)'])

    UF_list = df[['日期', '脫水量 (cc)']].groupby('日期')['脫水量 (cc)'].apply(list).reset_index()
    UF_list[['早上(排出)', '中午(排出)', '下午(排出)', '晚上(排出)']] = pd.DataFrame(UF_list['脫水量 (cc)'].tolist())
    UF_list = UF_list.drop(columns=['脫水量 (cc)'])


    merge_dfs = [BP_onSBPmin, weight_min, UF_list, DP_list]
    CAPD = merge_dfs[0]
    for merge_df in merge_dfs[1:]:
        CAPD = pd.merge(CAPD, merge_df, on='日期')
    CAPD = CAPD.replace({np.nan: None})

    return CAPD



class CAPDExcel():
    def __init__(self, uploadfile, downloadfile, start_date, end_date) -> None:
        super().__init__()
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

    
    def intervalSet(self, DP, input, output, r, c): # (透析液濃度, 注入量, 排出量, rowId, colId)
        self.workSheet.merge_range(f'{c}{r}:{c}{r+1}', DP, self.detail_DP_format)
        self.workSheet.write(f'{chr(ord(c)+1)}{r}', input, self.detail_based_format)
        self.workSheet.write(f'{chr(ord(c)+1)}{r+1}', output, self.detail_UF_format)
        self.workSheet.set_column(f'{c}:{c}', 3)
        self.workSheet.set_column(f'{chr(ord(c)+1)}:{chr(ord(c)+1)}', 7.88)


    def getDataframe(self):
        org_df = pd.read_excel(self.uploadfile, engine='openpyxl')
        org_df['日期'] = org_df['時間戳記'].dt.date
        org_df['時段'] = org_df['時間戳記'].dt.hour.apply(timePeriodTrans)
        org_df['脫水量 (cc)'] = org_df['脫水量 (cc)'].shift(-1)
        org_df = org_df[(org_df['日期'] >= self.start_date) & (org_df['日期'] < self.end_date)]
        org_df = org_df.drop(columns=['時間戳記'])

        pad_df = pd.DataFrame({'時間戳記': pd.date_range(start=self.start_date, end=self.end_date, freq='6H')})[:-1]
        pad_df['日期'] = pad_df['時間戳記'].dt.date
        pad_df['時段'] = pad_df['時間戳記'].dt.hour // 6 + 1
        pad_df = pad_df.drop(columns=['時間戳記'])

        df = pd.merge(org_df, pad_df, on=['日期' , '時段'], how='outer').sort_values(by=['日期', '時段']).reset_index(drop=True)

        weight_min = df[['日期', '體重 (kg)']].groupby('日期').min().reset_index()

        BP_onSBPmin = df[['日期', '收縮壓 (mmHg)', '舒張壓 (mmHg)']].groupby('日期').apply(getBPcomb, all_df=df).rename('血壓 (mm/Hg)').reset_index()

        DP_list = df[['日期', '透析液濃度 (%)']].groupby('日期')['透析液濃度 (%)'].apply(list).reset_index()
        DP_list[['早上(濃度)', '中午(濃度)', '下午(濃度)', '晚上(濃度)']] = pd.DataFrame(DP_list['透析液濃度 (%)'].tolist())
        DP_list = DP_list.drop(columns=['透析液濃度 (%)'])

        UF_list = df[['日期', '脫水量 (cc)']].groupby('日期')['脫水量 (cc)'].apply(list).reset_index()
        UF_list[['早上(排出)', '中午(排出)', '下午(排出)', '晚上(排出)']] = pd.DataFrame(UF_list['脫水量 (cc)'].tolist())
        UF_list = UF_list.drop(columns=['脫水量 (cc)'])


        merge_dfs = [BP_onSBPmin, weight_min, UF_list, DP_list]
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
        for idx, row in CAPD_df.iterrows():
            self.workSheet.merge_range(f'A{sidx}:A{sidx+1}', row['日期'], self.detail_date_format)
            self.workSheet.merge_range(f'B{sidx}:B{sidx+1}', f'=DAY(A{sidx})', self.detail_based_format)
            self.workSheet.merge_range(f'C{sidx}:C{sidx+1}', row['血壓 (mm/Hg)'], self.detail_based_format)
            self.workSheet.merge_range(f'D{sidx}:D{sidx+1}', row['體重 (kg)'], self.detail_based_format)
            self.intervalSet(row['早上(濃度)'], 2000, row['早上(排出)'], sidx, 'E')
            self.intervalSet(row['中午(濃度)'], 2000, row['中午(排出)'], sidx, 'G')
            self.intervalSet(row['下午(濃度)'], 2000, row['下午(排出)'], sidx, 'I')
            self.intervalSet(row['晚上(濃度)'], 2000, row['晚上(排出)'], sidx, 'K')
            self.intervalSet(None, None, None, sidx, 'M')

            self.workSheet.merge_range(f'O{sidx}:O{sidx+1}', f'=F{sidx+1}+H{sidx+1}+J{sidx+1}+L{sidx+1}-8000', self.detail_based_format)
            self.workSheet.set_row(sidx, 12)
            self.workSheet.set_row(sidx+1, 12)

            sidx += 2

        self.workSheet.print_area(f'B1:O{sidx-1}')
        self.workBook.close()
        


