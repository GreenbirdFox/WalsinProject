import pandas as pd
import numpy as np
import time
# filter warning
import warnings

warnings.simplefilter("ignore", UserWarning)

# 計算時間, 起始
start = time.time()


# 讀取主檔
def read_file():
    print('讀取排程主檔中...')
    dfs = pd.read_excel('0713MainFile(eng)_v1.xlsx')
    print('讀取排程主檔中... done')
    return dfs


# 中轉英
def to_translate(temp_df):
    # 中文欄位轉英文資料
    temp_df.rename(columns={
        '減面率': 'REDUCTION_RATE',
        '預計到料日': 'ARRIVAL_DATE',
        '預設調機狀態': 'STATUS',
        '預計產出時間': 'SCHEDULE_DATE',
        '排程群組': 'SCHE_GROUP',
        '排程順序': 'SCHEDULE_SEQ',
        '計畫量': 'WEIGHT',
        '預計投入重量': 'PLAN_WEIGHT_I',
        '預計產出重量': 'PLAN_WEIGHT_O',
        '產率': 'YIELD',
        '製序': 'ROUTING_SEQ',
        '現況製序': 'ROUTING_SEQ_NOW',
        '排程站別': 'SHOP_CODE_SCHE',
        '預設機台別': 'EQUIP_CODE',
        '預設機台工時': 'WORKING_HOURS',
        'ID備料進度': 'ID_PREP_STATE',
        '現況差異數': 'STAGE_DIFF_NO',
        '批次': 'SCHE_TYPE',
        '現況站別': 'SHOP_CODE_NOW',
        '放行碼': 'PROC_STATUS',
        '鋼種': 'STEEL_TYPE',
        '下站別': 'NEXT_SHOP_CODE',
        '下站機台': 'NEXT_EQUIP_CODE',
        '支數': 'PIECE_COUNT',
        '長度': 'ACTUAL_LENGTH',
        '總長度': 'TTL_LENGTH',
        '製程碼': 'PROCESS_CODE',
        'LINEUP_MIC_NO': 'LINEUP_MIC_NO',
        'LINEUP流程': 'LINEUP_PROCESS',
        'FINAL_MIC_NO': 'FINAL_MIC_NO',
        'FINAL流程': 'FINAL_PROCESS',
        '儲區': 'LOC',
        '訂單號碼': 'SALE_ORDER',
        '訂單項次': 'SALE_ITEM',
        '產品類型': 'PROD_TYPE',
        '生計交期': 'DATE_DELIVERY_PP',
        '營業交期': 'DATE_DELIVERY_SALES',
        '總重上限': 'MAX_OF_TOTAL_WEIGHT',
        '總重下限': 'MIN_OF_TOTAL_WEIGHT',
        '客戶': 'CUST_NAME',
        '料號': 'MTRL_NO',
        '最小尺寸': 'DIA_MIN',
        '最大尺寸': 'DIA_MAX',
        '訂單形狀': 'SALE_ORDER_SHAPE',
        '訂單尺寸': 'SALE_ORDER_DIA',
        '訂單長度': 'SALE_ORDER_LENGTH',
        '急單順序': 'URGENCY_ORDER',
        '急單日期': 'URGENCY_DATE',
        '急單說明': 'URGENCY_ORDER_DESC',
        '軋延尺寸': 'PP_SHAVE_SIZE',
        '投入型態': 'INPUT_SHAPE',
        '產出型態': 'OUTPUT_SHAPE',
        '投入尺寸': 'INPUT_DIA',
        '產出尺寸': 'TURN_DIA_MAX',
        '產品型態': 'PROD_CATEGORY',
        '頻率': 'FREQUENCE',
        '溫度': 'TEMP',
        '軋延日期': 'MILL_DATE',
        '機台_01': 'EQ_CODE_R01',
        '工時_01': 'W_HR_R01',
        '機台_02': 'EQ_CODE_R02',
        '工時_02': 'W_HR_R02',
        '機台_03': 'EQ_CODE_R03',
        '工時_03': 'W_HR_R03',
        '機台_04': 'EQ_CODE_R04',
        '工時_04': 'W_HR_R04',
        '機台_05': 'EQ_CODE_R05',
        '工時_05': 'W_HR_R05',
        '機台_06': 'EQ_CODE_R06',
        '工時_06': 'W_HR_R06',
        '機台_07': 'EQ_CODE_R07',
        '工時_07': 'W_HR_R07',
        '機台_08': 'EQ_CODE_R08',
        '工時_08': 'W_HR_R08',
        '機台_09': 'EQ_CODE_R09',
        '工時_09': 'W_HR_R09',
        '機台_10': 'EQ_CODE_R10',
        '工時_10': 'W_HR_R10',
        '機台調機狀態_01': 'STATUS_01',
        '機台調機狀態_02': 'STATUS_02',
        '機台調機狀態_03': 'STATUS_03',
        '機台調機狀態_04': 'STATUS_04',
        '機台調機狀態_05': 'STATUS_05',
        '機台調機狀態_06': 'STATUS_06',
        '機台調機狀態_07': 'STATUS_07',
        '機台調機狀態_08': 'STATUS_08',
        '機台調機狀態_09': 'STATUS_09',
        '機台調機狀態_10': 'STATUS_10',
        '最晚投入日(悲觀)': 'LATEST_INPUT_DN',
        '最晚投入日(樂觀)': 'LATEST_INPUT_DP',
        'SETUP_TIME': 'Tooling_change',
        '密度': 'DENSITY',
        '最短製程週期時間': 'WORKING_HOURS_MIN',
        '最長製程週期時間': 'WORKING_HOURS_MAX', }, inplace=True)
    return temp_df


# 英轉中
def to_translate_ch(temp_df):
    # 中文欄位轉英文資料
    temp_df.rename(columns={
        'REDUCTION_RATE': '減面率',
        'ARRIVAL_DATE': '預計到料日',
        'STATUS': '預設調機狀態',
        'SCHEDULE_DATE': '預計產出時間',
        'SCHE_GROUP': '排程群組',
        'SCHEDULE_SEQ': '排程順序',
        'WEIGHT': '計畫量',
        'PLAN_WEIGHT_I': '預計投入重量',
        'PLAN_WEIGHT_O': '預計產出重量',
        'YIELD': '產率',
        'ROUTING_SEQ': '製序',
        'ROUTING_SEQ_NOW': '現況製序',
        'SHOP_CODE_SCHE': '排程站別',
        'EQUIP_CODE': '預設機台別',
        'WORKING_HOURS': '預設機台工時',
        'ID_PREP_STATE': 'ID備料進度',
        'STAGE_DIFF_NO': '現況差異數',
        'SCHE_TYPE': '批次',
        'SHOP_CODE_NOW': '現況站別',
        'PROC_STATUS': '放行碼',
        'STEEL_TYPE': '鋼種',
        'NEXT_SHOP_CODE': '下站別',
        'NEXT_EQUIP_CODE': '下站機台',
        'PIECE_COUNT': '支數',
        'ACTUAL_LENGTH': '長度',
        'TTL_LENGTH': '總長度',
        'PROCESS_CODE': '製程碼',
        'LINEUP_MIC_NO': 'LINEUP_MIC_NO',
        'LINEUP_PROCESS': 'LINEUP流程',
        'FINAL_MIC_NO': 'FINAL_MIC_NO',
        'FINAL_PROCESS': 'FINAL流程',
        'LOC': '儲區',
        'SALE_ORDER': '訂單號碼',
        'SALE_ITEM': '訂單項次',
        'PROD_TYPE': '產品類型',
        'DATE_DELIVERY_PP': '生計交期',
        'DATE_DELIVERY_SALES': '營業交期',
        'MAX_OF_TOTAL_WEIGHT': '總重上限',
        'MIN_OF_TOTAL_WEIGHT': '總重下限',
        'CUST_NAME': '客戶',
        'MTRL_NO': '料號',
        'DIA_MIN': '最小尺寸',
        'DIA_MAX': '最大尺寸',
        'SALE_ORDER_SHAPE': '訂單形狀',
        'SALE_ORDER_DIA': '訂單尺寸',
        'SALE_ORDER_LENGTH': '訂單長度',
        'URGENCY_ORDER': '急單順序',
        'URGENCY_DATE': '急單日期',
        'URGENCY_ORDER_DESC': '急單說明',
        'PP_SHAVE_SIZE': '軋延尺寸',
        'INPUT_SHAPE': '投入型態',
        'OUTPUT_SHAPE': '產出型態',
        'INPUT_DIA': '投入尺寸',
        'TURN_DIA_MAX': '產出尺寸',
        'PROD_CATEGORY': '產品型態',
        'FREQUENCE': '頻率',
        'TEMP': '溫度',
        'MILL_DATE': '軋延日期',
        'EQ_CODE_R01': '機台_01',
        'W_HR_R01': '工時_01',
        'EQ_CODE_R02': '機台_02',
        'W_HR_R02': '工時_02',
        'EQ_CODE_R03': '機台_03',
        'W_HR_R03': '工時_03',
        'EQ_CODE_R04': '機台_04',
        'W_HR_R04': '工時_04',
        'EQ_CODE_R05': '機台_05',
        'W_HR_R05': '工時_05',
        'EQ_CODE_R06': '機台_06',
        'W_HR_R06': '工時_06',
        'EQ_CODE_R07': '機台_07',
        'W_HR_R07': '工時_07',
        'EQ_CODE_R08': '機台_08',
        'W_HR_R08': '工時_08',
        'EQ_CODE_R09': '機台_09',
        'W_HR_R09': '工時_09',
        'EQ_CODE_R10': '機台_10',
        'W_HR_R10': '工時_10',
        'STATUS_01': '機台調機狀態_01',
        'STATUS_02': '機台調機狀態_02',
        'STATUS_03': '機台調機狀態_03',
        'STATUS_04': '機台調機狀態_04',
        'STATUS_05': '機台調機狀態_05',
        'STATUS_06': '機台調機狀態_06',
        'STATUS_07': '機台調機狀態_07',
        'STATUS_08': '機台調機狀態_08',
        'STATUS_09': '機台調機狀態_09',
        'STATUS_10': '機台調機狀態_10',
        'LATEST_INPUT_DN': '最晚投入日(悲觀)',
        'LATEST_INPUT_DP': '最晚投入日(樂觀)',
        'SETUP_TIME': 'Tooling_change',
        'DENSITY': '密度',
        'ID_BAR_SHOP_CODE': '生產站別',
        'DATE_PP': '生計入庫日',
        'DATE_SALES': '營業入庫日',
        'FLAG_WHOLE_ORDER_SHIPMENT': '整單出貨',
        'EXPORT_CABINET_NO': '外貨貨櫃編號',
        'SALE_AREA_GROUP': '區別',
        'PP_CYCLE_NO': '軋延CYCLE',
        'WORKING_HOURS_MIN': '最短製程週期時間',
        'WORKING_HOURS_MAX': '最長製程週期時間', }, inplace=True)
    return temp_df


# 檢查排程是否有問題
def check_error(temp_df):
    print('檢查違反規則...')
    print("=====================================")
    # ==========================================================================
    # 1.  406(噴砂) 投入尺寸<30mm的棒材只能於CB0 機台生產，投入尺寸>=30mm的棒材只能於CB1 機台生產
    # CB0
    temp_df1 = temp_df.copy()
    temp_df1 = temp_df1.replace('None', np.nan)
    mask1_1 = (temp_df1['預設機台別'] != 'CB0') & (temp_df1['排程站別'] == 406) & (temp_df['投入尺寸'] != 'None')
    temp_df1 = temp_df1[mask1_1].reset_index(drop=False)
    err_1_1 = temp_df1.loc[temp_df1['投入尺寸'] < 30].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_1_1.loc[:, 'index'].tolist()
    for k in temp_index_list:
        temp_df.loc[k, 'VIOLATION'] = 'A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '投入尺寸<30mm的棒材只能於CB0 機台生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，投入尺寸<30mm的棒材只能於CB0 機台生產'
    # 報錯
    groups1_1 = err_1_1.groupby(by='ID_NO')
    groups_list1_1 = []
    for key in groups1_1.groups.keys():
        groups_list1_1.append(key)
    if len(groups_list1_1) != 0:
        print('錯誤1_1: 406站 發現投入尺寸<30mm的棒材但不在CB0機台生產 的資料')
        print('    發現該錯誤的ID_NO為: ', end=' ')
        print(*groups_list1_1, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list1_1)))
        print("=====================================")

    # CB1
    mask1_2 = (temp_df['預設機台別'] != 'CB1') & (temp_df['排程站別'] == 406) & (temp_df['投入尺寸'] != 'None')
    temp_df1 = temp_df[mask1_2].reset_index(drop=False)
    err_1_2 = temp_df1.loc[temp_df1['投入尺寸'] >= 30].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_1_2.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '投入尺寸>=30mm的棒材只能於CB1 機台生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，投入尺寸>=30mm的棒材只能於CB1 機台生產'
    # 報錯
    groups1_2 = err_1_2.groupby(by='ID_NO')
    groups_list1_2 = []
    for key in groups1_2.groups.keys():
        groups_list1_2.append(key)
    if len(groups_list1_2) != 0:
        print('錯誤1_2: 406站 發現投入尺寸>=30mm的棒材但不在CB1機台生產 的資料')
        print('    發現該錯誤的ID_NO為: ', end=' ')
        print(*groups_list1_2, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list1_2)))
        print("=====================================")
    # ==========================================================================

    # 2. 壓光站(451) 訂單尺寸30.8mm(含)以下最好排K4
    mask2_1 = (temp_df['排程站別'] == 451) & (temp_df['訂單尺寸'] != 'None')
    temp_df1 = temp_df[mask2_1]
    mask2_2 = (temp_df1['訂單尺寸'] <= 30.8)
    temp_df2 = temp_df1[mask2_2]
    err_2 = temp_df2.loc[temp_df2['預設機台別'] != 'K4'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_2.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸30.8mm(含)以下最好排K4'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸30.8mm(含)以下最好排K4'
    # 報錯
    groups2 = err_2.groupby(by='ID_NO')
    groups_list2 = []
    for key in groups2.groups.keys():
        groups_list2.append(key)
    if len(groups_list2) != 0:
        print('錯誤2: 451站發現 訂單尺寸<=30.8但預設機台別不是K4 的資料')
        print('    發現該錯誤的ID_NO為: ', end=' ')
        print(*groups_list2, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list2)))
        print("=====================================")
    # ==========================================================================

    # 3. 粗矯站(402) 400系鋼種HRAJ相關製程且長度為2500mm以下的ID，只能在KVS70做。
    mask3_1 = (temp_df['排程站別'] == 402) & (temp_df['鋼種'][3:] == 400) & (temp_df['FINAL流程'][:4] == 'HRAJ') & \
              (temp_df['長度'] != 'None')
    temp_df1 = temp_df[mask3_1]
    mask3_2 = (temp_df1['長度'] <= 2500)
    temp_df2 = temp_df1[mask3_2]
    err_3 = temp_df2.loc[temp_df2['預設機台別'] != 'KVS'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_3.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '400系鋼種HRAJ相關製程且長度為2500mm以下的ID只能在KVS70做'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，400系鋼種HRAJ相關製程且長度為2500mm以下的ID只能在KVS70做'
    # 報錯
    groups3 = err_3.groupby(by='ID_NO')
    groups_list3 = []
    for key in groups3.groups.keys():
        groups_list3.append(key)
    if len(groups_list3) != 0:
        print('錯誤3: 402站發現 400系鋼種HRAJ相關製程且長度為2500mm以下的ID但不在KVS70做 的資料')
        print('     發現該錯誤的ID_NO為: ', end=' ')
        print(*groups_list3, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list3)))
        print("=====================================")
    # ==========================================================================

    # 4. 削皮站(404) S20910鋼種優先在BTH60生產
    mask4 = (temp_df['排程站別'] == 404) & (temp_df['鋼種'] == 'S20910')
    temp_df1 = temp_df[mask4].reset_index(drop=False)
    err_4 = temp_df1.loc[temp_df1['預設機台別'] != 'BTH60'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_4.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = 'S20910鋼種優先在BTH60生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，S20910鋼種優先在BTH60生產'
    # 報錯
    groups4 = err_4.groupby(by='ID_NO')
    groups_list4 = []
    for key in groups4.groups.keys():
        groups_list4.append(key)
    if len(groups_list4) != 0:
        print('錯誤4: 404站發現 S20910鋼種預設機台別不是BTH60 的資料')
        print('    S20910鋼種不在BTH60生產 的ID_NO為: ', end=' ')
        print(*groups_list4, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list4)))
        print("=====================================")
    # ==========================================================================

    # 5. 研磨站(453) 訂單尺寸4.5mm~10mm,訂單長度大於3.6M(不含)須排I6/I17
    mask5_1 = (temp_df['排程站別'] == 453) & (temp_df['訂單長度'] != 'None') & (temp_df['訂單尺寸'] != 'None')
    temp_df1 = temp_df[mask5_1]
    mask5_2 = (temp_df1['訂單長度'] > 3600) & (temp_df1['訂單尺寸'] >= 4.5) & (temp_df1['訂單尺寸'] <= 10)
    temp_df2 = temp_df1[mask5_2]
    err_5 = temp_df2.loc[(temp_df2['預設機台別'] != 'I6') & (temp_df2['預設機台別'] != 'I17')].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_5.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸4.5mm~10mm且訂單長度大於3.6M(不含)須排I6或I17'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸4.5mm~10mm且訂單長度大於3.6M(不含)須排I6或I17'
    # 報錯
    groups5 = err_5.groupby(by='ID_NO')
    groups_list5 = []
    for key in groups5.groups.keys():
        groups_list5.append(key)
    if len(groups_list5) != 0:
        print('錯誤5: 453站發現 訂單尺寸4.5mm~10mm且訂單長度>3.6M不排在I6/I17 的資料')
        print('    不排在I6/I17 的ID_NO為: ', end=' ')
        print(*groups_list5, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list5)))
        print("=====================================")
    # ==========================================================================

    # 6. 矯直切斷站(433) 訂單尺寸>=4mm最好排CYB，<=2.5mm只能排CYA，<=3mm最好排CYA
    mask6_0 = (temp_df['排程站別'] == 433) & (temp_df['訂單尺寸'] != 'None')
    temp_df1 = temp_df[mask6_0]
    # 訂單尺寸>=4mm
    mask6_1 = (temp_df1['訂單尺寸'] >= 4)
    temp_df2 = temp_df1[mask6_1]
    err_6_1 = temp_df2.loc[temp_df2['預設機台別'] != 'CYB'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_6_1.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸4mm(含)以上最好排在CYB'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸4mm(含)以上最好排在CYB'
    # 報錯
    groups6_1 = err_6_1.groupby(by='ID_NO')
    groups_list6_1 = []
    for key in groups6_1.groups.keys():
        groups_list6_1.append(key)
    if len(groups_list6_1) != 0:
        print('錯誤6_1: 433站發現 訂單尺寸大於4mm但不排在CYB 的資料')
        print('    不排在CYB機台 的ID_NO為: ', end=' ')
        print(*groups_list6_1, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list6_1)))
        print("=====================================")

    # 訂單尺寸<=2.5mm
    mask6_2 = (temp_df1['訂單尺寸'] <= 2.5)
    temp_df2 = temp_df1[mask6_2]
    err_6_2 = temp_df2.loc[temp_df2['預設機台別'] != 'CYA'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_6_2.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸2.5mm(含)以下只能排在CYA'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸2.5mm(含)以下只能排在CYA'
    # 報錯
    groups6_2 = err_6_2.groupby(by='ID_NO')
    groups_list6_2 = []
    for key in groups6_2.groups.keys():
        groups_list6_2.append(key)
    if len(groups_list6_2) != 0:
        print('錯誤6_2: 433站發現 訂單尺寸小於2.5mm但不排在CYA 的資料')
        print('    不排在CYA機台 的ID_NO為: ', end=' ')
        print(*groups_list6_2, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list6_2)))
        print("=====================================")

    # 訂單尺寸<=3mm
    mask6_3 = (temp_df1['訂單尺寸'] <= 3)
    temp_df2 = temp_df1[mask6_3]
    err_6_3 = temp_df2.loc[temp_df2['預設機台別'] != 'CYA'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_6_3.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸3mm(含)以下最好排在CYA'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸3mm(含)以下最好排在CYA'
    # 報錯
    groups6_3 = err_6_3.groupby(by='ID_NO')
    groups_list6_3 = []
    for key in groups6_3.groups.keys():
        groups_list6_3.append(key)
    if len(groups_list6_3) != 0:
        print('錯誤6_3: 433站發現 訂單尺寸小於3mm但不排在CYA 的資料')
        print('    不排在CYA機台 的ID_NO為: ', end=' ')
        print(*groups_list6_3, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list6_3)))
        print("=====================================")
    # ==========================================================================

    # 7. 所有站後站的投入時間不得早於前站產出時間
    groups7 = temp_df.groupby(by='ID_NO')
    groups_list7 = []  # 用於遍歷ID_NO
    err_7_list = []  # 用於儲存發現錯誤的index
    for key in groups7.groups.keys():
        groups_list7.append(key)
    for ID in groups_list7:
        temp_df7_1 = temp_df[temp_df['ID_NO'] == ID].astype(str)
        temp_df7_2 = temp_df7_1.sort_values(by='製序')
        Sche_date = temp_df7_2.loc[:, '預計產出時間'].reset_index(drop=False)
        i = 0
        j = 1
        while j < len(Sche_date):
            # a為前一個值，b為後一個值
            if Sche_date.iloc[i, 1] != 'None':
                a = Sche_date.iloc[i, 1]
            else:
                a = 'None'
            if Sche_date.iloc[j, 1] != 'None':
                b = Sche_date.iloc[j, 1]
            else:
                b = 'None'
            if a != 'None' and b != 'None':
                if a > b:
                    err_7_list.append(Sche_date.iloc[j, 0])
            i += 1
            j += 1
    # 設置VIOLATION和AGAINST_RULE
    for k in err_7_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '預計產出時間早於前站'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，預計產出時間早於前站'
    # 報錯
    if len(err_7_list) != 0:
        print('錯誤7: 發現 相同ID中後站預計產出時間早於前站 的資料')
        print('    發現該錯誤的ID_NO, 製序, 排程站別分別為: \n', end=' ')
        # 設置是否顯示所有數據
        # pd.set_option('display.max_rows', None)
        print(temp_df.loc[err_7_list, ['ID_NO', '製序', '排程站別']].reset_index(drop=True))
        print('    共計 {} 個ID'.format(len(err_7_list)))
        print("=====================================")
    # #==========================================================================

    # 8. 直棒冷抽站（410） 成品抽之ID數量需佔該站所有排程ID數量的50%
    mask8_1 = (temp_df['排程站別'] == 410)
    temp_df1 = temp_df[mask8_1].reset_index(drop=False)
    temp_df1['總抽數'] = temp_df1['批次'].str[0]
    temp_df1['當前抽數'] = temp_df1['批次'].str[-1]
    # 計算總ID數量
    groups_list_ID = temp_df1['ID_NO'].tolist()
    num_ID = len(groups_list_ID)
    # 計算成品抽ID數量
    finish_bool = temp_df1.loc[temp_df1['總抽數'] == temp_df1['當前抽數']].reset_index(drop=False)
    groups_list8 = finish_bool['ID_NO'].tolist()
    num_finish = len(groups_list8)
    num_finish_rate = num_finish / num_ID
    # 報錯
    if num_finish < 0.5 * num_ID:
        print('錯誤8: 410站發現 成品抽之ID數量沒有佔該站所有排程ID數量的50%')
        print('    成品抽的數量為: ' + str(num_finish))
        print('    該站所有排程ID數量為: ' + str(num_ID))
        print('    成品抽佔該站總ID數量的 ' + str("%.2f" % num_finish_rate) + '%')
        print("=====================================")
    else:
        print('410直棒冷抽站: ')
        print('    成品抽的數量為: ' + str(num_finish))
        print('    該站所有排程ID數量為: ' + str(num_ID))
        print('    成品抽佔該站總ID數量的 ' + str("%.2f" % num_finish_rate) + '%')
        print("=====================================")
    temp_df1.drop(['總抽數', '當前抽數'], axis=1, inplace=True)
    # ==========================================================================

    # 9. 盤元冷抽站(420)
    # 只能在E1機台生產的條件：
    # （1）生產流程欄位為HRAPJ時
    mask9_1 = (temp_df['排程站別'] == 420) & (temp_df['FINAL流程'].str[:5] == 'HRAPJ')
    temp_df1 = temp_df[mask9_1]
    err_9_1 = temp_df1.loc[temp_df1['預設機台別'] != 'E1'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_1.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = 'FINAL流程欄位為HRAPJ時只能在E1生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，FINAL流程欄位為HRAPJ時只能在E1生產'
    # 報錯
    groups9_1 = err_9_1.groupby(by='ID_NO')
    groups_list9_1 = []
    for key in groups9_1.groups.keys():
        groups_list9_1.append(key)
    if len(groups_list9_1) != 0:
        print('錯誤9_1: 420站發現 FINAL流程為HRAPJ但不在E1機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_1, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_1)))
        print("=====================================")

    # （2）減面率高於20%之棒材時
    mask9_2_1 = (temp_df['排程站別'] == 420) & (temp_df['減面率'] != 'None')
    temp_df1 = temp_df[mask9_2_1]
    mask9_2_2 = (temp_df1['減面率'] > 0.2)
    temp_df2 = temp_df1[mask9_2_2]
    err_9_2 = temp_df2.loc[temp_df2['預設機台別'] != 'E1'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_2.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '減面率高於20%之棒材只能在E1生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，減面率高於20%之棒材只能在E1生產'
    # 報錯
    groups9_2 = err_9_2.groupby(by='ID_NO')
    groups_list9_2 = []
    for key in groups9_2.groups.keys():
        groups_list9_2.append(key)
    if len(groups_list9_2) != 0:
        print('錯誤9_2: 420站發現 減面率高於20%之棒材但不在E1機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_2, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_2)))
        print("=====================================")

    # （3）訂單尺寸欄位為12.65且鋼種欄位為S41501時
    mask9_3 = (temp_df['排程站別'] == 420) & (temp_df['訂單尺寸'] == 12.65) & (temp_df['鋼種'] == 'S41501')
    temp_df1 = temp_df[mask9_3]
    err_9_3 = temp_df1.loc[temp_df1['預設機台別'] != 'E1'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_3.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸欄位為12.65且鋼種欄位為41501只能在E1生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸欄位為12.65且鋼種欄位為41501只能在E1生產'
    # 報錯
    groups9_3 = err_9_3.groupby(by='ID_NO')
    groups_list9_3 = []
    for key in groups9_3.groups.keys():
        groups_list9_3.append(key)
    if len(groups_list9_3) != 0:
        print('錯誤9_3: 420站發現 訂單尺寸欄位為12.65且鋼種欄位為41501但不在E1機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_3, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_3)))
        print("=====================================")

    # （4）訂單尺寸欄位11.9以上且鋼種欄位為17400
    mask9_4_1 = (temp_df['排程站別'] == 420) & (temp_df['訂單尺寸'] != 'None') & (temp_df['鋼種'] == 'S17400')
    temp_df1 = temp_df[mask9_4_1]
    mask9_4_2 = (temp_df1['訂單尺寸'] > 11.9)
    temp_df2 = temp_df1[mask9_4_2]
    err_9_4 = temp_df2.loc[temp_df2['預設機台別'] != 'E1'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_4.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸欄位11.9以上且鋼種欄位為17400只能在E1生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸欄位11.9以上且鋼種欄位為17400只能在E1生產'
    # 報錯
    groups9_4 = err_9_4.groupby(by='ID_NO')
    groups_list9_4 = []
    for key in groups9_4.groups.keys():
        groups_list9_4.append(key)
    if len(groups_list9_4) != 0:
        print('錯誤9_4: 420站發現 訂單尺寸欄位11.9以上且鋼種欄位為17400但不在E1機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_4, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_4)))
        print("=====================================")

    # （5）訂型/尺寸欄位為S20的訂單
    mask9_5_1 = (temp_df['排程站別'] == 420) & (temp_df['訂單尺寸'] != 'None') & (temp_df['訂單形狀'] == 'S')
    temp_df1 = temp_df[mask9_5_1]
    mask9_5_2 = (temp_df1['訂單尺寸'] == 20)
    temp_df2 = temp_df1[mask9_5_2]
    err_9_5 = temp_df2.loc[temp_df2['預設機台別'] != 'E1'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_5.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂型/尺寸欄位為S20的訂單只能在E1生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂型/尺寸欄位為S20的訂單只能在E1生產'
    # 報錯
    groups9_5 = err_9_5.groupby(by='ID_NO')
    groups_list9_5 = []
    for key in groups9_5.groups.keys():
        groups_list9_5.append(key)
    if len(groups_list9_5) != 0:
        print('錯誤9_5: 420站發現 訂型/尺寸欄位為S20的訂單但不在E1機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_5, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_5)))
        print("=====================================")

    # 訂單尺寸欄位為8以上(含)且鋼種欄位為17400只能給E3生產
    mask9_6_1 = (temp_df['排程站別'] == 420) & (temp_df['訂單尺寸'] != 'None') & (temp_df['鋼種'] == 'S17400')
    temp_df1 = temp_df[mask9_6_1]
    mask9_6_2 = (temp_df1['訂單尺寸'] >= 8)
    temp_df2 = temp_df1[mask9_6_2]
    err_9_6 = temp_df2.loc[temp_df2['預設機台別'] != 'E3'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_6.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '訂單尺寸欄位為8以上(含)且鋼種欄位為17400只能給E3生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，訂單尺寸欄位為8以上(含)且鋼種欄位為17400只能給E3生產'
    # 報錯
    groups9_6 = err_9_6.groupby(by='ID_NO')
    groups_list9_6 = []
    for key in groups9_6.groups.keys():
        groups_list9_6.append(key)
    if len(groups_list9_6) != 0:
        print('錯誤9_6: 420站發現 訂單尺寸為8以上(含)且鋼種欄位為17400但不在E3機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_6, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_6)))
        print("=====================================")

    # 投入尺寸欄位23以下(含)只能在E4機台生產
    mask9_7_1 = (temp_df['排程站別'] == 420) & (temp_df['投入尺寸'] != 'None')
    temp_df1 = temp_df[mask9_7_1]
    mask9_7_2 = (temp_df1['投入尺寸'] <= 23)
    temp_df2 = temp_df1[mask9_7_2]
    err_9_7 = temp_df2.loc[temp_df2['預設機台別'] != 'E4'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_7.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'A'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', A'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '投入尺寸欄位23以下(含)只能在E4機台生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，投入尺寸欄位23以下(含)只能在E4機台生產'
    # 報錯
    groups9_7 = err_9_7.groupby(by='ID_NO')
    groups_list9_7 = []
    for key in groups9_7.groups.keys():
        groups_list9_7.append(key)
    if len(groups_list9_7) != 0:
        print('錯誤9_7: 420站發現 投入尺寸欄位23以下(含)但不在E4機台生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_7, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_7)))
        print("=====================================")

    # E1生產尺寸(投入尺寸)最好安排於13~20mm
    mask9_8_1 = (temp_df['排程站別'] == 420) & (temp_df['投入尺寸'] != 'None')
    temp_df1 = temp_df[mask9_8_1]
    mask9_8_2 = (temp_df1['投入尺寸'] >= 13) & (temp_df1['投入尺寸'] <= 20)
    temp_df2 = temp_df1[mask9_8_2]
    err_9_8 = temp_df2.loc[temp_df2['預設機台別'] != 'E1'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_8.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = 'E1生產尺寸(投入尺寸)最好安排於13~20mm'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，E1生產尺寸(投入尺寸)最好安排於13~20mm'
    # 報錯
    groups9_8 = err_9_8.groupby(by='ID_NO')
    groups_list9_8 = []
    for key in groups9_8.groups.keys():
        groups_list9_8.append(key)
    if len(groups_list9_8) != 0:
        print('錯誤9_8: 420站發現 投入尺寸為13~20mm但不在E1生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_8, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_8)))
        print("=====================================")

    # E3最好不要安排生產9mm以上的尺寸(投入尺寸)
    mask9_9_1 = (temp_df['排程站別'] == 420) & (temp_df['投入尺寸'] != 'None')
    temp_df1 = temp_df[mask9_9_1]
    mask9_9_2 = (temp_df1['投入尺寸'] > 9)
    temp_df2 = temp_df1[mask9_9_2]
    err_9_9 = temp_df2.loc[temp_df2['預設機台別'] == 'E3'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_9.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = 'E3最好不要安排生產9mm以上的尺寸(投入尺寸)'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，E3最好不要安排生產9mm以上的尺寸(投入尺寸)'
    # 報錯
    groups9_9 = err_9_9.groupby(by='ID_NO')
    groups_list9_9 = []
    for key in groups9_9.groups.keys():
        groups_list9_9.append(key)
    if len(groups_list9_9) != 0:
        print('錯誤9_9: 420站發現 投入尺寸9mm以上且在E3生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_9, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_9)))
        print("=====================================")

    # E5生產尺寸(投入尺寸)最好安排於18~32mm
    mask9_10_1 = (temp_df['排程站別'] == 420) & (temp_df['投入尺寸'] != 'None')
    temp_df1 = temp_df[mask9_10_1]
    mask9_10_2 = (temp_df1['投入尺寸'] >= 18) & (temp_df1['投入尺寸'] <= 32)
    temp_df2 = temp_df1[mask9_10_2]
    err_9_10 = temp_df2.loc[temp_df2['預設機台別'] != 'E5'].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_9_10.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = 'E5生產尺寸(投入尺寸)最好安排於18~32mm'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，E5生產尺寸(投入尺寸)最好安排於18~32mm'
    # 報錯
    groups9_10 = err_9_10.groupby(by='ID_NO')
    groups_list9_10 = []
    for key in groups9_10.groups.keys():
        groups_list9_10.append(key)
    if len(groups_list9_10) != 0:
        print('錯誤9_10: 420站發現 投入尺寸為18~32mm但不在E5生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list9_10, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list9_10)))
        print("=====================================")
    # ==========================================================================

    # 10. 直棒退火站（401）
    # TC站
    # TODO
    # (1) 生產流程(LINEUP MIC)為”HR”的174鋼種，需於三天內至401生產

    # (2) 鋼種為S17400與S20910只能在1060溫度下生產
    mask10_2_1 = (temp_df['排程站別'] == 401) & (temp_df['預設機台別'] == 'TC')
    temp_df1 = temp_df[mask10_2_1].copy()
    mask10_2_2 = (temp_df1['鋼種'] == 'S17400') | (temp_df1['鋼種'] == 'S20910')
    temp_df2 = temp_df1[mask10_2_2]
    err_10_2 = temp_df2.loc[temp_df2['溫度'] != 1060].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_10_2.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '鋼種為S17400與S20910只能在1060溫度下生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，鋼種為S17400與S20910只能在1060溫度下生產'
    # 報錯
    groups10_2 = err_10_2.groupby(by='ID_NO')
    groups_list10_2 = []
    for key in groups10_2.groups.keys():
        groups_list10_2.append(key)
    if len(groups_list10_2) != 0:
        print('錯誤10_2: 401-TC站發現 鋼種為S17400與S20910但不在1060溫度下生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list10_2, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list10_2)))
        print("=====================================")

    # (3) CD成品抽之ID數量需占總成品抽數量的50%
    temp_df1['總抽數'] = temp_df1['批次'].str[0]
    temp_df1['當前抽數'] = temp_df1['批次'].str[-1]
    # 計算總成品抽ID數量
    temp_df2 = temp_df1.loc[temp_df1['總抽數'] == temp_df1['當前抽數']].reset_index(drop=False)
    groups_list_ID = temp_df2['ID_NO'].tolist()
    Total_finish = len(groups_list_ID)
    # 計算CD成品抽ID數量
    mask10_3 = (temp_df2['產品型態'] == 'CD')
    temp_df3 = temp_df2[mask10_3]
    groups_list10_3 = temp_df3['ID_NO'].tolist()
    CD_finish = len(groups_list10_3)
    CD_finish_rate = CD_finish / Total_finish * 100
    if CD_finish < 0.5 * Total_finish:
        print('錯誤10_3: 401-TC站發現 CD成品抽之ID數量沒有佔總成品抽數量的50%')
        print('    CD成品抽ID數量為: ' + str(CD_finish))
        print('    總成品抽的數量為: ' + str(Total_finish))
        print('    CD成品抽佔總成品抽的 ' + str("%.2f" % CD_finish_rate) + '%')
        print("=====================================")
    else:
        print('401-TC站: ')
        print('    CD成品抽ID數量為: ' + str(CD_finish))
        print('    總成品抽的數量為: ' + str(Total_finish))
        print('    CD成品抽佔總成品抽的 ' + str("%.2f" % CD_finish_rate) + '%')
        print("=====================================")
    temp_df1.drop(['總抽數', '當前抽數'], axis=1, inplace=True)

    # (4) 排程群組間頻率差異<=5
    Frequency = temp_df1.loc[:, '頻率'].reset_index(drop=False)
    err_10_4_list = []
    i = 0
    j = 1
    while j < len(Frequency):
        if Frequency.iloc[i, 1] != 'None':
            a = float(Frequency.iloc[i, 1])
        else:
            a = 'None'
        if Frequency.iloc[j, 1] != 'None':
            b = float(Frequency.iloc[j, 1])
        else:
            b = 'None'
        if a != 'None' and b != 'None':
            if abs(b - a) > 5:
                err_10_4_list.append(Frequency.iloc[j, 0])
        i += 1
        j += 1
    # 設置VIOLATION和AGAINST_RULE
    for k in err_10_4_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'D'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', D'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '排程群組間頻率差異>5'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，排程群組間頻率差異>5'
    # 報錯
    if len(err_10_4_list) != 0:
        print('錯誤10_4: 401-TC站發現 排程群組間頻率差異>5 的資料')
        print('    發現該錯誤的ID_NO為: ', end=' ')
        print(*(temp_df.loc[err_10_4_list, 'ID_NO'].reset_index(drop=True).tolist()), sep=', ')
        print('    共計 {} 個ID'.format(len(err_10_4_list)))
        print("=====================================")

    # BA1站
    # TODO
    # (1) 生產流程(LINEUP MIC)為”HR”的174鋼種，需於三天內至401生產

    # (2) 鋼種為S17400與S20910只能在1060溫度下生產
    mask10_6_1 = (temp_df['排程站別'] == 401) & (temp_df['預設機台別'] == 'BA1')
    temp_df1 = temp_df[mask10_6_1].copy()
    mask10_6_2 = (temp_df1['鋼種'] == 'S17400') | (temp_df1['鋼種'] == 'S20910')
    temp_df2 = temp_df1[mask10_6_2]
    err_10_6 = temp_df2.loc[temp_df2['溫度'] != 1060].reset_index(drop=False)
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_10_6.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '鋼種為S17400與S20910只能在1060溫度下生產'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，鋼種為S17400與S20910只能在1060溫度下生產'
    # 報錯
    groups10_6 = err_10_6.groupby(by='ID_NO')
    groups_list10_6 = []
    for key in groups10_6.groups.keys():
        groups_list10_6.append(key)
    if len(groups_list10_6) != 0:
        print('錯誤10_6: 401-BA1站發現 鋼種為S17400與S20910但不在1060溫度下生產 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list10_6, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list10_6)))
        print("=====================================")
    # ==========================================================================

    # 11. 所有站預計產出時間不得早於生計交期
    mask11_1 = (temp_df['生計交期'] != 'None')
    temp_df1 = temp_df[mask11_1]
    mask11_2 = (temp_df['預計產出時間'] != 'None')
    temp_df2 = temp_df1[mask11_2]
    Delivery_date = temp_df2.loc[:, '生計交期'].reset_index(drop=False)
    Delivery_date['生計交期'] = pd.to_datetime(Delivery_date['生計交期'])
    Sche_date = temp_df2.loc[:, '預計產出時間'].reset_index(drop=False)
    Sche_date['預計產出時間'] = pd.to_datetime(Sche_date['預計產出時間'])
    mask11_3 = (Delivery_date['生計交期'] <= Sche_date['預計產出時間'])
    err_11 = Sche_date[mask11_3]
    # 設置VIOLATION和AGAINST_RULE
    temp_index_list = err_11.loc[:, 'index'].tolist()
    for k in temp_index_list:
        if temp_df.loc[k, 'VIOLATION'] == 'None':
            temp_df.loc[k, 'VIOLATION'] = 'C'
        else:
            temp_df.loc[k, 'VIOLATION'] += ', C'
        if temp_df.loc[k, 'AGAINST_RULE'] == 'None':
            temp_df.loc[k, 'AGAINST_RULE'] = '生計交期早於預計產出時間'
        else:
            temp_df.loc[k, 'AGAINST_RULE'] += '，生計交期早於預計產出時間'
    # 計算非空值的ID數量
    groups11_1 = temp_df2.groupby(by='ID_NO')
    groups_list11_1 = []
    for key in groups11_1.groups.keys():
        groups_list11_1.append(key)
    num_total_ID = len(groups_list11_1)
    # 計算生計交期早於預計產出時間的ID數量
    temp_df3 = temp_df.loc[temp_index_list]
    groups11_2 = temp_df3.groupby(by='ID_NO')
    groups_list11_2 = []
    for key in groups11_2.groups.keys():
        groups_list11_2.append(key)
    num_err_ID = len(groups_list11_2)
    Fine_ID_rate = float()
    if num_total_ID != 0:
        Fine_ID_rate = (num_total_ID - num_err_ID) / num_total_ID * 100
    # 報錯
    if len(groups_list11_2) != 0:
        print('錯誤11: 發現 生計交期早於預計產出時間 的資料')
        print('    發現該錯誤 的ID_NO為: ', end=' ')
        print(*groups_list11_2, sep=', ')
        print('    共計 {} 個ID'.format(len(groups_list11_2)))
        print('    符合交期的比率為 ' + str("%.2f" % Fine_ID_rate) + '%')
        print("=====================================")
    # ==========================================================================

    print('檢查違反規則... done')
    return temp_df


if __name__ == '__main__':
    df = read_file()
    df = df.fillna('None')
    df1 = to_translate_ch(df)
    df2 = check_error(df1)
    df_ans = to_translate(df2)
    df_ans.to_excel('0412完整主檔(版本2)_Check_rule.xlsx', index=False)
    # 計算時間, 結束
    end = time.time()
    # 總共耗時, 單位(秒)
    print('耗時' + str(round(end) - round(start)) + 's')
