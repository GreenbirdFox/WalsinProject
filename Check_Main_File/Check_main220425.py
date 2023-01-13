import pandas as pd
# import numpy as np
import time
import warnings

# ignore warning
warnings.filterwarnings('ignore')

# 計算時間, 起始
start = time.time()


# 讀取主檔
def read_file():
    print('讀取資料主檔中...')
    dfs = pd.read_excel('0425MainFile.xlsx')
    print('讀取資料主檔中... done')
    return dfs


# 欄位基本處理 中轉英
def to_translate_eng(temp_df):
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
        'LINEUP_MIC_NO': 'LINEUP_MIC_NO',
        'FINAL_MIC_NO': 'FINAL_MIC_NO',
        'FINAL流程': 'FINAL_PROCESS',
        '儲區': 'LOC',
        '訂單號碼': 'SALE_ORDER',
        '訂單項次': 'SALE_ITEM',
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
        '生產站別': 'ID_BAR_SHOP_CODE',
        '生計入庫日': 'DATE_PP',
        '營業入庫日': 'DATE_SALES',
        '整單出貨': 'FLAG_WHOLE_ORDER_SHIPMENT',
        '外貨貨櫃編號': 'EXPORT_CABINET_NO',
        '區別': 'SALE_AREA_GROUP',
        '軋延CYCLE': 'PP_CYCLE_NO',
        '最短製程週期時間': 'WORKING_HOURS_MIN',
        '最長製程週期時間': 'WORKING_HOURS_MAX', }, inplace=True)
    return temp_df


# 欄位基本處理 英轉中
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
        'SHOP_CODE': '現況站別',
        'PROC_STATUS': '放行碼',
        'STEEL_TYPE': '鋼種',
        'NEXT_SHOP_CODE': '下站別',
        'NEXT_EQUIP_CODE': '下站機台',
        'PIECE_COUNT': '支數',
        'ACTUAL_LENGTH': '長度',
        'TTL_LENGTH': '總長度',
        'LINEUP_MIC_NO': 'LINEUP_MIC_NO',
        'FINAL_MIC_NO': 'FINAL_MIC_NO',
        'FINAL_PROCESS': 'FINAL流程',
        'LOC': '儲區',
        'SALE_ORDER': '訂單號碼',
        'SALE_ITEM': '訂單項次',
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


# 檢查主檔是否有問題
def check_error(temp_df):
    print('檢查資料主檔中...')
    print("=====================================")
    # 1. 檢查是否有現況站別為 0 情形
    for err in temp_df['現況站別']:
        if err == 0:
            err_1 = temp_df.loc[temp_df['現況站別'] == 0, :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_1.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤1: 現況站別 發現有 0 的資料'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤1: 現況站別 發現有 0 的資料'
            # 報錯
            groups1 = err_1.groupby(by='ID_NO')
            groups_list1 = []
            for key in groups1.groups.keys():
                groups_list1.append(key)
            print('錯誤1: 現況站別 發現有 0 的資料')
            print('    現況站別 發現 0 的ID_NO為: ', end=' ')
            print(*groups_list1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list1)))
            break

    # 2. 檢查現況站別 401 下站別為 401 的錯誤
    mask2 = (temp_df['排程站別'] == 401)
    err_2 = temp_df[mask2].reset_index(drop=False)
    for err in err_2['下站別']:
        if err == 401:
            err_2_1 = err_2.loc[err_2['下站別'] == 401].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_2_1.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤2: 發現有現況站別 401 下站別 401 資料'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤2: 發現有現況站別 401 下站別 401 資料'
            # 報錯
            groups2 = err_2_1.groupby(by='ID_NO')
            groups_list2 = []
            for key in groups2.groups.keys():
                groups_list2.append(key)
            print('錯誤2: 發現有現況站別 401 下站別 401 資料')
            print('    存在該錯誤 的ID_NO為: ', end=' ')
            print(*groups_list2, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list2)))
            break

    # 4. 404、405 產出尺寸不能為空值
    mask4_1 = (temp_df['排程站別'] == 405)
    err_4_1 = temp_df[mask4_1].reset_index(drop=False)
    for err in err_4_1['產出尺寸']:
        if err == 'None':
            mask4_1_1 = (err_4_1['產出尺寸'] == 'None')
            temp = err_4_1[mask4_1_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤4_1: 405 站產出尺寸發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤4_1: 405 站產出尺寸發現 None'
            # 報錯
            groups4_1_1 = temp.groupby(by='ID_NO')
            groups_list4_1_1 = []
            for key in groups4_1_1.groups.keys():
                groups_list4_1_1.append(key)
            print('錯誤4_1: 405 站產出尺寸發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list4_1_1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list4_1_1)))
            break
    for err in err_4_1['產出尺寸']:
        if err == 0:
            mask4_1_2 = (err_4_1['產出尺寸'] == 0)
            temp = err_4_1[mask4_1_2].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤4_1: 405 站產出尺寸發現 0'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤4_1: 405 站產出尺寸發現 0'
            # 報錯
            groups4_1_2 = temp.groupby(by='ID_NO')
            groups_list4_1_2 = []
            for key in groups4_1_2.groups.keys():
                groups_list4_1_2.append(key)
            print('錯誤4_1: 405 站產出尺寸發現 0')
            print('    發現 0 的ID_NO為: ', end=' ')
            print(*groups_list4_1_2, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list4_1_2)))
            break

    mask4_2 = (temp_df['排程站別'] == 404)
    err_4_2 = temp_df[mask4_2].reset_index(drop=False)
    for err in err_4_2['產出尺寸']:
        if err == 'None':
            mask4_2_1 = (err_4_2['產出尺寸'] == 'None')
            temp = err_4_2[mask4_2_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤4_2: 404 站產出尺寸發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤4_2: 404 站產出尺寸發現 None'
            # 報錯
            groups4_2_1 = temp.groupby(by='ID_NO')
            groups_list4_2_1 = []
            for key in groups4_2_1.groups.keys():
                groups_list4_2_1.append(key)
            print('錯誤4_2: 404 站產出尺寸發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list4_2_1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list4_2_1)))
            break
    for err in err_4_2['產出尺寸']:
        if err == 0:
            mask4_2_2 = (err_4_2['產出尺寸'] == 0)
            temp = err_4_2[mask4_2_2].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤4_2: 404 站產出尺寸發現 0'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤4_2: 404 站產出尺寸發現 0'
            # 報錯
            groups4_2_2 = temp.groupby(by='ID_NO')
            groups_list4_2_2 = []
            for key in groups4_2_2.groups.keys():
                groups_list4_2_2.append(key)
            print('錯誤4_2: 404 站產出尺寸發現 0')
            print('    發現 0 的ID_NO為: ', end=' ')
            print(*groups_list4_2_2, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list4_2_2)))
            break

    # 5. 401 溫度與頻率不能為空值
    mask5_1 = (temp_df['排程站別'] == 401) & (temp_df['預設機台別'] == 'TC')
    err_5_1 = temp_df[mask5_1].reset_index(drop=False)
    for err in err_5_1['溫度']:
        if err == 'None':
            mask5_1_1 = (err_5_1['溫度'] == 'None')
            temp = err_5_1[mask5_1_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤5_1: 401 站溫度發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤5_1: 401 站溫度發現 None'
            # 報錯
            groups5_1_1 = temp.groupby(by='ID_NO')
            groups_list5_1_1 = []
            for key in groups5_1_1.groups.keys():
                groups_list5_1_1.append(key)
            print('錯誤5_1: 401 站溫度發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list5_1_1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list5_1_1)))
            break
    for err in err_5_1['溫度']:
        if err == 0:
            mask5_1_2 = (err_5_1['溫度'] == 0)
            temp = err_5_1[mask5_1_2].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤5_1: 401 站溫度發現 0'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤5_1: 401 站溫度發現 0'
            # 報錯
            groups5_1_2 = temp.groupby(by='ID_NO')
            groups_list5_1_2 = []
            for key in groups5_1_2.groups.keys():
                groups_list5_1_2.append(key)
            print('錯誤5_1: 401 站溫度發現 0')
            print('    發現 0 的ID_NO為: ', end=' ')
            print(*groups_list5_1_2, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list5_1_2)))
            break

    err_5_2 = temp_df[mask5_1].reset_index(drop=False)
    for err in err_5_2['頻率']:
        if err == 'None':
            mask5_2_1 = (err_5_2['頻率'] == 'None')
            temp = err_5_2[mask5_2_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤5_2: 401 站頻率發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤5_2: 401 站頻率發現 None'
            # 報錯
            groups5_2_1 = temp.groupby(by='ID_NO')
            groups_list5_2_1 = []
            for key in groups5_2_1.groups.keys():
                groups_list5_2_1.append(key)
            print('錯誤5_2: 401 站頻率發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list5_2_1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list5_2_1)))
            break
    for err in err_5_2['頻率']:
        if err == 0:
            mask5_2_2 = (err_5_2['頻率'] == 0)
            temp = err_5_2[mask5_2_2].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤5_2: 401 站頻率發現 0'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤5_2: 401 站頻率發現 0'
            # 報錯
            groups5_2_2 = temp.groupby(by='ID_NO')
            groups_list5_2_2 = []
            for key in groups5_2_2.groups.keys():
                groups_list5_2_2.append(key)
            print('錯誤5_2: 401 站頻率發現 0')
            print('    發現 0 的ID_NO為: ', end=' ')
            print(*groups_list5_2_2, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list5_2_2)))
            break

    # 6. 急單說明若有資料，急單日期不能為空值
    mask6 = (temp_df['急單說明'] != 'None')
    err_6 = temp_df[mask6].reset_index(drop=False)
    for err in err_6['急單日期']:
        if err == 'None':
            mask6_1 = (err_6['急單日期'] == 'None')
            temp = err_6[mask6_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤6: 急單說明有資料，但急單日期發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤6: 急單說明有資料，但急單日期發現 None'
            # 報錯
            groups6 = temp.groupby(by='ID_NO')
            groups_list6 = []
            for key in groups6.groups.keys():
                groups_list6.append(key)
            print('錯誤6: 急單說明有資料，但急單日期發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list6, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list6)))
            break

    # 7. 除 400 480 490 站外其餘預設機台工時, 機台_01 ~ 機台_10, 最小尺寸, 最大尺寸, 投入尺寸, 產出尺寸皆不能為空值
    mask7_0 = (temp_df['排程站別'] != 400) & (temp_df['排程站別'] != 480) & (temp_df['排程站別'] != 490)
    err_7_1 = temp_df[mask7_0].reset_index(drop=False)
    for err in err_7_1['預設機台工時']:
        if err == 'None':
            mask7_1 = (err_7_1['預設機台工時'] == 'None')
            temp = err_7_1[mask7_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤7_1: 預設機台工時 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤7_1: 預設機台工時 發現 None'
            # 報錯
            groups7_1 = temp.groupby(by='ID_NO')
            groups_list7_1 = []
            for key in groups7_1.groups.keys():
                groups_list7_1.append(key)
            print('錯誤7_1: 預設機台工時 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list7_1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list7_1)))
            break

    err_7_3 = temp_df[mask7_0].reset_index(drop=False)
    for err in err_7_3['最小尺寸']:
        if err == 'None':
            mask7_3 = (err_7_3['最小尺寸'] == 'None')
            temp = err_7_3[mask7_3].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤7_3: 最小尺寸 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤7_3: 最小尺寸 發現 None'
            # 報錯
            groups7_3 = temp.groupby(by='ID_NO')
            groups_list7_3 = []
            for key in groups7_3.groups.keys():
                groups_list7_3.append(key)
            print('錯誤7_3: 最小尺寸 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list7_3, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list7_3)))
            break

    err_7_4 = temp_df[mask7_0].reset_index(drop=False)
    for err in err_7_4['最大尺寸']:
        if err == 'None':
            mask7_4 = (err_7_4['最大尺寸'] == 'None')
            temp = err_7_4[mask7_4].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤7_4: 最大尺寸 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤7_4: 最大尺寸 發現 None'
            # 報錯
            groups7_4 = temp.groupby(by='ID_NO')
            groups_list7_4 = []
            for key in groups7_4.groups.keys():
                groups_list7_4.append(key)
            print('錯誤7_4: 最大尺寸 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list7_4, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list7_4)))
            break

    err_7_5 = temp_df[mask7_0].reset_index(drop=False)
    for err in err_7_5['投入尺寸']:
        if err == 'None':
            mask7_5 = (err_7_5['投入尺寸'] == 'None')
            temp = err_7_5[mask7_5].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤7_5: 投入尺寸 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤7_5: 投入尺寸 發現 None'
            # 報錯
            groups7_5 = temp.groupby(by='ID_NO')
            groups_list7_5 = []
            for key in groups7_5.groups.keys():
                groups_list7_5.append(key)
            print('錯誤7_5: 投入尺寸 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list7_5, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list7_5)))
            break

    err_7_6 = temp_df[mask7_0].reset_index(drop=False)
    for err in err_7_6['產出尺寸']:
        if err == 'None':
            mask7_6 = (err_7_6['產出尺寸'] == 'None')
            temp = err_7_6[mask7_6].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤7_6: 產出尺寸 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤7_6: 產出尺寸 發現 None'
            # 報錯
            groups7_6 = temp.groupby(by='ID_NO')
            groups_list7_6 = []
            for key in groups7_6.groups.keys():
                groups_list7_6.append(key)
            print('錯誤7_6: 產出尺寸 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list7_6, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list7_6)))
            break

    # 8. 除 400 站外其餘預設機台別皆不能為空值
    mask8 = (temp_df['排程站別'] != 400)
    err_8 = temp_df[mask8].reset_index(drop=False)
    for err in err_8['預設機台別']:
        if err == 'None':
            mask8_1 = (err_8['預設機台別'] == 'None')
            temp = err_8[mask8_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤8: 預設機台別 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤8: 預設機台別 發現 None'
            # 報錯
            groups8 = temp.groupby(by='ID_NO')
            groups_list8 = []
            for key in groups8.groups.keys():
                groups_list8.append(key)
            print('錯誤8: 預設機台別 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list8, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list8)))
            break

    # 9. 除 490 站外其餘下站別，下站機台皆不能為空值
    mask9_0 = (temp_df['排程站別'] != 490)

    err_9_1 = temp_df[mask9_0]
    for err in err_9_1['下站別']:
        if err == 'None':
            mask9_1 = (err_9_1['下站別'] == 'None')
            temp = err_9_1[mask9_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤9_1: 下站別 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤9_1: 下站別 發現 None'
            # 報錯
            groups9_1 = temp.groupby(by='ID_NO')
            groups_list9_1 = []
            for key in groups9_1.groups.keys():
                groups_list9_1.append(key)
            print('錯誤9_1: 下站別 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list9_1, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list9_1)))
            break

    err_9_2 = temp_df[mask9_0]
    for err in err_9_2['下站機台']:
        if err == 'None':
            mask9_2 = (err_9_2['下站機台'] == 'None')
            temp = err_9_2[mask9_2].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤9_2: 下站機台 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤9_2: 下站機台 發現 Nonee'
            # 報錯
            groups9_2 = temp.groupby(by='ID_NO')
            groups_list9_2 = []
            for key in groups9_2.groups.keys():
                groups_list9_2.append(key)
            print('錯誤9_2: 下站機台 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list9_2, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list9_2)))
            break

    # 11. 402 403 404 410 420 430 451 453 站預設調機狀態不能為空值
    mask11 = ((temp_df['排程站別'] == 402) | (temp_df['排程站別'] == 403) | (temp_df['排程站別'] == 404) |
              (temp_df['排程站別'] == 410) | (temp_df['排程站別'] == 420) | (temp_df['排程站別'] == 430) |
              (temp_df['排程站別'] == 451) | (temp_df['排程站別'] == 453))
    err_11 = temp_df[mask11].reset_index(drop=False)
    for err in err_11['預設調機狀態']:
        if err == 'None':
            mask11_1 = (err_11['預設調機狀態'] == 'None')
            temp = err_11[mask11_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤11: 預設調機狀態 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤11: 預設調機狀態 發現 None'
            # 報錯
            groups11 = temp.groupby(by='ID_NO')
            groups_list11 = []
            for key in groups11.groups.keys():
                groups_list11.append(key)
            print('錯誤11: 預設調機狀態 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list11, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list11)))
            break

    # 12. ID_NO不能為空值
    for err in temp_df['ID_NO']:
        if err == 'None':
            err_12 = temp_df.loc[temp_df['ID_NO'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_12.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤12: ID_NO 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤12: ID_NO 發現 None'
            # 報錯
            err_12 = temp_df.loc[temp_df['ID_NO'] == 'None', :]
            print('錯誤12: ID_NO 發現 None')
            print('    ID_NO 為 None 的行數為: ', end=' ')
            print(*err_12.index + 2, sep=', ')  # 加2是因為第一行為欄位名稱，第二行index為0
            print('    共計 {} 個ID'.format(len(err_12.index)))
            break

    # 13. 預計投入重量不能為空值
    for err in temp_df['預計投入重量']:
        if err == 'None':
            err_13 = temp_df.loc[temp_df['預計投入重量'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_13.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤13: 預計投入重量 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤13: 預計投入重量 發現 None'
            # 報錯
            groups13 = err_13.groupby(by='ID_NO')
            groups_list13 = []
            for key in groups13.groups.keys():
                groups_list13.append(key)
            print('錯誤13: 預計投入重量 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list13, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list13)))
            break

    # 14. 預計產出重量不能為空值
    for err in temp_df['預計產出重量']:
        if err == 'None':
            err_14 = temp_df.loc[temp_df['預計產出重量'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_14.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤14: 預計產出重量 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤14: 預計產出重量 發現 None'
            # 報錯
            groups14 = err_14.groupby(by='ID_NO')
            groups_list14 = []
            for key in groups14.groups.keys():
                groups_list14.append(key)
            print('錯誤14: 預計產出重量 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list14, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list14)))
            break

    # 15. 製序不能為空值
    for err in temp_df['製序']:
        if err == 'None':
            err_15 = temp_df.loc[temp_df['製序'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_15.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤15: 製序 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤15: 製序 發現 None'
            # 報錯
            groups15 = err_15.groupby(by='ID_NO')
            groups_list15 = []
            for key in groups15.groups.keys():
                groups_list15.append(key)
            print('錯誤15: 製序 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list15, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list15)))
            break

    # 16. 現況製序不能為空值
    for err in temp_df['現況製序']:
        if err == 'None':
            err_16 = temp_df.loc[temp_df['現況製序'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_16.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤16: 現況製序 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤16: 現況製序 發現 None'
            # 報錯
            groups16 = err_16.groupby(by='ID_NO')
            groups_list16 = []
            for key in groups16.groups.keys():
                groups_list16.append(key)
            print('錯誤16: 現況製序 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list16, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list16)))
            break

    # 17. 排程站別不能為空值
    for err in temp_df['排程站別']:
        if err == 'None':
            err_17 = temp_df.loc[temp_df['排程站別'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_17.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤17: 排程站別 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤17: 排程站別 發現 None'
            # 報錯
            groups17 = err_17.groupby(by='ID_NO')
            groups_list17 = []
            for key in groups17.groups.keys():
                groups_list17.append(key)
            print('錯誤17: 排程站別 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list17, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list17)))
            break

    # 18. 現況差異數不能為空值
    for err in temp_df['現況差異數']:
        if err == 'None':
            err_18 = temp_df.loc[temp_df['現況差異數'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_18.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤18: 現況差異數 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤18: 現況差異數 發現 None'
            # 報錯
            groups18 = err_18.groupby(by='ID_NO')
            groups_list18 = []
            for key in groups18.groups.keys():
                groups_list18.append(key)
            print('錯誤18: 現況差異數 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list18, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list18)))
            break

    # 19. 批次不能為空值
    for err in temp_df['批次']:
        if err == 'None':
            err_19 = temp_df.loc[temp_df['批次'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_19.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤19: 批次 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤19: 批次 發現 None'
            # 報錯
            groups19 = err_19.groupby(by='ID_NO')
            groups_list19 = []
            for key in groups19.groups.keys():
                groups_list19.append(key)
            print('錯誤19: 批次 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list19, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list19)))
            break

    # 20. 放行碼不能為空值
    for err in temp_df['放行碼']:
        if err == 'None':
            err_20 = temp_df.loc[temp_df['放行碼'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_20.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤20: 放行碼 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤20: 放行碼 發現 None'
            # 報錯
            groups20 = err_20.groupby(by='ID_NO')
            groups_list20 = []
            for key in groups20.groups.keys():
                groups_list20.append(key)
            print('錯誤20: 放行碼 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list20, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list20)))
            break

    # 21. 鋼種不能為空值
    for err in temp_df['鋼種']:
        if err == 'None':
            err_21 = temp_df.loc[temp_df['鋼種'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_21.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤21: 鋼種 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤21: 鋼種 發現 None'
            # 報錯
            groups21 = err_21.groupby(by='ID_NO')
            groups_list21 = []
            for key in groups21.groups.keys():
                groups_list21.append(key)
            print('錯誤21: 鋼種 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list21, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list21)))
            break

    # 22. LINEUP_MIC_NO不能為空值
    for err in temp_df['LINEUP_MIC_NO']:
        if err == 'None':
            err_22 = temp_df.loc[temp_df['LINEUP_MIC_NO'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_22.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤22: LINEUP_MIC_NO 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤22: LINEUP_MIC_NO 發現 None'
            # 報錯
            groups22 = err_22.groupby(by='ID_NO')
            groups_list22 = []
            for key in groups22.groups.keys():
                groups_list22.append(key)
            print('錯誤22: LINEUP_MIC_NO 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list22, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list22)))
            break

    # 24. FINAL_MIC_NO不能為空值
    for err in temp_df['FINAL_MIC_NO']:
        if err == 'None':
            err_24 = temp_df.loc[temp_df['FINAL_MIC_NO'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_24.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤24: FINAL_MIC_NO 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤24: FINAL_MIC_NO 發現 None'
            # 報錯
            groups24 = err_24.groupby(by='ID_NO')
            groups_list24 = []
            for key in groups24.groups.keys():
                groups_list24.append(key)
            print('錯誤24: FINAL_MIC_NO 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list24, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list24)))
            break

    # 25. FINAL流程不能為空值
    for err in temp_df['FINAL流程']:
        if err == 'None':
            err_25 = temp_df.loc[temp_df['FINAL流程'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_25.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤25: FINAL流程 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤25: FINAL流程 發現 None'
            # 報錯
            groups25 = err_25.groupby(by='ID_NO')
            groups_list25 = []
            for key in groups25.groups.keys():
                groups_list25.append(key)
            print('錯誤25: FINAL流程 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list25, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list25)))
            break

    # 26. 儲區不能為空值
    for err in temp_df['儲區']:
        if err == 'None':
            err_26 = temp_df.loc[temp_df['儲區'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_26.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤26: 儲區 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤26: 儲區 發現 None'
            # 報錯
            groups26 = err_26.groupby(by='ID_NO')
            groups_list26 = []
            for key in groups26.groups.keys():
                groups_list26.append(key)
            print('錯誤26: 儲區 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list26, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list26)))
            break

    # 27. 訂單號碼不能為空值
    for err in temp_df['訂單號碼']:
        if err == 'None':
            err_27 = temp_df.loc[temp_df['訂單號碼'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_27.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤27: 訂單號碼 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤27: 訂單號碼 發現 None'
            # 報錯
            groups27 = err_27.groupby(by='ID_NO')
            groups_list27 = []
            for key in groups27.groups.keys():
                groups_list27.append(key)
            print('錯誤27: 訂單號碼 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list27, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list27)))
            break

    # 28. 訂單項次不能為空值
    for err in temp_df['訂單項次']:
        if err == 'None':
            err_28 = temp_df.loc[temp_df['訂單項次'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_28.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤28: 訂單項次 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤28: 訂單項次 發現 None'
            # 報錯
            groups28 = err_28.groupby(by='ID_NO')
            groups_list28 = []
            for key in groups28.groups.keys():
                groups_list28.append(key)
            print('錯誤28: 訂單項次 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list28, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list28)))
            break

    # 29. 生計交期不能為空值
    for err in temp_df['生計交期']:
        if err == 'None':
            err_29 = temp_df.loc[temp_df['生計交期'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_29.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤29: 生計交期 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤29: 生計交期 發現 None'
            # 報錯
            groups29 = err_29.groupby(by='ID_NO')
            groups_list29 = []
            for key in groups29.groups.keys():
                groups_list29.append(key)
            print('錯誤29: 生計交期 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list29, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list29)))
            break

    # 30. 營業交期不能為空值
    for err in temp_df['營業交期']:
        if err == 'None':
            err_30 = temp_df.loc[temp_df['營業交期'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_30.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤30: 營業交期 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤30: 營業交期 發現 None'
            # 報錯
            groups30 = err_30.groupby(by='ID_NO')
            groups_list30 = []
            for key in groups30.groups.keys():
                groups_list30.append(key)
            print('錯誤30: 營業交期 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list30, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list30)))
            break

    # 31. 總重上限不能為空值
    for err in temp_df['總重上限']:
        if err == 'None':
            err_31 = temp_df.loc[temp_df['總重上限'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_31.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤31: 總重上限 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤31: 總重上限 發現 None'
            # 報錯
            groups31 = err_31.groupby(by='ID_NO')
            groups_list31 = []
            for key in groups31.groups.keys():
                groups_list31.append(key)
            print('錯誤31: 總重上限 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list31, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list31)))
            break

    # 32. 總重下限不能為空值
    for err in temp_df['總重下限']:
        if err == 'None':
            err_32 = temp_df.loc[temp_df['總重下限'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_32.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤32: 總重下限 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤32: 總重下限 發現 None'
            # 報錯
            groups32 = err_32.groupby(by='ID_NO')
            groups_list32 = []
            for key in groups32.groups.keys():
                groups_list32.append(key)
            print('錯誤32: 總重下限 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list32, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list32)))
            break

    # 33. 客戶不能為空值
    for err in temp_df['客戶']:
        if err == 'None':
            err_33 = temp_df.loc[temp_df['客戶'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_33.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤33: 客戶 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤33: 客戶 發現 None'
            # 報錯
            groups33 = err_33.groupby(by='ID_NO')
            groups_list33 = []
            for key in groups33.groups.keys():
                groups_list33.append(key)
            print('錯誤33: 客戶 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list33, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list33)))
            break

    # 34. 料號不能為空值
    for err in temp_df['料號']:
        if err == 'None':
            err_34 = temp_df.loc[temp_df['料號'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_34.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤34: 料號 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤34: 料號 發現 None'
            # 報錯
            groups34 = err_34.groupby(by='ID_NO')
            groups_list34 = []
            for key in groups34.groups.keys():
                groups_list34.append(key)
            print('錯誤34: 料號 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list34, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list34)))
            break

    # 35. 訂單形狀不能為空值
    for err in temp_df['訂單形狀']:
        if err == 'None':
            err_35 = temp_df.loc[temp_df['訂單形狀'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_35.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤35: 訂單形狀 發現 Non'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤35: 訂單形狀 發現 Non'
            # 報錯
            groups35 = err_35.groupby(by='ID_NO')
            groups_list35 = []
            for key in groups35.groups.keys():
                groups_list35.append(key)
            print('錯誤35: 訂單形狀 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list35, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list35)))
            break

    # 36. 訂單尺寸不能為空值
    for err in temp_df['訂單尺寸']:
        if err == 'None':
            err_36 = temp_df.loc[temp_df['訂單尺寸'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_36.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤36: 訂單尺寸 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤36: 訂單尺寸 發現 None'
            # 報錯
            groups36 = err_36.groupby(by='ID_NO')
            groups_list36 = []
            for key in groups36.groups.keys():
                groups_list36.append(key)
            print('錯誤36: 訂單尺寸 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list36, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list36)))
            break

    # 38. 軋延尺寸不能為空值
    for err in temp_df['軋延尺寸']:
        if err == 'None':
            err_38 = temp_df.loc[temp_df['軋延尺寸'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_38.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤38: 軋延尺寸 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤38: 軋延尺寸 發現 None'
            # 報錯
            groups38 = err_38.groupby(by='ID_NO')
            groups_list38 = []
            for key in groups38.groups.keys():
                groups_list38.append(key)
            print('錯誤38: 軋延尺寸 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list38, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list38)))
            break

    # 39. 投入型態不能為空值
    for err in temp_df['投入型態']:
        if err == 'None':
            err_39 = temp_df.loc[temp_df['投入型態'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_39.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤39: 投入型態 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤39: 投入型態 發現 None'
            # 報錯
            groups39 = err_39.groupby(by='ID_NO')
            groups_list39 = []
            for key in groups39.groups.keys():
                groups_list39.append(key)
            print('錯誤39: 投入型態 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list39, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list39)))
            break

    # 40. 產品型態不能為空值
    for err in temp_df['產品型態']:
        if err == 'None':
            err_40 = temp_df.loc[temp_df['產品型態'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_40.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤40: 產品型態 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤40: 產品型態 發現 None'
            # 報錯
            groups40 = err_40.groupby(by='ID_NO')
            groups_list40 = []
            for key in groups40.groups.keys():
                groups_list40.append(key)
            print('錯誤40: 產品型態 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list40, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list40)))
            break

    # 41. 減面率不能為空值
    for err in temp_df['減面率']:
        if err == 'None':
            err_41 = temp_df.loc[temp_df['減面率'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_41.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤41: 減面率 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤41: 減面率 發現 None'
            # 報錯
            groups41 = err_41.groupby(by='ID_NO')
            groups_list41 = []
            for key in groups41.groups.keys():
                groups_list41.append(key)
            print('錯誤41: 減面率 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list41, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list41)))
            break

    # 42. 最晚投入日(悲觀)不能為空值
    for err in temp_df['最晚投入日(悲觀)']:
        if err == 'None':
            err_42 = temp_df.loc[temp_df['最晚投入日(悲觀)'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_42.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤42: 最晚投入日(悲觀) 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤42: 最晚投入日(悲觀) 發現 None'
            # 報錯
            groups42 = err_42.groupby(by='ID_NO')
            groups_list42 = []
            for key in groups42.groups.keys():
                groups_list42.append(key)
            print('錯誤42: 最晚投入日(悲觀) 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list42, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list42)))
            break

    # 43. 最晚投入日(樂觀)不能為空值
    for err in temp_df['最晚投入日(樂觀)']:
        if err == 'None':
            err_43 = temp_df.loc[temp_df['最晚投入日(樂觀)'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_43.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤43: 最晚投入日(樂觀) 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤43: 最晚投入日(樂觀) 發現 None'
            # 報錯
            groups43 = err_43.groupby(by='ID_NO')
            groups_list43 = []
            for key in groups43.groups.keys():
                groups_list43.append(key)
            print('錯誤43: 最晚投入日(樂觀) 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list43, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list43)))
            break

    # 44. 生產站別不能為空值
    for err in temp_df['生產站別']:
        if err == 'None':
            err_44 = temp_df.loc[temp_df['生產站別'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_44.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤44: 生產站別 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤44: 生產站別 發現 None'
            # 報錯
            groups44 = err_44.groupby(by='ID_NO')
            groups_list44 = []
            for key in groups44.groups.keys():
                groups_list44.append(key)
            print('錯誤44: 生產站別 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list44, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list44)))
            break

    # 45. 區別不能為空值
    for err in temp_df['區別']:
        if err == 'None':
            err_45 = temp_df.loc[temp_df['區別'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_45.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤45: 區別 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤45: 區別 發現 None'
            # 報錯
            groups45 = err_45.groupby(by='ID_NO')
            groups_list45 = []
            for key in groups45.groups.keys():
                groups_list45.append(key)
            print('錯誤45: 區別 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list45, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list45)))
            break

    # 46. 計畫量不能為空值
    for err in temp_df['計畫量']:
        if err == 'None':
            err_46 = temp_df.loc[temp_df['計畫量'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_46.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤46: 計畫量 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤46: 計畫量 發現 None'
            # 報錯
            groups46 = err_46.groupby(by='ID_NO')
            groups_list46 = []
            for key in groups46.groups.keys():
                groups_list46.append(key)
            print('錯誤46: 計畫量 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list46, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list46)))
            break

    # 47. 產率不能為空值, 除480與490站外
    mask47 = (temp_df['排程站別'] != 480) & (temp_df['排程站別'] != 490)
    err_47_1 = temp_df[mask47].reset_index(drop=False)
    for err in err_47_1['產率']:
        if err == 'None':
            err_47_2 = err_47_1.loc[temp_df['產率'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_47_2.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤47: 產率 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤47: 產率 發現 None'
            # 報錯
            groups47 = err_47_2.groupby(by='ID_NO')
            groups_list47 = []
            for key in groups47.groups.keys():
                groups_list47.append(key)
            print('錯誤47: 產率 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list47, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list47)))
            break

    # 48. 支數不能為空值, 420、430、403、431、432除外
    mask48 = ((temp_df['排程站別'] != 420) & (temp_df['排程站別'] != 430) & (temp_df['排程站別'] != 403) &
              (temp_df['排程站別'] != 431) & (temp_df['排程站別'] != 432))
    err_48_1 = temp_df[mask48].reset_index(drop=False)
    for err in err_48_1['支數']:
        if err == 'None':
            err_48_2 = err_48_1.loc[temp_df['支數'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_48_2.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤48: 支數 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤48: 支數 發現 None'
            # 報錯
            groups48 = err_48_2.groupby(by='ID_NO')
            groups_list48 = []
            for key in groups48.groups.keys():
                groups_list48.append(key)
            print('錯誤48: 支數 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list48, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list48)))
            break

    # 49. 長度不能為空值
    for err in temp_df['長度']:
        if err == 'None':
            err_49 = temp_df.loc[temp_df['長度'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_49.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤49: 長度 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤49: 長度 發現 None'
            # 報錯
            groups49 = err_49.groupby(by='ID_NO')
            groups_list49 = []
            for key in groups49.groups.keys():
                groups_list49.append(key)
            print('錯誤49: 長度 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list49, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list49)))
            break

    # 50. 總長度不能為空值
    for err in temp_df['總長度']:
        if err == 'None':
            err_50 = temp_df.loc[temp_df['總長度'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_50.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤50: 總長度 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤50: 總長度 發現 None'
            # 報錯
            groups50 = err_50.groupby(by='ID_NO')
            groups_list50 = []
            for key in groups50.groups.keys():
                groups_list50.append(key)
            print('錯誤50: 總長度 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list50, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list50)))
            break

    # 53. 除490站外產出型態不能為空值
    mask53 = (temp_df['排程站別'] != 490)
    err_53 = temp_df[mask53].reset_index(drop=False)
    for err in err_53['產出型態']:
        if err == 'None':
            mask53_1 = (err_53['產出型態'] == 'None')
            temp = err_53[mask53_1].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = temp.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤53: 產出型態 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤53: 產出型態 發現 None'
            # 報錯
            groups53 = temp.groupby(by='ID_NO')
            groups_list53 = []
            for key in groups53.groups.keys():
                groups_list53.append(key)
            print('錯誤53: 產出型態 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list53, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list53)))
            break

    # 54. 最短製程週期時間不能為空值
    for err in temp_df['最短製程週期時間']:
        if err == 'None':
            err_54 = temp_df.loc[temp_df['最短製程週期時間'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_54.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤54: 最短製程週期時間 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤54: 最短製程週期時間 發現 None'
            # 報錯
            groups54 = err_54.groupby(by='ID_NO')
            groups_list54 = []
            for key in groups54.groups.keys():
                groups_list54.append(key)
            print('錯誤54: 最短製程週期時間 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list54, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list54)))
            break

    # 55. 最長製程週期時間不能為空值
    for err in temp_df['最長製程週期時間']:
        if err == 'None':
            err_55 = temp_df.loc[temp_df['最長製程週期時間'] == 'None', :].reset_index(drop=False)
            # 設置SCHEDULE_ABLE和ERROR_FACT
            temp_index_list = err_55.loc[:, 'index'].tolist()
            for k in temp_index_list:
                temp_df.loc[k, 'SCHEDULE_ABLE'] = 0
                if temp_df.loc[k, 'ERROR_FACT'] == 'None':
                    temp_df.loc[k, 'ERROR_FACT'] = '錯誤55: 最長製程週期時間 發現 None'
                else:
                    temp_df.loc[k, 'ERROR_FACT'] += '，錯誤55: 最長製程週期時間 發現 None'
            # 報錯
            groups55 = err_55.groupby(by='ID_NO')
            groups_list55 = []
            for key in groups55.groups.keys():
                groups_list55.append(key)
            print('錯誤55: 最長製程週期時間 發現 None')
            print('    發現 None 的ID_NO為: ', end=' ')
            print(*groups_list55, sep=', ')
            print('    共計 {} 個ID'.format(len(groups_list55)))
            break


    print("=====================================")

    print('檢查資料主檔... done')
    return temp_df


if __name__ == '__main__':
    df = read_file()
    df = df.fillna('None')
    df1 = df['SCHEDULE_ABLE'] = 1
    df2 = to_translate_ch(df)
    df3 = check_error(df2)
    df_ans = to_translate_eng(df3)
    df_ans.to_excel("0425主檔_Check_main.xlsx", index=False)
    # 計算時間, 結束
    end = time.time()
    # 總共耗時, 單位(秒)
    print('耗時' + str(round(end) - round(start)) + 's')
