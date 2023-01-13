import pandas as pd
import numpy as np
import time


# Start timing
start = time.time()


def read_file():
    df_main = pd.read_excel('MasterFile0803.xlsx')
    df_downtime = pd.read_excel('MachineDowntime.xlsx')
    df_machine_cp = pd.read_excel('MachineCapacity0902.xlsx')
    return df_main, df_downtime, df_machine_cp


# Translate the column name main file from Chinese to English
def to_translate(temp_df):
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


# Translate the column name main file from English to Chinese
def to_translate_ch(temp_df):
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


# Get the data within a certain time
def filter_time(temp_df, shop_code, N):
    temp_df = temp_df.fillna('None')
    mask_1 = (temp_df['現況差異數'] != 'None') & (temp_df['產出尺寸'] != 'None')
    temp_df = temp_df[mask_1].reset_index(drop=True)
    mask_2 = (temp_df['排程站別'] == shop_code) & (temp_df['現況差異數'] <= N)
    temp_df = temp_df[mask_2].reset_index(drop=True)

    return temp_df


# Process the column
def process_column(temp_df):
    # Drop EQUIP_CODE column
    temp_df.drop('預設機台別', inplace=True, axis=1)
    col_name = temp_df.columns.tolist()
    col_name.insert(col_name.index('排程站別') + 1, '預設機台別')
    temp_df = temp_df.reindex(columns=col_name)
    temp_df['預設機台別'] = 'None'

    return temp_df


def assign_available_machine(temp_df, temp_cp, shop_code):
    # Process the main file first
    if shop_code == 401 or 402 or 403 or 411 or 420:
        # Add a new column for steel type and group the steel type
        temp_df['鋼種_type'] = 'None'
        for i in range(len(temp_df)):
            if temp_df.loc[i, '鋼種'][:2] == 'S2':
                temp_df.loc[i, '鋼種_type'] = 'S2XX'
            elif temp_df.loc[i, '鋼種'][:2] == 'S3':
                temp_df.loc[i, '鋼種_type'] = 'S3XX'
            elif temp_df.loc[i, '鋼種'][:2] == 'S4':
                temp_df.loc[i, '鋼種_type'] = 'S4XX'
            elif temp_df.loc[i, '鋼種'][:4] == 'S174':
                temp_df.loc[i, '鋼種_type'] = 'S174'
    # Add a new column for assigning available machine for the order
    temp_df['可分配機台'] = str()

    # Process the machine capacity file
    # Create a dataframe with index
    df_unstack = temp_cp.reset_index()
    # Loop over SHOP_CODE(a), PROCESS_CODE(b), SHAPE_TYPE(c), STEEL_TYPE(d) in sequence
    # Loop over SHOP_CODE(a)
    Shop_code_list = df_unstack['SHOP_CODE'].drop_duplicates().to_list()
    for a in Shop_code_list:
        # Loop over PROCESS_CODE(b) when SHOP_CODE is same
        Process_code_series = df_unstack.loc[df_unstack['SHOP_CODE'] == a, 'PROCESS_CODE']
        Process_code_list = Process_code_series.drop_duplicates().to_list()
        for b in Process_code_list:
            # Loop over SHAPE_TYPE(c) when PROCESS_CODE is same
            Shape_type_series = df_unstack.loc[
                (df_unstack['SHOP_CODE'] == a) & (df_unstack['PROCESS_CODE'] == b), 'SHAPE_TYPE']
            Shape_type_list = Shape_type_series.drop_duplicates().to_list()
            for c in Shape_type_list:
                # Loop over STEEL_TYPE(d) when SHAPE_TYPE is same
                Grade_group_series = df_unstack.loc[
                    (df_unstack['SHOP_CODE'] == a) & (df_unstack['PROCESS_CODE'] == b) &
                    (df_unstack['SHAPE_TYPE'] == c), 'GRADE_GROUP']
                Grade_group_list = Grade_group_series.drop_duplicates().to_list()
                # When there is no STEEL_TYPE
                if pd.isnull(Grade_group_list).any():
                    # Read the min and max of diameter and length of the order
                    Dia_min = temp_cp.loc[(a, b, c, None), 'CAPABILITY_DIA_MIN'].to_list()
                    Dia_max = temp_cp.loc[(a, b, c, None), 'CAPABILITY_DIA_MAX'].to_list()
                    Length_min = temp_cp.loc[(a, b, c, None), 'CAPABILITY_LENGTH_MIN'].to_list()
                    Length_max = temp_cp.loc[(a, b, c, None), 'CAPABILITY_LENGTH_MAX'].to_list()
                    Equip_code = temp_cp.loc[(a, b, c, None), 'EQUIP_CODE'].to_list()

                    # Filter the order file
                    mask = (temp_df['排程站別'] == a) & (temp_df['製程碼'] == b) & (temp_df['產出型態'] == c)
                    temp_df_1 = temp_df.loc[mask]
                    # Loop the number of min length
                    for i in range(len(Length_min)):
                        # Loop the order file
                        for j in temp_df_1.index:
                            # Assign the order ID to the machine that are capable to produce it
                            if (temp_df.loc[j, '產出尺寸'] >= Dia_min[i]) & \
                                    (temp_df.loc[j, '產出尺寸'] <= Dia_max[i]) & \
                                    (temp_df.loc[j, '長度'] >= Length_min[i]) & \
                                    (temp_df.loc[j, '長度'] <= Length_max[i]):
                                # 若可分配機台中已有該機台則不再添加
                                if Equip_code[i] not in temp_df.loc[j, '可分配機台']:
                                    temp_df.loc[j, '可分配機台'] += Equip_code[i]

                # When there is a STEEL_TYPE
                else:
                    for d in Grade_group_list:
                        # Read the min and max of diameter and length of the order
                        Dia_min = temp_cp.loc[(a, b, c, d), 'CAPABILITY_DIA_MIN'].to_list()
                        Dia_max = temp_cp.loc[(a, b, c, d), 'CAPABILITY_DIA_MAX'].to_list()
                        Length_min = temp_cp.loc[(a, b, c, d), 'CAPABILITY_LENGTH_MIN'].to_list()
                        Length_max = temp_cp.loc[(a, b, c, d), 'CAPABILITY_LENGTH_MAX'].to_list()
                        Equip_code = temp_cp.loc[(a, b, c, d), 'EQUIP_CODE'].to_list()

                        # Filter the order file
                        mask = (temp_df['排程站別'] == a) & (temp_df['製程碼'] == b) & (temp_df['產出型態'] == c) & \
                               (temp_df['鋼種_type'] == d)
                        temp_df_1 = temp_df.loc[mask]
                        # Loop the number of min length
                        for i in range(len(Length_min)):
                            # Loop the order file
                            for j in temp_df_1.index:
                                # Assign the order ID to the machine that are capable to produce it
                                if (temp_df.loc[j, '產出尺寸'] >= Dia_min[i]) & \
                                        (temp_df.loc[j, '產出尺寸'] <= Dia_max[i]) & \
                                        (temp_df.loc[j, '長度'] >= Length_min[i]) & \
                                        (temp_df.loc[j, '長度'] <= Length_max[i]):
                                    if Equip_code[i] not in temp_df.loc[j, '可分配機台']:
                                        temp_df.loc[j, '可分配機台'] += Equip_code[i]
                                else:
                                    temp_df.loc[j, '可分配機台'] = temp_df.loc[j, '預設機台別']
    temp_df = process_column(temp_df)
    return temp_df


# Calculate the number of times that machine is down and its total downtime
def cal_downtime(temp_downtime, shop_code, machine_list, hr_list, down_times, period):
    if shop_code == 420:
        E1_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E1', '停機累計時間(hr)'])
        E2_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E2', '停機累計時間(hr)'])
        E3_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E3', '停機累計時間(hr)'])
        E4_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E4', '停機累計時間(hr)'])
        E5_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E5', '停機累計時間(hr)'])
        for i in hr_list:
            if machine_list[hr_list.index(i)] == 'E1':
                if (down_times[0] == 1) & (i / (60 * (24*period-E1_down) * down_times[0]) >= 1):
                    hr_list[machine_list.index('E1')] += E1_down * 60
                    down_times[0] += 1
                elif (down_times[0] > 1) & ((i - 60 * 24 * period * (down_times[0] - 1)) /
                                            (60 * (24*period-E1_down) * down_times[0]) >= 1):
                    hr_list[machine_list.index('E1')] += E1_down * 60
                    down_times[0] += 1
            elif machine_list[hr_list.index(i)] == 'E2':
                if (down_times[1] == 1) & (i / (60 * (24*period-E2_down) * down_times[1]) >= 1):
                    hr_list[machine_list.index('E2')] += E2_down * 60
                    down_times[1] += 1
                elif (down_times[1] > 1) & ((i - 60 * 24 * period * (down_times[1] - 1)) /
                                            (60 * (24*period-E2_down) * down_times[1]) >= 1):
                    hr_list[machine_list.index('E2')] += E2_down * 60
                    down_times[1] += 1
            elif machine_list[hr_list.index(i)] == 'E3':
                if (down_times[2] == 1) & (i / (60 * (24*period-E3_down) * down_times[2]) >= 1):
                    hr_list[machine_list.index('E3')] += E3_down * 60
                    down_times[2] += 1
                elif (down_times[2] > 1) & ((i - 60 * 24 * period *
                                             (down_times[2] - 1)) / (60 * (24*period-E3_down) * down_times[2]) >= 1):
                    hr_list[machine_list.index('E3')] += E3_down * 60
                    down_times[2] += 1
            elif machine_list[hr_list.index(i)] == 'E4':
                if (down_times[3] == 1) & (i / (60 * (24*period-E4_down) * down_times[3]) >= 1):
                    hr_list[machine_list.index('E4')] += E4_down * 60
                    down_times[3] += 1
                elif (down_times[3] > 1) & ((i - 60 * 24 * period *
                                             (down_times[3] - 1)) / (60 * (24*period-E4_down) * down_times[3]) >= 1):
                    hr_list[machine_list.index('E4')] += E4_down * 60
                    down_times[3] += 1
            elif machine_list[hr_list.index(i)] == 'E5':
                if (down_times[4] == 1) & (i / (60 * (24*period-E5_down) * down_times[4]) >= 1):
                    hr_list[machine_list.index('E5')] += E5_down * 60
                    down_times[4] += 1
                elif (down_times[4] > 1) & ((i - 60 * 24 * period *
                                             (down_times[4] - 1)) / (60 * (24*period-E5_down) * down_times[4]) >= 1):
                    hr_list[machine_list.index('E3')] += E5_down * 60
                    down_times[4] += 1

    return hr_list, down_times


# Assign the order to machine and calculate the workload in hour for each machine
def cal_working_hours(temp_df, temp_downtime, shop_code, down_period):
    if shop_code == 420:
        # Initialize the workload of each machine in this station
        E1 = 0
        E2 = 0
        E3 = 0
        E4 = 0
        E5 = 0

        # Initialize the number of time(s) that the machine is down
        n1 = 1
        n2 = 1
        n3 = 1
        n4 = 1
        n5 = 1

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['E1', 'E2', 'E3', 'E4', 'E5']
        hr_list = [E1, E2, E3, E4, E5]
        down_times = [n1, n2, n3, n4, n5]

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'E':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'E'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                # Define lists to store the available machine and workload
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            hr_list, down_times = cal_downtime(temp_downtime, shop_code, machine_list, hr_list,
                                                               down_times, down_period)
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']
                    hr_list, down_times = cal_downtime(temp_downtime, shop_code, machine_list, hr_list, down_times,
                                                       down_period)

        down_times[:] = [x - 1 for x in down_times]
        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台', '鋼種_type'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 453:
        # Initialize the workload of each machine in this station
        I01 = 0
        I02 = 0
        I13 = 0
        I9 = 0
        I6 = 0
        I17 = 0
        I18 = 0
        I19 = 0
        I20 = 0
        I21 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['I01', 'I02', 'I13', 'I9', 'I6', 'I17', 'I18', 'I19', 'I20', 'I21']
        hr_list = [I01, I02, I13, I9, I6, I17, I18, I19, I20, I21]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'I':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'I'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 452:
        # Initialize the workload of each machine in this station
        I5 = 0
        I15 = 0
        I16 = 0
        BF2 = 0
        BF3 = 0
        BF5 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['I5', 'I15', 'I16', 'BF2', 'BF3', 'BF5']
        hr_list = [I5, I15, I16, BF2, BF3, BF5]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (df_2.loc[row, '可分配機台'][i] == 'I') or (df_2.loc[row, '可分配機台'][i] == 'B'):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'I') or (temp_df.loc[row, '可分配機台'][j] == 'B')):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 421:
        # Initialize the workload of each machine in this station
        CHS = 0
        CH0 = 0
        Manual = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['CHS', 'CH0', '手動']
        hr_list = [CHS, CH0, Manual]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (temp_df.loc[row, '可分配機台'][i] == 'C') or (temp_df.loc[row, '可分配機台'][i] == '手'):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'C') or (temp_df.loc[row, '可分配機台'][j] == '手')):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 422:
        # Initialize the workload of each machine in this station
        DB = 0
        EC = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['DB', 'EC']
        hr_list = [DB, EC]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (temp_df.loc[row, '可分配機台'][i] == 'D') or (temp_df.loc[row, '可分配機台'][i] == 'E'):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'D') or (temp_df.loc[row, '可分配機台'][j] == 'E')):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 460:
        # Initialize the workload of each machine in this station
        C1 = 0
        C3 = 0
        C6 = 0
        C7 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['C1', 'C3', 'C6', 'C7']
        hr_list = [C1, C3, C6, C7]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'C':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'C'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 461:
        # Initialize the workload of each machine in this station
        BF0 = 0
        BF1 = 0
        BF4 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['BF0', 'BF1', 'BF4']
        hr_list = [BF0, BF1, BF4]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'B':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'B'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 401:
        # Initialize the workload of each machine in this station
        TC = 0
        BA1 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['TC', 'BA1']
        hr_list = [TC, BA1]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (temp_df.loc[row, '可分配機台'][i] == 'T') or (temp_df.loc[row, '可分配機台'][i] == 'B'):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'T') or (temp_df.loc[row, '可分配機台'][j] == 'B')):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台', '鋼種_type'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 402:
        # Initialize the workload of each machine in this station
        KVS = 0
        SM80 = 0
        SM165 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['KVS', 'SM80', 'SM165']
        hr_list = [KVS, SM80, SM165]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (temp_df.loc[row, '可分配機台'][i] == 'K') or \
                            ((i <= len(temp_df.loc[row, '可分配機台']) - 2) and (
                                    temp_df.loc[row, '可分配機台'][i:i + 2] == 'SM')):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'K') or
                                   ((j <= len(temp_df.loc[row, '可分配機台']) - 2) and (
                                           temp_df.loc[row, '可分配機台'][j:j + 2] == 'SM'))):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台', '鋼種_type'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 404:
        # Initialize the workload of each machine in this station
        BTH60 = 0
        S80 = 0
        PM160 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['BTH60', 'S80', 'PM160']
        hr_list = [BTH60, S80, PM160]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (temp_df.loc[row, '可分配機台'][i] == 'B') or \
                            (temp_df.loc[row, '可分配機台'][i] == 'S') or (temp_df.loc[row, '可分配機台'][i] == 'P'):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'B') or
                                   (temp_df.loc[row, '可分配機台'][j] == 'S') or (temp_df.loc[row, '可分配機台'][j] == 'P')):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 405:
        # Initialize the workload of each machine in this station
        C2 = 0
        A4 = 0
        A6 = 0
        A8 = 0
        CF = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['C2', 'A4', 'A6', 'A8', 'CF']
        hr_list = [C2, A4, A6, A8, CF]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if (temp_df.loc[row, '可分配機台'][i] == 'C') or (temp_df.loc[row, '可分配機台'][i] == 'A'):
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & ((temp_df.loc[row, '可分配機台'][j] == 'C') or (temp_df.loc[row, '可分配機台'][j] == 'A')):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 406:
        # Initialize the workload of each machine in this station
        CB0 = 0
        CB1 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['CB0', 'CB1']
        hr_list = [CB0, CB1]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'C':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'C'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 410:
        # Initialize the workload of each machine in this station
        D2 = 0
        D5 = 0
        D6 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['D2', 'D5', 'D6']
        hr_list = [D2, D5, D6]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'D':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'D'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 450:
        # Initialize the workload of each machine in this station
        H9 = 0
        H10 = 0
        H11 = 0
        H12 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['H9', 'H10', 'H11', 'H12']
        hr_list = [H9, H10, H11, H12]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'H':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'H'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 451:
        # Initialize the workload of each machine in this station
        K4 = 0
        K5 = 0
        K8 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['K4', 'K5', 'K8']
        hr_list = [K4, K5, K8]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'K':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'K'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 430:
        # Initialize the workload of each machine in this station
        E10 = 0
        E11 = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['E10', 'E11']
        hr_list = [E10, E11]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'E':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'E'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times

    if shop_code == 433:
        # Initialize the workload of each machine in this station
        CYA = 0
        CYB = 0

        # Sort the data by the number of available machine
        temp_df['LEN'] = np.nan
        for row_loc in range(len(temp_df)):
            if temp_df.loc[row_loc, '可分配機台'] != 'None':
                temp_df.loc[row_loc, 'LEN'] = len(temp_df.loc[row_loc, '可分配機台'])
        temp_df.sort_values(by='LEN', inplace=True)

        machine_list = ['CYA', 'CYB']
        hr_list = [CYA, CYB]
        down_times = []

        for row in temp_df.index:
            if (temp_df.loc[row, '可分配機台'] != 'None') & (temp_df.loc[row, '預設機台工時'] != 'None'):
                index_start = []
                index_end = []
                for i in range(len(temp_df.loc[row, '可分配機台'])):
                    if temp_df.loc[row, '可分配機台'][i] == 'C':
                        index_start.append(i)
                for j in range(len(temp_df.loc[row, '可分配機台'])):
                    if (j != 0) & (temp_df.loc[row, '可分配機台'][j] == 'C'):
                        index_end.append(j)
                    elif j == len(temp_df.loc[row, '可分配機台']) - 1:
                        index_end.append(j + 1)
                ava_machine = []
                ava_hr = []
                for ele in range(len(index_start)):
                    m = index_start[ele]
                    n = index_end[ele]
                    ava_machine.append(temp_df.loc[row, '可分配機台'][m:n])
                    ava_hr.append(hr_list[machine_list.index(str(temp_df.loc[row, '可分配機台'][m:n]))])
                # Check if this is an order ID which has the same output size and shape as the previous IDs
                # If yes, then assign this order to the same machine
                sameID_bool = False
                if not sameID_bool:
                    for pre_row in temp_df.index[:temp_df.index.get_loc(row)]:
                        if (temp_df.loc[pre_row, '預設機台別'] in ava_machine) \
                                & (temp_df.loc[row, '產出尺寸'] == temp_df.loc[pre_row, '產出尺寸']):
                            hr_list[machine_list.index(str(temp_df.loc[pre_row, '預設機台別']))] += temp_df.loc[
                                row, '預設機台工時']
                            temp_df.loc[row, '預設機台別'] = temp_df.loc[pre_row, '預設機台別']
                            sameID_bool = True
                            break
                # If not, then assign this order to the machine with lightest workload
                if not sameID_bool:
                    # Find the machine that has the lightest workload
                    min_hr = min(ava_hr)
                    equip_code = machine_list[hr_list.index(min_hr)]
                    temp_df.loc[row, '預設機台別'] = equip_code
                    hr_list[machine_list.index(equip_code)] += temp_df.loc[row, '預設機台工時']

        temp_df.sort_values(['ID_NO'], inplace=True)
        temp_df.drop(['LEN', '可分配機台'], axis=1, inplace=True)

        return temp_df, hr_list, down_times


# Add the data after being processed and balanced to the original order file
def add_to_main(df_final, df_main, shop_code, N):
    mask_1 = (df_main['現況差異數'] != 'None')
    df_main_1 = df_main[mask_1].reset_index(drop=True)
    mask_2 = (df_main_1['排程站別'] == shop_code) & (df_main_1['現況差異數'] <= N)
    df_main_2 = df_main.drop(df_main_1[mask_2].index)

    temp_df = pd.concat([df_final, df_main_2], ignore_index=True)
    temp_df.sort_values(['ID_NO'], inplace=True)
    temp_df = temp_df.reset_index(drop=True)

    return temp_df


def print_result(hr_list, temp_downtime, shop_code, down_times):
    if shop_code == 420:
        E1_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E1', '停機累計時間(hr)'])
        E2_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E2', '停機累計時間(hr)'])
        E3_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E3', '停機累計時間(hr)'])
        E4_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E4', '停機累計時間(hr)'])
        E5_down = float(temp_downtime.loc[temp_downtime['機台'] == 'E5', '停機累計時間(hr)'])
        table = [['E1', int(hr_list[0]), down_times[0], format(down_times[0] * E1_down * 60, '.1f')],
                 ['E2', int(hr_list[1]), down_times[1], format(down_times[1] * E2_down * 60, '.1f')],
                 ['E3', int(hr_list[2]), down_times[2], format(down_times[2] * E3_down * 60, '.1f')],
                 ['E4', int(hr_list[3]), down_times[3], format(down_times[3] * E4_down * 60, '.1f')],
                 ['E5', int(hr_list[4]), down_times[4], format(down_times[4] * E5_down * 60, '.1f')]]
        df = pd.DataFrame(table, columns=['機台', '總工時(min)', '停機次數', '停機累計時間(min)'])
        print(df)

    if shop_code == 453:
        print('I01機台總工時：' + str(int(hr_list[0])))
        print('I02機台總工時：' + str(int(hr_list[1])))
        print('I13機台總工時：' + str(int(hr_list[2])))
        print('I9 機台總工時：' + str(int(hr_list[3])))
        print('I6 機台總工時：' + str(int(hr_list[4])))
        print('I17機台總工時：' + str(int(hr_list[5])))
        print('I18機台總工時：' + str(int(hr_list[6])))
        print('I19機台總工時：' + str(int(hr_list[7])))
        print('I20機台總工時：' + str(int(hr_list[8])))
        print('I21機台總工時：' + str(int(hr_list[9])))

    if shop_code == 452:
        print('I5機台總工時：' + str(int(hr_list[0])))
        print('I15機台總工時：' + str(int(hr_list[1])))
        print('I16機台總工時：' + str(int(hr_list[2])))
        print('BF2機台總工時：' + str(int(hr_list[3])))
        print('BF3機台總工時：' + str(int(hr_list[4])))
        print('BF5機台總工時：' + str(int(hr_list[5])))

    if shop_code == 421:
        print('CHS機台總工時：' + str(int(hr_list[0])))
        print('CH0機台總工時：' + str(int(hr_list[1])))
        print('手動機台總工時：' + str(int(hr_list[2])))

    if shop_code == 422:
        print('DB機台總工時：' + str(int(hr_list[0])))
        print('EC機台總工時：' + str(int(hr_list[1])))

    if shop_code == 460:
        print('C1機台總工時：' + str(int(hr_list[0])))
        print('C3機台總工時：' + str(int(hr_list[1])))
        print('C6機台總工時：' + str(int(hr_list[2])))
        print('C7機台總工時：' + str(int(hr_list[3])))

    if shop_code == 461:
        print('BF0機台總工時：' + str(int(hr_list[0])))
        print('BF1機台總工時：' + str(int(hr_list[1])))
        print('BF4機台總工時：' + str(int(hr_list[2])))

    if shop_code == 401:
        print('TC 機台總工時：' + str(int(hr_list[0])))
        print('BA1機台總工時：' + str(int(hr_list[1])))

    if shop_code == 402:
        print('KVS70機台總工時：' + str(int(hr_list[0])))
        print('SM80 機台總工時：' + str(int(hr_list[1])))
        print('SM165機台總工時：' + str(int(hr_list[2])))

    if shop_code == 404:
        print('BTH60機台總工時：' + str(int(hr_list[0])))
        print('S80  機台總工時：' + str(int(hr_list[1])))
        print('PM160機台總工時：' + str(int(hr_list[2])))

    if shop_code == 405:
        print('C2機台總工時：' + str(int(hr_list[0])))
        print('A4機台總工時：' + str(int(hr_list[1])))
        print('A6機台總工時：' + str(int(hr_list[2])))
        print('A8機台總工時：' + str(int(hr_list[3])))
        print('CF機台總工時：' + str(int(hr_list[4])))

    if shop_code == 406:
        print('CB0機台總工時：' + str(int(hr_list[0])))
        print('CB1機台總工時：' + str(int(hr_list[1])))

    if shop_code == 410:
        print('D2機台總工時：' + str(int(hr_list[0])))
        print('D5機台總工時：' + str(int(hr_list[1])))
        print('D6機台總工時：' + str(int(hr_list[2])))

    if shop_code == 450:
        print('H9 機台總工時：' + str(int(hr_list[0])))
        print('H10機台總工時：' + str(int(hr_list[1])))
        print('H11機台總工時：' + str(int(hr_list[2])))
        print('H12機台總工時：' + str(int(hr_list[3])))

    if shop_code == 451:
        print('K4機台總工時：' + str(int(hr_list[0])))
        print('K5機台總工時：' + str(int(hr_list[1])))
        print('K8機台總工時：' + str(int(hr_list[2])))

    if shop_code == 430:
        print('E10機台總工時：' + str(int(hr_list[0])))
        print('E11機台總工時：' + str(int(hr_list[1])))

    if shop_code == 433:
        print('CYA機台總工時：' + str(int(hr_list[0])))
        print('CYB機台總工時：' + str(int(hr_list[1])))


if __name__ == '__main__':
    # Read the imported Excel file and convert it to DataFrame
    df_0, df_down, df_cp = read_file()
    # Process the file of machines' capacity
    df_cp.set_index(['SHOP_CODE', 'PROCESS_CODE', 'SHAPE_TYPE', 'GRADE_GROUP'], inplace=True)
    df_cp.sort_index(inplace=True)

    # Translate the column name of the order file from English to Chinese
    df_main_ch = to_translate_ch(df_0)
    # User input the shop code that needs to be balanced and difference No. as a list here
    shop_code_list = [[420, 2], [453, 2], [452, 2], [421, 2], [422, 2],
                      [460, 2], [461, 2], [401, 2], [402, 2], [404, 2],
                      [405, 2], [406, 2], [410, 2], [450, 2], [451, 2],
                      [430, 2], [433, 2]]
    # shop_code_list = [[420, 2], [410, 2], [433, 2], [451, 2], [453, 2], [461, 2]]
    for code in shop_code_list:
        print('平衡%s站機台...' % code[0])
        print("=====================================")
        # Get the data within a certain time
        df_1 = filter_time(df_main_ch, shop_code=code[0], N=code[1])
        # Assign available machines
        df_2 = assign_available_machine(df_1, df_cp, shop_code=code[0])
        # Calculate the workload
        df_3, time_list, downtimes_list = cal_working_hours(df_2, df_down, shop_code=code[0], down_period=7)
        # Add the data after being processed and balanced to the original order file
        df_4 = add_to_main(df_3, df_main_ch, shop_code=code[0], N=code[1])
        # Update the machine
        df_main_ch = df_4
        # Print the result
        print_result(time_list, df_down, code[0], downtimes_list)
        print("=====================================")
        print('平衡%s站機台...done\n' % code[0])
    # Translate the column name of the order file from Chinese to English
    df_final_eng = to_translate(df_main_ch)
    df_final_eng.to_excel("0803主檔_現況012(平衡後)_eng.xlsx", index=False)

    # End timing
    end = time.time()
    # Total time used in second
    print('耗時' + str(round(end) - round(start)) + 's')
