"""
====== BLOCK OF FUNCTIONS ======
"""
import xosrm

import os
from IPython.display import display
import pandas as pd


def df_first(data):
    df_first = pd.DataFrame(data['первая половина месяца']).T
    df_first.fillna(0, inplace=True)
    return df_first

def df_second(data):
    df_second = pd.DataFrame(data['вторая половина месяца']).T
    df_second.fillna(0, inplace=True)
    return df_second

def open_excel_file(file_path):
    '''Open Excel file and create DataFrame
    input: filepath
    output: DataFrame'''
    data = pd.read_excel(file_path)
    return data

def schedule(df):
    '''
    Crete df for each day for even and odd week
    even - четная неделя; 
    odd - нечентная неделя
    input: DataFrame (schedule)
    output: dict with coords
    '''
    # Create nested dicts
    monthly_coords = {}
    monthly_coords['первая половина месяца'] = {}
    monthly_coords['вторая половина месяца'] = {}
    
    monthly_coords['первая половина месяца']['четная нед.'] = {}
    monthly_coords['первая половина месяца']['не четная нед.'] = {}
    
    monthly_coords['вторая половина месяца']['четная нед.'] = {}
    monthly_coords['вторая половина месяца']['не четная нед.'] = {}
    
    
    '''First part of the month'''
    # Crete df for each day of even week and add to dict
    first_part = df[df['№п/п четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,4]))]
    even_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п четная нед.'])
    monthly_coords['первая половина месяца']['четная нед.']['1-ПН'] = even_mon.to_dict('records')
    even_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п четная нед.'])
    monthly_coords['первая половина месяца']['четная нед.']['2-ВТ'] = even_tue.to_dict('records')
    even_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п четная нед.'])
    monthly_coords['первая половина месяца']['четная нед.']['3-СР'] = even_wed.to_dict('records')
    even_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п четная нед.'])
    monthly_coords['первая половина месяца']['четная нед.']['4-ЧТ'] = even_thu.to_dict('records')
    even_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п четная нед.'])
    monthly_coords['первая половина месяца']['четная нед.']['5-ПТ'] = even_fri.to_dict('records')
    
    # Crete df for each day of odd week and add to dict
    first_part = df[df['№п/п не четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,4]))]
    odd_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['первая половина месяца']['не четная нед.']['1-ПН'] = odd_mon.to_dict('records')
    odd_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['первая половина месяца']['не четная нед.']['2-ВТ'] = odd_tue.to_dict('records')
    odd_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['первая половина месяца']['не четная нед.']['3-СР'] = odd_wed.to_dict('records')
    odd_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['первая половина месяца']['не четная нед.']['4-ЧТ'] = odd_thu.to_dict('records')
    odd_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['первая половина месяца']['не четная нед.']['5-ПТ'] = odd_fri.to_dict('records')
    
    '''Second part of the month'''
    # Crete df for each day of even week and add to dict
    first_part = df[df['№п/п четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,8]))]
    even_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п четная нед.'])
    monthly_coords['вторая половина месяца']['четная нед.']['1-ПН'] = even_mon.to_dict('records')
    even_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п четная нед.'])
    monthly_coords['вторая половина месяца']['четная нед.']['2-ВТ'] = even_tue.to_dict('records')
    even_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п четная нед.'])
    monthly_coords['вторая половина месяца']['четная нед.']['3-СР'] = even_wed.to_dict('records')
    even_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п четная нед.'])
    monthly_coords['вторая половина месяца']['четная нед.']['4-ЧТ'] = even_thu.to_dict('records')
    even_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п четная нед.'])
    monthly_coords['вторая половина месяца']['четная нед.']['5-ПТ'] = even_fri.to_dict('records')
    
    # Crete df for each day of odd week and add to dict
    first_part = df[df['№п/п не четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,8]))]
    odd_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['вторая половина месяца']['не четная нед.']['1-ПН'] = odd_mon.to_dict('records')
    odd_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['вторая половина месяца']['не четная нед.']['2-ВТ'] = odd_tue.to_dict('records')
    odd_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['вторая половина месяца']['не четная нед.']['3-СР'] = odd_wed.to_dict('records')
    odd_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['вторая половина месяца']['не четная нед.']['4-ЧТ'] = odd_thu.to_dict('records')
    odd_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['вторая половина месяца']['не четная нед.']['5-ПТ'] = odd_fri.to_dict('records')
    
    return monthly_coords



def calc_routes(full_dict):
    ''' Calculate SIMPLE_ROUTE for coords in a given order
    Calculate distances and durations for each day of the week
    Output: dict of geometries for each day
    '''
    result_dict = full_dict.copy()
    for key0, value0 in result_dict.items():
        print('================== {} =================='.format(key0))
        for key, value in value0.items():
            print('================== {} =================='.format(key))
            for key2, value2 in value.items():
                print(key2)
                if len(value2) > 1: # If list has more than 1 point
                    coords = []
                    for point in value2:
                        coords.append((point['Долгота'], point['Широта'])) 

                    source = coords[0]
                    dest = coords[0]
                    coords = coords[1:]
                    result = xosrm.simple_route(source, dest, coords,
                                        output='full', overview="full", geometry="geojson")
                    print(result['routes'][0]['distance']/1000, 'km')
                    print(result['routes'][0]['duration']/60.026, 'min')

                    result_dict[key0][key][key2].append(result['routes'][0]['geometry'])
                else:
                    pass
                
    return result_dict

def calc_dist(full_dict):
    ''' Calculate distances for each day of the week
    Output: dict of distances
    '''
    result_dict = full_dict.copy()
    # Create dict of distances
    dist_dict = {}
    for key0, value0 in result_dict.items():
        dist_dict[key0] = {}
        for key, value in value0.items():
            dist_dict[key0][key]= {}
            for key2, value2 in value.items():
                if len(value2) > 1: # If list has more than 1 point
                    coords = []
                    for point in value2:
                        coords.append([point['Долгота'], point['Широта']])

                    source = coords[0]
                    dest = coords[0]
                    coords = coords[1:]
                    result = xosrm.simple_route(source, dest, coords,
                                        output='full', overview="full", geometry="geojson")
                    
                    # Add distances to dict
                    distance = result['routes'][0]['distance']/1000
                    dist_dict[key0][key][key2] = distance
                else:
                    dist_dict[key0][key][key2] = 0

    return dist_dict
    
def calc_trips(full_dict):
    ''' Calculate TRIP_ROUTE
    Calculate distances and durations for each day of the week
    Output: dict of geometries for each day
    '''    
    # Dict of geometries of route by day
    result_dict = full_dict.copy()
    # Create dict of distances
    dist_dict = {}
    for key0, value0 in result_dict.items():
        dist_dict[key0] = {}
        for key, value in value0.items():
            dist_dict[key0][key]= {}
            for key2, value2 in value.items():
                if len(value2) > 1: # If list has more than 1 point
                    coords = []
                    for point in value2:
                        coords.append([point['Долгота'], point['Широта']])

                    result = xosrm.trip(coords, source='first', roundtrip=True,
                                        output='full', overview="full", geometry="geojson")
                    distance = result['trips'][0]['distance']/1000
                    dist_dict[key0][key][key2] = distance

                    # Get waypoinnt order and replace origin order in the dict
                    for index, point in enumerate(value2):
                        if key == 'четная нед.':
                            point.update({'№п/п четная нед.': result['waypoints'][index]['waypoint_index'] + 1})
                            point.update({'№п/п не четная нед.': None})
                        if key == 'не четная нед.':
                            point.update({'№п/п не четная нед.': result['waypoints'][index]['waypoint_index'] + 1})
                            point.update({'№п/п четная нед.': None})  


                else:
                    dist_dict[key0][key][key2] = 0
                    
    return dist_dict, result_dict

def write_to_excel(data, short_file_name):
    writer = pd.ExcelWriter('Расписания ТП расчетные.xlsx', engine = 'xlsxwriter')

    short_file_name = short_file_name.split(" ")[0]

    full_schedule = pd.DataFrame()

    for key0, value0 in data[1].items():
        for key, value in value0.items():
            for key2, value2 in value.items():
                df = pd.DataFrame.from_dict(value2, orient='columns')
                full_schedule = pd.concat([full_schedule, df], ignore_index=True)
    # Get uniq rows
    even_part = full_schedule[['Внешний ID ТТ', '№п/п четная нед.']].drop_duplicates(subset='Внешний ID ТТ', keep='first', inplace=False)
    # Get only not null rows for even week
    even_part = even_part[even_part['№п/п четная нед.'].notnull()]

    # Create table with uniq rows WITHOUT even week
    odd_part = full_schedule[['Внешний ID ТТ','Клиент', 'Адрес', 'Долгота', 'Широта', 'Интервал повторений', 'Дни недели', 
                            '№п/п не четная нед.']].drop_duplicates(subset=['Внешний ID ТТ', 'Дни недели'], keep='last', inplace=False)
    # Merge two tables on ID TT
    merge_table = pd.merge(odd_part, even_part, on='Внешний ID ТТ', how='left')

    merge_table.to_excel(writer, sheet_name = short_file_name + ' расписание', index=False)

    dist_table1 = pd.DataFrame.from_dict(data[0]['первая половина месяца'], orient='index')
    dist_table2 = pd.DataFrame.from_dict(data[0]['вторая половина месяца'], orient='index')
    dist_table1.to_excel(writer, sheet_name = short_file_name + " пробег", startrow=2, startcol=1) 
    dist_table2.to_excel(writer, sheet_name = short_file_name + " пробег", startrow=2, startcol=dist_table1.shape[1]+2, index=False) 

    writer.save()
    writer.close()



def write_to_excel2(writer, data, file_name, index_num):

    short_file_name = file_name.split(" ")[0]

    full_schedule = pd.DataFrame()

    for key0, value0 in data[1].items():
        for key, value in value0.items():
            for key2, value2 in value.items():
                df = pd.DataFrame.from_dict(value2, orient='columns')
                full_schedule = pd.concat([full_schedule, df], ignore_index=True)
    # Get uniq rows
    even_part = full_schedule[['Внешний ID ТТ', '№п/п четная нед.']].drop_duplicates(subset='Внешний ID ТТ', keep='first', inplace=False)
    # Get only not null rows for even week
    even_part = even_part[even_part['№п/п четная нед.'].notnull()]

    # Create table with uniq rows WITHOUT even week
    odd_part = full_schedule[['Внешний ID ТТ','Клиент', 'Адрес', 'Долгота', 'Широта', 'Интервал повторений', 'Дни недели', 
                            '№п/п не четная нед.']].drop_duplicates(subset=['Внешний ID ТТ', 'Дни недели'], keep='last', inplace=False)
    # Merge two tables on ID TT
    merge_table = pd.merge(odd_part, even_part, on='Внешний ID ТТ', how='left')

    merge_table.to_excel(writer, sheet_name = short_file_name + ' расписание', index=False)

    
    dist_table1 = pd.DataFrame.from_dict(data[0]['первая половина месяца'], orient='index')
    dist_table2 = pd.DataFrame.from_dict(data[0]['вторая половина месяца'], orient='index')
    dist_table1.iloc[[1]].to_excel(writer, sheet_name = 'Пробег', startrow=3 + index_num, startcol=1, index=False, header=False)  # first part add week
    dist_table1.iloc[[0]].to_excel(writer, sheet_name = 'Пробег', startrow=3 + index_num, startcol=6, index=False, header=False)  # first part even week
    dist_table2.iloc[[1]].to_excel(writer, sheet_name = 'Пробег', startrow=3 + index_num, startcol=11, index=False, header=False)  
    dist_table2.iloc[[0]].to_excel(writer, sheet_name = 'Пробег', startrow=3 + index_num, startcol=16, index=False, header=False) 
    

    worksheet = writer.sheets['Пробег']
    worksheet.write(3 + index_num, 0, file_name)



def calc_folder(path):
    '''Calculate trip for all files in folder'''
    files = []
    # r=root, d=directories, f = files
    for r, _, f in os.walk(path):
        for file in f:
            if '.xls' in file:
                files.append(os.path.join(r, file))

    for f in files:
        print(f.split('\\')[-2], ' - ',f.split('\\')[-1].split('.')[0])
        schedule_data = open_excel_file(f)
        full_dict = schedule(schedule_data)
        data = calc_trips(full_dict)
        print('Первая половина месяца:')
        display(df_first(data))
        print('Вторая половина месяца:')
        display(df_second(data))



def calc_trips_for_folder(path):
    '''Calculate trips order for all files in folder'''
    files = []
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
        for file in f:
            if '.xls' in file:
                files.append(os.path.join(r, file))
    # Open excel file
    writer = pd.ExcelWriter('Расписания ТП расчетные.xlsx', engine = 'xlsxwriter')

    for f in files:
        short_file_name = f.split('\\')[-1].split('.')[0]
        
        schedule_data = open_excel_file(f)
        full_dict = schedule_odd_even(schedule_data)
        # data consists of dict of distances and dict of salepoints order
        data = calc_trips_order(full_dict)

        # Create empty dataFrame
        full_schedule = pd.DataFrame()
        # dicts transformation to write in excel
        for key0, value0 in data[1].items():
            for key2, value2 in value0.items():
                df = pd.DataFrame.from_dict(value2, orient='columns')
                full_schedule = pd.concat([full_schedule, df], ignore_index=True)
        # Get uniq rows
        even_part = full_schedule[['Внешний ID ТТ', '№п/п четная нед.']].drop_duplicates(subset='Внешний ID ТТ', keep='first', inplace=False)
        # Get only not null rows for even week
        even_part = even_part[even_part['№п/п четная нед.'].notnull()]

        # Create table with uniq rows WITHOUT even week
        odd_part = full_schedule[['Внешний ID ТТ','Клиент', 'Адрес', 'Долгота', 'Широта', 'Интервал повторений', 'Дни недели', 
                                '№п/п не четная нед.']].drop_duplicates(subset=['Внешний ID ТТ', 'Дни недели'], keep='last', inplace=False)
        # Merge two tables on ID TT
        merge_table = pd.merge(odd_part, even_part, on='Внешний ID ТТ', how='left')
        # Add sheet to file
        merge_table.to_excel(writer, sheet_name = short_file_name)

        dist_table = pd.DataFrame.from_dict(data[0])
        dist_table.to_excel(writer, sheet_name = short_file_name + " пробег") 


        print("Расписание {} готово".format(short_file_name))
        
    
    writer.save()
    writer.close()
    print('Файл создан')