# import core
import xosrm

import os
import pandas as pd
from IPython.display import display

def open_excel_file(file_path):
    '''Open Excel file and create DataFrame
    input: filepath
    output: DataFrame'''
    data = pd.read_excel(file_path)
    return data

def schedule_odd_even(df):
    '''Create dict with odd and even parts of the week. Devided by days'''
    week = {}
    week['четная нед.'] = {}
    week['не четная нед.'] = {}
    
    even = df[df['№п/п четная нед.'].notnull()].sort_values(by=['№п/п четная нед.'])
    week['четная нед.']['1-ПН'] = even[even['Дни недели']==1].to_dict('records')
    week['четная нед.']['2-ВТ'] = even[even['Дни недели']==2].to_dict('records')
    week['четная нед.']['3-СР'] = even[even['Дни недели']==3].to_dict('records')
    week['четная нед.']['4-ЧТ'] = even[even['Дни недели']==4].to_dict('records')
    week['четная нед.']['5-ПТ'] = even[even['Дни недели']==5].to_dict('records')
    
    odd = df[df['№п/п не четная нед.'].notnull()].sort_values(by=['№п/п не четная нед.'])        
    week['не четная нед.']['1-ПН'] = odd[odd['Дни недели']==1].to_dict('records')
    week['не четная нед.']['2-ВТ'] = odd[odd['Дни недели']==2].to_dict('records')
    week['не четная нед.']['3-СР'] = odd[odd['Дни недели']==3].to_dict('records')
    week['не четная нед.']['4-ЧТ'] = odd[odd['Дни недели']==4].to_dict('records')
    week['не четная нед.']['5-ПТ'] = odd[odd['Дни недели']==5].to_dict('records')
    return week

def df2table(data):
    '''Convert dict of distances to table'''
    table = pd.DataFrame(data).T
    table.fillna(0, inplace=True)
    return table

def calc_trips_order(full_dict):
    ''' Calculate TRIP_ROUTE
    Calculate distances for each day of the week
    Output: dict of dictances and dict of waypoints order
    '''    
    # Dict of geometries of route by day
    result_dict = full_dict.copy()
    # Create dict of distances
    dist_dict = {}
    for key0, value0 in result_dict.items():
        dist_dict[key0] = {}
        for key2, value2 in value0.items():
            if value2: # If list is empty (empty day)
                coords = []
                for point in value2:
                    coords.append([point['Долгота'], point['Широта']])
                result = xosrm.trip(coords, source='first', roundtrip=True,
                                    output='full', overview="full", geometry="geojson")
                distance = result['trips'][0]['distance']/1000
                dist_dict[key0][key2] = distance
                
                # Get waypoinnt order and replace origin order in the dict
                for index, point in enumerate(value2):
                    if key0 == 'четная нед.':
                        point.update({'№п/п четная нед.': result['waypoints'][index]['waypoint_index'] + 1})
                        point.update({'№п/п не четная нед.': None})
                    if key0 == 'не четная нед.':
                        point.update({'№п/п не четная нед.': result['waypoints'][index]['waypoint_index'] + 1})
                        point.update({'№п/п четная нед.': None})    
            else:
                dist_dict[key0][key2] = 0
            
    return dist_dict, result_dict


def write_to_excel(data):
    writer = pd.ExcelWriter('Расписания ТП расчетные.xlsx', engine = 'xlsxwriter')

    full_schedule = pd.DataFrame()

    for key0, value0 in data.items():
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

    merge_table.to_excel(writer, sheet_name = 'Расписание ТП')
    #full_schedule.to_excel(writer, sheet_name = 'full')

    writer.save()
    writer.close()



def calc_order_for_folder(path):
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

        dist_table = pd.DataFrame.from_dict(data[0], orient='columns')
        dist_table.to_excel(writer, sheet_name = short_file_name + " пробег") 


        print("Расписание {} готово".format(short_file_name))
        
    
    writer.save()
    writer.close()
    print('Файл создан')
