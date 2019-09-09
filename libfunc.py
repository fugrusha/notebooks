


"""
====== BLOCK OF FUNCTIONS ======
"""
import core
import xosrm

import os
from IPython.display import display
import pandas as pd


def df_first(data):
    df_first = pd.DataFrame(data['first_month_part']).T
    df_first.fillna(0, inplace=True)
    return df_first

def df_second(data):
    df_second = pd.DataFrame(data['second_month_part']).T
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
    monthly_coords['first_month_part'] = {}
    monthly_coords['second_month_part'] = {}
    
    monthly_coords['first_month_part']['even_part_of_coords'] = {}
    monthly_coords['first_month_part']['odd_part_of_coords'] = {}
    
    monthly_coords['second_month_part']['even_part_of_coords'] = {}
    monthly_coords['second_month_part']['odd_part_of_coords'] = {}
    
    
    '''First part of the month'''
    # Crete df for each day of even week and add to dict
    first_part = df[df['№п/п четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,4]))]
    even_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п четная нед.'])
    monthly_coords['first_month_part']['even_part_of_coords']['1-monday'] = even_mon.to_dict('records')
    even_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п четная нед.'])
    monthly_coords['first_month_part']['even_part_of_coords']['2-tuesday'] = even_tue.to_dict('records')
    even_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п четная нед.'])
    monthly_coords['first_month_part']['even_part_of_coords']['3-wednesday'] = even_wed.to_dict('records')
    even_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п четная нед.'])
    monthly_coords['first_month_part']['even_part_of_coords']['4-thursday'] = even_thu.to_dict('records')
    even_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п четная нед.'])
    monthly_coords['first_month_part']['even_part_of_coords']['5-friday'] = even_fri.to_dict('records')
    
    # Crete df for each day of odd week and add to dict
    first_part = df[df['№п/п не четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,4]))]
    odd_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['first_month_part']['odd_part_of_coords']['1-monday'] = odd_mon.to_dict('records')
    odd_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['first_month_part']['odd_part_of_coords']['2-tuesday'] = odd_tue.to_dict('records')
    odd_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['first_month_part']['odd_part_of_coords']['3-wednesday'] = odd_wed.to_dict('records')
    odd_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['first_month_part']['odd_part_of_coords']['4-thursday'] = odd_thu.to_dict('records')
    odd_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['first_month_part']['odd_part_of_coords']['5-friday'] = odd_fri.to_dict('records')
    
    '''Second part of the month'''
    # Crete df for each day of even week and add to dict
    first_part = df[df['№п/п четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,8]))]
    even_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п четная нед.'])
    monthly_coords['second_month_part']['even_part_of_coords']['1-monday'] = even_mon.to_dict('records')
    even_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п четная нед.'])
    monthly_coords['second_month_part']['even_part_of_coords']['2-tuesday'] = even_tue.to_dict('records')
    even_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п четная нед.'])
    monthly_coords['second_month_part']['even_part_of_coords']['3-wednesday'] = even_wed.to_dict('records')
    even_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п четная нед.'])
    monthly_coords['second_month_part']['even_part_of_coords']['4-thursday'] = even_thu.to_dict('records')
    even_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п четная нед.'])
    monthly_coords['second_month_part']['even_part_of_coords']['5-friday'] = even_fri.to_dict('records')
    
    # Crete df for each day of odd week and add to dict
    first_part = df[df['№п/п не четная нед.'].notnull() &
                                          (df['Интервал повторений'].isin([1,2,8]))]
    odd_mon = first_part[first_part['Дни недели']==1].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['second_month_part']['odd_part_of_coords']['1-monday'] = odd_mon.to_dict('records')
    odd_tue = first_part[first_part['Дни недели']==2].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['second_month_part']['odd_part_of_coords']['2-tuesday'] = odd_tue.to_dict('records')
    odd_wed = first_part[first_part['Дни недели']==3].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['second_month_part']['odd_part_of_coords']['3-wednesday'] = odd_wed.to_dict('records')
    odd_thu = first_part[first_part['Дни недели']==4].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['second_month_part']['odd_part_of_coords']['4-thursday'] = odd_thu.to_dict('records')
    odd_fri = first_part[first_part['Дни недели']==5].sort_values(by=['№п/п не четная нед.'])
    monthly_coords['second_month_part']['odd_part_of_coords']['5-friday'] = odd_fri.to_dict('records')
    
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
                coords = []
                for point in value2:
                    coords.append((point['Долгота'], point['Широта'])) 
                result = xosrm.simple_route([30.6508, 46.4289], [30.6508, 46.4289], coords,
                                    output='full', overview="full", geometry="geojson")
                print(result['routes'][0]['distance']/1000, 'km')
                print(result['routes'][0]['duration']/60.026, 'min')

                result_dict[key0][key][key2].append(result['routes'][0]['geometry'])
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
                coords = []
                for point in value2:
                    coords.append([point['Долгота'], point['Широта']])
                # Office Vinnitsa
                result = xosrm.simple_route([28.5673, 49.2226], [28.5673, 49.2226], coords,
                                    output='full', overview="full", geometry="geojson")
                #print(result['routes'][0]['distance']/1000, 'km')
                #print(result['routes'][0]['duration']/60.026, 'min')
                # Add distances to dict
                distance = result['routes'][0]['distance']/1000
                dist_dict[key0][key][key2] = distance
                
    return dist_dict
    
def calc_trips(dict_of_coords_by_day):
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
                coords = []
                for point in value2:
                    coords.append([point['Долгота'], point['Широта']])
                result = xosrm.trip(coords, source='first', roundtrip=True,
                                    output='full', overview="full", geometry="geojson")
                distance = result['trips'][0]['distance']/1000
                dist_dict[key0][key][key2] = distance
                    
    return dist_dict
    
def calc_folder(path):
    '''Calculate trip for all files in folder'''
    files = []
    # r=root, d=directories, f = files
    for r, d, f in os.walk(path):
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