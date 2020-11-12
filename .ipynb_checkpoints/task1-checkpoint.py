# -*- coding: utf-8 -*-
import pandas as pd
import xlrd
import numpy as np
from collections import Counter
import os
import xlsxwriter
import time
import re


pd.options.mode.chained_assignment = None

def dropTopLeftRight(spec):

    # finding elements from varsNamesSet
    start_norm = time.time()
    spec = spec.replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    spec = spec.replace('nan', np.nan)
    end_norm = time.time()
#     print('Таблица нормализована за: {}'.format(end_norm-start_norm))
    normTime = (end_norm-start_norm)

    start_norm = time.time()
    specTrueFalseMap = spec.isin(varsNamesSet).replace(False, np.nan)
    if specTrueFalseMap.count(axis='columns').sum() == 0:
#         print('Не удалось обработать файл')
        return  pd.DataFrame()


    # находим строку
    heading_row = specTrueFalseMap.count(axis='columns').idxmax()+2
    # находим левую границу
    a = spec.iloc[specTrueFalseMap.count(axis='columns').idxmax()].isnull().replace(True,1)
    if 'Unnamed: ' in a.idxmin():
        heading_left_end = int(a.idxmin().split('Unnamed: ')[1])+1
    else:
        heading_left_end = a.index.get_loc(a.idxmin())+1
    # находим правую границу
    min_a = min(list(a))
    len_a = len(list(a))
    if [i for i, j in enumerate(list(a)) if j == min_a][-1] < len_a-1:
        heading_right_end = [i for i, j in enumerate(list(a)) if j == min_a][-1]+1
    else:
        heading_right_end = len_a

    RC_title = ((heading_row,heading_left_end),(heading_row,heading_right_end))

    # находим строку с номерами столбцов
    # в doesContainNumbersRow тупл с bool и координатами(тоже тупл)
    doesContainNumbersRow = (False,())

    try:
        if int(spec.iloc[heading_row-1][1])+1 == int(spec.iloc[heading_row-1][2]):
            doesContainNumbersRow = (True,((heading_row+1,heading_left_end),(heading_row+1,heading_right_end)))

    except:
        pass

    if not doesContainNumbersRow[0]:
        try:
            if int(spec.iloc[heading_row][1])+1 == int(spec.iloc[heading_row][2]):
                doesContainNumbersRow = (True,((heading_row+2,heading_left_end),(heading_row+2,heading_right_end)))
        except:
            pass

    # находим количество заголовков
    numberOfTitleNames = (~spec.iloc[heading_row-2].isnull()).replace(True,1).sum()

    # находим повторяющиеся заголовки и их количество
    titelNames = Counter(spec.iloc[heading_row-2])
    repeatedTitleNames = Counter(el for el in titelNames.elements() if titelNames[el] >= 2)
#     print(repeatedTitleNames)
    amountOFRepeatedTitleNames = sum(repeatedTitleNames.values())


    # выясняем двойной ли заголовок
    titleIsDoubled = False

    if doesContainNumbersRow[0]:
        if heading_row == doesContainNumbersRow[1][0][0]-2:
            titleIsDoubled = True
    else:
        smallHead = list(spec.iloc[heading_row-1:].head(3).count(axis=1))
        if smallHead[0] < sum(smallHead)/len(smallHead):
            titleIsDoubled = True

    # находим вторую (английскую) строку заголовков
    secondTitle = (False,())

    if titleIsDoubled:
        try:
            if (~spec.iloc[heading_row-4].isnull()).replace(True,1).sum() > 3:
                secondTitle = (True, ((heading_row-2,heading_left_end),(heading_row-2,heading_right_end)))
        except:
            pass
    else:
        try:
            if (~spec.iloc[heading_row-3].isnull()).replace(True,1).sum() > 3:
                secondTitle = (True, ((heading_row-1,heading_left_end),(heading_row-1,heading_right_end)))
        except:
            pass


    # # дропаем строку с номерами
    # if doesContainNumbersRow[0]:
    #     if doesContainNumbersRow[1][0][0] - heading_row == 1:
    #         spec.drop(heading_row+1, inplace = True)
    #     elif doesContainNumbersRow[1][0][0] - heading_row == 1:
    #         spec.drop(heading_row, inplace = True)

    spec = spec.iloc[specTrueFalseMap.count(axis='columns').idxmax():].dropna(how='all')
    # print(spec)
    listOfEmptyTitleFields = [i for i, j in spec.iloc[0].isnull().replace(False,np.nan).to_dict().items() if j == 1]
    spec = spec.drop(columns=listOfEmptyTitleFields)
    spec.columns = spec.iloc[0]
    spec = spec.iloc[1:]
    spec.reset_index(inplace = True,drop=True)
    spec.columns.name = ''

    end_norm = time.time()
#     print('Чтение диапазона заголовков выполнено за: {}'.format(end_norm-start_norm))
    titleReadingTime = end_norm-start_norm
    return spec, RC_title, numberOfTitleNames, repeatedTitleNames, amountOFRepeatedTitleNames, titleIsDoubled, secondTitle,doesContainNumbersRow, titleReadingTime, normTime

def renameColumnsNames(spec):
    start_norm = time.time()
    changedColumnsNames = []
    changedColumnsNamesDict = {}

    for i,col in enumerate(spec):
        if col in varsNamesSet:
            for j in varsDF:
                if col in list(varsDF[j]):
                    spec.rename(columns={col: j}, inplace=True)
                    changedColumnsNames.append(j)
                    try:
                        if col in list(copy_columns[j]):
                            changedColumnsNamesDict[col] = ((j,'copy_columns'))
                    except:
                        pass

                    try:
                        if col in list(text_columns[j]):
                            changedColumnsNamesDict[col] = ((j,'text_columns'))
                    except:
                        pass

                    try:
                        if col in list(eng_vars_columns[j]):
                            changedColumnsNamesDict[col] = ((j,'eng_vars_columns'))
                    except:
                        pass

                    break
        else:
            changedColumnsNamesDict[col] = (('не обработан',''))



    # numberOfColumnsAppearences = Counter(changedColumnsNames)
    changedColumnsNames = list(set(changedColumnsNames))
    end_norm = time.time()
#     print('Обработка диапазона заголовков рабочей таблицы выполнена за: {}'.format(end_norm-start_norm))
    return spec, changedColumnsNames, changedColumnsNamesDict

def dropSpecsBottom(spec,titleIsDoubled,doesContainNumbersRow):
    apearencesOfNanInRows = spec.isnull().sum(axis=1)

    # h = Counter(apearencesOfNanInRows.head(int(len(apearencesOfNanInRows)/2))).most_common()

    # listOfRowIndicesToDelete = list(apearencesOfNanInRows.loc[apearencesOfNanInRows > 1 + h[0][0]].index)


    if doesContainNumbersRow[0] &  titleIsDoubled:
        apearenceOfNanInRows = spec.iloc[2].isnull().sum()
        # spec.drop([0,1], inplace=True)
    elif doesContainNumbersRow[0] | titleIsDoubled:
        apearenceOfNanInRows = spec.iloc[1].isnull().sum()
        # spec.drop(0, inplace=True)
    else:
        apearenceOfNanInRows = spec.iloc[0].isnull().sum()

    # apearenceOfNanInRows = spec.iloc[0].isnull().sum()
    listOfRowIndicesToDelete = list(apearencesOfNanInRows.loc[apearencesOfNanInRows > 1 + apearenceOfNanInRows].index)

    spec.drop(listOfRowIndicesToDelete,inplace=True)

    if doesContainNumbersRow[0] &  titleIsDoubled:
        # apearenceOfNanInRows = spec.iloc[2].isnull().sum()
        try:
            spec.drop([0,1], inplace=True)
        except:
            pass
    elif doesContainNumbersRow[0] | titleIsDoubled:
        # apearenceOfNanInRows = spec.iloc[1].isnull().sum()
        try:
            spec.drop(0, inplace=True)
        except:
            pass

    return spec

def findSameColumnNames(spec):
    # если есть одноименные столбцы мы их мержим. в astype можно указать другой тип данных
    # к сожалению этот код меняет порядок в столбцов
    def sjoin(x): return ';'.join(x[x.notnull()].astype(str))
    spec = spec.groupby(level=0, axis=1).apply(lambda x: x.apply(sjoin, axis=1))
    return spec

def changingColumnsValues(spec):
    columnsData = {}
    start_changing_columns_values = time.time()
    for (columnName, columnData) in spec.iteritems():
        columnsData[columnName] = {}
        columnsData[columnName]['amountOfEmptyFields'] = spec[columnName].isnull().sum()
        columnsData[columnName]['amountOfChangedFields'] = 0
        columnsData[columnName]['amountOfUnchangedFields'] = len(columnData)
        columnsData[columnName]['unchangedFields'] = set()

        if columnName in list(copy_columns.columns):
            columnsData[columnName]['amountOfChangedFields'] = columnData.astype(str).str.contains(',', regex=False).replace(np.nan, False).sum()
            columnsData[columnName]['amountOfUnchangedFields'] = len(columnData) - columnsData[columnName]['amountOfChangedFields']
            # print(spec[columnName].loc[~(spec[columnName].astype(str).str.contains(',', regex=False).replace(np.nan, True))])
            # print(set(spec[columnName].loc[~(spec[columnName].astype(str).str.contains(',', regex=False).replace(np.nan, True))]))
            columnsData[columnName]['unchangedFields'] = set(spec[columnName].loc[~(spec[columnName].astype(str).str.contains(',', regex=False).replace(np.nan, True))])
            spec[columnName] = spec[columnName].astype(str).str.replace(',','.')

        elif columnName in list(eng_vars_columns.columns):
    #         for val in columnData:
            for k in eng_vars_Connection_columns.keys():
                columnsData[columnName]['unchangedFields'] = columnsData[columnName]['unchangedFields'].union(set(spec[columnName].loc[~spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)]))
                columnsData[columnName]['amountOfChangedFields'] += len(spec[columnName].loc[spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)])
                columnsData[columnName]['amountOfUnchangedFields'] -= len(spec[columnName].loc[spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)])
                spec[columnName] = spec[columnName].astype(str).str.replace(k,str(eng_vars_Connection_columns[k][0]),regex=False)
            for k in eng_vars_Material_columns.keys():
                columnsData[columnName]['unchangedFields'] = columnsData[columnName]['unchangedFields'].union(set(spec[columnName].loc[~spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)]))
                columnsData[columnName]['amountOfChangedFields'] += len(spec[columnName].loc[spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)])
                columnsData[columnName]['amountOfUnchangedFields'] -= len(spec[columnName].loc[spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)])
                spec[columnName] = spec[columnName].astype(str).str.replace(k,str(eng_vars_Material_columns[k][0]),regex=False)
            for k in eng_vars_SeismoCat_columns.keys():
                columnsData[columnName]['unchangedFields'] = columnsData[columnName]['unchangedFields'].union(set(spec[columnName].loc[~spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)]))
                columnsData[columnName]['amountOfChangedFields'] += len(spec[columnName].loc[spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)])
                columnsData[columnName]['amountOfUnchangedFields'] -= len(spec[columnName].loc[spec[columnName].astype(str).str.contains(k, regex=False).replace(np.nan, True)])
                spec[columnName] = spec[columnName].astype(str).str.replace(k,str(eng_vars_SeismoCat_columns[k][0]),regex=False)

            columnsData[columnName]['unchangedFields'] = list(set(columnsData[columnName]['unchangedFields']))
    # text_Type
    if 'Type' in spec.columns:

        if not 'Bellow' in spec.columns:
            spec['Bellow'] = ''
        if not 'Actuator type' in spec.columns:
            spec['Actuator type'] = ''
        if not 'RPI' in spec.columns:
            spec['RPI'] = ''
        if not 'DC' in spec.columns:
            spec['DC'] = ''
        if not 'Fluid' in spec.columns:
            spec['Fluid'] = ''
        if not 'Connection' in spec.columns:
            spec['Connection'] = ''
        if not 'Material' in spec.columns:
            spec['Material'] = ''

        columnsData['Type'] = {}
        columnsData['Type']['amountOfEmptyFields'] = spec['Type'].isnull().sum()
        columnsData['Type']['amountOfChangedFields'] = 0
        columnsData['Type']['amountOfUnchangedFields'] = len(spec['Type'])
        columnsData['Type']['unchangedFields'] = set()

        for index, item in spec['Type'].items():


            triggersInItem = text_Type.loc[text_Type['Тригер'].apply(lambda x: x in item)]
            triggersInItem = triggersInItem.replace('nan','')
            triggersInItem = triggersInItem.loc[~(triggersInItem['Weight']=='')]
            triggersInItem['Weight'] = triggersInItem['Weight'].apply(lambda x: float(x))
            valuesfForTypeField = ''
            valuesForActuatorType = ''
            valuesForBellow = ''
            valuesForConnection = ''
            valuesForMaterial = ''
            valuesForRPI = ''
            valuesForFluid = ''
            valuesForDC = ''
            weight = -1

            for i, val in triggersInItem.iterrows():

                if (valuesForBellow == '') & (val['Bellow'] != ''):
                    valuesForBellow = val['Bellow']

                if (valuesForRPI == '') & (val['RPI'] != ''):
                    valuesForRPI = val['RPI']

                if (valuesForDC == '') & (val['DC'] != ''):
                    valuesForDC = val['DC']

                if not val['Actuator type'] in valuesForActuatorType:
                    if weight == -1:
                        valuesForActuatorType += val['Actuator type']
                    else:
                        valuesForActuatorType += ', ' + val['Actuator type']


                if not val['Connection'] in valuesForConnection:
                    if weight == -1 :
                        valuesForConnection += val['Connection']
                    else:
                        valuesForConnection += ', ' + val['Connection']

                if not val['Material'] in valuesForMaterial:
                    if weight == -1 :
                        valuesForMaterial += val['Material']
                    else:
                        valuesForMaterial += ', ' + val['Material']

                if not val['Fluid'] in valuesForFluid:
                    if weight == -1 :
                        valuesForFluid += val['Fluid']
                    else:
                        valuesForFluid += ', ' + val['Fluid']

                if val['Weight'] >= weight:
                    if weight == -1:
                        valuesfForTypeField += val['Type']
                    else:
                        valuesfForTypeField += ', ' + val['Type']
                    weight = val['Weight']

            if valuesfForTypeField != '':
                spec['Type'].loc[index] = valuesfForTypeField
                columnsData['Type']['amountOfChangedFields'] += 1
                columnsData['Type']['amountOfUnchangedFields'] -= 1
            else:
                columnsData['Type']['unchangedFields'].add(spec['Type'].loc[index])

            spec['Bellow'].loc[index] = valuesForBellow
            spec['RPI'].loc[index] = valuesForRPI

            if spec['DC'].loc[index] == '':
                spec['DC'].loc[index] = valuesForDC
            else:
                if valuesForDC != '':
                    spec['DC'].loc[index] += ', ' + valuesForDC


            if spec['Connection'].loc[index] == '':
                spec['Connection'].loc[index] = valuesForConnection
            else:
                if valuesForConnection != '':
                    spec['Connection'].loc[index] += ', ' + valuesForConnection


            if spec['Material'].loc[index] == '':
                spec['Material'].loc[index] = valuesForMaterial
            else:
                if valuesForMaterial != '':
                    spec['Material'].loc[index] += ', ' + valuesForMaterial


            if spec['Fluid'].loc[index] == '':
                spec['Fluid'].loc[index] = valuesForFluid
            else:
                if valuesForFluid != '':
                    spec['Fluid'].loc[index] += ', ' + valuesForFluid


            if spec['Actuator type'].loc[index] == '':
                spec['Actuator type'].loc[index] = valuesForActuatorType
            else:
                if valuesForActuatorType != '':
                    spec['Actuator type'].loc[index] += ', ' + valuesForActuatorType

    # text_Time
    if 'Time' in spec.columns:
        columnsData['Time'] = {}
        columnsData['Time']['amountOfEmptyFields'] = spec['Time'].isnull().sum()
        columnsData['Time']['amountOfChangedFields'] = 0
        columnsData['Time']['amountOfUnchangedFields'] = len(spec['Time'])
        columnsData['Time']['unchangedFields'] = set()
        for index, item in spec['Time'].items():
            if spec['Time'].loc[index] == re.sub("[^0-9]", "", str(spec['Time'].loc[index])):
                columnsData['Time']['unchangedFields'].add(spec['Time'].loc[index])
            else:
                columnsData['Time']['amountOfChangedFields'] += 1
                columnsData['Time']['amountOfUnchangedFields'] -= 1
            spec['Time'].loc[index] = re.sub("[^0-9]", "", str(spec['Time'].loc[index]))


    # Text_NC
    if 'NC' in spec.columns:
        if not 'GroupNC' in spec.columns:
            spec['GroupNC'] = ''
        if not 'Dostup' in spec.columns:
            spec['Dostup'] = ''
        if not 'Pletter' in spec.columns:
            spec['Pletter'] = ''
        if not 'SeismoCat' in spec.columns:
            spec['SeismoCat'] = ''

        columnsData['NC'] = {}
        columnsData['NC']['amountOfEmptyFields'] = spec['NC'].isnull().sum()
        columnsData['NC']['amountOfChangedFields'] = 0
        columnsData['NC']['amountOfUnchangedFields'] = len(spec['NC'])
        columnsData['NC']['unchangedFields'] = set()

        for index, item in spec['NC'].items():
            triggersInItem = text_NC.loc[text_NC['Тригер'].str.contains(item)]
            triggersInItem = triggersInItem.replace('nan','')
            triggersInItem['Weight'] = triggersInItem['Weight'].apply(lambda x: float(x))
            NC_values = ''
            GroupNC_values = ''
            Dostup_values = ''
            Pletter_values = ''
            SeismoCat_values = ''

            # на случай предпоследнего варианта с весом 4000
            if len(re.sub("[^0-9]", "", item)) == len(item):
                NC_values = item
                continue

            # для основной массы
            for i, val in triggersInItem.iterrows():
                if NC_values == '':
                    NC_values = val['NC']
                else:
                    NC_values += ', ' + val['NC']

                if GroupNC_values == '':
                    GroupNC_values = val['GroupNC']
                else:
                    GroupNC_values += ', ' + val['GroupNC']

                if Dostup_values == '':
                    Dostup_values = val['Dostup']
                else:
                    Dostup_values += ', ' + val['Dostup']

                if Pletter_values == '':
                    Pletter_values = val['Pletter']
                else:
                    Pletter_values += ', ' + val['Pletter']

                if SeismoCat_values == '':
                    SeismoCat_values = val['SeismoCat']
                else:
                    SeismoCat_values += ', ' + val['SeismoCat']


            # для последнего варианта с весом 4000
            for el in re.findall(r'[0-9]+/[^0-9]+', item):
                NC_values += ', ' + el

            # для вариантов с весом 3000
            for el in re.findall(r'[^0-9], i', item):
                if SeismoCat_values == '':
                    SeismoCat_values = '1'
                else:
                    SeismoCat_values += ', 1'

            for el in re.findall(r'[^0-9], ii', item):
                if SeismoCat_values == '':
                    SeismoCat_values = '2'
                else:
                    SeismoCat_values += ', 2'

            for el in re.findall(r'[^0-9], iii', item):
                if SeismoCat_values == '':
                    SeismoCat_values = '3'
                else:
                    SeismoCat_values += ', 3'

            for el in re.findall(r'[^0-9]/i', item):
                if SeismoCat_values == '':
                    SeismoCat_values = '1'
                else:
                    SeismoCat_values += ', 1'

            for el in re.findall(r'[^0-9]/ii', item):
                if SeismoCat_values == '':
                    SeismoCat_values = '2'
                else:
                    SeismoCat_values += ', 2'

            for el in re.findall(r'[^0-9]/iii', item):
                if SeismoCat_values == '':
                    SeismoCat_values = '3'
                else:
                    SeismoCat_values += ', 3'

            if NC_values != '':
                spec['NC'].loc[index] = NC_values
                columnsData['NC']['amountOfChangedFields'] += 1
                columnsData['NC']['amountOfUnchangedFields'] -= 1
            else:
                columnsData['NC']['unchangedFields'].add(spec['NC'].loc[index])

            if spec['GroupNC'].loc[index] == '':
                spec['GroupNC'].loc[index] = GroupNC_values
            else:
                if GroupNC_values != '':
                    spec['GroupNC'].loc[index] += ', ' + GroupNC_values

            if spec['Dostup'].loc[index] == '':
                spec['Dostup'].loc[index] = Dostup_values
            else:
                if Dostup_values != '':
                    spec['Dostup'].loc[index] += ', ' + Dostup_values

            if spec['Pletter'].loc[index] == '':
                spec['Pletter'].loc[index] = Pletter_values
            else:
                if Pletter_values != '':
                    spec['Pletter'].loc[index] += ', ' + Pletter_values

            if spec['SeismoCat'].loc[index] == '':
                spec['SeismoCat'].loc[index] = SeismoCat_values
            else:
                if SeismoCat_values != '':
                    spec['SeismoCat'].loc[index] += ', ' + SeismoCat_values

    # text_Kv
    if 'Kv' in spec.columns:
        if not 'F' in spec.columns:
            spec['F'] = ''
        columnsData['Kv'] = {}
        columnsData['Kv']['amountOfEmptyFields'] = spec['Kv'].isnull().sum()
        columnsData['Kv']['amountOfChangedFields'] = 0
        columnsData['Kv']['amountOfUnchangedFields'] = len(spec['Kv'])
        columnsData['Kv']['unchangedFields'] = set()
        for index, item in spec['Kv'].items():
            Kv_values = ''
            F_values = ''
            el = ''

            try:
                el = re.findall(r'kv=[0-9]+ fmin=[0-9]+ см2', item)[0]
                numbersFromElement = re.findall(r'[0-9]+',el)
                Kv_values = numbersFromElement[0]
                F_values = numbersFromElement[1]
            except:
                pass

            if len(re.sub("[^0-9]", "", str(item))) == len(str(item)):
                Kv_values = item

            if Kv_values != '':
                columnsData['Kv']['amountOfChangedFields'] += 1
                columnsData['Kv']['amountOfUnchangedFields'] -= 1
                spec['Kv'].loc[index] = Kv_values
                spec['F'].loc[index] = F_values
            else:
                columnsData['Kv']['unchangedFields'].add(spec['Kv'].loc[index])

    # text_Gmin
    if 'Gmin under ∆Pmax' in spec.columns:
        columnsData['Gmin under ∆Pmax'] = {}
        columnsData['Gmin under ∆Pmax']['amountOfEmptyFields'] = spec['Gmin under ∆Pmax'].isnull().sum()
        columnsData['Gmin under ∆Pmax']['amountOfChangedFields'] = 0
        columnsData['Gmin under ∆Pmax']['amountOfUnchangedFields'] = len(spec['Gmin under ∆Pmax'])
        columnsData['Gmin under ∆Pmax']['unchangedFields'] = set()
        for index, item in spec['Gmin under ∆Pmax'].items():
            Gmin_value = ''
            el = ''
            try:
                el = re.findall(r'[0-9]+ нм3/ч', item)[0]
                Gmin_value = re.findall(r'[0-9]+',el)[0]
            except: pass

            if len(re.sub("[^0-9]", "", item)) == len(item):
                Gmin_value = item

            if Gmin_value != '':
                columnsData['Gmin under ∆Pmax']['amountOfChangedFields'] += 1
                columnsData['Gmin under ∆Pmax']['amountOfUnchangedFields'] -= 1
                spec['Gmin under ∆Pmax'].loc[index] = Gmin_value
            else:
                columnsData['Gmin under ∆Pmax']['unchangedFields'].add(spec['Gmin under ∆Pmax'].loc[index])

    # text_Gmax
    if 'Gmax under ∆Pmin' in spec.columns:
        columnsData['Gmax under ∆Pmin'] = {}
        columnsData['Gmax under ∆Pmin']['amountOfEmptyFields'] = spec['Gmax under ∆Pmin'].isnull().sum()
        columnsData['Gmax under ∆Pmin']['amountOfChangedFields'] = 0
        columnsData['Gmax under ∆Pmin']['amountOfUnchangedFields'] = len(spec['Gmax under ∆Pmin'])
        columnsData['Gmax under ∆Pmin']['unchangedFields'] = set()
        for index, item in spec['Gmax under ∆Pmin'].items():
            Gmax_value = ''
            el = ''
            try:
                el = re.findall(r'[0-9]+ нм3/ч', item)[0]
                Gmax_value = re.findall(r'[0-9]+',el)[0]
            except: pass

            if len(re.sub("[^0-9]", "", item)) == len(item):
                Gmax_value = item

            if Gmax_value != '':
                columnsData['Gmax under ∆Pmin']['amountOfChangedFields'] += 1
                columnsData['Gmax under ∆Pmin']['amountOfUnchangedFields'] -= 1
                spec['Gmax under ∆Pmin'].loc[index] = Gmax_value
            else:
                columnsData['Gmax under ∆Pmin']['unchangedFields'].add(spec['Gmax under ∆Pmin'].loc[index])

    # text_Fluid
    if 'Fluid' in spec.columns:
        columnsData['Fluid'] = {}
        columnsData['Fluid']['amountOfEmptyFields'] = spec['Fluid'].isnull().sum()
        columnsData['Fluid']['amountOfChangedFields'] = 0
        columnsData['Fluid']['amountOfUnchangedFields'] = len(spec['Fluid'])
        columnsData['Fluid']['unchangedFields'] = set()
        if columnsData['Fluid']['amountOfUnchangedFields'] != len(spec['Fluid'].loc[spec['Fluid'] == '']):
            for index, item in spec['Fluid'].items():
                triggersInItem = text_Fluid.loc[text_Fluid['Тригер'].str.contains(item)]
                if len(triggersInItem) > 0:
                    columnsData['Fluid']['amountOfChangedFields'] += 1
                    columnsData['Fluid']['amountOfUnchangedFields'] -= 1
                    spec['Fluid'].loc[index] = ', '.join(set(triggersInItem['Fluid'].astype(str)))
                else:
                    columnsData['Fluid']['unchangedFields'].add(spec['Fluid'].loc[index])

    # Text_Connection_pipeline
    if 'Connection_pipeline' in spec.columns:
        columnsData['Connection_pipeline'] = {}
        columnsData['Connection_pipeline']['amountOfEmptyFields'] = spec['Connection_pipeline'].isnull().sum()
        columnsData['Connection_pipeline']['amountOfChangedFields'] = 0
        columnsData['Connection_pipeline']['amountOfUnchangedFields'] = len(spec['Connection_pipeline'])
        columnsData['Connection_pipeline']['unchangedFields'] = set()
        for index, item in spec['Connection_pipeline'].items():
            Connection_pipeline_value = ''
            el = ''
            try:
                el = re.findall(r'[0-9]+x[0-9]+', item)[0]
                Connection_pipeline_value = el
            except: pass


            if Connection_pipeline_value != '':
                columnsData['Connection_pipeline']['amountOfChangedFields'] += 1
                columnsData['Connection_pipeline']['amountOfUnchangedFields'] -= 1
                spec['Connection_pipeline'].loc[index] = Connection_pipeline_value
            else:
                columnsData['Connection_pipeline']['unchangedFields'].add(spec['Connection_pipeline'].loc[index])

    # text_Actuator_type
    if 'Actuator type' in spec.columns:
        if not 'RPI' in spec.columns:
            spec['RPI'] = ''
        if not 'under containment' in spec.columns:
            spec['under containment'] = ''

        columnsData['Actuator type'] = {}
        columnsData['Actuator type']['amountOfEmptyFields'] = spec['Actuator type'].isnull().sum()
        columnsData['Actuator type']['amountOfChangedFields'] = 0
        columnsData['Actuator type']['amountOfUnchangedFields'] = len(spec['Actuator type'])
        columnsData['Actuator type']['unchangedFields'] = set()

        for index, item in spec['Actuator type'].items():
            Actuator_type_values = ''
            RPI_values = ''
            under_containment_values = ''
            triggersInItem = text_Actuator_type.loc[text_Actuator_type['Тригер'].str.contains(item)]
            triggersInItem = triggersInItem.replace('nan','')
            triggersInItem['Weight'] = triggersInItem['Weight'].apply(lambda x: float(x))

            if len(triggersInItem) > 0:
                triggersWithBigWeight = triggersInItem.loc[triggersInItem['Weight'] == triggersInItem['Weight'].iloc[0]]
            else: continue

            triggersSet = set(triggersWithBigWeight['Actuator type'])
            triggersSet.discard('')
            Actuator_type_values = ', '.join(set(triggersSet))

            RPI_values = set(triggersInItem['RPI'])
            RPI_values.discard('')
            under_containment_values = set(triggersInItem['under containment'])
            under_containment_values.discard('')
            if Actuator_type_values != '':
                columnsData['Actuator type']['amountOfChangedFields'] += 1
                columnsData['Actuator type']['amountOfUnchangedFields'] -= 1
                spec['Actuator type'].loc[index] = Actuator_type_values
                if len(under_containment_values) > 0:
                    under_containment_value = next(iter(under_containment_values))
                    spec['under containment'].loc[index] = under_containment_value
                if len(RPI_values) > 0:
                    RPI_value = next(iter(RPI_values))
                    spec['RPI'].loc[index] = RPI_value
            else:
                columnsData['Actuator type']['unchangedFields'].add(spec['Actuator type'].loc[index])

    stop_changing_columns_values = time.time()
    # print('Значения в таблице заменены за: {}'.format(start_changing_columns_values-stop_changing_columns_values))

    return spec, columnsData

def findingTable(spec):
    stats = {}
    # droppig unwanted columns and rows from top and left
    spec, stats['RC_title'], stats['numberOfTitleNames'], stats['repeatedTitleNames'], stats['amountOFRepeatedTitleNames'], stats['titleIsDoubled'], stats['secondTitle'],stats['doesContainNumbersRow'],stats['titleReadingTime'], stats['normTime'] = dropTopLeftRight(spec)

    # changing columns names
    spec, changedColumnsNames, stats['changedColumnsNamesDict'] = renameColumnsNames(spec)
    stats['changedColumnsNamesAmount'] = len(changedColumnsNames)


    # finding the bottom of spec
    start_read = time.time()
    spec = dropSpecsBottom(spec,stats['titleIsDoubled'],stats['doesContainNumbersRow'])

    # находим диапазон значений
    stats['RC_rows'] = ( ( stats['RC_title'][0][0] + 1 , stats['RC_title'][0][1] ) , ( stats['RC_title'][0][0] + 1 , stats['RC_title'][1][1] ) )

    if stats['doesContainNumbersRow'][0] &  stats['titleIsDoubled']:
        stats['RC_rows'] = ((stats['RC_title'][0][0]+3,stats['RC_title'][0][1]),(spec.tail(1).index[0] + stats['RC_title'][0][0] + 1,stats['RC_title'][1][1]))
    elif stats['doesContainNumbersRow'][0] | stats['titleIsDoubled']:
        if stats['doesContainNumbersRow'][0]:
            stats['RC_rows'] = ((stats['RC_title'][0][0]+2,stats['RC_title'][0][1]),(spec.tail(1).index[0] + stats['RC_title'][0][0] + 1,stats['RC_title'][1][1]))
        else:
            stats['RC_rows'] = ((stats['RC_title'][0][0]+2,stats['RC_title'][0][1]),(spec.tail(1).index[0] + stats['RC_title'][0][0] + 1,stats['RC_title'][1][1]))


    stats['amountOfRows'] = stats['RC_rows'][1][0] - stats['RC_rows'][0][0] + 1

    end_read = time.time()
#     print('Чтение диапазона характеристик выполнено за: {}'.format(end_read-start_read))
    stats['chReadingTime'] = end_read-start_read

    spec = spec.fillna('REPLACEMENT')

    # find and merge columns with same name
    spec = findSameColumnNames(spec)
    copiesFreqDF = pd.DataFrame(spec.value_counts(subset=changedColumnsNames))

    spec['countClones'] = 0
    for i,r in enumerate(spec.iterrows()):
        spec['countClones'].iloc[i] = copiesFreqDF.loc[tuple(r[1].loc[changedColumnsNames])][0]
        copiesFreqDF.loc[tuple(r[1].loc[changedColumnsNames])][0] = 0

    spec = spec.loc[spec['countClones']!=0]

    if 'DN' in spec.columns:
        spec = spec.sort_values(by=['DN'])

    spec = spec.replace('REPLACEMENT', np.nan)

    spec , stats['columnsStats']= changingColumnsValues(spec)
    return spec, stats


def saveSpec(PATH, stats, PATH_TO_SAVE=''):

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(PATH_TO_SAVE + 'table_'+PATH.split('/')[-1], engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    spec.sort_index(inplace = True)
    spec.to_excel(writer, sheet_name='Sheet1')

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet1 = writer.sheets['Sheet1']
    worksheet2 = workbook.add_worksheet()
    worksheet3 = workbook.add_worksheet()

    # worksheet2.write('A2', pd.DataFrame.from_dict(stats))
    dictForFirstTaskKeys = {
        'RC_title': 'RC:RC для строки заголовков',
        'numberOfTitleNames':'количество заголовков',
        'repeatedTitleNames':'повторяющиеся заголовки',
        'amountOFRepeatedTitleNames':'количество повторяющихся заголовков',
        'titleIsDoubled':'двойной заголовок',
        'secondTitle':'второй (английский ) заголовок',
        'doesContainNumbersRow':'содержит строку с номерами столбцов',
        'titleReadingTime':'вреся на чтение диапазона заголовков',
        'normTime':'время на нормализацию',
        'changedColumnsNamesDict':'заменен или нет заголовок',
        'changedColumnsNamesAmount':'замененных заголовков',
        'RC_rows':'RC:RC диапазона характеристик',
        'chReadingTime':'время на чтение диапазона характеристик',
        'amountOfRows':'количество строк в диапазоне характеристик'
    }


    row = 2
    for key in stats:
        if key == 'columnsStats':
            continue
        worksheet2.write('A' + str(row), dictForFirstTaskKeys[key])
        worksheet2.write('B' + str(row), str(stats[key]))
        row = row + 1


    col = 2

    listOfUnchangedValues = set()

    def isNaN(num):
        return num != num
    for key in stats['columnsStats']:
        worksheet3.write(2, col, key)
        worksheet3.write(3, col, 'amountOfEmptyFields')
        worksheet3.write(3, col+1, stats['columnsStats'][key]['amountOfEmptyFields'])
        worksheet3.write(4, col, 'amountOfChangedFields')
        worksheet3.write(4, col+1, stats['columnsStats'][key]['amountOfChangedFields'])
        worksheet3.write(5, col, 'amountOfUnchangedFields')
        worksheet3.write(5, col+1, stats['columnsStats'][key]['amountOfUnchangedFields'])
        worksheet3.write(6, col, 'unchangedFields')
        k = 0
        listOfUnchangedValues = listOfUnchangedValues.union(set(stats['columnsStats'][key]['unchangedFields']))
        for el in set(stats['columnsStats'][key]['unchangedFields']):
            if not isNaN(el):
                worksheet3.write(6+k, col+1,el)
                k += 1

        col += 2

    listOfUnchangedValues.add('')
    listOfUnchangedValues.add(np.nan)
    # print(listOfUnchangedValues)

    # Add a format. Light red fill with dark red text.
    format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                   'font_color': '#9C0006'})

    # Add a header format.
    copy_columns_header_cell_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1})

    text_columns_header_cell_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E55В',
        'border': 1})

    eng_vars_columns_header_cell_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7EССС',
        'border': 1})

    usuals_columns_header_cell_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#FFFFFF',
        'border': 1})

    unrecognized_cell_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#444444',
        'border': 1})


    # Write the column headers with the defined format.
    for col_num, value in enumerate(spec.columns.values):
        if value in set(copy_columns.columns):
            worksheet1.write(0, col_num + 1, value, copy_columns_header_cell_format)
        elif value in set(text_columns.columns):
            worksheet1.write(0, col_num + 1, value, text_columns_header_cell_format)
        elif value in set(eng_vars_columns.columns):
            worksheet1.write(0, col_num + 1, value, eng_vars_columns_header_cell_format)
        else:
            worksheet1.write(0, col_num + 1, value, usuals_columns_header_cell_format)

    for (columnName, columnData) in spec.iteritems():
        if not columnName in stats['changedColumnsNamesDict'].keys():

            # print(list(columnData.loc[columnData.isin(listOfUnchangedValues)].index))
            indicesList = list(columnData.loc[columnData.isin(listOfUnchangedValues)].index)
            # if columnName == 'KKS' :
                # print('{}: {}'.format(columnName,columnData.isin(listOfUnchangedValues)))
            indicesList.sort()
            # print(indicesList)
            for value in indicesList:

                # print(spec[columnName][spec.index.get_loc(value)])
                # print(isNaN(spec[columnName][spec.index.get_loc(value)]))
                if not isNaN(spec.loc[value][columnName]):
                    # print(spec[columnName][spec.index.get_loc(value)])
                    worksheet1.write(columnData.index.get_loc(value)+1, spec.columns.get_loc(columnName)+1, spec.loc[value][columnName], unrecognized_cell_format)

    # print(listOfUnchangedValues)
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()



if __name__ == '__main__':
    print('Введите путь к файлу')

    PATH = input()

    print('Введите путь для сохранения файла')

    PATH_TO_SAVE = input()

    if PATH_TO_SAVE != '':
        PATH_TO_SAVE += '/'


    xls = pd.ExcelFile('vars_for_columns.xlsx')

    copy_columns = pd.read_excel(xls, 'copy_columns').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    text_columns = pd.read_excel(xls, 'text_columns').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    eng_vars_columns = pd.read_excel(xls, 'eng_vars_columns').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())

    eng_vars_Connection_columns = pd.read_excel('english_vars_Connection.xlsx').replace('\s+', ' ', regex=True)
    eng_vars_Connection_columns.connection_ru = eng_vars_Connection_columns.connection_ru.str.lower()
    eng_vars_Connection_columns = eng_vars_Connection_columns[['connection_ru','connection']]
    eng_vars_Connection_columns = eng_vars_Connection_columns.set_index('connection_ru').T.to_dict('list')

    eng_vars_Material_columns = pd.read_excel('english_vars_Material.xlsx').replace('\s+', ' ', regex=True)
    eng_vars_Material_columns.material_ru = eng_vars_Material_columns.material_ru.str.lower()
    eng_vars_Material_columns = eng_vars_Material_columns[['material_ru','material']]
    eng_vars_Material_columns = eng_vars_Material_columns.set_index('material_ru').T.to_dict('list')

    eng_vars_SeismoCat_columns = pd.read_excel('english_vars_SeismoCat.xlsx').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    eng_vars_SeismoCat_columns = eng_vars_SeismoCat_columns[['SeismoCat_ru','SeismoCat']]
    eng_vars_SeismoCat_columns = eng_vars_SeismoCat_columns.set_index('SeismoCat_ru').T.to_dict('list')

    varsDF = pd.concat([copy_columns,text_columns,eng_vars_columns],axis=1)

    varsNamesSet = set(copy_columns.stack().tolist())|set(text_columns.stack().tolist())|set(eng_vars_columns.stack().tolist())
    varsNamesSet.discard('nan')

    text_xls = pd.ExcelFile('Text_dict_ver.2.xlsx')

    text_Type = pd.read_excel(text_xls, 'text_Type').replace('\s+', ' ', regex=True).astype(str)
    text_Type['Тригер'] = text_Type['Тригер'].apply(lambda x: x.lower())
    text_Time = pd.read_excel(text_xls, 'Text_Time').replace('\s+', ' ', regex=True).astype(str)
    text_Time['Тригер'] = text_Time['Тригер'].apply(lambda x: x.lower())
    text_NC = pd.read_excel(text_xls, 'Text_NC').replace('\s+', ' ', regex=True).astype(str)
    text_NC['Тригер'] = text_NC['Тригер'].apply(lambda x: x.lower())
    text_Kv = pd.read_excel(text_xls, 'Text_Kv').replace('\s+', ' ', regex=True).astype(str)
    text_Kv['Тригер'] = text_Kv['Тригер'].apply(lambda x: x.lower())
    text_Gmin = pd.read_excel(text_xls, 'Text_Gmin').replace('\s+', ' ', regex=True).astype(str)
    text_Gmin['Тригер'] = text_Gmin['Тригер'].apply(lambda x: x.lower())
    text_Gmax = pd.read_excel(text_xls, 'Text_Gmax').replace('\s+', ' ', regex=True).astype(str)
    text_Gmax['Тригер'] = text_Gmax['Тригер'].apply(lambda x: x.lower())
    text_Fluid = pd.read_excel(text_xls, 'Text_Fluid').replace('\s+', ' ', regex=True).astype(str)
    text_Fluid['Тригер'] = text_Fluid['Тригер'].apply(lambda x: x.lower())
    text_Gmax = pd.read_excel(text_xls, 'Text_Gmax').replace('\s+', ' ', regex=True).astype(str)
    text_Gmax['Тригер'] = text_Gmax['Тригер'].apply(lambda x: x.lower())
    text_Connection_pipeline = pd.read_excel(text_xls, 'Text_Connection_pipeline').replace('\s+', ' ', regex=True).astype(str)
    text_Connection_pipeline['Тригер'] = text_Connection_pipeline['Тригер'].apply(lambda x: x.lower())
    text_Actuator_type = pd.read_excel(text_xls, 'Text_Actuator_type').replace('\s+', ' ', regex=True).astype(str)
    text_Actuator_type['Тригер'] = text_Actuator_type['Тригер'].apply(lambda x: x.lower())

    if (PATH.split('.')[-1] == 'xls' )|(PATH.split('.')[-1] == 'xlsx' ):
        try:
            spec = pd.read_excel(PATH)
            spec, stats = findingTable(spec)
            if not spec.empty:
                saveSpec(PATH,stats,PATH_TO_SAVE)
            else:
                print(PATH)
        except FileNotFoundError:
            print('File not found')
        except xlrd.biffh.XLRDError:
            print('Wrong format')
        except xlsxwriter.exceptions.FileCreateError:
            checkIfDirContainsFiles = True
            print('No such directory to save file')
    else:
        checkIfDirContainsFiles = False
        try:
            for entry in os.listdir(PATH):
                if os.path.isfile(os.path.join(PATH, entry)):
                    if (entry.split('.')[-1] == 'xls' )|(entry.split('.')[-1] == 'xlsx' ):
                        checkIfDirContainsFiles = True
                        spec = pd.read_excel(os.path.join(PATH, entry))
                        spec, stats = findingTable(spec)
                        if not spec.empty:
                            saveSpec(entry, stats, PATH_TO_SAVE)
                        else:
                            print(os.path.join(PATH, entry))
        except FileNotFoundError:
            checkIfDirContainsFiles = True
            print('No such directory')
        except xlsxwriter.exceptions.FileCreateError:
            checkIfDirContainsFiles = True
            print('No such directory to save file')
        if not checkIfDirContainsFiles:
            print('No Files in directory')

            
            
            
def run():
    print('Введите путь к файлу')

    PATH = input()

    print('Введите путь для сохранения файла')

    PATH_TO_SAVE = input()

    if PATH_TO_SAVE != '':
        PATH_TO_SAVE += '/'


    xls = pd.ExcelFile('vars_for_columns.xlsx')

    copy_columns = pd.read_excel(xls, 'copy_columns').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    text_columns = pd.read_excel(xls, 'text_columns').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    eng_vars_columns = pd.read_excel(xls, 'eng_vars_columns').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())

    eng_vars_Connection_columns = pd.read_excel('english_vars_Connection.xlsx').replace('\s+', ' ', regex=True)
    eng_vars_Connection_columns.connection_ru = eng_vars_Connection_columns.connection_ru.str.lower()
    eng_vars_Connection_columns = eng_vars_Connection_columns[['connection_ru','connection']]
    eng_vars_Connection_columns = eng_vars_Connection_columns.set_index('connection_ru').T.to_dict('list')

    eng_vars_Material_columns = pd.read_excel('english_vars_Material.xlsx').replace('\s+', ' ', regex=True)
    eng_vars_Material_columns.material_ru = eng_vars_Material_columns.material_ru.str.lower()
    eng_vars_Material_columns = eng_vars_Material_columns[['material_ru','material']]
    eng_vars_Material_columns = eng_vars_Material_columns.set_index('material_ru').T.to_dict('list')

    eng_vars_SeismoCat_columns = pd.read_excel('english_vars_SeismoCat.xlsx').replace('\s+', ' ', regex=True).astype(str).apply(lambda x: x.str.lower())
    eng_vars_SeismoCat_columns = eng_vars_SeismoCat_columns[['SeismoCat_ru','SeismoCat']]
    eng_vars_SeismoCat_columns = eng_vars_SeismoCat_columns.set_index('SeismoCat_ru').T.to_dict('list')

    varsDF = pd.concat([copy_columns,text_columns,eng_vars_columns],axis=1)

    global varsNamesSet = set(copy_columns.stack().tolist())|set(text_columns.stack().tolist())|set(eng_vars_columns.stack().tolist())
    varsNamesSet.discard('nan')

    text_xls = pd.ExcelFile('Text_dict_ver.2.xlsx')

    text_Type = pd.read_excel(text_xls, 'text_Type').replace('\s+', ' ', regex=True).astype(str)
    text_Type['Тригер'] = text_Type['Тригер'].apply(lambda x: x.lower())
    text_Time = pd.read_excel(text_xls, 'Text_Time').replace('\s+', ' ', regex=True).astype(str)
    text_Time['Тригер'] = text_Time['Тригер'].apply(lambda x: x.lower())
    text_NC = pd.read_excel(text_xls, 'Text_NC').replace('\s+', ' ', regex=True).astype(str)
    text_NC['Тригер'] = text_NC['Тригер'].apply(lambda x: x.lower())
    text_Kv = pd.read_excel(text_xls, 'Text_Kv').replace('\s+', ' ', regex=True).astype(str)
    text_Kv['Тригер'] = text_Kv['Тригер'].apply(lambda x: x.lower())
    text_Gmin = pd.read_excel(text_xls, 'Text_Gmin').replace('\s+', ' ', regex=True).astype(str)
    text_Gmin['Тригер'] = text_Gmin['Тригер'].apply(lambda x: x.lower())
    text_Gmax = pd.read_excel(text_xls, 'Text_Gmax').replace('\s+', ' ', regex=True).astype(str)
    text_Gmax['Тригер'] = text_Gmax['Тригер'].apply(lambda x: x.lower())
    text_Fluid = pd.read_excel(text_xls, 'Text_Fluid').replace('\s+', ' ', regex=True).astype(str)
    text_Fluid['Тригер'] = text_Fluid['Тригер'].apply(lambda x: x.lower())
    text_Gmax = pd.read_excel(text_xls, 'Text_Gmax').replace('\s+', ' ', regex=True).astype(str)
    text_Gmax['Тригер'] = text_Gmax['Тригер'].apply(lambda x: x.lower())
    text_Connection_pipeline = pd.read_excel(text_xls, 'Text_Connection_pipeline').replace('\s+', ' ', regex=True).astype(str)
    text_Connection_pipeline['Тригер'] = text_Connection_pipeline['Тригер'].apply(lambda x: x.lower())
    text_Actuator_type = pd.read_excel(text_xls, 'Text_Actuator_type').replace('\s+', ' ', regex=True).astype(str)
    text_Actuator_type['Тригер'] = text_Actuator_type['Тригер'].apply(lambda x: x.lower())

    if (PATH.split('.')[-1] == 'xls' )|(PATH.split('.')[-1] == 'xlsx' ):
        try:
            spec = pd.read_excel(PATH)
            spec, stats = findingTable(spec)
            if not spec.empty:
                saveSpec(PATH,stats,PATH_TO_SAVE)
            else:
                print(PATH)
        except FileNotFoundError:
            print('File not found')
        except xlrd.biffh.XLRDError:
            print('Wrong format')
        except xlsxwriter.exceptions.FileCreateError:
            checkIfDirContainsFiles = True
            print('No such directory to save file')
    else:
        checkIfDirContainsFiles = False
        try:
            for entry in os.listdir(PATH):
                if os.path.isfile(os.path.join(PATH, entry)):
                    if (entry.split('.')[-1] == 'xls' )|(entry.split('.')[-1] == 'xlsx' ):
                        checkIfDirContainsFiles = True
                        spec = pd.read_excel(os.path.join(PATH, entry))
                        spec, stats = findingTable(spec)
                        if not spec.empty:
                            saveSpec(entry, stats, PATH_TO_SAVE)
                        else:
                            print(os.path.join(PATH, entry))
        except FileNotFoundError:
            checkIfDirContainsFiles = True
            print('No such directory')
        except xlsxwriter.exceptions.FileCreateError:
            checkIfDirContainsFiles = True
            print('No such directory to save file')
        if not checkIfDirContainsFiles:
            print('No Files in directory')

