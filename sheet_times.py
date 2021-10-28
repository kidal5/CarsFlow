import xlsxwriter as xw
import xlsxwriter.utility as xwu
import pandas as pd


def createSheetTimes(df, xlsxWriter, params):
    sheet_name = 'Časové údaje'
    workbook = xlsxWriter.book
    worksheet = workbook.add_worksheet(sheet_name)

    currentColumnShift = 0

    params = params['sheet_times']
    for key in params.keys():
        selectedCombinedIndexes = params[key]['selected_directions']
        startTime = params[key]['time_start']
        endTime = params[key]['time_end']

        data = computeData(df, selectedCombinedIndexes, startTime, endTime)
        writeTemplate(workbook, worksheet, selectedCombinedIndexes, startTime, endTime, currentColumnShift)
        writeData(workbook, worksheet, data, currentColumnShift)
        currentColumnShift = currentColumnShift + 20

        print(key)

    xlsxWriter.close()


def computeData(df, selectedCombinedIndexes, startTime, endTime):
    combinedIndexesString = "_".join([f'{i}' for i in selectedCombinedIndexes])

    def combinedIndexesFilter(x):
        string = "_".join([f'{i}' for i in x])
        return combinedIndexesString in string

    # filter camera spots
    temp = df
    temp = temp.drop(columns=['Place sheet name'])
    # print(len(temp))

    temp = temp[temp['Place combined index'].isin(selectedCombinedIndexes)]
    # print(len(df))

    # filter time
    df.set_index('Capture Time').between_time(startTime, endTime).reset_index()
    # print(len(temp))

    # filter unknown
    temp = temp[temp['License Plate Number'] != 'unknown']
    # print(len(temp))

    # filter cars that do not have enough transit data
    temp_platesFilter = temp.groupby(by='License Plate Number').nunique().reset_index()
    temp_platesFilter = temp_platesFilter[temp_platesFilter['Place combined index'] >= len(selectedCombinedIndexes)]
    temp = temp[temp['License Plate Number'].isin(temp_platesFilter['License Plate Number'])]
    # print(len(temp))

    # filter cars that have transit in wrong order
    temp_wrongOrderFilter = temp.sort_values(['License Plate Number', 'Capture Time'], ascending=True).groupby(
        by='License Plate Number')
    temp_wrongOrderFilter = temp_wrongOrderFilter[['Place combined index']].agg(combinedIndexesFilter).reset_index()
    temp_wrongOrderFilter = temp_wrongOrderFilter[temp_wrongOrderFilter['Place combined index']]
    temp = temp[temp['License Plate Number'].isin(temp_wrongOrderFilter['License Plate Number'])]
    # print(len(temp))

    # create wide format
    temp = temp.sort_values(['License Plate Number', 'Capture Time'], ascending=True).reset_index(drop=True)
    temp['index'] = temp.index
    temp = temp.rename(
        columns={"Capture Time": f"Capture Time_{0}", "Place combined index": f"Place combined index_{0}"})

    temp_save = temp

    for i in range(1, len(selectedCombinedIndexes)):
        temp_moved = temp_save.copy(deep=True)
        temp_moved['index'] = temp_moved.index - i
        temp_moved = temp_moved.rename(
            columns={"Capture Time_0": f"Capture Time_{i}", "Place combined index_0": f"Place combined index_{i}"})

        temp = temp.merge(temp_moved, how='inner', on=['index', 'License Plate Number'])

    temp = temp.drop(columns=['index'])

    def wideFormatCombinedIndexesFilter(row):

        toCheckIndexes = [i * 2 + 2 for i in range(len(selectedCombinedIndexes))]
        toCheck = row[toCheckIndexes].tolist()

        return selectedCombinedIndexes == toCheck

    # filter wide format
    temp = temp[temp.apply(wideFormatCombinedIndexesFilter, axis=1)]

    # add time difference
    for i in range(1, len(selectedCombinedIndexes)):
        temp[f'time_{i - 1}'] = temp[f'Capture Time_{i}'] - temp[f'Capture Time_{i - 1}']

    # remove unused columns
    for i in range(0, len(selectedCombinedIndexes)):
        temp = temp.drop(columns=[f'Capture Time_{i}', f'Place combined index_{i}'])

    # convert time columns to better number
    for i in range(0, len(selectedCombinedIndexes) - 1):
        temp[f'time_{i}'] = temp[f'time_{i}'].dt.seconds / 60

    return temp


def writeData(workbook, worksheet, data, colShift=10):
    SPZ_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'font_size': 11,
        'bg_color': '#D9E1F2',
    })

    number_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'font_size': 11,
        'num_format': '0.0'
    })

    SUM_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'font_size': 11,
        'num_format': '0',
        'bg_color': '#B7DEE8'
    })

    for row in range(data.shape[0]):
        worksheet.write(row + 3, colShift, data.iloc[row, 0], SPZ_format)

    for i in range(1, data.shape[1]):
        for row in range(data.shape[0]):
            worksheet.write(row + 3, colShift + i, data.iloc[row, i], number_format)

    # write sums
    for row in range(data.shape[0]):
        a = xwu.xl_col_to_name(colShift + 1)
        b = xwu.xl_col_to_name(colShift + data.shape[1] - 1)

        worksheet.write(row + 3, colShift + data.shape[1], f'=SUM({a}{row + 4}:{b}{row + 4})', SUM_format)


def writeTemplate(workbook, worksheet, selectedCombinedIndexes, startTime, endTime, colShift=10):
    header_format = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter',
        'text_wrap': True,
    })

    SPZ_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#D9E1F2',
        'text_wrap': True,
    })

    SUM_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'text_wrap': True,
        'bg_color': '#B7DEE8'
    })

    DIRECTION_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#FCE4D6',
        'text_wrap': True,
    })

    # write header
    worksheet.set_row(0, 40)
    header_text = "Označené směry " + ", ".join([str(i) for i in selectedCombinedIndexes])
    header_text = f'{header_text}\nPočátek: {startTime}, Konec: {endTime}'

    worksheet.merge_range(0, colShift, 0, colShift + len(selectedCombinedIndexes), header_text, header_format)
    worksheet.merge_range(1, colShift, 2, colShift, "SPZ", SPZ_format)
    worksheet.merge_range(1, colShift + len(selectedCombinedIndexes), 2, colShift + len(selectedCombinedIndexes),
                          "Celkem [min]", SUM_format)
    if len(selectedCombinedIndexes) == 2:
        worksheet.write(1, colShift + 1, 'Označení směru', DIRECTION_format)
    else:
        worksheet.merge_range(1, colShift + 1, 1, colShift + len(selectedCombinedIndexes) - 1, "Označení směru",
                              DIRECTION_format)

    for i in range(len(selectedCombinedIndexes) - 1):
        text = f'{selectedCombinedIndexes[i]}-{selectedCombinedIndexes[i + 1]} [min]'
        worksheet.write(2, colShift + 1 + i, text, DIRECTION_format)
