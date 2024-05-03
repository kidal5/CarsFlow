import xlsxwriter as xw
import xlsxwriter.utility as xwu
import pandas as pd

from TimeStruct import *


def createSheetTimes(df, xlsxWriter, params):
    sheet_name = 'Časové údaje'
    workbook = xlsxWriter.book
    worksheet = workbook.add_worksheet(sheet_name)

    currentColumnShift = 0

    for item in params['sheet_times'].values():
        selectedDirections = item['selected_directions']

        time = TimeStruct.createFromDict(item, df)
        sc = item.get('selected_categories', [])
        if len(sc) == 0:
            sc = None

        data = computeData(df, selectedDirections, time, sc)
        writeTemplate(workbook, worksheet, selectedDirections, time, sc, currentColumnShift)
        writeData(workbook, worksheet, selectedDirections, data, currentColumnShift)
        currentColumnShift = currentColumnShift + len(selectedDirections) * 2 + 2


def computeData(df, selectedDirections, time: TimeStruct, selectedCategories):
    combinedIndexesString = "_".join([f'{i}' for i in selectedDirections])

    def checkForCorrectDirections(x):
        values = [str(x[f'Direction_{i}']) for i in range(len(selectedDirections))]
        return combinedIndexesString == "_".join(values)

    temp = df.copy(deep=True)
    temp = temp[temp['License_plate'] != 'unknown']

    # filter time
    temp = temp.set_index('Capture_time').sort_index()
    temp = temp[time.dateTimeStart: time.dateTimeEnd]
    if selectedCategories is not None:
        temp = temp[temp['Vehicle_category'].isin(selectedCategories)]

    # drop vehicle category column as it is not needed anymore and merging down below does not work
    temp = temp.drop(columns=['Vehicle_category'], errors='ignore')
    temp = temp.reset_index()

    # add fake data to assure that everything goes smoothly
    fake_data = {'License_plate': [], 'Capture_time': [], 'Direction': []}
    for i, dire in enumerate(selectedDirections):
        fake_data['License_plate'].append('FakeSPZ')
        fake_data['Capture_time'].append(pd.to_datetime(f'00:{i:02d}'))
        fake_data['Direction'].append(dire)
    fake_df = pd.DataFrame().from_dict(fake_data)
    temp = pd.concat([temp, fake_df], ignore_index=True)

    # create wide format
    temp = temp.sort_values(['License_plate', 'Capture_time'], ascending=True).reset_index(drop=True)
    temp['index'] = temp.index
    temp = temp.rename(
        columns={"Capture_time": f"Capture_time_{0}", "Direction": f"Direction_{0}"})

    temp_save = temp
    for i in range(1, len(selectedDirections)):
        temp_moved = temp_save.copy(deep=True)
        temp_moved['index'] = temp_moved.index - i
        temp_moved = temp_moved.rename(
            columns={"Capture_time_0": f"Capture_time_{i}", "Direction_0": f"Direction_{i}"})

        temp = temp.merge(temp_moved, how='inner', on=['index', 'License_plate'])

    temp = temp.drop(columns=['index'])
    temp = temp[temp.agg(checkForCorrectDirections, axis=1)]

    # remove fake data
    temp = temp[temp['License_plate'] != 'FakeSPZ']

    # add time difference
    for i in range(1, len(selectedDirections)):
        temp[f'Diff_time_{i - 1}'] = temp[f'Capture_time_{i}'] - temp[f'Capture_time_{i - 1}']

    # remove unused columns
    for i in range(0, len(selectedDirections)):
        temp = temp.drop(columns=[f'Direction_{i}'])

    # convert time columns to better number
    for i in range(0, len(selectedDirections) - 1):
        temp[f'Diff_time_{i}'] = temp[f'Diff_time_{i}'].dt.total_seconds() / 60

    # convert capture time to string, so excel writer do not mess it up...
    for i in range(0, len(selectedDirections)):
        temp[f'Capture_time_{i}'] = temp[f'Capture_time_{i}'].dt.strftime('%d/%m/%Y - %H:%M:%S')

    # reorder columns
    columns = ['License_plate'] + [f'Capture_time_{i}' for i in range(len(selectedDirections))] \
              + [f'Diff_time_{i}' for i in range(len(selectedDirections) - 1)]
    temp = temp[columns]

    return temp


def writeData(workbook, worksheet, selectedDirections, data, colShift=10):
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
        a = xwu.xl_col_to_name(colShift + 1 + len(selectedDirections))
        b = xwu.xl_col_to_name(colShift + data.shape[1] - 1)

        worksheet.write(row + 3, colShift + data.shape[1], f'=SUM({a}{row + 4}:{b}{row + 4})', SUM_format)


def writeTemplate(workbook, worksheet, selectedDirections, time: TimeStruct, selectedCategories, colShift=10):
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
    worksheet.set_row(0, 60)
    worksheet.set_column(colShift, colShift + len(selectedDirections) * 2, 10)  # make size of all columns to 10
    worksheet.set_column(colShift + 1, colShift + len(selectedDirections), 20)  # make size of datetime columns to 20

    header_text = "Označené směry " + ", ".join([str(i) for i in selectedDirections])
    categoriesStr = f'Vybrané typy vozidel: {", ".join(selectedCategories) if selectedCategories is not None else "Všechny kategorie"}'
    header_text = f'{header_text}\n{time.fullSheetName}\n{categoriesStr}'

    worksheet.merge_range(0, colShift, 0, colShift + len(selectedDirections) * 2, header_text, header_format)
    worksheet.merge_range(1, colShift, 2, colShift, "SPZ", SPZ_format)
    worksheet.merge_range(1, colShift + len(selectedDirections) * 2, 2, colShift + len(selectedDirections) * 2,
                          "Celkem [min]", SUM_format)

    worksheet.merge_range(1, colShift + 1, 1, colShift + len(selectedDirections) * 2 - 1, "Označení směru",
                          DIRECTION_format)

    for i in range(len(selectedDirections)):
        text = f'Průjezd {selectedDirections[i]}'
        worksheet.write(2, colShift + 1 + i, text, DIRECTION_format)

    for i in range(len(selectedDirections) - 1):
        text = f'{selectedDirections[i]}-{selectedDirections[i + 1]} [min]'
        worksheet.write(2, colShift + 1 + i + len(selectedDirections), text, DIRECTION_format)
