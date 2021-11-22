import xlsxwriter as xw
import xlsxwriter.utility as xwu
import pandas as pd

from TimeStruct import *


def createSheetNumberOfCars(df, xlsxWriter, params):
    def createSheetNumberOfCarsInner(sheet_name_in, time_in: TimeStruct, addDataCheck):
        data = computeData(df, params, time_in, addDataCheck)
        writeData(xlsxWriter, sheet_name_in, data, params, addDataCheck)
        writeTemplate(xlsxWriter, sheet_name_in, params, time_in, addDataCheck)

    sheet_name = 'Počty vozidel'
    createSheetNumberOfCarsInner(sheet_name, TimeStruct.createFromStartAndEndTime('00:00', '23:59', df), True)

    for key in params['sheet_cars_count']:
        item = params['sheet_cars_count'][key]

        time = TimeStruct.createFromDict(item, df)

        sheet_name_edited = f'{sheet_name} {time.sheetName}'
        createSheetNumberOfCarsInner(sheet_name_edited, time, False)


def computeData(df, params, time: TimeStruct, addDataCheck=True):
    # compute number of items on every input sheet, for data validation
    dataCheck = df.groupby(by='Direction').count().drop(columns=['License_plate'])

    # filter time
    df = df.set_index('Capture_time').sort_index()
    df = df[time.dateTimeStart: time.dateTimeEnd]
    df = df.reset_index()

    # add fake data to force dataframe layout, aka have at least one entry for every possible combination
    fake_df = createFakeDataset(params['number_of_cameras'])
    df = df.append(fake_df, ignore_index=True)

    # separate dataset into two parts based on license plate count
    plate_counts = df["License_plate"].value_counts()
    plate_counts['unknown'] = 0  # fake condition value of unknown type, so they are selected into single dataset
    plate_count = plate_counts[plate_counts > 1]

    df_multiple = df[df['License_plate'].isin(plate_count.index)]
    df_single = df[~df['License_plate'].isin(plate_count.index)]

    # make sense only in full dataset...
    if addDataCheck:
        df_multiple = df_multiple.append(createFakeEndStopForDataValidityCheck(df_multiple))

    # create another dataset moved by one and merge it. Aka create pairs of all following directions
    df_multiple = df_multiple.sort_values(by=['License_plate', 'Capture_time']).reset_index(drop=True)
    df_multiple['joinIndex'] = df_multiple.index
    df_multiple_moved = df_multiple.copy(deep=True)
    df_multiple_moved['joinIndex'] = df_multiple_moved.index - 1
    df_multiple_join = df_multiple.merge(df_multiple_moved, how='inner', on=['joinIndex', 'License_plate'],
                                         suffixes=['_from', '_to'])

    crosstab = pd.crosstab(df_multiple_join['Direction_from'],
                           df_multiple_join['Direction_to'])
    crosstab = crosstab - 1  # remove fake data

    # create half cross tab from single data
    halfcrosstab = df_single.groupby('Direction').count()
    halfcrosstab = halfcrosstab - 1  # remove fake datax;
    halfcrosstab = halfcrosstab.drop(columns=['Capture_time'])
    halfcrosstab_to = halfcrosstab[halfcrosstab.index % 2 == 1].T
    halfcrosstab_from = halfcrosstab[halfcrosstab.index % 2 == 0].T

    return {'multiple': crosstab, 'singleTo': halfcrosstab_to, 'singleFrom': halfcrosstab_from, 'dataCheck': dataCheck}


def createFakeDataset(N):
    fake_df = {'Capture_time': [], 'License_plate': [], 'Direction': []}
    fake_start_time = pd.to_datetime('00:00')
    fake_end_time = pd.to_datetime('01:00')

    for i in range(1, N * 2 + 1):
        fake_df['Capture_time'].append(fake_start_time)
        fake_df['License_plate'].append(f'fake_{i}')
        fake_df['Direction'].append(i)

        for k in range(1, N * 2 + 1):
            fake_df['Capture_time'].append(fake_start_time)
            fake_df['License_plate'].append(f'fake_{i}_{k}')
            fake_df['Direction'].append(i)
            fake_df['Capture_time'].append(fake_end_time)
            fake_df['License_plate'].append(f'fake_{i}_{k}')
            fake_df['Direction'].append(k)

    return pd.DataFrame.from_dict(fake_df)


def createFakeEndStopForDataValidityCheck(df):
    endStopDf = pd.DataFrame.from_dict({'Capture_time': [], 'License_plate': [], 'Direction': []})

    endStopDf['License_plate'] = df['License_plate'].unique()
    endStopDf['Direction'] = 100
    endStopDf['Capture_time'] = pd.to_datetime('23:59')
    endStopDf = endStopDf[~endStopDf['License_plate'].str.contains('fake_')]
    return endStopDf


def writeData(xlsxWriter, sheet_name, data, params, addDataCheck):
    # Position the dataframes in the worksheet.

    N = params['number_of_cameras']

    data['multiple'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=5, startcol=2, header=False, index=False,
                              columns=[i + 1 for i in range(N * 2)])

    data['singleTo'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=5 + N * 2 + 3, startcol=2, header=False,
                              index=False)
    data['singleFrom'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=5 + N * 2 + 7, startcol=2, header=False,
                                index=False)

    if addDataCheck:
        data['multiple'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=5, startcol=N * 2 + 3, header=False,
                                  index=False, columns=[100])
        data['dataCheck'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=5, startcol=N * 2 + 5, header=False,
                                   index=False)

    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    border_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
    })

    good_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#D8E4BC'
    })

    wrong_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'bg_color': '#FF0000'
    })

    cond = {'type': 'cell', 'criteria': 'greater than', 'value': -1, 'format': border_format}
    cond_good = {'type': 'cell', 'criteria': '<>', 'value': 0, 'format': wrong_format}
    cond_wrong = {'type': 'cell', 'criteria': '=', 'value': 0, 'format': good_format}

    # this should not be conditional formatting, but i could not found way hwo to do it properly...
    worksheet.conditional_format(5, 2, N * 2 + 4, N * 2 + 1, cond)
    if addDataCheck:
        worksheet.conditional_format(5, N * 2 + 3, N * 2 + 4, N * 2 + 5, cond)
        worksheet.conditional_format(5, N * 2 + 6, N * 2 + 4, N * 2 + 6, cond_good)
        worksheet.conditional_format(5, N * 2 + 6, N * 2 + 4, N * 2 + 6, cond_wrong)

    worksheet.conditional_format(5 + N * 2 + 3, 2, 4 + N * 2 + 4, N + 1, cond)
    worksheet.conditional_format(5 + N * 2 + 7, 2, 4 + N * 2 + 8, N + 1, cond)


def writeTemplate(xlsxWriter, sheet_name, params, time: TimeStruct, addDataCheck):
    N = params['number_of_cameras']

    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    # column width and formats
    worksheet.set_column(0, N * 2 + 1, 9)
    worksheet.set_column(N * 2 + 3, N * 2 + 3, 20)
    worksheet.set_column(N * 2 + 4, N * 2 + 4, 15)
    worksheet.set_column(N * 2 + 5, N * 2 + 6, 12)

    # formats
    first_line_format = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'font_size': 12
    })

    orange_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'num_format': '0',
        'bg_color': '#FCE4D6'
    })

    blue_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'font_size': 11,
        'num_format': '0',
        'bg_color': '#D9E1F2'
    })

    border_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
    })

    # write first and second row
    worksheet.merge_range(0, 0, 0, N * 2 + 1, "Základní výstupní tabulky (počty všech vozidel)", first_line_format)
    worksheet.merge_range(1, 0, 1, N * 2 + 1, f"Vybraný čas: {time.fullSheetName}", first_line_format)

    # write third row and first column
    worksheet.write(3, 0, "Sčítací bod", orange_format)
    worksheet.write(3, 1, "", orange_format)

    for i in range(N):
        worksheet.merge_range(3, i * 2 + 2, 3, i * 2 + 3, i + 1, orange_format)
        worksheet.merge_range(i * 2 + 5, 0, i * 2 + 6, 0, i + 1, orange_format)

    # write fourth row and second column
    worksheet.write(4, 0, "", orange_format)
    worksheet.write(4, 1, "Směr", blue_format)
    for i in range(N * 2):
        worksheet.write(4, i + 2, i + 1, blue_format)
        worksheet.write(i + 5, 1, i + 1, blue_format)

    # write single cars template
    worksheet.merge_range(N * 2 + 6, 0, N * 2 + 6, N + 1, "Pouze DO města ", first_line_format)
    worksheet.merge_range(N * 2 + 10, 0, N * 2 + 10, N + 1, "Pouze Z města ", first_line_format)

    worksheet.merge_range(N * 2 + 7, 0, N * 2 + 7, 1, "Sčítací bod", orange_format)
    worksheet.merge_range(N * 2 + 11, 0, N * 2 + 11, 1, "Sčítací bod", orange_format)

    worksheet.merge_range(N * 2 + 8, 0, N * 2 + 8, 1, "DO města", orange_format)
    worksheet.merge_range(N * 2 + 12, 0, N * 2 + 12, 1, "Z města", orange_format)

    for i in range(N):
        worksheet.write(N * 2 + 7, i + 2, f'{i + 1} ({i * 2 + 1})', orange_format)
        worksheet.write(N * 2 + 11, i + 2, f'{i + 1} ({i * 2 + 2})', orange_format)

    # write data validity checks
    if addDataCheck:
        worksheet.write(3, N * 2 + 3, "Zástupný koncový bod", orange_format)
        worksheet.write(3, N * 2 + 4, "Pomocný součet", orange_format)
        worksheet.write(3, N * 2 + 5, "Cílový součet", orange_format)
        worksheet.write(3, N * 2 + 6, "Validace dat", orange_format)

        worksheet.write(4, N * 2 + 3, "", blue_format)
        worksheet.write(4, N * 2 + 4, "", blue_format)
        worksheet.write(4, N * 2 + 5, "", blue_format)
        worksheet.write(4, N * 2 + 6, "", blue_format)

        # write data into Pomocný součet column
        for i in range(N * 2):
            a = xwu.xl_col_to_name(2)
            b = xwu.xl_col_to_name(N * 2 + 1)
            c = xwu.xl_col_to_name(N * 2 + 3)

            d = xwu.xl_col_to_name(2 + i // 2)
            di = 5 + N * 2 + 4

            if i % 2 == 1:
                di = di + 4

            cellText = f'=SUM({a}{6 + i}:{b}{6 + i}) + {c}{6 + i} + {d}{di} + 1'
            worksheet.write(5 + i, N * 2 + 4, cellText, border_format)

        # write data into Validace dat column
        for i in range(N * 2):
            a = xwu.xl_col_to_name(N * 2 + 4)
            b = xwu.xl_col_to_name(N * 2 + 5)

            cellText = f'={a}{6 + i} - {b}{6 + i}'
            worksheet.write(5 + i, N * 2 + 6, cellText, border_format)
