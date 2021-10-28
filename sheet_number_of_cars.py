import xlsxwriter as xw
import xlsxwriter.utility as xwu
import pandas as pd

import load_utils


def createSheetNumberOfCars(df, xlsxWriter, params):
    sheet_name = 'Počty vozidel'

    data = computeData(df)
    writeData(xlsxWriter, sheet_name, data, params)
    writeTemplate(xlsxWriter, sheet_name, params)


def computeData(df, startTime='00:00', endTime='23:59'):
    # filter time
    df.set_index('Capture Time').between_time(startTime, endTime).reset_index()

    # separate dataset into two parts based on license plate count
    plate_counts = df["License Plate Number"].value_counts()
    plate_counts['unknown'] = 0  # fake unknown values, so they are selected into single dataset
    plate_count = plate_counts[plate_counts > 1]

    df_multiple = df[df['License Plate Number'].isin(plate_count.index)]
    df_single = df[df['License Plate Number'].isin(plate_count.index)]

    # create cross tab from multiple data
    df_multiple_first = df_multiple.sort_values('Capture Time').groupby('License Plate Number').first().reset_index()
    df_multiple_last = df_multiple.sort_values('Capture Time').groupby('License Plate Number').last().reset_index()

    df_multiple_firstlast = df_multiple_first.join(df_multiple_last, lsuffix='_first', rsuffix='_last')

    crosstab = pd.crosstab(df_multiple_firstlast['Place combined index_first'],
                           df_multiple_firstlast['Place combined index_last'])

    # create half cross tab from single data
    halfcrosstab = df_single.groupby('Place combined index').count()
    halfcrosstab = halfcrosstab.drop(columns=['Capture Time', 'Place sheet name'])
    halfcrosstab_to = halfcrosstab[halfcrosstab.index % 2 == 1].T
    halfcrosstab_from = halfcrosstab[halfcrosstab.index % 2 == 0].T

    return {'multiple': crosstab, 'singleTo': halfcrosstab_to, 'singleFrom': halfcrosstab_from}


def writeData(xlsxWriter, sheet_name, data, params):
    # Position the dataframes in the worksheet.

    N = params['number_of_cameras']

    data['multiple'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=4, startcol=2, header=False, index=False)
    data['singleTo'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=4 + N * 2 + 3, startcol=2, header=False,
                              index=False)
    data['singleFrom'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=4 + N * 2 + 7, startcol=2, header=False,
                                index=False)

    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    border_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
    })

    cond = {'type': 'cell', 'criteria': 'greater than', 'value': -1, 'format': border_format}

    # this should not be conditional formatting, but i could not found way hwo to do it properly...
    worksheet.conditional_format(4, 2, N * 2 + 3, N * 2 + 1, cond)
    worksheet.conditional_format(4 + N * 2 + 3, 2, 4 + N * 2 + 3, N + 1, cond)
    worksheet.conditional_format(4 + N * 2 + 7, 2, 4 + N * 2 + 7, N + 1, cond)


def writeTemplate(xlsxWriter, sheet_name, params):
    N = params['number_of_cameras']

    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    # column width and formats
    worksheet.set_column(0, N * 2 + 4, 9)

    # formats
    simple_format = workbook.add_format({
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
    })

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

    # write first row
    worksheet.merge_range(0, 0, 0, N * 2 + 1, "Základní výstupní tabulky (počty všech vozidel)", first_line_format)

    # write third row and first column
    worksheet.write(2, 0, "Sčítací bod", orange_format)
    worksheet.write(2, 1, "", orange_format)
    for i in range(N):
        worksheet.merge_range(2, i * 2 + 2, 2, i * 2 + 3, i + 1, orange_format)
        worksheet.merge_range(i * 2 + 4, 0, i * 2 + 5, 0, i + 1, orange_format)

    # write fourth row and second column
    worksheet.write(3, 0, "", orange_format)
    worksheet.write(3, 1, "Směr", blue_format)
    for i in range(N * 2):
        worksheet.write(3, i + 2, i + 1, blue_format)
        worksheet.write(i + 4, 1, i + 1, blue_format)

    # write single cars template
    worksheet.merge_range(N * 2 + 5, 0, N * 2 + 5, N + 1, "Pouze DO města ", first_line_format)
    worksheet.merge_range(N * 2 + 9, 0, N * 2 + 9, N + 1, "Pouze Z města ", first_line_format)

    worksheet.merge_range(N * 2 + 6, 0, N * 2 + 6, 1, "Sčítací bod", orange_format)
    worksheet.merge_range(N * 2 + 10, 0, N * 2 + 10, 1, "Sčítací bod", orange_format)

    worksheet.merge_range(N * 2 + 7, 0, N * 2 + 7, 1, "DO města", orange_format)
    worksheet.merge_range(N * 2 + 11, 0, N * 2 + 11, 1, "Z města", orange_format)

    for i in range(N):
        worksheet.write(N * 2 + 6, i + 2, f'{i+1} ({i*2+1})', orange_format)
        worksheet.write(N * 2 + 10, i + 2, f'{i+1} ({i*2+2})', orange_format)
