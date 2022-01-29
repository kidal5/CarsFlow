import numpy as np
import xlsxwriter as xw
import xlsxwriter.utility as xwu
import pandas as pd

from TimeStruct import TimeStruct


def createSheetNumberOfTravels(df, xlsxWriter, params):
    def createSheetNumberOfTravelsInner(sheet_name_in, time_in: TimeStruct):
        data = computeData(df, time_in, N)
        writeData(xlsxWriter, sheet_name_in, data, N)
        writeTemplate(xlsxWriter, sheet_name_in, data, N, time_in)

    N = params['number_of_cameras']
    sheet_name = 'Počty průjezdů'
    createSheetNumberOfTravelsInner(sheet_name, TimeStruct.createFromStartAndEndTime('00:00', '23:59', df))

    for item in params['sheet_number_of_travels'].values():
        time = TimeStruct.createFromDict(item, df)
        createSheetNumberOfTravelsInner(time.findUnusedSheetName(sheet_name, xlsxWriter), time)


def computeData(df, time, N):
    # filter time
    df = df.set_index('Capture_time').sort_index()
    df = df[time.dateTimeStart: time.dateTimeEnd]
    df = df.reset_index()

    # compute number of items on every input sheet, for data validation
    dataCheck = getDataCheckColumn(df, N)

    uniqueSPZ = df['License_plate'].nunique()

    dick = {
        'Direction': [],
        'SPZ': [],
        'Count': []
    }

    # add fake column to be removed later, order of operations should be kept because SPZs ordering matter
    for spz in range(1, 11):  # always show at least 10 SPZs
        dick['Direction'].append('remove me')
        dick['SPZ'].append(spz)
        dick['Count'].append(0)

    for direction in df['Direction'].unique():
        temp = df[df['Direction'] == direction]

        # solve unknown spz
        temp_unknown = temp[temp['License_plate'] == 'unknown']
        dick['Direction'].append(direction)
        dick['SPZ'].append(9999)
        dick['Count'].append(len(temp_unknown))

        # solve correct spz
        temp = temp[temp['License_plate'] != 'unknown']
        temp = temp.groupby(by='License_plate').count().groupby(by='Capture_time').count()
        temp = temp.reset_index()

        # column names lost their original meaning...
        for spz, count in zip(temp['Capture_time'], temp['Direction']):
            dick['Direction'].append(direction)
            dick['SPZ'].append(spz)
            dick['Count'].append(count)

    # add data even when there were none loaded
    missing = np.setdiff1d(list(range(1, N*2 + 1)), df['Direction'].unique())
    for direction in missing:
        dick['Direction'].append(direction)
        dick['SPZ'].append(1)
        dick['Count'].append(0)

    out = pd.DataFrame.from_dict(dick)
    SPZs = out['SPZ'].unique()
    SPZs.sort()
    SPZs = SPZs.tolist()
    SPZs[-1] = 'unknown'

    out = out.pivot(index='SPZ', columns='Direction', values='Count').fillna(0)
    out = out.drop(columns=['remove me'])

    return {'data': out, 'SPZs': SPZs, 'dataCheck': dataCheck, 'uniqueSPZ': uniqueSPZ}


def getDataCheckColumn(df, N):
    dataCheck = df.groupby(by='Direction').count().drop(columns=['License_plate'])

    fake_df = {'Direction': [], 'Capture_time': []}
    for i in range(1, N * 2 + 1):
        fake_df['Capture_time'].append(0)
        fake_df['Direction'].append(i)

    fake_df = pd.DataFrame.from_dict(fake_df)
    fake_df = fake_df.set_index('Direction')
    fake_df.update(dataCheck)

    return fake_df


def writeData(xlsxWriter, sheet_name, data, N):
    # Position the dataframes in the worksheet.
    data['data'].to_excel(xlsxWriter, sheet_name=sheet_name, startrow=5, startcol=1, header=False, index=False)
    data['dataCheck'].T.to_excel(xlsxWriter, sheet_name=sheet_name, startrow=len(data["SPZs"]) + 8, startcol=1,
                                 header=False, index=False)

    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    yellow_format = workbook.add_format({
        'border': 1,
        'font_color': 'red',
        'bg_color': '#FFF2CC',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'text_wrap': True,
    })

    xd_format = workbook.add_format({
        'bold': 1,
        'border': 1,
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

    cond_good = {'type': 'cell', 'criteria': '<>', 'value': 0, 'format': wrong_format}
    cond_wrong = {'type': 'cell', 'criteria': '=', 'value': 0, 'format': good_format}

    # this should not be conditional formatting, but I could not find way how to do it properly...

    worksheet.conditional_format(5, 1, len(data['SPZs']) + 4, N * 2, {'type': 'cell',
                                                                      'criteria': 'greater than',
                                                                      'value': -1,
                                                                      'format': yellow_format})

    worksheet.conditional_format(len(data["SPZs"]) + 8, 1, len(data['SPZs']) + 8, N * 2, {'type': 'cell',
                                                                                          'criteria': 'greater than',
                                                                                          'value': -1,
                                                                                          'format': xd_format})

    worksheet.conditional_format(len(data["SPZs"]) + 9, 1, len(data['SPZs']) + 9, N * 2, cond_good)
    worksheet.conditional_format(len(data["SPZs"]) + 9, 1, len(data['SPZs']) + 9, N * 2, cond_wrong)


def writeTemplate(xlsxWriter, sheet_name, data, N, time: TimeStruct):
    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    # formats
    yellow_format = workbook.add_format({
        'border': 1,
        'font_color': 'red',
        'bg_color': '#FFF2CC',
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'text_wrap': True,
    })

    first_line_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'font_size': 12
    })

    second_line_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'text_wrap': True,
    })

    percent_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'text_wrap': True,
        'num_format': '0.00%'
    })

    # column width and formats
    worksheet.set_column(1, N * 2 + 4, 13)
    worksheet.set_column(0, 0, 16)

    # write first and second row
    worksheet.merge_range(0, 0, 0, N * 2 + 1, "Počty průjezdů jednotlivých stanovišť", first_line_format)
    worksheet.merge_range(1, 0, 1, N * 2 + 1, f"Vybraný čas: {time.fullSheetName}", first_line_format)

    # write first line
    worksheet.write(3, 0, "Počet průjezdů", first_line_format)
    for i in range(1, N + 1):
        text = f'Sčít. stanoviště {i}'
        worksheet.merge_range(3, i * 2 - 1, 3, i * 2, text, first_line_format)

    # write second line
    worksheet.write(4, 0, "s totožnou SPZ", second_line_format)
    for i in range(1, N + 1):
        text_do = f'Do města ({i * 2 - 1})'
        text_z = f'Z města ({i * 2})'
        worksheet.write(4, i * 2 - 1, text_do, second_line_format)
        worksheet.write(4, i * 2, text_z, second_line_format)

    # write first/second line
    worksheet.merge_range(3, N * 2 + 1, 4, N * 2 + 1, 'Celkem [počet]', second_line_format)
    worksheet.merge_range(3, N * 2 + 2, 4, N * 2 + 2, 'Poměr', second_line_format)
    worksheet.merge_range(3, N * 2 + 3, 4, N * 2 + 3, 'Celkový počet unikátních SPZ ', second_line_format)
    worksheet.write(5, N * 2 + 3, data['uniqueSPZ'], yellow_format)

    # write data lines
    for i, spz in enumerate(data['SPZs']):
        worksheet.write(5 + i, 0, spz, second_line_format)

        a = xwu.xl_col_to_name(1)
        b = xwu.xl_col_to_name(N * 2)
        c = xwu.xl_col_to_name(N * 2 + 1)

        # solve case when spz is unknown
        if spz == 'unknown':
            worksheet.write(5 + i, N * 2 + 1, f'=SUM({a}{6 + i}:{b}{6 + i}) * 1', second_line_format)
            worksheet.write(5 + i, N * 2 + 2, f'={c}{6 + i}/${c}${len(data["SPZs"]) + 7}', percent_format)
        else:
            worksheet.write(5 + i, N * 2 + 1, f'=SUM({a}{6 + i}:{b}{6 + i}) * A{6 + i}', second_line_format)
            worksheet.write(5 + i, N * 2 + 2, f'={c}{6 + i}/${c}${len(data["SPZs"]) + 7}', percent_format)

    # write last line, not anymore, xd
    worksheet.write(len(data['SPZs']) + 6, 0, "Celkem [počet]", second_line_format)
    for i in range(N * 2):
        a = xwu.xl_col_to_name(i + 1)
        worksheet.write(len(data['SPZs']) + 6, i + 1,
                        f'=SUMPRODUCT({a}{6}:{a}{len(data["SPZs"]) + 4},A{6}:A{len(data["SPZs"]) + 4}) + {a}{len(data["SPZs"]) + 5}',
                        second_line_format)

    for i in range(N * 2, N * 2 + 2):
        a = xwu.xl_col_to_name(i + 1)

        form = second_line_format
        if i == N * 2 + 1:
            form = percent_format

        worksheet.write(len(data["SPZs"]) + 6, i + 1, f'=SUM({a}{6}:{a}{len(data["SPZs"]) + 5})', form)

    worksheet.write(len(data["SPZs"]) + 8, 0, "Cílový součet [počet]", second_line_format)
    worksheet.write(len(data["SPZs"]) + 9, 0, "Validace dat", second_line_format)

    for i in range(N * 2):
        a = xwu.xl_col_to_name(i + 1)

        worksheet.write(len(data["SPZs"]) + 9, i + 1, f'={a}{len(data["SPZs"]) + 7} - {a}{len(data["SPZs"]) + 9}',
                        second_line_format)
