import xlsxwriter as xw
import xlsxwriter.utility as xwu
import pandas as pd


def createSheetNumberOfTravels(df, xlsxWriter, params):
    N = params['number_of_cameras']
    sheet_name = 'Počty průjezdů'

    data, SPZs = computeData(df)
    writeData(xlsxWriter, sheet_name, data, SPZs, N)
    writeTemplate(xlsxWriter, sheet_name, data, SPZs, N)


def computeData(df):
    dick = {
        'Direction': [],
        'SPZ': [],
        'Count': []
    }

    # add fake column to be removed later, order of operations should be kept because SPZs ordering matter
    for spz in range(1, 11):  # always show atleast 10 SPZs
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

        # column names lost their oringal meaning...
        for spz, count in zip(temp['Capture_time'], temp['Direction']):
            dick['Direction'].append(direction)
            dick['SPZ'].append(spz)
            dick['Count'].append(count)

    out = pd.DataFrame.from_dict(dick)
    SPZs = out['SPZ'].unique()
    SPZs.sort()
    SPZs = SPZs.tolist()
    SPZs[-1] = 'unknown'

    out = out.pivot(index='SPZ', columns='Direction', values='Count').fillna(0)
    out = out.drop(columns=['remove me'])

    return out, SPZs


def writeData(xlsxWriter, sheet_name, data, SPZs, N):
    # Position the dataframes in the worksheet.
    data.to_excel(xlsxWriter, sheet_name=sheet_name, startrow=2, startcol=1, header=False, index=False)

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

    # this should not be conditional formatting, but i could not found way how to do it properly...

    worksheet.conditional_format(2, 1, len(SPZs) + 1, N * 2, {'type': 'cell',
                                                              'criteria': 'greater than',
                                                              'value': -1,
                                                              'format': yellow_format})


def writeTemplate(xlsxWriter, sheet_name, data, SPZs, N):
    workbook = xlsxWriter.book
    worksheet = xlsxWriter.sheets[sheet_name]

    # column width and formats
    worksheet.set_column(1, N * 2 + 4, 13)
    worksheet.set_column(0, 0, 16)

    # formats
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

    # write first line
    worksheet.write(0, 0, "Počet průjezdů", first_line_format)
    for i in range(1, N + 1):
        text = f'Sčít. stanoviště {i}'
        worksheet.merge_range(0, i * 2 - 1, 0, i * 2, text, first_line_format)

    # write second line
    worksheet.write(1, 0, "s totožnou SPZ", second_line_format)
    for i in range(1, N + 1):
        text_do = f'Do města ({i * 2 - 1})'
        text_z = f'Z města ({i * 2})'
        worksheet.write(1, i * 2 - 1, text_do, second_line_format)
        worksheet.write(1, i * 2, text_z, second_line_format)

    # write first/second line
    worksheet.merge_range(0, N * 2 + 1, 1, N * 2 + 1, 'Celkem [počet]', second_line_format)
    worksheet.merge_range(0, N * 2 + 2, 1, N * 2 + 2, 'Poměr', second_line_format)
    worksheet.merge_range(0, N * 2 + 3, 1, N * 2 + 3, 'Počet SPZ', second_line_format)

    # write data lines
    for i, spz in enumerate(SPZs):
        worksheet.write(2 + i, 0, spz, second_line_format)

        a = xwu.xl_col_to_name(1)
        b = xwu.xl_col_to_name(N * 2)
        c = xwu.xl_col_to_name(N * 2 + 1)

        # solve case when spz is unknown
        if spz == 'unknown':
            worksheet.write(2 + i, N * 2 + 1, f'=SUM({a}{3 + i}:{b}{3 + i}) * 1', second_line_format)
            worksheet.write(2 + i, N * 2 + 2, f'={c}{3 + i}/${c}${len(SPZs) + 4}', percent_format)
            worksheet.write(2 + i, N * 2 + 3, f'={c}{3 + i}', second_line_format)
        else:
            worksheet.write(2 + i, N * 2 + 1, f'=SUM({a}{3 + i}:{b}{3 + i}) * A{3 + i}', second_line_format)
            worksheet.write(2 + i, N * 2 + 2, f'={c}{3 + i}/${c}${len(SPZs) + 4}', percent_format)
            worksheet.write(2 + i, N * 2 + 3, f'={c}{3 + i} / A{3 + i}', second_line_format)

    # write last line
    worksheet.write(len(SPZs) + 3, 0, "Celkem [počet]", second_line_format)
    for i in range(N * 2):
        a = xwu.xl_col_to_name(i + 1)
        worksheet.write(len(SPZs) + 3, i + 1, f'=SUMPRODUCT({a}{3}:{a}{len(SPZs) + 1},A{3}:A{len(SPZs) + 1}) + {a}{len(SPZs) + 2}', second_line_format)

    for i in range(N * 2, N * 2 + 3):
        a = xwu.xl_col_to_name(i + 1)

        form = second_line_format
        if i == N * 2 + 1:
            form = percent_format

        worksheet.write(len(SPZs) + 3, i + 1, f'=SUM({a}{3}:{a}{len(SPZs) + 2})', form)
