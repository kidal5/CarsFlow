from itertools import chain
import pandas as pd
from pathlib import Path

load_columns = ['License Plate Number', 'Capture Time', 'Kategorie']
rename_columns = {'License Plate Number': 'License_plate', 'Capture Time': 'Capture_time', 'Kategorie': 'Vehicle_category'}


def loadExcel(fname, N):
    print('\tLoading excel file:', Path(fname).absolute())

    xlsx = pd.ExcelFile(fname)
    sheet_names = getSheetNames(N)
    sheet_names_place_index = getSheetNamesPlaceIndexes(N)
    sheet_names_place_direction = getSheetNamesPlaceDirections(N)
    sheet_names_place_combined_index = getSheetNamesPlaceCombinedIndexes(N)

    dfs = []

    for name, index, direction, combined_index in zip(sheet_names, sheet_names_place_index, sheet_names_place_direction,
                                                      sheet_names_place_combined_index):
        print(f'\tLoading sheet "{name}"')
        df_read = pd.read_excel(xlsx, sheet_name=name)
        if df_read.empty:
            print(f'\t\tWarning, skipping empty sheet "{name}".')
            continue

        if 'Kategorie' not in df_read.columns:
            print(f'\t\tWarning, sheet "{name}" does not contain column "Kategorie". Vehicle category filtering is limited.')
            df_read['Kategorie'] = 'Dummy_category'
        df_read = df_read[load_columns]

        df_read['Direction'] = combined_index
        df_read = df_read.rename(columns=rename_columns)

        # drop empty rows
        df_read = df_read[~df_read['Capture_time'].isnull()]

        if not df_read.empty:
            dfs.append(df_read)

    df = pd.concat(dfs)
    df = df.reset_index(drop=True)
    df = df.astype({'License_plate': 'str', 'Vehicle_category': 'str'})
    return df


def getSheetNamesPlaceIndexes(N):
    return list(chain.from_iterable((x, x) for x in range(1, N + 1)))


def getSheetNamesPlaceDirections(N):
    return ['to', 'from'] * N


def getSheetNamesPlaceDirectionsCZE(N):
    return ['do', 'z'] * N


def getSheetNamesPlaceCombinedIndexes(N):
    return list(range(1, N * 2 + 1))


def getSheetNames(N):
    ind = getSheetNamesPlaceIndexes(N)
    directions = getSheetNamesPlaceDirectionsCZE(N)
    comb_ind = getSheetNamesPlaceCombinedIndexes(N)

    return [f'{i}_{d}({ci})' for (i, d, ci) in zip(ind, directions, comb_ind)]
