from itertools import chain
import pandas as pd

load_columns = ['License Plate Number', 'Country', 'Capture Time']
load_columns = ['License Plate Number', 'Capture Time']


def loadExcel(fname, N=10):
    xlsx = pd.ExcelFile(fname)
    sheet_names = getSheetNames(N)
    sheet_names_place_index = getSheetNamesPlaceIndexes(N)
    sheet_names_place_direction = getSheetNamesPlaceDirections(N)
    sheet_names_place_combined_index = getSheetNamesPlaceCombinedIndexes(N)

    df = None

    for name, index, direction, combined_index in zip(sheet_names, sheet_names_place_index, sheet_names_place_direction,
                                                      sheet_names_place_combined_index):
        df_read = pd.read_excel(xlsx, sheet_name=name, usecols=load_columns)
        df_read['Place sheet name'] = name
        # df_read['Place index'] = index
        # df_read['Place direction'] = direction
        df_read['Place combined index'] = combined_index

        if df is None:
            df = df_read
        else:
            df = df.append(df_read)

    df = df.reset_index(drop=True)

    return df


def getSheetNamesPlaceIndexes(N=10):
    return list(chain.from_iterable((x, x) for x in range(1, N + 1)))


def getSheetNamesPlaceDirections(N=10):
    return ['to', 'from'] * N


def getSheetNamesPlaceDirectionsCZE(N=10):
    return ['do', 'z'] * N


def getSheetNamesPlaceCombinedIndexes(N=10):
    return list(range(1, N * 2 + 1))


def getSheetNames(N=10):
    ind = getSheetNamesPlaceIndexes(N)
    directions = getSheetNamesPlaceDirectionsCZE(N)
    comb_ind = getSheetNamesPlaceCombinedIndexes(N)

    return [f'{i}_{d}({ci})' for (i, d, ci) in zip(ind, directions, comb_ind)]