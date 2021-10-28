import yaml
import os
import os.path as osp
import pandas as pd

from load_utils import loadExcel
from sheet_times import createSheetTimes
from sheet_number_of_travels import createSheetNumberOfTravels
from sheet_number_of_cars import createSheetNumberOfCars

if __name__ == '__main__':
    with open('parameters.yaml') as f:
        params = yaml.safe_load(f)

    print('Loading data')
    df = loadExcel(params['input_file'], params['number_of_cameras'])

    if osp.isfile(params['output_file']):
        os.remove(params['output_file'])

    writer = pd.ExcelWriter(params['output_file'], engine='xlsxwriter')

    print('Creating sheet Počty vozidel')
    createSheetNumberOfCars(df.copy(True), writer, params)

    print('Creating sheet Počty průjezdů')
    createSheetNumberOfTravels(df.copy(True), writer, params)

    print('Creating sheet Časové údaje')
    createSheetTimes(df.copy(True), writer, params)
    writer.close()


