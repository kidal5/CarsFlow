import yaml
import os
import os.path as osp
import pandas as pd
from pathlib import Path

from load_utils import loadExcel
from sheet_times import createSheetTimes
from sheet_number_of_travels import createSheetNumberOfTravels
from sheet_number_of_cars import createSheetNumberOfCars

import argparse


def parseInput():
    parser = argparse.ArgumentParser(
        description='Helper script, that reads excel data, processes them in pandas and writes them back into excel.')
    parser.add_argument('-p', '--paramsFile', type=str,
                        help='Optional path for parameters file. Defaulted to ./parameters.yaml',
                        default='parameters.yaml', required=False)

    arguments = parser.parse_args()
    arguments = vars(arguments)
    return arguments


if __name__ == '__main__':
    args = parseInput()
    with open(args['paramsFile']) as f:
        params = yaml.safe_load(f)

    print('Loading data')
    df = loadExcel(params['input_file'], params['number_of_cameras'])

    print('Delete previous output file, if present.')
    os.makedirs(Path(params['output_file']).parent, exist_ok=True)
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
