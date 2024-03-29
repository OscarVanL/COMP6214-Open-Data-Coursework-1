from utils import ModelUtils
import pandas as pd
import numpy as np


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_sample_size(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Sample Size", header=None)

    # Fix all industries total error
    xlsx.loc[19, 2] = int(xlsx.loc[19, 2])

    # Fix total error
    xlsx.loc[27, 4] = xlsx.loc[27, 1] + xlsx.loc[27, 2] + xlsx.loc[27, 3]

    # Make workforce size label consistent
    xlsx.loc[3, 3] = "All Size Bands"
    xlsx.loc[26, 4] = "All Size Bands"

    # Merge 0-99 and 100-250 into single < 250 category
    xlsx.loc[26, 2] = "Workforce Size < 250"
    xlsx.loc[27, 2] = xlsx.loc[27, 1] + xlsx.loc[27, 2]
    # Remove now redundant heading
    xlsx.loc[26, 1] = np.NaN
    xlsx.loc[27, 1] = np.NaN

    return xlsx


def _write_sample_size(xlsx, file):
    file.write('# Sample Size Dataset #1\n')

    # Create DataSet definition
    file.write(':ss1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[2][1] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:comment "' + xlsx.iloc[22][0] + '".\n\n')

    # Create industry sample size data
    data = xlsx.iloc[4:20, 0:4]
    data.columns = xlsx.iloc[3][0:4]
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':ss1_{}_{} '.format(index, col+1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value ' + str(rows[col+1]) + ';\n')
            file.write('	 qb:dataSet :ss1;\n')
            file.write('	 qb:dimension :TP2020;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col+1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')

    # ==============

    file.write('# Sample Size Dataset #2\n')

    # Create DataSet definition
    file.write(':ss2 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[24, 0] + '".\n\n')

    # Create sample workforce size data
    for val in range(2, 5):
        file.write(':ss2_{}_{} '.format(27, val) + 'rdf:type qb:Observation;\n')
        file.write('	 rdf:value {}'.format(xlsx.iloc[27][val])+ ';\n')
        file.write('	 qb:dataSet :ss2;\n')
        file.write('	 qb:dimension :TP2020;\n')
        file.write('	 qb:dimension :' + ModelUtils.clean_label(xlsx.iloc[26][val]) + ';\n')
        file.write('	 qb:dimension :' + ModelUtils.clean_label(xlsx.iloc[27][0]) + '.\n\n')