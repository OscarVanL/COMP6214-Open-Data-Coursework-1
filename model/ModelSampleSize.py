from utils import ModelUtils
import pandas as pd
import math


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_sample_size(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Sample Size", header=None)
    print(xlsx.loc[18].keys())

    print(xlsx.loc[19, 2])
    xlsx.loc[19, 2] = math.ceil(xlsx.loc[19, 2])
    print(xlsx.loc[19][2])
    return xlsx


def _write_sample_size(xlsx, file):
    file.write('# Sample Size Dataset #1\n')

    # Create DataSet definition
    file.write(':ss1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[2][1] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:comment "' + xlsx.iloc[22][0] + '".\n\n')

    # Create SurveyedMetric for each workforce size heading
    for cell in xlsx.iloc[3][1:4]:
        file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type :SurveyedMetric;\n')
        file.write('	 dc:title "' + cell + '".\n\n')

    # Create industry sample size data
    data = xlsx.iloc[4:20, 0:4]
    data.columns = xlsx.iloc[3][0:4]
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':ss1_{}_{} '.format(index, col+1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value ' + str(rows[col+1]) + ';\n')
            file.write('	 qb:dataSet :ss1;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col+1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')

    # ==============

    file.write('# Sample Size Dataset #2\n')

    # Create DataSet definition
    file.write(':ss2 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[24, 0] + '".\n\n')

    # Create SurveyedMetric for each workforce size heading
    for cell in xlsx.iloc[26][1:5]:
        file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type :SurveyedMetric;\n')
        file.write('	 dc:title "' + cell + '".\n\n')

    # Create sample workforce size data
    for val in range(1, 5):
        file.write(':ss2_{}_{} '.format(27, val) + 'rdf:type qb:Observation;\n')
        file.write('	 rdf:value ' + str(xlsx.iloc[27][val]) + ';\n')
        file.write('	 qb:dataSet :ss2;\n')
        file.write('	 qb:dimension :' + ModelUtils.clean_label(xlsx.iloc[26][val]) + ';\n')
        file.write('	 qb:dimension :' + ModelUtils.clean_label(xlsx.iloc[27][0]) + '.\n\n')