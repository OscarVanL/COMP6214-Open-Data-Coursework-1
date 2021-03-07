from utils import ModelUtils
import pandas as pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_sample_response_rates(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Response Rates")
    # Todo: Dataset cleaning
    return xlsx


def _write_sample_response_rates(xlsx, file):
    pd.set_option("display.max_colwidth", None)

    file.write('# Response Rate Dataset #1\n')
    # Create DataSet definition
    file.write(':rr1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[1][1] + '";\n')
    file.write('	 rdfs:label "' + xlsx.columns[0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(list(map(str.strip, xlsx.iloc[18:21, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create industry response rate data
    data = xlsx.iloc[3:16, 0:4]
    data.columns = xlsx.iloc[2][0:4]
    print(data)
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':rr1_{}_{} '.format(index, col+1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value ' + str(rows[col+1]) + ';\n')
            file.write('	 qb:dataSet :rr1;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col+1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')

    # ==============
    file.write('# Response Rate Dataset #2\n')
    print(xlsx.iloc[1][7])
    # Create DataSet definition
    file.write(':rr2 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[1][7] + '";\n')
    file.write('	 rdfs:label "' + xlsx.columns[6] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(list(map(str.strip, xlsx.iloc[18:21, 0].to_string(index=False).split('\n')))) + '".\n\n')

    data = xlsx.iloc[3:16, 6:10]
    data.columns = xlsx.iloc[2][6:10]
    print(data)
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':rr2_{}_{} '.format(index, col+7) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value ' + str(rows[col+1]) + ';\n')
            file.write('	 qb:dataSet :rr2;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col+1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')