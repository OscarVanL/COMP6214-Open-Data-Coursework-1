from utils import ModelUtils
import pandas as pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_trading_status(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Government Schemes", header=None)

    return xlsx


def _write_trading_status(xlsx, file):
    pd.set_option("display.max_colwidth", None)

    file.write('# Government Schemes Dataset #1\n')
    # Create DataSet definition
    file.write(':gs1_1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[1][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[21:24, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create SurveyedMetric for trading status columns
    for cell in xlsx.iloc[3][1:8]:
        file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type :SurveyedMetric;\n')
        file.write('	 dc:title "' + cell + '".\n\n')

    # Create data
    data = xlsx.iloc[4:17, 0:8]
    data.columns = xlsx.iloc[3][0:8]
    for index, rows in data.iterrows():
        for col in range(0, 7):
            file.write(':gs1_1_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :gs1_1;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')

    # ==========

    file.write('# Government Schemes Dataset #2\n')
    # Create DataSet definition
    file.write(':gs1_2 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[25][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[35:37, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create data
    data = xlsx.iloc[28:31, 0:8]
    data.columns = xlsx.iloc[27][0:8]
    for index, rows in data.iterrows():
        for col in range(0, 7):
            file.write(':gs1_2_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :gs1_2;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :WorkforceSize' + ModelUtils.clean_label(rows[0]) + '.\n\n')


    # ==========

    file.write('# Government Schemes Dataset #3\n')
    # Create DataSet definition
    file.write(':gs1_3 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[38][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[50:53, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create data
    data = xlsx.iloc[41:46, 0:8]
    data.columns = xlsx.iloc[40][0:8]
    for index, rows in data.iterrows():
        for col in range(0, 7):
            file.write(':gs1_3_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :gs1_3;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')