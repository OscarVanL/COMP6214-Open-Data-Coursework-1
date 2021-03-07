from utils import ModelUtils
import pandas as pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_trading_status(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Trading Status ", header=None)

    # Replace asterisks with actual data
    for index, rows in xlsx.iloc[4:17, 0:4].iterrows():
        missing_val = 1.0 - rows[1] - rows[3]
        xlsx.loc[index, 2] = abs(missing_val)

    for index, rows in xlsx.iloc[30:33, 0:4].iterrows():
        missing_val = 1.0 - rows[1] - rows[3]
        xlsx.loc[index, 2] = abs(missing_val)

    for index, rows in xlsx.iloc[45:51, 0:4].iterrows():
        missing_val = 1.0 - rows[1] - rows[3]
        xlsx.loc[index, 2] = abs(missing_val)

    return xlsx


def _write_trading_status(xlsx, file):
    pd.set_option("display.max_colwidth", None)

    file.write('# Trading Status Questionnaire Dataset #1\n')
    # Create DataSet definition
    file.write(':ts1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[1][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[21:26, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create SurveyedMetric for trading status columns
    for cell in xlsx.iloc[3][1:4]:
        file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type :SurveyedMetric;\n')
        file.write('	 dc:title "' + cell + '".\n\n')

    # Create industry response rate data
    data = xlsx.iloc[4:17, 0:4]
    data.columns = xlsx.iloc[3][0:4]
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':ts1_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :ts1;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')

    # ===============

    file.write('# Trading Status Questionnaire Dataset #2\n')

    # Create DataSet definition
    file.write(':ts2 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[27][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[37:41, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create SurveyedMetric for trading status columns
    for cell in xlsx.iloc[30:33, 0]:
        file.write(':WorkforceSize' + ModelUtils.clean_label(cell) + ' rdf:type :SurveyedMetric;\n')
        file.write('	 dc:title "' + cell + '".\n\n')

    # Create industry response rate data
    data = xlsx.iloc[30:33, 0:4]
    data.columns = xlsx.iloc[29][0:4]
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':ts2_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :ts2;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :WorkforceSize' + ModelUtils.clean_label(rows[0]) + '.\n\n')

    # ===============

    file.write('# Trading Status Questionnaire Dataset #3\n')

    # Create DataSet definition
    file.write(':ts3 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[27][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[54:59, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create Country for each represented country
    for cell in xlsx.iloc[45:50, 0]:
        print(cell)
        file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type :Country;\n')
        file.write('	 dc:title "' + cell + '".\n\n')

    data = xlsx.iloc[45:50, 0:4]
    data.columns = xlsx.iloc[44][0:4]
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':ts3_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :ts3;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')