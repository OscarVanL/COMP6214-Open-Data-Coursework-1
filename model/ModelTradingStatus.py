from utils import ModelUtils
import pandas as pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_trading_status(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Trading Status ", header=None)
    # Todo: Dataset cleaning
    return xlsx


def _write_trading_status(xlsx, file):
    pd.set_option("display.max_colwidth", None)

    file.write('# Trading Status Questionnaire Dataset #1\n')
    # Create DataSet definition
    file.write(':ts1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[1][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[20:25, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create industry response rate data
    data = xlsx.iloc[4:17, 0:4]
    data.columns = xlsx.iloc[3][0:4]
    for index, rows in data.iterrows():
        for col in range(0, 3):
            file.write(':ts1_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value ' + str(rows[col + 1]) + ';\n')
            file.write('	 qb:dataSet :ts1;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')