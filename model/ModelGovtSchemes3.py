from utils import ModelUtils
import pandas as pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_trading_status(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Government Schemes (3)", header=None)

    # Make workforce size label consistent with other sheets.
    xlsx.loc[29, 0] = 'Workforce Size < 250'
    xlsx.loc[30, 0] = 'Workforce Size 250 +'

    # Replace asterisk with 0%
    xlsx.loc[13, 5] = 0

    # Remove signs from values
    xlsx.loc[6, 7] = 0.422
    xlsx.loc[16, 1] = 0.368

    return xlsx


def _write_trading_status(xlsx, file):
    pd.set_option("display.max_colwidth", None)

    file.write('# Government Schemes (3) Dataset #1\n')
    # Create DataSet definition
    file.write(':gs3_1 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[1][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[21:25, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create data
    data = xlsx.iloc[4:17, 0:8]
    data.columns = xlsx.iloc[3][0:8]
    for index, rows in data.iterrows():
        for col in range(0, 7):
            file.write(':gs3_1_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :gs3_1;\n')
            file.write('	 qb:dimension :TP2020;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')


    file.write('# Government Schemes (3) Dataset #2\n')
    # Create DataSet definition
    file.write(':gs3_2 rdf:type qb:DataSet;\n')
    file.write('	 dc:title "' + xlsx.iloc[0][0] + '";\n')
    file.write('	 rdfs:label "' + xlsx.iloc[26][0] + '";\n')
    file.write('	 rdfs:comment "' + '; '.join(
        list(map(str.strip, xlsx.iloc[36:39, 0].to_string(index=False).split('\n')))) + '".\n\n')

    # Create data
    data = xlsx.iloc[29:32, 0:8]
    data.columns = xlsx.iloc[28][0:8]
    for index, rows in data.iterrows():
        for col in range(0, 7):
            file.write(':gs3_2_{}_{} '.format(index, col + 1) + 'rdf:type qb:Observation;\n')
            file.write('	 rdf:value %.3f' % rows[col + 1] + ';\n')
            file.write('	 qb:dataSet :gs3_2;\n')
            file.write('	 qb:dimension :TP2020;\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows.index[col + 1]) + ';\n')
            file.write('	 qb:dimension :' + ModelUtils.clean_label(rows[0]) + '.\n\n')


