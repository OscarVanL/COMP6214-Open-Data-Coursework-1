from utils import ModelUtils
import pandas as pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_trading_status(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Government Schemes (2)", header=None)

    # Make workforce size label consistent with other sheets.
    xlsx.loc[29, 0] = 'Workforce Size < 250'
    xlsx.loc[30, 0] = 'Workforce Size 250 +'
    print(xlsx.loc[:,0])

    return xlsx


def _write_trading_status(xlsx, file):
    pd.set_option("display.max_colwidth", None)
    pass
