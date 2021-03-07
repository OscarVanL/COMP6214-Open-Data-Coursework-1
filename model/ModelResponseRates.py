from utils import ModelUtils
from pandas import pd


def model_data(filename, out_file):
    xlsx = _clean_sheet(filename)
    _write_sample_response_rates(xlsx, out_file)


def _clean_sheet(filename):
    xlsx = pd.read_excel(io=filename, sheet_name="Response Rates")
    # Todo: Dataset cleaning
    return xlsx


def _write_sample_response_rates(xlsx, file):
    pass