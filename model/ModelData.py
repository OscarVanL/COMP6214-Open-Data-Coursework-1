from utils import ModelUtils
from model import ModelSampleSize, ModelResponseRates
import pandas as pd


class ModelCovidImpacts:

    def __init__(self, filename):
        self.filename = filename
        # self.xlsx = load_workbook(filename=filename)
        # print('Opening dataset:', filename)
        # print(self.xlsx.get_sheet_names())
        # assert self.xlsx['Read me']['A1'].value == 'Business impacts of COVID-19 data'

    def create_rdf_model(self):
        with open('CovidImpactsSchema.ttl', 'w', encoding='utf-8') as schema:
            self._create_schema(schema)

        with open('CovidImpactsData.ttl', 'w', encoding='utf-8') as data:
            self._create_data_header(data)
            self._write_industry_types(data)
            ModelSampleSize.model_data(self.filename, data)
            ModelResponseRates.model_data(self.filename, data)

    def _create_schema(self, file):
        file.write('@prefix : <http://oscarvl.synology.me/schema/covid-impacts/> .\n')
        file.write('@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> .\n')
        file.write('@prefix dc: <http://purl.org/dc/elements/1.1/> .\n')
        file.write('@prefix owl: <http://www.w3.org/2002/07/owl#> .\n')
        file.write('@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .\n')
        file.write('@prefix rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .\n')
        file.write('@prefix qb: <http://purl.org/linked-data/cube#> .\n\n')

        file.write('# What industry the data element represents\n')
        file.write(':Industry rdfs:subClassOf owl:Class ;\n')
        file.write('	 rdfs:subClassOf qb:DimensionProperty .\n\n')

        file.write('# Surveyed metric represented by the data\n')
        file.write(':SurveyedMetric rdf:type owl:Class ;\n')
        file.write('	 rdfs:subClassOf qb:DimensionProperty .\n\n')

    def _create_data_header(self, file):
        file.write('@prefix : <http://oscarvl.synology.me/data/> .\n')
        file.write('@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> .\n')
        file.write('@prefix dc: <http://purl.org/dc/elements/1.1/> .\n')
        file.write('@prefix owl: <http://www.w3.org/2002/07/owl#> .\n')
        file.write('@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .\n')
        file.write('@prefix rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .\n')
        file.write('@prefix qb: <http://purl.org/linked-data/cube#> .\n')
        file.write('@prefix covid-impacts: <http://oscarvl.synology.me/schema/covid-impacts/> .\n\n')

    def _write_industry_types(self, file):
        xlsx = pd.read_excel(io=self.filename, sheet_name="Sample Size")
        for industry in xlsx.iloc[:,0][3:19]:
            file.write(':' + ModelUtils.clean_label(industry) + ' rdf:type covid-impacts:Industry;\n')
            file.write('	 dc:title "' + industry + '".\n\n')


