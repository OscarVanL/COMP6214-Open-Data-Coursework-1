from utils import ModelUtils
from model import ModelSampleSize, ModelResponseRates, ModelTradingStatus, ModelGovtSchemes1, ModelGovtSchemes2, \
    ModelGovtSchemes3
import pandas as pd


class ModelCovidImpacts:

    def __init__(self, filename):
        self.filename = filename

    def create_rdf_model(self):
        with open('CovidImpactsSchema.ttl', 'w', encoding='utf-8') as schema:
            self._create_schema(schema)

        with open('CovidImpactsData.ttl', 'w', encoding='utf-8') as data:
            self._create_data_header(data)
            self._write_time_range(data)
            self._write_industry_types(data)
            self._write_govt_initiatives(data)
            self._write_country_types(data)
            self._write_trading_status_types(data)
            self._write_workforce_size_types(data)
            ModelSampleSize.model_data(self.filename, data)
            ModelResponseRates.model_data(self.filename, data)
            ModelTradingStatus.model_data(self.filename, data)
            ModelGovtSchemes1.model_data(self.filename, data)
            ModelGovtSchemes2.model_data(self.filename, data)
            ModelGovtSchemes3.model_data(self.filename, data)

    def _create_schema(self, file):
        file.write('@prefix : <http://oscarvl.synology.me/schema/covid-impacts/> .\n')
        file.write('@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> .\n')
        file.write('@prefix dc: <http://purl.org/dc/elements/1.1/> .\n')
        file.write('@prefix owl: <http://www.w3.org/2002/07/owl#> .\n')
        file.write('@prefix xsd: <http://www.w3.org/2001/XMLSchema#> .\n')
        file.write('@prefix rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .\n')
        file.write('@prefix qb: <http://purl.org/linked-data/cube#> .\n\n')

        file.write('# Time represented within the data\n')
        file.write(':TimeRange rdf:type owl:Class ;\n')
        file.write('	rdfs:subClassOf qb:DimensionProperty .\n\n')

        file.write('# What industry the data element represents\n')
        file.write(':Industry rdfs:subClassOf owl:Class ;\n')
        file.write('	 rdfs:subClassOf qb:DimensionProperty .\n\n')

        file.write('# Government initiative applied for\n')
        file.write(':GovernmentInitiative rdf:type owl:Class ;\n')
        file.write('	 rdfs:subClassOf qb:DimensionProperty .\n\n')

        file.write('# Country represented in the data\n')
        file.write(':Country rdf:type owl:Class ;\n')
        file.write('	 rdfs:subClassOf qb:DimensionProperty .\n\n')

        file.write('# Surveyed company trading status\n')
        file.write(':TradingStatus rdf:type owl:Class ;\n')
        file.write('	 rdfs:subClassOf qb:DimensionProperty .\n\n')

        file.write('# Surveyed company workforce size\n')
        file.write(':WorkforceSize rdf:type owl:Class ;\n')
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

    def _write_time_range(self, file):
        # Create dataset time range description
        file.write(':TP2020 rdf:type :TimeRange;\n')
        file.write('	 dc:title "Survey reference period: 6 April 2020 to 19 April 2020" .\n\n')

    def _write_industry_types(self, file):
        # Create industry category types
        xlsx = pd.read_excel(io=self.filename, sheet_name="Sample Size")
        for industry in xlsx.iloc[:,0][3:19]:
            file.write(':' + ModelUtils.clean_label(industry) + ' rdf:type covid-impacts:Industry;\n')
            file.write('	 dc:title "' + industry + '".\n\n')

    def _write_govt_initiatives(self, file):
        # Create government COVID-19 relief initiative types
        xlsx = pd.read_excel(io=self.filename, sheet_name="Government Schemes")
        for scheme in xlsx.iloc[2][1:8]:
            file.write(':' + ModelUtils.clean_label(scheme) + ' rdf:type covid-impacts:GovernmentInitiative;\n')
            file.write('	 dc:title "' + scheme + '".\n\n')

        xlsx = pd.read_excel(io=self.filename, sheet_name="Government Schemes (2)")
        file.write(':' + ModelUtils.clean_label(xlsx.iloc[2][7]) + ' rdf:type covid-impacts:GovernmentInitiative;\n')
        file.write('	 dc:title "' + xlsx.iloc[2][7] + '".\n\n')

        xlsx = pd.read_excel(io=self.filename, sheet_name="Government Schemes (3)")
        file.write(':' + ModelUtils.clean_label(xlsx.iloc[2][7]) + ' rdf:type covid-impacts:GovernmentInitiative;\n')
        file.write('	 dc:title "' + xlsx.iloc[2][7] + '".\n\n')

    def _write_country_types(self, file):
        # Create Country for each represented country
        xlsx = pd.read_excel(io=self.filename, sheet_name="Trading Status ", header=None)
        for cell in xlsx.iloc[45:50, 0]:
            file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type covid-impacts:Country;\n')
            file.write('	 dc:title "' + cell + '".\n\n')

    def _write_trading_status_types(self, file):
        # Create TradingStatus for surveyed companies
        xlsx = pd.read_excel(io=self.filename, sheet_name="Trading Status ", header=None)
        for cell in xlsx.iloc[3][1:4]:
            file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type covid-impacts:TradingStatus;\n')
            file.write('	 dc:title "' + cell + '".\n\n')

    def _write_workforce_size_types(self, file):
        xlsx = ModelResponseRates._clean_sheet(self.filename)
        # Create WorkforceSize for each workforce size heading
        for cell in xlsx.iloc[3][1:4]:
            file.write(':' + ModelUtils.clean_label(cell) + ' rdf:type covid-impacts:WorkforceSize;\n')
            file.write('	 dc:title "' + cell + '".\n\n')



