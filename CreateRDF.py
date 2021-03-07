from model.ModelData import ModelCovidImpacts

if __name__ == '__main__':
    model = ModelCovidImpacts('CW1-BusinessImpactsOfCovid19Data.xlsx')
    model.create_rdf_model()