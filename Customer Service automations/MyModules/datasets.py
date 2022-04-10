from MyModules import utils


# manage data storage and data types here

class ExcelDataset:
    def get_path(self, filename):
        raise NotImplementedError("Define your dataset get_path")

    def read_data(self, filename):
        return utils.write_read_data().read_excel(self.get_path(filename))
    
    def write_data(self, filename, data):
        return utils.write_read_data().write_excel(self.get_path(filename), data)

    def read_data_pickle(self, filename):
        return utils.write_read_data().read_pickle(self.get_path(filename))

    def write_data_pickle(self, filename, data):
        return utils.write_read_data().write_pickle(self.get_path(filename), data)


class Dataset_20200521(ExcelDataset):
    def __init__(self, root=None):
        if root is None:
            from mlmailclassify.trash import DataSource
            root = DataSource.path
        self.root = root
        self.datatypes = {
            'cleanlabelcut': 'ShipExprLabledCut.xlsx',
            'livedata': 'LiveRawData.xlsx',
            'MyTeam' : 'ExportTeam.xlsx',
            'Plant_with_country' : 'PlantCountry.xlsx',
            'initmails' : 'df.pickle',
            'cleanedmails': 'df2.pickle',
            'ML_model': 'ML_model',
            'Vectorizer': 'Vectorizer',

        }

    def get_path(self, datatype):
        return self.root + self.datatypes[datatype]
