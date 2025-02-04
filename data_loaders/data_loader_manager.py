
from data_loaders.pdf_loader import PdfLoader
from data_loaders.csv_loader import CsvLoader
from data_loaders.docx_loader import DocxLoader

class DataLoaderManager:
    def __init__(self):
        self.loaders = {
            #'docx': DocxLoader(),
            'pdf': PdfLoader()
        }

    def load_data(self, data_path, file_type):
        loader = self.loaders.get(file_type.lower())
        if loader:
            return loader.load_data(data_path)
        else:
            raise ValueError(f"No data loader available for file type: {file_type}")
        
    def get_supported_data_loaders(self):
        return list(self.loaders.keys())