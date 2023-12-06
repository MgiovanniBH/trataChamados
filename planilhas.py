import pandas as pd
import xlrd

class leXls:
    def __init__(self, varq):
        self.arquivo = varq

    def open(self):
        try:
            data = pd.read_excel(self.arquivo, usecols='B:L')
            return data
        except:
            return None

