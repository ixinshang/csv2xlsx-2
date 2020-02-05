import csv
from reader import IReader

class CsvReader(IReader):

    def load(self):
        data = []
        with open(self.filename, 'r') as fp:
            cr = csv.reader(fp)
            for row in cr:
                data.append(row)
        return data
