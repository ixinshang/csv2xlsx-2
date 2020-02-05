import os

class Convert():
    source = []
    type = "xlsx"
    merge = False

    data = []

    def __init__(self, source, folder = ""):
        sources = []
        folder = folder.rstrip("/")
        default_source_file = folder + "/" + source
        if os.path.exists(default_source_file):
            sources.append(default_source_file)
        elif "," in source:
            sources = [folder + "/" + r for r in source.split(",") if os.path.exists(folder + "/" + r)]
        if len(sources) == 0:
            raise Exception("Unable to load the specified file: %s" % default_source_file)

        self.source = sources

    def __str__(self):
        msg = ""
        for row in self.data:
            msg += row['source'] + " -> " + row['dest'] + "\t" + ("OK" if row['read'] and row['write'] else "Failed") + "\r\n"
        return msg

    def read(self):
        for f in self.source:
            t = {
                "source": f,
                "data": [],
                "suffix": os.path.splitext(f)[1].lower(),
                "read": False,
                "write": False,
                "dest": "",
            }
            if t["suffix"] == ".csv":
                from reader.csv import CsvReader
                cr = CsvReader(f)
                t["data"] = cr.load()
                if len(t["data"]) > 0:
                    t["read"] = True
            self.data.append(t)
        return self

    def write(self, type = "xlsx", merge = False):
        from writer.xlsx import XlsxWriter
        for row in self.data:
            if row["read"] == False:
                continue
            xw = XlsxWriter(row['source'])
            xw.write(row['data'])
            row['write'] = True
            row['dest'] = xw.filename
        return self