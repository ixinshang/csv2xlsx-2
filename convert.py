import os
from writer.xlsx import XlsxWriter

class Convert:

    def __init__(self, source, **kwargs):
        self.source = []
        self.data = []
        self.dest = ""
        self.merge = 0
        self.debug = True
        self._init_param(source, **kwargs)

    def __str__(self):
        msg = ""
        for row in self.data:
            msg += row['source'] + " -> " + row['dest'] + "\t" + (
                "OK" if row['read'] and row['write'] else "Failed") + "\r\n"
        return msg

    def _init_param(self, source, **kwargs):
        sources = []
        if "dest" in kwargs and kwargs["dest"] != "":
            self.dest = kwargs["dest"]
        if "merge" in kwargs and kwargs['merge'] > 0:
            self.merge = kwargs["merge"]
        if "folder" in kwargs:
            folder = kwargs['folder']
        else:
            folder = ""
        folder = folder.rstrip("/")
        default_source_file = folder + "/" + source.lstrip("/") if folder != "" else source
        if os.path.exists(default_source_file):
            sources.append(default_source_file)
        elif "," in source:
            if folder != "":
                sources = [folder + "/" + r.lstrip("/") for r in source.split(",") if os.path.exists(folder + "/" + r.lstrip("/"))]
            else:
                sources = [r for r in source.split(",") if os.path.exists(r)]
        if len(sources) == 0:
            raise Exception("Unable to load the specified file: %s" % default_source_file)

        self.source = sources
        if self.debug:
            print("---------------------------------------------")
            print("Source File List:")
            for row in self.source:
                print(row)
            print("---------------------------------------------")

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
            else:
                continue
            self.data.append(t)
        
        if self.debug:
            print("Read File Info:")
            for row in self.data:
                print(row['source'], row['read'], len(row['data']))
            print("---------------------------------------------")
        return self

    def write(self, **kwargs):
        if self.debug:
            print("Result:")
        self._write_to_excel()
        return self

    def _write_to_excel(self):
        if self.merge == 1:
            self._write_to_excel_1()
        elif self.merge == 2:
            self._write_to_excel_2()
        else:
            self._write_to_excel_0()
        return self

    def _write_to_excel_0(self):
        cnt = len(self.data)
        for row in self.data:
            if not row['read']:
                continue
            if cnt == 1 and self.dest != "":
                xw = XlsxWriter(self.dest)
            else:
                xw = XlsxWriter(row['source'])
            xw.write(row['data'])
            row['write'] = True
            row['dest'] = xw.filename

    def _write_to_excel_1(self):
        filename = ''
        data = []
        for row in self.data:
            if not row['read']:
                continue
            if filename == "":
                filename = row['source'] + ".m"
            data.extend(row['data'])
        if self.dest != "":
            filename = self.dest
        xw = XlsxWriter(filename)
        xw.write(data)
        for row in self.data:
            if not row['read']:
                continue
            row['write'] = True
            row['dest'] = xw.filename

    def _write_to_excel_2(self):
        filename = ''
        data = []
        for row in self.data:
            if not row['read']:
                continue
            if filename == "":
                filename = row['source'] + ".s"
            data.append({
                "title": os.path.basename(row['source']).lower(),
                "data": row['data']
            })
        if self.dest != "":
            filename = self.dest
        xw = XlsxWriter(filename)  
        xw.write_multi(data)
        for row in self.data:
            if not row['read']:
                continue
            row['write'] = True
            row['dest'] = xw.filename
