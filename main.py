#! /usr/bin/env python
from optparse import OptionParser
from convert import Convert

parser = OptionParser("./main.py -s t.csv")
parser.add_option("-f", "--folder", action="store", dest="folder", type="string",
                  default="", help="源文件目录名，多个文件时用于指定目录")
parser.add_option("-s", "--source", action="store", dest="source", type="string",
                  default="", help="源文件名，多个文件用英文逗号分隔，如:1.csv,2.csv")
parser.add_option("-m", "--merge", action="store", dest="merge", type="int",
                  default=0, help="合并方式，默认不合并。1合并数据到一个工作表，2合并到同一个文件的多个工作表中")
options, arg = parser.parse_args()

try:
    c = Convert(options.source, options.folder)
    print(c.read().write(merge=options.merge))
except Exception as e:
    print(e)
