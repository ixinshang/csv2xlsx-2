#! /usr/bin/env python
from optparse import OptionParser
from csv2xlsx import convert
import traceback

parser = OptionParser("./main.py -s t.csv")
parser.add_option("-f", "--folder", action="store", dest="folder", type="string",
                  default="", help="源文件目录名，多个文件时用于指定目录")
parser.add_option("-s", "--source", action="store", dest="source", type="string",
                  default="", help="源文件名，多个文件用英文逗号分隔，如:1.csv,2.csv")
parser.add_option("-m", "--merge", action="store", dest="merge", type="int",
                  default=0, help="合并方式，默认不合并。1合并数据到一个工作表，2合并到同一个文件的多个工作表中")
parser.add_option("-d", "--destination", action="store", dest="dest", type="string", 
                  default="", help="生成的文件名")
options, arg = parser.parse_args()

try:
    c = convert(options.source, folder=options.folder, merge=options.merge, dest=options.dest)
    print(c.read().write())
except:
    print(traceback.format_exc())
