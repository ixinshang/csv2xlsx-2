#! /usr/bin/env python
import os
from optparse import OptionParser
from convert import Convert

parser = OptionParser("./main.py -s t.csv")
parser.add_option("-f", "--folder", action="store", dest="folder", type="string", default="", help="source folder")
parser.add_option("-s", "--source", action="store", dest="source", type="string", default="", help="source filename. eg:1.csv,2.csv")
parser.add_option("-t", "--type", action="store", dest="type", type="string", default="xlsx", help="output filetype，default:xlsx, options:[csv、xlsx、txt]")
parser.add_option("-m", "--merge", action="store_true", dest="merge", default=False, help="merge all source file")
options, arg = parser.parse_args()

try:
    c = Convert(options.source, options.folder)
    print(c.read().write(options.type, options.merge))
except Exception as e:
    print(e)
