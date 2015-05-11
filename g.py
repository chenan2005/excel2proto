import os
import sys
from export_generator import *
from exporter_helper import *

from google.protobuf.internal import enum_type_wrapper
from google.protobuf import descriptor as _descriptor
from google.protobuf import message as _message
from google.protobuf import reflection as _reflection
from google.protobuf import symbol_database as _symbol_database
from google.protobuf import descriptor_pb2

import xlrd

config_data = {"package_name" : "iod.excel_generate", "path_proto_out" : "./output", "path_cpp_out" : "./output", "path_py_out" : "./output", "path_data_out:" : "./output"}

def check_dir(dirname):
        if os.path.isfile(dirname):
                print dirname + " is a file!"
                os.makedirs(dirname)  #raise exception by illegal calling
                
        try:
                os.makedirs(dirname)
                print dirname + " create"
        except:
                #print dirname + " exists"
                pass
                

def parconfig(filename):
        wb = xlrd.open_workbook(filename)
        try:
                ws = wb.sheet_by_name("#Config")
                for i in range(1, ws.nrows):
                        config_data[ws.row_values(i)[0]] = ws.row_values(i)[1]

        except:
                print "#Config sheet not found, use default config"

        check_dir(config_data["path_proto_out"])
        check_dir(config_data["path_cpp_out"])
        check_dir(config_data["path_py_out"])
        check_dir(config_data["path_data_out"])

        print "Config:", config_data

def generate_py_protocol(filename):
        cmdline=".\protoc.exe --python_out=" + config_data["path_py_out"] + " -I" + config_data["path_proto_out"] + " " + JoinPath(config_data["path_proto_out"], GetName(filename) + ".proto")
        print cmdline
	os.system(cmdline)

def generate_cpp_protocol(filename):
        cmdline=".\protoc.exe --cpp_out=" + config_data["path_cpp_out"] + " -I" + config_data["path_proto_out"] + " " + JoinPath(config_data["path_proto_out"], GetName(filename) + ".proto")
        print cmdline
	os.system(cmdline)

def genexport(filename):
	
        parconfig(filename)
        generate_proto(filename, config_data["package_name"], config_data["path_proto_out"])
        generate_py_protocol(filename)
        generate_cpp_protocol(filename)
        generate_export(filename, config_data["path_py_out"], config_data["path_data_out"])

        sys.path.append(config_data["path_py_out"])
        __import__(GetName(filename) +"_export")

def print_usage(name) :
        print "-----------------------------------"
        print "usage:"
        print name, "[excel_file|dir]"
        print "-----------------------------------"

print_usage(sys.argv[0])

if len(sys.argv) == 1:

        from Tkinter import Tk
        root = Tk()
        root.withdraw()

        import tkFileDialog
        filename = tkFileDialog.askopenfilename(parent=None, title="select excel file", initialdir = ".\\", filetypes=[("excel file", "*.xls;*.xlsx;*.xlsm")])

        print_usage(sys.argv[0])
        
        genexport(filename)

elif (os.path.isdir(sys.argv[1])) :

        """for dirpath, dirnames, filenames in os.walk(sys.argv[1]) :
                for f in filenames :
                        print os.path.join(dirpath, f)"""
        for f in os.listdir(sys.argv[1]):
                filename = os.path.join(sys.argv[1], f)
                if (os.path.isfile(filename)) and (filename.endswith(".xlsm") or filename.endswith(".xlsx") or filename.endswith(".xls")) :                        
                        genexport(filename)

elif (os.path.isfile(sys.argv[1])) :

        genexport(sys.argv[1])
        
#print filename

#genexport(filename)


raw_input("press enter to continue...")
