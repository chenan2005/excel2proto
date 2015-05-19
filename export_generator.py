import sys
import io
import xlrd
from exporter_helper import *

reload(sys)
sys.setdefaultencoding( "utf-8" )

def has_sheet_name(wb, sheet_name):
        try:
                print "has sheet name ", sheet_name
                wb.sheet_by_name(sheet_name)
                return True
        except:
                print "has no sheet name ", sheet_name
                return False

def generate_export(excel_filename, path_py_out, path_data_out) :
	dataname = GetName(excel_filename)
	filename = JoinPath(path_py_out, dataname + "_export.py")
	pyfile = open(filename, "w+")
	
	pyfile.write("\"\"\"auto generate by generate_exporter.py	\"\"\"\r\n")

	pyfile.write("""
import sys
import os
import xlrd
from exporter_helper import *
""")

	pyfile.write("import " + dataname + "_pb2\r\n")
	
	wb = xlrd.open_workbook(excel_filename)
	for tb in wb.sheets() :
		if tb.name.find("#P_") == 0 :
			hasdatasheet = True
			try:
				datasheet = wb.sheet_by_name(tb.name[3:])
			except:
				hasdatasheet = False
				
			print "generate converter of " + tb.name[3:]
			pyfile.write("\r\n")
			pyfile.write("def Set" + tb.name[3:] + "Data(msg, field_data_array) :\r\n")
			pyfile.write("    print 'Set" + tb.name[3:] + "Data', field_data_array\r\n")
			data_idx = 0
			for i in range(2,tb.nrows):
				
				field_repeat_count = 1
				if hasdatasheet:
					field_repeat_count = GetElemCount(datasheet.col_values(0), tb.row_values(i)[0])
				elif len(str(tb.row_values(i)[4])) > 0 :
					field_repeat_count = int(tb.row_values(i)[4])
					
				print "	field info:" , tb.row_values(i), "repeat count:", field_repeat_count
				
				set_values = []
				if tb.row_values(i)[1] == 'int32' \
					or tb.row_values(i)[1] == 'int64' \
					or tb.row_values(i)[1] == 'sint32' \
					or tb.row_values(i)[1] == 'sint64' \
					or tb.row_values(i)[1] == 'uint32' \
					or tb.row_values(i)[1] == 'uint64' \
					or tb.row_values(i)[1] == 'fixed32' \
					or tb.row_values(i)[1] == 'fixed64' \
					or tb.row_values(i)[1] == 'sfixed32' \
					or tb.row_values(i)[1] == 'sfixed64' \
					or has_sheet_name(wb, "#E_" + tb.row_values(i)[1]) :
					if tb.row_values(i)[2] != "repeated" :
						str_set_value = "msg." + tb.row_values(i)[0] + " = " + "int(float(field_data_array[" + str(data_idx) + "]))"
						set_values.append(str_set_value)
						data_idx+=1
					else :
                                                if field_repeat_count == 1:
                                                    str_set_value = "for i in SplitEx(field_data_array[" + str(data_idx) + "], ',') :\r\n"
                                                    str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "int(float(i))" + ")"
                                                    set_values.append(str_set_value)
                                                    data_idx+=1
                                                else:
                                                    for j in range(field_repeat_count) :
                                                        str_set_value = "if len(field_data_array[" + str(data_idx) + "]) > 0 :\r\n"
                                                        str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "int(float(field_data_array[" + str(data_idx) + "]))" + ")"
                                                        set_values.append(str_set_value)
                                                        data_idx+=1
				elif tb.row_values(i)[1] == 'float' \
					or tb.row_values(i)[1] == 'double' :
					if tb.row_values(i)[2] != "repeated" :
						str_set_value = "msg." + tb.row_values(i)[0] + " = " + "float(field_data_array[" + str(data_idx) + "])"
						set_values.append(str_set_value)
						data_idx+=1
					else :
                                                if field_repeat_count == 1:
                                                    str_set_value = "for i in SplitEx(field_data_array[" + str(data_idx) + "], ',') :\r\n"
                                                    str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "float(i)" + ")"
                                                    set_values.append(str_set_value)
                                                    data_idx+=1
                                                else:
                                                    for j in range(field_repeat_count) :
                                                        str_set_value = "if len(field_data_array[" + str(data_idx) + "]) > 0 :\r\n"
                                                        str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "float(field_data_array[" + str(data_idx) + "])" + ")"
                                                        set_values.append(str_set_value)
                                                        data_idx+=1
				elif tb.row_values(i)[1] == 'bool' :
					if tb.row_values(i)[2] != "repeated" :
						str_set_value = "msg." + tb.row_values(i)[0] + " = " + "bool(int(float(field_data_array[" + str(data_idx) + "])))"
						set_values.append(str_set_value)
						data_idx+=1
					else :
                                                if field_repeat_count == 1:
                                                    str_set_value = "for i in SplitEx(field_data_array[" + str(data_idx) + "], ',') :\r\n"
                                                    str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "bool(int(float(i)))" + ")"
                                                    set_values.append(str_set_value)
                                                    data_idx+=1
                                                else:
                                                    for j in range(field_repeat_count) :
                                                        str_set_value = "if len(field_data_array[" + str(data_idx) + "]) > 0 :\r\n"
                                                        str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "bool(int(float(field_data_array[" + str(data_idx) + "])))" + ")"
                                                        set_values.append(str_set_value)
                                                        data_idx+=1
				elif tb.row_values(i)[1] == 'string' \
					or tb.row_values(i)[1] == 'bytes' :
					if tb.row_values(i)[2] != "repeated" :
						str_set_value = "msg." + tb.row_values(i)[0] + " = " + "field_data_array[" + str(data_idx) + "]"
						set_values.append(str_set_value)
						data_idx+=1
					else :
                                                if field_repeat_count == 1:
                                                    str_set_value = "for i in SplitEx(field_data_array[" + str(data_idx) + "], ',') :\r\n"
                                                    str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "i" + ")"
                                                    set_values.append(str_set_value)
                                                    data_idx+=1
                                                else:
                                                    for j in range(field_repeat_count) :
                                                        str_set_value = "if len(field_data_array[" + str(data_idx) + "]) > 0 :\r\n"
                                                        str_set_value = str_set_value + "        msg." + tb.row_values(i)[0] + ".append(" + "field_data_array[" + str(data_idx) + "]" + ")"
                                                        set_values.append(str_set_value)
                                                        data_idx+=1
				elif has_sheet_name(wb, "#P_" + tb.row_values(i)[1]):
					if tb.row_values(i)[2] != "repeated" :
						str_set_value = "Set" + tb.row_values(i)[1] + "Data(msg." + tb.row_values(i)[0] + ", SplitEx(field_data_array[" + str(data_idx) + "], ';'))"
						set_values.append(str_set_value)
						data_idx+=1
					else:
						print field_repeat_count
						for j in range(field_repeat_count) :
                                                    str_set_value = "if len(field_data_array[" + str(data_idx) + "]) > 0 :\r\n"
                                                    str_set_value = str_set_value + "        Set" + tb.row_values(i)[1] + "Data(msg." + tb.row_values(i)[0] + ".add(), SplitEx(field_data_array[" + str(data_idx) + "], ';'))"
                                                    set_values.append(str_set_value)
                                                    data_idx+=1				
				for str_set_value in set_values:
					pyfile.write("    " + str_set_value + "\r\n")
	
	pyfile.write("\r\n")
	
	pyfile.write("wb = xlrd.open_workbook('" + excel_filename + "')\r\n")
	for tb in wb.sheets() :
		if tb.name[0] != "#" and tb.name[0] != "!":
			dataname = tb.name
			pyfile.write("ws = wb.sheet_by_name('" + dataname + "')\r\n")
			pyfile.write("data_container = " + GetName(excel_filename) + "_pb2." + dataname + "_Container()\r\n")
			pyfile.write("for i in range(1,ws.ncols):\r\n")			
			pyfile.write("    Set" + dataname + "Data(data_container.list_of_" + dataname + ".add(), [str(j) for j in ws.col_values(i)])\r\n")
			#pyfile.write("    for j in ws.col_values(i):\r\n")
			#pyfile.write("        print j, ', ', str(j), type(j), len(str(j))\r\n")
			dbfilename = JoinPath(path_data_out, dataname + ".db")
			pyfile.write("f = open('" + dbfilename + "', 'wb+')\r\n")
			pyfile.write("f.write(data_container.SerializeToString())\r\n")
			pyfile.write("f.close()\r\n")
			pyfile.write("print 'file " + dbfilename + " saved'\r\n\r\n")
			pyfile.write("print 'try readback from file " + dbfilename + "'\r\n")
			pyfile.write("f = open('" + dbfilename + "', 'rb')\r\n")
			pyfile.write("data_container = " + GetName(excel_filename) + "_pb2." + dataname + "_Container()\r\n")
			pyfile.write("data_container.ParseFromString(f.read())\r\n")
			pyfile.write("f.close()\r\n")
			pyfile.write("f = open('" + dbfilename + ".txt', 'w+')\r\n")			
			pyfile.write("print >> f, data_container\r\n")
			pyfile.write("print 'done! you can open file " + dbfilename + ".txt to check the read result from .db file'\r\n")
	pyfile.close()
	
	
def generate_proto(excel_filename, pkg_name, output_path) :
	filename = JoinPath(output_path, GetName(excel_filename) + ".proto")
	pyfile = open(filename, "w+")
	pyfile.write("option optimize_for = CODE_SIZE;\r\n\r\n")
	pyfile.write("package " + pkg_name + ";\r\n\r\n")
	wb = xlrd.open_workbook(excel_filename)
	for tb in wb.sheets() :
		if tb.name.find("#P_") == 0 :			
			pyfile.write("///" + tb.row_values(0)[0] + "\r\n")
			pyfile.write("message " + tb.name[3:] + "{\r\n")
			for i in range(2, tb.nrows) :
				rowdata = tb.row_values(i)
				pyfile.write("	" + rowdata[2] + " " + rowdata[1] + " " + rowdata[0] + " = " + str(i - 1) + "; //" + rowdata[3] + "\r\n")
			pyfile.write("}\r\n")			
			pyfile.write("///container of " + tb.name[3:] + "\r\n")
			pyfile.write("message " + tb.name[3:] + "_Container {\r\n")
			pyfile.write("	repeated " + tb.name[3:] + " list_of_" + tb.name[3:] + " = 1; \r\n")
			pyfile.write("}\r\n\r\n")
		elif tb.name.find("#E_") == 0 :	
			#print tb.row_values(0)[0]
			pyfile.write("///" + tb.row_values(0)[0] + "\r\n")
			pyfile.write("enum " + tb.name[3:] + "{\r\n")
			for i in range(2, tb.nrows) :
				rowdata = tb.row_values(i)
				#print rowdata
				pyfile.write("	" + rowdata[0] + " = " + str(int(rowdata[1])) + "; //" + rowdata[2] + "\r\n")
			pyfile.write("}\r\n\r\n")	
	pyfile.close()
