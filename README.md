functional:

	1, read message definition info from excel workbook and export <workbook_name>.proto
	
	2, read data from excel workbook and export <workbook_name>.db

usage:

	g.py [workbook_name | dir]

need: 

	1, python 2.7
	
	2, protoc.exe to generate python protocol <workbook_name>_pb2.py
	
	3, protobuf python lib (the same version with protoc.exe)
	