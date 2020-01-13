# -*- coding: UTF-8 -*- 
# test in Python2.7.11
# this is used to parse Microsoft Excel worksheet which save as XML Spreadsheet 2003 (.xml) format
import xml.dom.minidom

filename = "1.xml".decode('utf-8')

dom = xml.dom.minidom.parse(filename)
root = dom.documentElement
for Worksheet in root.getElementsByTagName("Worksheet"):
    for Table in Worksheet.getElementsByTagName("Table"):
        for RowIndex, Row in enumerate(Table.getElementsByTagName("Row")):
            print "RowIndex: %s" % RowIndex
            for CellIndex, Cell in enumerate(Row.getElementsByTagName("Cell")):
                for Data in Cell.getElementsByTagName("Data"):
                    for childNodes in Data.childNodes:
                        print "CellIndex: %s" % CellIndex, "->", childNodes.data