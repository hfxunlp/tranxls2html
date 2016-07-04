#encoding: utf-8

import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )

import xlrd

wb=xlrd.open_workbook("src.xlsx")
table=wb.sheet_by_index(0)

with open("rs.html","a") as fwrt:
	lind=table.row_values(0)
	fwrt.write("<tr>\n".encode("utf-8"))
	for linu in lind:
		fwrt.write(("<th>"+linu+"</th>\n").encode("utf-8"))
	fwrt.write("</tr>\n".encode("utf-8"))
	for i in xrange(1,table.nrows):
		fwrt.write("<tr>\n".encode("utf-8"))
		lind=table.row_values(i)
		for linu in lind:
			fwrt.write(("<td>"+linu+"</td>\n").encode("utf-8"))
		fwrt.write("</tr>\n".encode("utf-8"))
	fwrt.write("</table>\n</body>\n</html>".encode("utf-8"))
