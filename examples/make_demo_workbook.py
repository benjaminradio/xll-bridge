"""Creates a plain demo workbook whose formulas call the Lua functions in demo.lua.

Usage: python make_demo_workbook.py out.xlsx
Then embed the scripts:  xllb-pack embed out.xlsx demo.xlsx demo.lua

Formulas use the `_xll.` prefix, which is how Excel stores add-in functions in files.
"""
import sys

from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Demo"
ws["A1"], ws["B1"] = 2, 3
ws["A2"] = "add(A1, B1)"
ws["B2"] = '=_xll.LUA.CALL("add", A1, B1)'
ws["A3"] = "safe_div(1, 0)"
ws["B3"] = '=_xll.LUA.CALL("safe_div", 1, 0)'
ws["A4"] = "hostname()"
ws["B4"] = '=_xll.LUA.CALL("hostname")'
ws["A5"] = "eval"
ws["B5"] = '=_xll.LUA.EVAL("... * 10", A1)'
ws["A6"] = "status"
ws["B6"] = "=_xll.LUA.STATUS()"
for r, (x, y) in enumerate([(1, 2), (3, 4), (5, 6)], start=8):
    ws.cell(row=r, column=1, value=x)
    ws.cell(row=r, column=2, value=y)
ws["D7"] = "transpose(A8:B10)  (Ctrl+Shift+Enter on pre-365 Excel)"
ws["D8"] = "=_xll.LUA.CALL(\"transpose\", A8:B10)"
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 18
wb.save(sys.argv[1])
