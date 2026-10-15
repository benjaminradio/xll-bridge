-- Demo scripts for xll-bridge. Embedded into demo.xlsx by examples/make_demo_workbook.py.

function add(a, b)
  return a + b
end

-- Receives a range as a table of rows; returns the transposed range.
function transpose(r)
  local out = {}
  for i = 1, #r[1] do
    out[i] = {}
    for j = 1, #r do out[i][j] = r[j][i] end
  end
  return out
end

-- Returns an Excel error for bad input.
function safe_div(a, b)
  if b == 0 then return xl.DIV0 end
  return a / b
end

-- Runs an external program (no shell, no console window) and returns its first output line.
function hostname()
  local r = xl.run{ "hostname", timeout = 10 }
  if not r then return xl.NA end
  return (r.stdout:gsub("%s+$", ""))
end

-- Slow on purpose (about 3 seconds). Call it as =LUA.ASYNC("slow_hostname") and Excel stays
-- responsive while the cell shows #GETTING_DATA; =LUA.CALL("slow_hostname") would freeze Excel.
function slow_hostname()
  local r = xl.shell("ping -n 4 127.0.0.1 > nul & hostname", { timeout = 30 })
  if not r then return xl.NA end
  return (r.stdout:gsub("%s+$", ""))
end

-- Uses a module stored in the lib/ folder of the embedded file tree.
local mathx = require("lib.mathx")
function hypot(a, b)
  return mathx.hypot(a, b)
end

-- Reading cells. Unqualified addresses mean the active sheet of this workbook; a range passed
-- as an argument (=LUA.CALL("total", A1:A10)) is better because Excel then tracks the dependency.
function total_of(address)
  local sum = 0
  for _, row in ipairs(xl.range(address):values()) do
    for _, v in ipairs(row) do
      if type(v) == "number" then sum = sum + v end
    end
  end
  return sum
end

-- Writing cells is only allowed from background calls and macros, never from LUA.CALL (Excel
-- forbids a worksheet function from changing the sheet). =LUA.ASYNC("stamp", "Sheet1!E1")
function stamp(address)
  xl.range(address):set({ { "written by Lua", os.date("%Y-%m-%d %H:%M:%S") } })
  return "ok"
end

-- A macro: shows up in the ribbon's Lua > Macros menu. It runs in the background and may change
-- cells (unlike LUA.CALL). This one writes the selection's values into a "Copy" sheet if there
-- is one, at the same address.
xl.macro("Copy selection to 'Copy' sheet", function()
  local sel = xl.selection()
  local cells = sel:address():match("!(.*)$")
  xl.range("Copy!" .. cells):set(sel:values())
end)

-- A live stream: =LUA.STREAM("clock") shows the time and keeps updating. A stream function pushes
-- values with xl.emit(); xl.sleep() waits without blocking Excel. It stops when no cell uses it.
function clock()
  while true do
    xl.emit(os.date("%H:%M:%S"))
    xl.sleep(1)
  end
end
