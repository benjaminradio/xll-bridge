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
