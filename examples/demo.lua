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

-- Runs an external program (Windows) and returns its first output line.
function hostname()
  local p = io.popen("hostname")
  local s = p:read("*l")
  p:close()
  return s
end
