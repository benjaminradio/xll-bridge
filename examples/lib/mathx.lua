-- A module in a folder: reachable as require("lib.mathx") from the root scripts.
local M = {}

function M.hypot(a, b)
  return math.sqrt(a * a + b * b)
end

return M
