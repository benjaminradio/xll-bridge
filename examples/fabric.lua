-- fabric.lua: a local, shared cache for DAX queries against a Microsoft Fabric / Power BI
-- semantic model, pulled with DAX Studio's command-line tool (dscmd).
--
-- Put this file in the machine folder (%APPDATA%\xll-bridge\local) and use it from any workbook:
--
--     local fabric = require("fabric")
--
--     -- =LUA.ASYNC("gl_summary", "2024-06"): runs in the background, shows #GETTING_DATA
--     -- while dscmd works, and is cached by Excel until the arguments (hence the DAX) change.
--     function gl_summary(month)
--       return fabric.table(([[
--         EVALUATE SUMMARIZECOLUMNS('BU'[Business Unit], 'Calendar'[Month],
--           FILTER(ALL('Calendar'[Month]), 'Calendar'[Month] = DATE(%s)),
--           "Amount", [Total Amount])]]):format(month:gsub("-", ",")))
--     end
--
--     -- A macro that loads a big detailed pull into the machine database "fabric" once a day,
--     -- so every report can query it with SQL.QUERY ("... from [db:fabric].gl ...").
--     xl.macro("Refresh Fabric cache", function()
--       fabric.load("fabric", "gl", [[EVALUATE ...detailed query...]],
--                   { dates = { "Month" }, text = { "Account" }, index = { "BU", "Account" } })
--     end)
--
-- Every distinct DAX text is pulled once and kept as a CSV file in the machine folder, shared by
-- every workbook, until it is older than the daily refresh time. Fabric is only hit when the file
-- is missing or stale (or when you pass { force = true }).

local M = {}

M.config = {
  -- Path to DAX Studio's command-line tool.
  dscmd = [[C:\Program Files\DAX Studio\dscmd.exe]],
  server = "powerbi://api.powerbi.com/v1.0/myorg/YOUR WORKSPACE",
  database = "YOUR DATASET",
  -- The model refreshes once a day: a pull older than the most recent HH:MM (local time) is stale.
  refresh_after = "06:30",
  -- Give dscmd time for the sign-in popup and a long query (seconds).
  timeout = 900,

  -- ADAPT THIS to your dscmd syntax. It must return the command as a list of words that runs the
  -- DAX in `dax_file` against the model and writes the result as CSV to `out_file`. The words
  -- below are placeholders, not dscmd's real flags.
  command = function(c, dax_file, out_file)
    return { c.dscmd, "--server", c.server, "--database", c.database,
             "--query-file", dax_file, "--output", "csv", "--output-file", out_file }
  end,
}

local function sep() return package.config:sub(1, 1) end

local function cache_dir()
  local dir = xl.local_dir() .. sep() .. "fabric-cache"
  assert(xl.mkdir(dir))
  return dir
end

-- Seconds since 1970 of the latest HH:MM that has already passed today or yesterday.
local function last_refresh_time()
  local h, m = M.config.refresh_after:match("^(%d+):(%d+)$")
  local now = os.time()
  local d = os.date("*t", now)
  local t = os.time { year = d.year, month = d.month, day = d.day, hour = tonumber(h), min = tonumber(m) }
  if t > now then t = t - 24 * 3600 end
  return t
end

function M.stale(file)
  local st = xl.stat(file)
  return not st or st.mtime < last_refresh_time()
end

--- Path of a CSV holding the result of `dax`; pulls it only when there is no fresh copy.
function M.csv(dax, opts)
  local c = M.config
  local base = cache_dir() .. sep() .. xl.hash(c.server .. "\n" .. c.database .. "\n" .. dax):sub(1, 24)
  local csv = base .. ".csv"
  if not (opts and opts.force) and not M.stale(csv) then return csv end

  local dax_file, part = base .. ".dax", base .. ".csv.part"
  local f = assert(io.open(dax_file, "wb"))
  f:write(dax)
  f:close()
  os.remove(part)

  -- The sign-in popup (if any) is dscmd's own; this waits for it.
  local r, err = xl.run(c.command(c, dax_file, part))
  if not r then error("cannot start dscmd: " .. tostring(err), 0) end
  if r.timed_out or not r.ok then
    error(("dscmd failed (exit %s): %s"):format(tostring(r.code), (r.stderr ~= "" and r.stderr or r.stdout):sub(1, 400)), 0)
  end
  if not xl.stat(part) then error("dscmd produced no output file", 0) end
  os.remove(csv)
  assert(os.rename(part, csv))
  return csv
end

--- The result as a table of rows (column names first), ready to return from LUA.ASYNC.
function M.table(dax, opts)
  local o = { dates = opts and opts.dates, text = opts and opts.text }
  return xl.read_csv(M.csv(dax, opts), o)
end

--- Loads the result into the table `name` of the machine database `dbname` (created if needed),
--- replacing the table atomically. Returns { rows = n, columns = {...} }.
function M.load(dbname, name, dax, opts)
  local db = xl.db.open(dbname, { create = "local" })
  return db:import_csv(M.csv(dax, opts), name, opts)
end

return M
