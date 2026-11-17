-- admonitions.lua
-- Turns pandoc fenced Divs into dAIgnostics callout boxes (LaTeX output only).
--
-- Usage in markdown:
--   ::: note
--   Body text...
--   :::
--
--   ::: {.warning title="Heads up"}
--   Body text...
--   :::
--
-- Supported classes: note, tip, info, warning, important
-- Optional attribute `title="..."` overrides the box heading.

local map = {
  note      = { env = "dgnote",     default = "Note" },
  tip       = { env = "dgtip",      default = "Tip" },
  info      = { env = "dgtip",      default = "Info" },
  warning   = { env = "dgwarning",  default = "Warning" },
  important  = { env = "dgwarning", default = "Important" },
}

-- escape user-supplied title for LaTeX
local function tex_escape(s)
  return (s:gsub("[\\%%%${}&#_^~]", {
    ["\\"]="\\textbackslash{}", ["%"]="\\%%", ["$"]="\\$", ["{"]="\\{",
    ["}"]="\\}", ["&"]="\\&", ["#"]="\\#", ["_"]="\\_",
    ["^"]="\\textasciicircum{}", ["~"]="\\textasciitilde{}",
  }))
end

function Div(el)
  if not FORMAT:match("latex") then return nil end
  for _, cls in ipairs(el.classes) do
    local spec = map[cls]
    if spec then
      local title = el.attributes["title"] or spec.default
      local opener = pandoc.RawBlock("latex",
        "\\begin{" .. spec.env .. "}[" .. tex_escape(title) .. "]")
      local closer = pandoc.RawBlock("latex", "\\end{" .. spec.env .. "}")
      local out = { opener }
      for _, b in ipairs(el.content) do table.insert(out, b) end
      table.insert(out, closer)
      return out
    end
  end
  return nil
end
