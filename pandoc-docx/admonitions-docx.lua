-- admonitions-docx.lua
-- Turns pandoc fenced Divs into dAIgnostics callout boxes in DOCX output.
-- Same markdown syntax as ../pandoc-document/admonitions.lua (LaTeX):
--
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
--
-- Word has no "box around arbitrary blocks", so each callout becomes a title
-- paragraph + body paragraphs in custom paragraph styles from reference.docx
-- (Callout Note/Tip/Warning [Title]). Those styles share one border/shading
-- definition, which Word and LibreOffice merge into a single shaded box with
-- a coloured left spine.

if not FORMAT:match("docx") then return {} end

local map = {
  note      = { style = "Callout Note",    default = "Note" },
  tip       = { style = "Callout Tip",     default = "Tip" },
  info      = { style = "Callout Tip",     default = "Info" },
  warning   = { style = "Callout Warning", default = "Warning" },
  important = { style = "Callout Warning", default = "Important" },
}

-- Word list numbering carries its own indents, which would split the box, so
-- lists inside a callout become plain paragraphs with a literal marker.
local function flatten_list(items, marker)
  local out = {}
  for i, item in ipairs(items) do
    local first = true
    for _, b in ipairs(item) do
      if first and (b.t == "Para" or b.t == "Plain") then
        local inl = pandoc.Inlines({ pandoc.Str(marker(i) .. "\u{2002}") })
        inl:extend(b.content)
        out[#out + 1] = pandoc.Para(inl)
      else
        if first then out[#out + 1] = pandoc.Para({ pandoc.Str(marker(i)) }) end
        out[#out + 1] = b
      end
      first = false
    end
  end
  return out
end

local list_filter = {
  BulletList = function(l)
    return flatten_list(l.content, function() return "\u{2022}" end)
  end,
  OrderedList = function(l)
    local start = l.listAttributes and l.listAttributes.start or 1
    return flatten_list(l.content, function(i) return tostring(start + i - 1) .. "." end)
  end,
}

function Div(el)
  for _, cls in ipairs(el.classes) do
    local spec = map[cls]
    if spec then
      local title = el.attributes["title"] or spec.default
      return {
        pandoc.Div({ pandoc.Para(pandoc.Inlines(title)) },
                   { ["custom-style"] = spec.style .. " Title" }),
        pandoc.Div(el.content:walk(list_filter), { ["custom-style"] = spec.style }),
        pandoc.RawBlock("openxml", '<w:p><w:pPr><w:pStyle w:val="CalloutEnd"/></w:pPr></w:p>'),
      }
    end
  end
  return nil
end
