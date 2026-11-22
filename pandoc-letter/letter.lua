-- letter.lua
-- dAIgnostics letter / memorandum: one Markdown file -> PDF and DOCX.
--
-- Normalises the letter metadata for both outputs (labels by language,
-- defaults for place, date, opening, closing, organisation and signer).
-- For LaTeX the template daignostics-letter.latex lays the letter out from
-- that metadata; for DOCX this filter builds the same blocks around the
-- body, styled by reference.docx:
--
--   recipient (first line bold) / N/R
--                                                   place, date
--   Predmet: subject
--   Poštovani,
--   body ...
--   S poštovanjem,
--   DAIGNOSTICS d.o.o. / signer / signer title
--   Potpis: ____________
--   Prilozi: ...
--
-- Company letterhead and footer (logo, address, OIB, IBAN) live in the
-- LaTeX template and in reference.docx.

local stringify = pandoc.utils.stringify

local MONTHS_EN = { "January", "February", "March", "April", "May", "June", "July",
                    "August", "September", "October", "November", "December" }

local labels = {
  hr = { subject = "Predmet", attention = "N/R", signature = "Potpis",
         enclosures = "Prilozi", opening = "Poštovani,", closing = "S poštovanjem,",
         date = function() return os.date("%d.%m.%Y.") end },
  en = { subject = "Subject", attention = "Attn.", signature = "Signature",
         enclosures = "Enclosures", opening = "Dear Sir or Madam,", closing = "Kind regards,",
         date = function()
           local t = os.date("*t")
           return string.format("%d %s %d", t.day, MONTHS_EN[t.month], t.year)
         end },
}

local DEFAULT_PLACE = "Bobovje"
local DEFAULT_ORGANISATION = "DAIGNOSTICS d.o.o."

local function inlines_of(v)
  if v == nil then return nil end
  local t = pandoc.utils.type(v)
  if t == "Inlines" then return v end
  if t == "Blocks" then return pandoc.utils.blocks_to_inlines(v) end
  return pandoc.Inlines(stringify(v))
end

local function list_of(v)
  if v == nil then return {} end
  if pandoc.utils.type(v) == "List" then
    local out = {}
    for _, x in ipairs(v) do out[#out + 1] = inlines_of(x) end
    return out
  end
  return { inlines_of(v) }
end

local function falsy(v)
  if v == nil then return false end
  if type(v) == "boolean" then return not v end
  local s = stringify(v):lower()
  return s == "false" or s == "no" or s == "0"
end

local function para(style, inl)
  return pandoc.Div({ pandoc.Para(inl) }, { ["custom-style"] = style })
end

local function tab()
  return pandoc.RawInline("openxml", "<w:r><w:tab/></w:r>")
end

local function docx_blocks(meta, body)
  local out = pandoc.Blocks({})

  -- recipient
  local rec = pandoc.Inlines({})
  local to = list_of(meta.to)
  for i, line in ipairs(to) do
    if i == 1 then rec:insert(pandoc.Strong(line)) else
      rec:insert(pandoc.LineBreak()); rec:extend(line) end
  end
  if meta.attention then
    if #rec > 0 then rec:insert(pandoc.LineBreak()) end
    rec:insert(pandoc.Span(pandoc.Inlines(stringify(meta["label-attention"]) .. ": "),
                           { ["custom-style"] = "Letter Label" }))
    rec:extend(inlines_of(meta.attention))
  end
  if #rec > 0 then out:insert(para("Letter Recipient", rec)) end

  -- place, date
  local pd = pandoc.Inlines({})
  pd:extend(inlines_of(meta.place)); pd:insert(pandoc.Str(", ")); pd:extend(inlines_of(meta.date))
  out:insert(para("Letter Date", pd))

  -- subject
  if meta.subject then
    local s = pandoc.Inlines({
      pandoc.Span(pandoc.Inlines(stringify(meta["label-subject"]) .. ":"),
                  { ["custom-style"] = "Letter Label" }),
      pandoc.Str("\u{2002}"),
      pandoc.Strong(inlines_of(meta.subject)),
    })
    out:insert(para("Letter Subject", s))
  end

  out:insert(para("Letter Opening", inlines_of(meta.opening)))
  out:extend(body)

  -- closing + signature block (keep-with-next chain in reference.docx)
  out:insert(para("Letter Closing", inlines_of(meta.closing)))
  out:insert(para("Letter Signer", pandoc.Inlines({ pandoc.Strong(inlines_of(meta.organisation)) })))
  if meta.signer then out:insert(para("Letter Signer", inlines_of(meta.signer))) end
  if meta["signer-title"] then out:insert(para("Letter Signer Title", inlines_of(meta["signer-title"]))) end
  if meta["signature-line"] then
    out:insert(para("Letter Signature",
      pandoc.Inlines({ pandoc.Str(stringify(meta["label-signature"]) .. ":"), tab() })))
  end

  -- enclosures
  local enc = list_of(meta.enclosures)
  if #enc > 0 then
    out:insert(para("Letter Enclosures Title",
      pandoc.Inlines({ pandoc.Strong(pandoc.Inlines(stringify(meta["label-enclosures"]) .. ":")) })))
    for _, e in ipairs(enc) do
      local inl = pandoc.Inlines({ pandoc.Str("–\u{2002}") }); inl:extend(e)
      out:insert(para("Letter Enclosure", inl))
    end
  end
  return out
end

function Pandoc(doc)
  local meta = doc.meta
  if not meta.lang then meta.lang = pandoc.MetaString("hr") end
  local lang = stringify(meta.lang):lower():sub(1, 2)
  local L = labels[lang] or labels.en

  meta["label-subject"] = pandoc.Inlines(L.subject)
  meta["label-attention"] = pandoc.Inlines(L.attention)
  meta["label-signature"] = pandoc.Inlines(L.signature)
  meta["label-enclosures"] = pandoc.Inlines(L.enclosures)

  meta.place = meta.place or pandoc.Inlines(DEFAULT_PLACE)
  meta.date = meta.date or pandoc.Inlines(L.date())
  meta.opening = meta.opening or pandoc.Inlines(L.opening)
  meta.closing = meta.closing or pandoc.Inlines(L.closing)
  meta.organisation = meta.organisation or pandoc.Inlines(DEFAULT_ORGANISATION)
  if not meta.signer and meta.author then
    local a = meta.author
    if pandoc.utils.type(a) == "List" then a = a[1] end
    meta.signer = inlines_of(a)
  end
  if falsy(meta["signature-line"]) then
    meta["signature-line"] = nil
  else
    meta["signature-line"] = true
  end

  -- recipient split for the LaTeX template (first line bold)
  local to = list_of(meta.to)
  if #to > 0 then
    meta["to-first"] = to[1]
    local rest = {}
    for i = 2, #to do rest[#rest + 1] = pandoc.MetaInlines(to[i]) end
    meta["to-rest"] = pandoc.MetaList(rest)
  end
  -- enclosures as a list for the LaTeX template
  if meta.enclosures and pandoc.utils.type(meta.enclosures) ~= "List" then
    meta.enclosures = pandoc.MetaList({ meta.enclosures })
  end

  if FORMAT:match("docx") then
    doc.blocks = docx_blocks(meta, doc.blocks)
  end
  doc.meta = meta
  return doc
end
