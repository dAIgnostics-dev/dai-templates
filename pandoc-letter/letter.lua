--[[
letter.lua -- dAIgnostics business letter (dopis), shared by PDF and DOCX.

* Fills language-dependent labels (hr / en) unless the document sets them.
* Sets `lang-hr` / `lang-en` flags for the LaTeX template.
* For DOCX: prepends the address / reference block, subject and appends the
  closing, signature, enclosures and cc list (the LaTeX template does this
  itself for PDF).
]]

local LABELS = {
  hr = {
    ["label-our-ref"]    = "Naš znak",
    ["label-your-ref"]   = "Vaš znak",
    ["label-date"]       = "Datum",
    ["label-contact"]    = "Kontakt",
    ["label-subject"]    = "Predmet",
    ["label-enclosures"] = "Prilozi",
    ["label-cc"]         = "Na znanje",
    ["label-page"]       = "Stranica",
    ["label-court"]      = "Upisano kod",
    ["label-mbs"]        = "MBS",
    ["label-oib"]        = "OIB",
    ["label-capital"]    = "Temeljni kapital",
    ["label-management"] = "Direktor",
    ["label-bank"]       = "Banka",
    ["closing"]          = "S poštovanjem,",
    ["funding"]          = "Projekt je sufinancirala Europska unija.",
  },
  en = {
    ["label-our-ref"]    = "Our ref.",
    ["label-your-ref"]   = "Your ref.",
    ["label-date"]       = "Date",
    ["label-contact"]    = "Contact",
    ["label-subject"]    = "Subject",
    ["label-enclosures"] = "Enclosures",
    ["label-cc"]         = "cc",
    ["label-page"]       = "Page",
    ["label-court"]      = "Registered at",
    ["label-mbs"]        = "Reg. no. (MBS)",
    ["label-oib"]        = "VAT ID",
    ["label-capital"]    = "Share capital",
    ["label-management"] = "Managing director",
    ["label-bank"]       = "Bank",
    ["closing"]          = "Kind regards,",
    ["funding"]          = "Co-funded by the European Union.",
  },
}

local function s(v)
  if v == nil then return nil end
  return pandoc.utils.stringify(v)
end

local function lang_of(meta)
  local l = (s(meta.lang) or "hr"):lower()
  if l:sub(1, 2) == "en" then return "en" end
  return "hr"
end

---------------------------------------------------------------------------
-- DOCX helpers
---------------------------------------------------------------------------
local function esc(t)
  return (t:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"))
end

local function run(text, opts)
  opts = opts or {}
  local rpr = ""
  if opts.bold then rpr = rpr .. "<w:b/>" end
  if opts.color then rpr = rpr .. '<w:color w:val="' .. opts.color .. '"/>' end
  if opts.size then
    rpr = rpr .. '<w:sz w:val="' .. opts.size .. '"/><w:szCs w:val="' .. opts.size .. '"/>'
  end
  return '<w:r><w:rPr>' .. rpr .. '</w:rPr><w:t xml:space="preserve">' ..
         esc(text) .. '</w:t></w:r>'
end

local function para(style, runs)
  return '<w:p><w:pPr><w:pStyle w:val="' .. style .. '"/></w:pPr>' ..
         runs .. '</w:p>'
end

local function list_of(v)
  local out = {}
  if v == nil then return out end
  if type(v) == "table" and v.t == nil and #v > 0 then
    for _, x in ipairs(v) do out[#out + 1] = s(x) end
  elseif type(v) == "table" and (v.t == "MetaList") then
    for _, x in ipairs(v) do out[#out + 1] = s(x) end
  else
    out[1] = s(v)
  end
  return out
end

-- Address window (left) + reference block (right) as a borderless table.
local function docx_head(meta)
  local L = function(k) return s(meta[k]) end
  local left = {}
  left[#left + 1] = para("Return Address", run(
    (L("company-short") or "") .. " · " .. (L("company-street") or "") ..
    " · " .. (L("company-city") or "") ..
    (L("lang") == "en-GB" and (" · " .. (L("company-country-en") or "")) or "")))
  for _, line in ipairs(list_of(meta.recipient)) do
    left[#left + 1] = para("Recipient", run(line))
  end

  local right = {}
  local function info(label, value)
    if value and value ~= "" then
      right[#right + 1] = para("Reference Label", run(label))
      right[#right + 1] = para("Reference Value", run(value))
    end
  end
  info(L("label-our-ref"), L("our-ref"))
  info(L("label-your-ref"), L("your-ref"))
  local d = L("date")
  if d and L("place") then d = L("place") .. ", " .. d end
  info(L("label-date"), d)
  local c = L("contact")
  if c then
    local extra = {}
    if L("contact-email") then extra[#extra + 1] = L("contact-email") end
    if L("contact-phone") then extra[#extra + 1] = L("contact-phone") end
    right[#right + 1] = para("Reference Label", run(L("label-contact")))
    right[#right + 1] = para("Reference Value", run(c))
    for _, e in ipairs(extra) do
      right[#right + 1] = para("Reference Value", run(e))
    end
  end
  if #right == 0 then right[1] = '<w:p/>' end

  local nob = '<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/>' ..
              '<w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'
  local tbl =
    '<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/>' ..
    '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>' ..
    '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>' ..
    '<w:tblLayout w:type="fixed"/>' ..
    '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>' ..
    '</w:tblPr><w:tblGrid><w:gridCol w:w="5670"/><w:gridCol w:w="3968"/></w:tblGrid>' ..
    '<w:tr><w:trPr><w:trHeight w:val="2835" w:hRule="atLeast"/><w:cantSplit/></w:trPr>' ..
    '<w:tc><w:tcPr><w:tcW w:w="5670" w:type="dxa"/>' .. nob .. '</w:tcPr>' ..
    table.concat(left) .. '</w:tc>' ..
    '<w:tc><w:tcPr><w:tcW w:w="3968" w:type="dxa"/>' .. nob .. '</w:tcPr>' ..
    table.concat(right) .. '</w:tc></w:tr></w:tbl>'

  local blocks = {
    pandoc.RawBlock("openxml", para("Address Spacer", "")),
    pandoc.RawBlock("openxml", tbl),
  }
  if L("subject") then
    blocks[#blocks + 1] = pandoc.RawBlock("openxml",
      para("Letter Subject", run(L("label-subject") .. ": " .. L("subject"))))
  end
  if L("opening") then
    blocks[#blocks + 1] = pandoc.Para(pandoc.Inlines(L("opening")))
  end
  return blocks
end

local function docx_tail(meta)
  local L = function(k) return s(meta[k]) end
  local blocks = {}
  blocks[#blocks + 1] = pandoc.Div(pandoc.Para(pandoc.Inlines(L("closing"))),
    pandoc.Attr("", {}, {{"custom-style", "Closing"}}))
  local sig = L("signature")
  if sig then
    blocks[#blocks + 1] = pandoc.Div(
      pandoc.Para({pandoc.Image({}, sig, "", pandoc.Attr("", {}, {{"height", "1.6cm"}}))}),
      pandoc.Attr("", {}, {{"custom-style", "Signature Image"}}))
  else
    blocks[#blocks + 1] = pandoc.Div(pandoc.Para({pandoc.Str(" ")}),
      pandoc.Attr("", {}, {{"custom-style", "Signature Space"}}))
  end
  if L("signatory") then
    blocks[#blocks + 1] = pandoc.Div(pandoc.Para(pandoc.Inlines(L("signatory"))),
      pandoc.Attr("", {}, {{"custom-style", "Signatory"}}))
  end
  local t = L("signatory-title")
  if t then
    blocks[#blocks + 1] = pandoc.Div(pandoc.Para(pandoc.Inlines(t)),
      pandoc.Attr("", {}, {{"custom-style", "Signatory Title"}}))
  end
  blocks[#blocks + 1] = pandoc.Div(pandoc.Para(pandoc.Inlines(L("company-short") or "")),
    pandoc.Attr("", {}, {{"custom-style", "Signatory Company"}}))
  local function lst(label, items)
    if #items == 0 then return end
    blocks[#blocks + 1] = pandoc.Div(pandoc.Para(pandoc.Inlines(label .. ":")),
      pandoc.Attr("", {}, {{"custom-style", "Enclosure Label"}}))
    for _, it in ipairs(items) do
      blocks[#blocks + 1] = pandoc.Div(pandoc.Para(pandoc.Inlines("– " .. it)),
        pandoc.Attr("", {}, {{"custom-style", "Enclosure Item"}}))
    end
  end
  lst(L("label-enclosures"), list_of(meta.enclosures))
  lst(L("label-cc"), list_of(meta.cc))
  if meta["eu-funding"] then
    local logo = L("eu-logo") or "eu-logo.png"
    blocks[#blocks + 1] = pandoc.Div(pandoc.Para({
        pandoc.Image({}, logo, "", pandoc.Attr("", {}, {{"height", "0.9cm"}})),
        pandoc.Space(), pandoc.Str(L("funding")) }),
      pandoc.Attr("", {}, {{"custom-style", "Funding Note"}}))
  end
  return blocks
end

---------------------------------------------------------------------------
function Pandoc(doc)
  local meta = doc.meta
  local lang = lang_of(meta)
  for k, v in pairs(LABELS[lang]) do
    if meta[k] == nil then meta[k] = pandoc.MetaString(v) end
  end
  meta["lang-" .. lang] = true
  -- letterhead variant: a | b | c (default c, matching the e-mail signature);
  -- build.sh picks the DOCX reference document by the same rule
  local lh = (s(meta.letterhead) or "c"):lower():sub(1, 1)
  if not lh:match("^[abc]$") then lh = "c" end
  meta["layout-" .. lh] = true
  if lang == "hr" then meta.lang = pandoc.MetaString("hr-HR")
  else meta.lang = pandoc.MetaString("en-GB") end

  if FORMAT:match("docx") then
    local blocks = docx_head(meta)
    for _, b in ipairs(doc.blocks) do blocks[#blocks + 1] = b end
    for _, b in ipairs(docx_tail(meta)) do blocks[#blocks + 1] = b end
    doc.blocks = blocks
    -- the letter block already shows these; keep pandoc from adding a
    -- title block, but store the subject as the document title property
    meta.date = nil
    meta.title = nil
    meta.subtitle = nil
    meta.author = nil
  end
  doc.meta = meta
  return doc
end
