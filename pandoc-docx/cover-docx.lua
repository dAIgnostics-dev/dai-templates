-- cover-docx.lua
-- dAIgnostics cover page, table of contents and section numbering for DOCX.
--
-- pandoc's docx writer only knows Title/Subtitle/Author/Date paragraphs, so
-- this filter builds the cover of ../pandoc-document/daignostics.latex from
-- the same YAML keys and prepends it to the body:
--
--   logo ................................................. project
--   ════════════════════════════ (red bar) ══════════════════════
--   DOCTYPE
--   Title
--   Subtitle
--   ──────
--   Author(s)  ...
--   Version    ...
--   Date       ...
--   funding text (bottom)                              [EU emblem]
--
-- The cover is page 1 and uses reference.docx's first-page header/footer
-- (no running header; funding text from the `funding` key + large emblem).
-- Pandoc's own title block is hidden by the reference.docx styles but still
-- fills the document properties (title, author).
--
-- Also handled here (because pandoc would put them *before* the cover):
--   toc: true / toc-depth / toc-title   -> TOC field after the cover
--   abstract                            -> after the cover, before the TOC
--   numbersections: true                -> red "1.2" heading numbers
--   confidential: true                  -> "Confidential" in the footer
--   lang: hr                            -> Croatian cover labels / TOC title

if not FORMAT:match("docx") then return {} end

local stringify = pandoc.utils.stringify

-- Official logo (registered figurative mark HR Z20251009), next to this
-- filter. Used as supplied - only scaled by height, never redrawn, stretched
-- or recoloured. Without the file the cover keeps the text wordmark.
local BRAND_LOGO = pandoc.path.join({ pandoc.path.directory(PANDOC_SCRIPT_FILE),
                                      "daignostics-logo.png" })

local function file_exists(path)
  local f = io.open(path, "rb")
  if f then f:close() return true end
  return false
end

local labels = {
  en = { author = "Author(s)", version = "Version", date = "Date",
         contents = "Contents", confidential = "Confidential",
         doctype = "Software Documentation",
         funding = "This project has received funding from the European Union. " ..
           "Views and opinions expressed are however those of the author(s) only " ..
           "and do not necessarily reflect those of the European Union." },
  hr = { author = "Autor(i)", version = "Verzija", date = "Datum",
         contents = "Sadržaj", confidential = "Povjerljivo",
         doctype = "Softverska dokumentacija",
         funding = "Ovaj projekt financiran je sredstvima Europske unije. " ..
           "Izneseni stavovi i mišljenja su stavovi i mišljenja autora i ne " ..
           "odražavaju nužno stavove i mišljenja Europske unije." },
}

local function truthy(v)
  if v == nil then return false end
  if type(v) == "boolean" then return v end
  local s = stringify(v):lower()
  return s ~= "" and s ~= "false" and s ~= "no" and s ~= "0"
end

local function styled(style, blocks)
  return pandoc.Div(blocks, { ["custom-style"] = style })
end

local function span(style, inl)
  return pandoc.Span(inl, { ["custom-style"] = style })
end

local function tab()
  return pandoc.RawInline("openxml", "<w:r><w:tab/></w:r>")
end

local function inlines_of(v)
  if v == nil then return nil end
  if pandoc.utils.type(v) == "Inlines" then return v end
  if type(v) == "table" and v.t == "MetaInlines" then return v end
  if pandoc.utils.type(v) == "Blocks" then return pandoc.utils.blocks_to_inlines(v) end
  return pandoc.Inlines(stringify(v))
end

local function xml_escape(s)
  return (s:gsub("&", "&amp;"):gsub("<", "&lt;"):gsub(">", "&gt;"))
end

-- TOC field. Word refreshes it on open (w:dirty) and adds page numbers; the
-- cached entries (headings without page numbers) are what viewers that do
-- not update fields, e.g. LibreOffice, show until the index is updated.
local function toc_block(title, depth, blocks)
  local entries = {}
  blocks:walk({
    Header = function(h)
      if h.level <= depth and not h.classes:includes("unlisted") then
        local text = xml_escape(stringify(h.content))
        local inner = '<w:r><w:t xml:space="preserve">' .. text .. '</w:t></w:r>'
        if h.identifier ~= "" then
          inner = '<w:hyperlink w:anchor="' .. xml_escape(h.identifier) ..
            '" w:history="1">' .. inner .. '</w:hyperlink>'
        end
        entries[#entries + 1] = { h.level, inner }
      end
    end,
  })
  local begin = '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' ..
    '<w:r><w:instrText xml:space="preserve">TOC \\o "1-' .. depth ..
    '" \\h \\z \\u</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  local fin = '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
  local paras = {}
  if #entries == 0 then
    paras[1] = '<w:p>' .. begin .. fin .. '</w:p>'
  else
    for i, e in ipairs(entries) do
      paras[i] = '<w:p><w:pPr><w:pStyle w:val="TOC' .. e[1] .. '"/></w:pPr>' ..
        (i == 1 and begin or "") .. e[2] .. (i == #entries and fin or "") .. '</w:p>'
    end
  end
  return pandoc.RawBlock("openxml",
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' ..
    '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' ..
    '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t xml:space="preserve">' ..
    xml_escape(title) .. '</w:t></w:r></w:p>' .. table.concat(paras) ..
    '</w:sdtContent></w:sdt>')
end

local function number_headings(blocks, maxdepth)
  local counters = { 0, 0, 0, 0, 0, 0 }
  return blocks:walk({
    Header = function(h)
      if h.classes:includes("unnumbered") then return nil end
      local lvl = h.level
      counters[lvl] = counters[lvl] + 1
      for i = lvl + 1, #counters do counters[i] = 0 end
      if lvl > maxdepth then return nil end
      local parts = {}
      for i = 1, lvl do parts[#parts + 1] = tostring(counters[i]) end
      local num = table.concat(parts, ".")
      -- en space rather than a tab: survives in TOC entries and needs no tab stops
      local content = pandoc.Inlines({ span("Section Number", pandoc.Inlines(num)),
                                       pandoc.Str("\u{2002}") })
      content:extend(h.content)
      h.content = content
      return h
    end,
  })
end

function Pandoc(doc)
  local meta = doc.meta
  local lang = stringify(meta.lang or "en"):lower():sub(1, 2)
  local L = labels[lang] or labels.en
  local front = pandoc.Blocks({})

  -- ---- cover --------------------------------------------------------
  if meta.title then
    local brand
    if file_exists(BRAND_LOGO) then
      brand = pandoc.Inlines({
        pandoc.Image(pandoc.Inlines("dAIgnostics"), BRAND_LOGO, "", { height = "0.9cm" }),
      })
    else
      brand = pandoc.Inlines({
        span("Wordmark", pandoc.Inlines("d")),
        span("Brand Accent", pandoc.Inlines("AI")),
        span("Wordmark", pandoc.Inlines("gnostics")),
      })
    end
    if meta.project then
      brand:insert(tab())
      brand:extend(inlines_of(meta.project))
    end
    front:insert(styled("Cover Brand", { pandoc.Para(brand) }))
    front:insert(styled("Cover Doctype",
      { pandoc.Para(meta.doctype and inlines_of(meta.doctype) or pandoc.Inlines(L.doctype)) }))
    front:insert(styled("Cover Title", { pandoc.Para(inlines_of(meta.title)) }))
    if meta.subtitle then
      front:insert(styled("Cover Subtitle", { pandoc.Para(inlines_of(meta.subtitle)) }))
    end
    front:insert(pandoc.RawBlock("openxml",
      '<w:p><w:pPr><w:pStyle w:val="CoverRule"/></w:pPr></w:p>'))

    local rows = {}
    if meta.author then
      local names = {}
      local authors = meta.author
      if pandoc.utils.type(authors) ~= "List" then authors = { authors } end
      for _, a in ipairs(authors) do names[#names + 1] = stringify(a) end
      rows[#rows + 1] = { L.author, pandoc.Inlines(table.concat(names, ", ")) }
    end
    if meta.version then rows[#rows + 1] = { L.version, inlines_of(meta.version) } end
    if meta.date then rows[#rows + 1] = { L.date, inlines_of(meta.date) } end
    for _, r in ipairs(rows) do
      local inl = pandoc.Inlines({ span("Cover Label", pandoc.Inlines(r[1])), tab() })
      inl:extend(r[2])
      front:insert(styled("Cover Meta", { pandoc.Para(inl) }))
    end

    -- end of the cover page. The funding text and the large EU emblem are
    -- in reference.docx's first-page footer (DOCPROPERTY "funding").
    front:insert(pandoc.RawBlock("openxml",
      '<w:p><w:pPr><w:pStyle w:val="CoverBreak"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'))
  end

  -- ---- abstract (pandoc would print it above the cover) ---------------
  if meta.abstract then
    local ab = meta.abstract
    local blocks = pandoc.utils.type(ab) == "Blocks" and ab or pandoc.Blocks({ pandoc.Para(inlines_of(ab)) })
    front:insert(styled("Abstract", blocks))
    meta.abstract = nil
  end

  local blocks = doc.blocks
  -- ---- section numbers -----------------------------------------------
  if truthy(meta.numbersections) and not PANDOC_WRITER_OPTIONS.number_sections then
    blocks = number_headings(blocks, 6)
  end

  -- ---- table of contents ---------------------------------------------
  local depth = tonumber(meta["toc-depth"] and stringify(meta["toc-depth"]) or "") or 3
  if truthy(meta.toc) and not PANDOC_WRITER_OPTIONS.table_of_contents then
    local title = meta["toc-title"] and stringify(meta["toc-title"]) or L.contents
    front:insert(toc_block(title, depth, blocks))
    meta.toc = false
  elseif PANDOC_WRITER_OPTIONS.table_of_contents then
    io.stderr:write("cover-docx.lua: --toc given on the command line; pandoc will place " ..
      "the TOC before the cover. Use `toc: true` in the YAML instead.\n")
  end

  -- ---- footer-left custom property (DOCPROPERTY field in the footer) --
  local fl
  if truthy(meta.confidential) then
    fl = L.confidential:upper()
  elseif meta.project then
    fl = stringify(meta.project)
  else
    fl = "\u{00A0}"
  end
  meta["footer-left"] = pandoc.MetaString(fl)
  -- cover funding acknowledgement (first-page footer). Word limits custom
  -- property strings to 255 characters.
  meta.funding = pandoc.MetaString(meta.funding and stringify(meta.funding) or L.funding)

  front:extend(blocks)
  doc.blocks = front
  doc.meta = meta
  return doc
end
