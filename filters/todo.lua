-- Make [[TODO: ...]] gaps impossible to miss in a draft: yellow highlight in
-- the PDF, and a "Todo" character style in Word (docx_fix.py colours it).
-- The checker still lists every TODO and a --final build refuses them.

local function is_open(el) return el.t == "Str" and el.text:match("^%[%[TODO") end
local function is_close(el) return el.t == "Str" and el.text:match("%]%]%p*$") end

local function wrap(inlines)
  if FORMAT:match("typst") then
    local out = pandoc.List({ pandoc.RawInline("typst", "#highlight(fill: rgb(\"#fff176\"))[") })
    out:extend(inlines)
    out:insert(pandoc.RawInline("typst", "]"))
    return out
  end
  return pandoc.List({ pandoc.Span(inlines, pandoc.Attr("", {}, { ["custom-style"] = "Todo" })) })
end

function Inlines(inlines)
  local out, i = pandoc.List(), 1
  while i <= #inlines do
    local el = inlines[i]
    if is_open(el) then
      local j = i
      while j <= #inlines and not is_close(inlines[j]) do j = j + 1 end
      if j > #inlines then j = #inlines end
      -- keep trailing punctuation after "]]" outside the highlight
      local last = inlines[j]
      local tail = nil
      if last.t == "Str" then
        local body, punct = last.text:match("^(.-%]%])(%p*)$")
        if body and punct ~= "" then
          inlines[j] = pandoc.Str(body)
          tail = pandoc.Str(punct)
        end
      end
      local span = pandoc.List()
      for k = i, j do span:insert(inlines[k]) end
      out:extend(wrap(span))
      if tail then out:insert(tail) end
      i = j + 1
    else
      out:insert(el)
      i = i + 1
    end
  end
  return out
end
