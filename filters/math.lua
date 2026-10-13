-- Normalise LaTeX norms so Typst gets lr(‖x‖)^2 instead of a stray ‖
-- glyph carrying the exponent. \|x\|, \left\|x\right\| and
-- \left\lVert x\right\rVert all become \lVert x\rVert.
local function fix(s)
  s = s:gsub("\\left\\lVert", "\\lVert"):gsub("\\right\\rVert", "\\rVert")
  s = s:gsub("\\left\\|", "\\lVert"):gsub("\\right\\|", "\\rVert")
  local open = true
  s = s:gsub("\\|", function()
    local r = open and "\\lVert " or "\\rVert "
    open = not open
    return r
  end)
  return s
end

function Math(el)
  el.text = fix(el.text)
  return el
end
