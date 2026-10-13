-- Quarto drops the body of a custom "alg" float in DOCX unless the body is
-- wrapped in its own Div. Authors write the natural form:
--   ::: {#alg-x}
--   1. step ...
--
--   Caption text.
--   :::
-- and this filter adds the wrapper. The last paragraph stays the caption.
function Div(el)
  if not el.identifier:match("^alg%-") then return nil end
  local n = #el.content
  if n < 2 then return nil end
  local body = pandoc.List()
  for i = 1, n - 1 do body:insert(el.content[i]) end
  el.content = pandoc.List({ pandoc.Div(body, pandoc.Attr("", { "alg-body" })), el.content[n] })
  return el
end
