-- Bold every header-row cell (including multi-row merged headers), in all
-- output formats, so neither the Typst nor the Word template has to guess
-- how many header rows a table has.
local function embolden(blocks)
  return blocks:walk({
    Plain = function(b) return pandoc.Plain({ pandoc.Strong(b.content) }) end,
    Para = function(b) return pandoc.Para({ pandoc.Strong(b.content) }) end,
  })
end

function Table(tbl)
  for _, row in ipairs(tbl.head.rows) do
    for _, cell in ipairs(row.cells) do
      cell.contents = embolden(cell.contents)
    end
  end
  return tbl
end
