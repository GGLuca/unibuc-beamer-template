-- Header levels inside a slide (slide_level: 2).
-- Level 3: navy bold sub-heading via \subheading{} (no block).
-- Level 4: beamer block, built here as raw LaTeX so that it closes at the
--          next header of level <= 4 (pandoc's own block logic would let a
--          level-3 sub-heading fall inside the preceding block).
local function latex(s) return pandoc.RawInline('latex', s) end

function Pandoc(doc)
  local out, open = pandoc.List(), false
  local function close()
    if open then out:insert(pandoc.Plain({latex('\\end{block}')})); open = false end
  end
  for _, el in ipairs(doc.blocks) do
    if el.t == 'Header' and el.level <= 4 then close() end
    if el.t == 'Header' and el.level == 3 then
      local inl = pandoc.List({latex('\\subheading{')}); inl:extend(el.content); inl:insert(latex('}'))
      out:insert(pandoc.Para(inl))
    elseif el.t == 'Header' and el.level == 4 then
      local inl = pandoc.List({latex('\\begin{block}{')}); inl:extend(el.content); inl:insert(latex('}'))
      out:insert(pandoc.Plain(inl)); open = true
    else
      out:insert(el)
    end
  end
  close()
  doc.blocks = out
  return doc
end
