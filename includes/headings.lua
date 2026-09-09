-- Header levels inside a slide (slide_level: 2).
-- Level 3: a navy bold sub-heading, typeset with \subheading{} from
--          includes/presentation_setup.tex (no block).
-- Level 4: demoted to level 3 so that pandoc turns it into a beamer block.
function Header(el)
  if el.level == 3 then
    local inlines = pandoc.List({pandoc.RawInline('latex', '\\subheading{')})
    inlines:extend(el.content)
    inlines:insert(pandoc.RawInline('latex', '}'))
    return pandoc.Para(inlines)
  elseif el.level == 4 then
    el.level = 3
    return el
  end
end
