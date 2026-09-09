# University of Bucharest Beamer template (R Markdown)

My version of a Beamer template in University of Bucharest colours, driven by R Markdown / pandoc.

![Screenshot](screenshot.png)

## Files

- `index.Rmd` — the template: YAML header plus example slides that document the conventions. Delete the examples and write your own slides.
- `includes/presentation_setup.tex` — colours, title page, logo, background, and the `\subheading{}` macro.
- `includes/headings.lua` — pandoc filter: level-3 headers become sub-headings, level-4 headers become blocks.
- `includes/logo.png` — UB crest (title page).
- `includes/background.png` — title-page artwork (lower-right swoosh).
- `referinte.bib` — bibliography (BibTeX).

## Rendering

Requires R with `rmarkdown` and a XeLaTeX installation (e.g. `tinytex::install_tinytex()`).

``` r
rmarkdown::render("index.Rmd")
```

Or from the RStudio Knit button. Handout mode: uncomment `classoption: handout` in the YAML.

## Writing slides

- `# Title` — section (no section slide is produced; `\AtBeginSection{}` is empty).
- `## Title` — new slide (`slide_level: 2`).
- `### Title` — sub-heading inside a slide (navy bold, no box).
- `#### Title` — block inside a slide.
- `## Title {.allowframebreaks}` — let a long slide (e.g. the bibliography) break across frames.
- Citations: `@Key` or `[@Key]` with keys from `referinte.bib` (one example entry, `Viechtbauer2010`); the list is printed on the last slide, `## References {.allowframebreaks}`.
- Tables: `knitr::kable()` / `kableExtra` in an R chunk with `echo=FALSE`.

## Colours

- `bluep` — RGB(2, 33, 105), UB navy (titles, items, TOC, bibliography).
- `tealp` — RGB(0, 174, 202), teal (primary palette / footer).
