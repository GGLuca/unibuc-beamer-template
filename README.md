# University of Bucharest Beamer template (using R Markdown)

This repo contains my version of a Beamer template using the University of Bucharest colours, driven by R Markdown / pandoc. It is based on the CambridgeUS default and modified to match the official guidelines.

![Screenshot](screenshot.png)

## Files

- `index.Rmd`. The template containing the YAML header plus example slides that document the conventions. Just delete the examples and write your own content.
- `includes/presentation_setup.tex`. This contais the colours, title page, logo, background, and the `\subheading{}` macro.
- `includes/headings.lua`. The pandoc filter where I modified the level-3 headers to become sub-headings. Level-4 headers become blocks.
- `includes/logo.png`. The UB Logo (only on the title page).
- `includes/background.png`.  Some title-page artwork I found on the official ,PPT file. 
- `referinte.bib`. The bibliography file (BibTeX).

## Rendering

Requires R with `rmarkdown` and a XeLaTeX installation (I recommend `tinytex::install_tinytex()`).

``` r
rmarkdown::render("index.Rmd")
```

Or you can do it from the RStudio Knit button. For the handout mode: uncomment `classoption: handout` in the YAML.

## Slide generation 

- `# Title` Creates a section (but no section slide is produced; `\AtBeginSection{}` is empty).
- `## Title`. This is a new slide (`slide_level: 2`).
- `### Title` Sub-heading inside a slide.
- `#### Title` Creates a block inside a slide.
- Citations. Classic BibTex citations. Use `@Key` or `[@Key]` with keys from `referinte.bib` (one example entry, `Viechtbauer2010`); the list is printed on the last slide, `## References {.allowframebreaks}`.
- Tables. Use `knitr::kable()` / `kableExtra` in an R chunk with `echo=FALSE`. But of course, Markdown tables work a swell.
