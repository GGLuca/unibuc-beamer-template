# University of Bucharest Beamer template (using R Markdown)

This repo contains my version of a Beamer template using the University of Bucharest colours, driven by R Markdown / pandoc. It is based on the CambridgeUS default and modified to match the official guidelines. It comes as a small R package, `unibucbeamer`, with two templates: one in Romanian and one in English.

![Screenshot](screenshot.png)

## Installation

``` r
# install.packages("remotes")
remotes::install_github("GGLuca/unibuc-beamer-template")
```

You also need a functioning XeLaTeX installation (I recommend `tinytex::install_tinytex()`).

## Usage

In RStudio: **File → New File → R Markdown… → From Template**, then choose

- **Prezentare Beamer UB (română)**, or
- **UB Beamer presentation (English)**.

RStudio creates a new folder with the `.Rmd` and the `includes/` folder. Knit it as it is to see the examples, then delete them and write your own content.

Without RStudio:

``` r
rmarkdown::draft("my-talk.Rmd", template = "ub-beamer-en", package = "unibucbeamer")
rmarkdown::render("my-talk/my-talk.Rmd")
```

For the handout mode: uncomment `classoption: handout` in the YAML.

## Files in each new presentation

- `<name>.Rmd`. The template containing the YAML header plus example slides that document the conventions.
- `includes/presentation_setup.tex`. This contains the colours, title page, logo, background, and the `\subheading{}` macro.
- `includes/headings.lua`. The pandoc filter where I modified the level-3 headers to become sub-headings. Level-4 headers become blocks.
- `includes/logo.png`. The UB Logo (only on the title page).
- `includes/background.png`. Some title-page artwork I found on the official .PPT file.
- `referinte.bib` / `references.bib`. The bibliography file (BibTeX).

The templates themselves live in `inst/rmarkdown/templates/`.

## Basics

- `# Title` Creates a section (but no section slide is produced).
- `## Title` This is a new slide (`slide_level: 2`).
- `### Title` Sub-heading inside a slide.
- `#### Title` Creates a block inside a slide.
- Citations. Classic BibTeX citations. Use `@Key` or `[@Key]` with keys from the `.bib` file (one example entry, `Viechtbauer2010`); the list is printed on the last slide.
- Tables. Use `knitr::kable()` in an R chunk with `echo=FALSE`. But of course, Markdown tables work as well.
- Language. `lang: ro` or `lang: en-GB` in the YAML sets the automatic labels (e.g., "Tabel" vs. "Table").

## Keynote

A Keynote theme with the same design is in `Keynote (MacOS)/`. It is not part of the R package; download it from the repo.

## License

The code and text are MIT-licensed (see `LICENSE.md`). The University of Bucharest logo and title-page artwork are not covered by this license; they belong to the University of Bucharest.
