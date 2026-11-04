#' University of Bucharest Beamer presentation
#'
#' An R Markdown output format for Beamer slides in the University of
#' Bucharest colours. All the technical setup (theme, Source Sans 3 font
#' (bundled), XeLaTeX, logo,
#' title-page artwork, link colours, and the Lua filter for sub-headings and
#' blocks) is built in, so the YAML header of a presentation only needs the
#' title, author, and date.
#'
#' Header levels: `#` starts a section (shown in the top bar, no slide of its
#' own), `##` starts a slide, `###` is a sub-heading inside a slide, and
#' `####` is a block.
#'
#' Any pandoc variable set in the YAML header (e.g., `lang`, `fontsize`,
#' `aspectratio`, `urlcolor`, `toc-title`) takes precedence over the defaults
#' of this format.
#'
#' @param lang Language of the automatic labels ("Table", "Contents", ...):
#'   `"en"` (British English) or `"ro"` (Romanian).
#' @param toc Add a table of contents after the title slide?
#' @param handout Produce handouts (no overlays/pauses)?
#' @param keep_tex Keep the intermediate `.tex` file?
#' @param slide_level Header level that starts a new slide. The sub-heading
#'   and block conventions assume the default of 2.
#' @param includes Additional files to include, see [rmarkdown::includes()].
#' @param pandoc_args Additional command-line arguments for pandoc.
#' @param ... Further arguments passed to [rmarkdown::beamer_presentation()].
#'
#' @return An R Markdown output format.
#'
#' @examples
#' # The format object used by rmarkdown::render()
#' fmt <- ub_presentation()
#' class(fmt)
#'
#' # Romanian labels
#' fmt_ro <- ub_prezentare()
#'
#' \donttest{
#' # Create a presentation from the English template and render it
#' # (requires pandoc and XeLaTeX)
#' if (rmarkdown::pandoc_available()) {
#'   dir <- tempfile("ub-talk-")
#'   dir.create(dir)
#'   rmd <- rmarkdown::draft(file.path(dir, "talk.Rmd"),
#'                           template = "ub-beamer-en",
#'                           package = "unibucbeamer", edit = FALSE)
#'   pdf <- try(rmarkdown::render(rmd, quiet = TRUE))
#'   unlink(dir, recursive = TRUE)
#' }
#' }
#' @export
ub_presentation <- function(lang = c("en", "ro"),
                            toc = FALSE,
                            handout = FALSE,
                            keep_tex = FALSE,
                            slide_level = 2,
                            includes = NULL,
                            pandoc_args = NULL,
                            ...) {
  lang <- match.arg(lang)

  resources <- system.file("resources", package = "unibucbeamer", mustWork = TRUE)
  res <- function(f) file.path(resources, f)

  # Tell LaTeX where the logo, background and fonts live (forward slashes so
  # that Windows paths work as well).
  res_tex <- gsub("\\\\", "/", normalizePath(resources))
  graphics_header <- tempfile("ub-graphicspath-", fileext = ".tex")
  writeLines(c(sprintf("\\graphicspath{{%s/}}", res_tex),
               sprintf("\\newcommand{\\ubfontdir}{%s/fonts/}", res_tex)),
             graphics_header)

  includes <- as.list(includes)
  includes$in_header <- c(graphics_header, res("presentation_setup.tex"),
                          includes$in_header)

  # Defaults, applied only when not set in the YAML header. Link colours must
  # be pandoc variables: pandoc writes \hypersetup{hidelinks} after the
  # header includes, so setting them in presentation_setup.tex has no effect.
  metadata_defaults <- list(
    lang = if (lang == "ro") "ro" else "en-GB",
    `toc-title` = if (lang == "ro") "Cuprins" else "Contents"
  )
  variable_defaults <- list(
    fontsize = "11pt",
    mathspec = "true",
    aspectratio = "169",
    colorlinks = "true",
    linkcolor = ".",   # internal links keep the surrounding colour
    urlcolor = "tealp" # web links in UB teal
  )
  if (handout) variable_defaults$classoption <- "handout"

  pre_processor <- function(metadata, input_file, runtime, knit_meta,
                            files_dir, output_dir) {
    unset <- function(x) x[!names(x) %in% names(metadata)]
    md <- unset(metadata_defaults)
    vars <- unset(variable_defaults)
    c(unlist(Map(rmarkdown::pandoc_metadata_arg, names(md), md), use.names = FALSE),
      unlist(Map(rmarkdown::pandoc_variable_arg, names(vars), vars), use.names = FALSE))
  }

  base <- rmarkdown::beamer_presentation(
    theme = "CambridgeUS",
    fonttheme = "professionalfonts",
    latex_engine = "xelatex",
    toc = toc,
    slide_level = slide_level,
    keep_tex = keep_tex,
    includes = includes,
    pandoc_args = c(rmarkdown::pandoc_lua_filter_args(res("headings.lua")),
                    pandoc_args),
    ...
  )

  rmarkdown::output_format(
    knitr = NULL,
    pandoc = NULL,
    pre_processor = pre_processor,
    base_format = base
  )
}

#' @rdname ub_presentation
#' @description `ub_prezentare()` is the same format with Romanian labels.
#' @export
ub_prezentare <- function(...) {
  ub_presentation(lang = "ro", ...)
}
