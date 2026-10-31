# ============================================================================
# NDC body-only keyword analysis: final Appendix dictionary
#
# Safeguards:
#   1. Search variants belong to one canonical term and are never added twice.
#   2. Longest-match priority prevents nested double counting. For example,
#      "health benefits" is retained and overlapping "health" is rejected.
#   3. Text is analysed page by page, preserving page number and context.
#   4. No result files are created unless SAVE_OUTPUTS <- TRUE.
# ============================================================================

SAVE_OUTPUTS <- FALSE
OPEN_PREVIEW_TABLES <- TRUE
STRICT_CHECKS <- TRUE

# Input validation reflects the files inspected on 23 October 2026.
EXPECTED_PDFS <- 101L
EXPECTED_CANONICAL_TERMS <- 100L

keyword_file <- file.path("Keywords_Excel_27.08.xlsx")
pdf_dir <- file.path(
  "Data Set", "body_only_pdf", "body_only_pdfs"
)
metadata_file <- file.path(
  "Data Set", "NDC Report.xlsx"
)
output_dir <- file.path(
  #  project_dir,
  paste0(format(Sys.time(), "%Y-%m-%d_%H%M%S"), "_body_only_keyword_analysis")
)

required_packages <- c(
  "readxl", "pdftools", "dplyr", "tidyr", "purrr", "stringr",
  "stringi", "tibble", "ggplot2", "scales", "writexl", "readr",
  "sf", "rnaturalearth", "rnaturalearthdata", "countrycode", "ggrepel"
)
missing_packages <- required_packages[
  !vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)
]
if (length(missing_packages) > 0L) {
  stop("Install missing packages first: install.packages(c(",
       paste(sprintf("\"%s\"", missing_packages), collapse = ", "),
       "))")
}
invisible(lapply(required_packages, library, character.only = TRUE))

if (!file.exists(keyword_file)) stop("Keyword file not found:\n", keyword_file)
if (!dir.exists(pdf_dir)) stop("Body-only PDF folder not found:\n", pdf_dir)
if (!file.exists(metadata_file)) stop("Metadata workbook not found:\n", metadata_file)

normalize_text <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x |>
    stringi::stri_trans_nfkc() |>
    stringr::str_replace_all("[\u2010\u2011\u2012\u2013\u2014\u2212]", "-") |>
    stringr::str_replace_all("\u00a0", " ") |>
    stringr::str_replace_all("\\s+", " ") |>
    stringr::str_trim()
}

normalize_country <- function(x) {
  x <- x |>
    stringi::stri_trans_general("Latin-ASCII") |>
    stringr::str_to_lower() |>
    stringr::str_replace_all("[^a-z0-9]+", " ") |>
    stringr::str_squish()
  dplyr::recode(
    x,
    "czehia" = "czechia",
    "holy see or vatican" = "holy see",
    "norways" = "norway",
    "united kingdom" = "united kingdom of great britain and northern ireland",
    .default = x
  )
}

column_key <- function(x) {
  x |>
    stringr::str_to_lower() |>
    stringr::str_replace_all("[^a-z0-9]+", "")
}

find_column <- function(data, candidates, required = TRUE) {
  position <- match(column_key(candidates), column_key(names(data)))
  position <- position[!is.na(position)]
  if (length(position) == 0L) {
    if (required) {
      stop(
        "Required keyword column not found. Expected one of: ",
        paste(candidates, collapse = ", "),
        ". Available columns: ", paste(names(data), collapse = ", ")
      )
    }
    return(NULL)
  }
  names(data)[position[[1]]]
}

parse_logical <- function(x, default = FALSE) {
  value <- stringr::str_to_upper(normalize_text(x))
  result <- value %in% c("TRUE", "T", "1", "YES", "Y")
  result[value == ""] <- rep_len(default, length(value))[value == ""]
  result
}

escape_regex <- function(x) {
  characters <- strsplit(x, "", fixed = TRUE)[[1]]
  special <- c("\\", ".", "^", "$", "|", "(", ")", "[", "]", "{", "}", "*", "+", "?")
  paste0(
    ifelse(characters %in% special, paste0("\\", characters), characters),
    collapse = ""
  )
}

term_to_regex <- function(term) {
  escaped <- escape_regex(normalize_text(term))
  escaped <- paste(strsplit(escaped, " ", fixed = TRUE)[[1]], collapse = "\\s+")
  paste0("(?<![\\p{L}\\p{N}_])(?:", escaped, ")(?![\\p{L}\\p{N}_])")
}

split_variants <- function(canonical_term, alias_text) {
  aliases <- normalize_text(alias_text)
  aliases <- if (!nzchar(aliases)) {
    character()
  } else {
    unlist(stringr::str_split(aliases, "\\s*[|;\\n]\\s*"), use.names = FALSE)
  }
  forms <- unique(normalize_text(c(canonical_term, aliases)))
  forms[nzchar(forms)]
}

# Read and validate the final Appendix keyword dictionary.
keyword_sheets <- tryCatch(
  readxl::excel_sheets(keyword_file),
  error = function(error) {
    stop(
      "The keyword workbook exists but could not be opened. Make it available ",
      "offline in OneDrive and rerun.\nDetails: ", conditionMessage(error)
    )
  }
)
keyword_sheet <- if ("Appendix FINAL" %in% keyword_sheets) {
  "Appendix FINAL"
} else if ("Updated" %in% keyword_sheets) {
  warning(
    "Sheet 'Appendix FINAL' was not found; reading 'Updated'. ",
    "Formal analysis should normally use Appendix FINAL."
  )
  "Updated"
} else {
  stop(
    "Neither 'Appendix FINAL' nor 'Updated' exists. Available sheets: ",
    paste(keyword_sheets, collapse = ", ")
  )
}

keyword_raw <- readxl::read_excel(
  keyword_file, sheet = keyword_sheet, .name_repair = "unique"
)
term_column <- find_column(keyword_raw, c("Term", "Canonical Term"))
category_column <- find_column(keyword_raw, "Category")
level_column <- find_column(keyword_raw, "Level")
variant_column <- find_column(
  keyword_raw,
  c("Approved Search Variants", "Search Variants", "Variants", "Aliases"),
  required = FALSE
)
case_column <- find_column(
  keyword_raw,
  c("Case Sensitive", "Case_sensitive", "CaseSensitive"),
  required = FALSE
)
status_column <- find_column(keyword_raw, "Status", required = FALSE)

keyword_staging <- tibble::tibble(
  canonical_term = normalize_text(keyword_raw[[term_column]]),
  category = dplyr::na_if(normalize_text(keyword_raw[[category_column]]), ""),
  level = dplyr::na_if(
    stringr::str_to_upper(normalize_text(keyword_raw[[level_column]])),
    ""
  ),
  search_aliases = if (is.null(variant_column)) {
    rep("", nrow(keyword_raw))
  } else {
    normalize_text(keyword_raw[[variant_column]])
  },
  case_value = if (is.null(case_column)) {
    rep("", nrow(keyword_raw))
  } else {
    normalize_text(keyword_raw[[case_column]])
  },
  status = if (is.null(status_column)) {
    rep("", nrow(keyword_raw))
  } else {
    normalize_text(keyword_raw[[status_column]])
  }
) |>
  tidyr::fill(category, level, .direction = "down") |>
  dplyr::filter(nzchar(canonical_term)) |>
  dplyr::mutate(level = stringr::str_replace(level, "^LEVE\\s*2$", "LEVEL 2"))

if (any(nzchar(keyword_staging$status))) {
  keyword_staging <- keyword_staging |>
    dplyr::filter(stringr::str_to_upper(status) == "APPROVED")
}

default_case_sensitive <- keyword_staging$canonical_term %in% c("WHO", "UHC")
keywords <- keyword_staging |>
  dplyr::mutate(
    case_sensitive = parse_logical(case_value, default_case_sensitive)
  ) |>
  dplyr::transmute(
    keyword_id = dplyr::row_number(),
    canonical_term = canonical_term,
    category = category,
    level = level,
    search_aliases = search_aliases,
    case_sensitive = case_sensitive
  )

if (nrow(keywords) != EXPECTED_CANONICAL_TERMS) {
  stop(
    "Expected ", EXPECTED_CANONICAL_TERMS, " approved canonical terms, but found ",
    nrow(keywords), " on sheet '", keyword_sheet, "'. Check that the workbook ",
    "is the final Appendix version."
  )
}
if (
  any(is.na(keywords$category) | !nzchar(keywords$category)) ||
  any(is.na(keywords$level) | !nzchar(keywords$level))
) {
  stop("At least one approved term has a missing Category or Level.")
}
if (!all(keywords$level %in% c("LEVEL 1", "LEVEL 2"))) {
  stop(
    "Unexpected Level values: ",
    paste(setdiff(unique(keywords$level), c("LEVEL 1", "LEVEL 2")), collapse = ", ")
  )
}

duplicate_terms <- keywords |>
  dplyr::mutate(term_key = stringr::str_to_lower(canonical_term)) |>
  dplyr::count(term_key, name = "n") |>
  dplyr::filter(n > 1L)
if (nrow(duplicate_terms) > 0L) stop("Duplicate canonical terms remain.")

variant_table <- keywords |>
  dplyr::mutate(
    search_form = purrr::map2(
      canonical_term, search_aliases, split_variants
    )
  ) |>
  tidyr::unnest_longer(search_form) |>
  dplyr::mutate(
    search_form = normalize_text(search_form),
    variant_key = stringr::str_to_lower(search_form),
    pattern = vapply(search_form, term_to_regex, character(1)),
    variant_characters = nchar(search_form),
    variant_words = stringi::stri_count_words(search_form)
  )

variant_collisions <- variant_table |>
  dplyr::group_by(variant_key) |>
  dplyr::summarise(
    canonical_terms = paste(sort(unique(canonical_term)), collapse = " | "),
    n_canonical_terms = dplyr::n_distinct(canonical_term),
    .groups = "drop"
  ) |>
  dplyr::filter(n_canonical_terms > 1L)
if (nrow(variant_collisions) > 0L) {
  stop(
    "Search forms map to multiple canonical terms: ",
    paste(variant_collisions$canonical_terms, collapse = "; ")
  )
}

# Read metadata and locate the 101 body-only PDFs.
metadata <- readxl::read_excel(metadata_file, sheet = "CLEAN FINAL VER") |>
  dplyr::transmute(
    country_metadata = stringr::str_squish(as.character(Country)),
    country_key = normalize_country(Country),
    submission_year = `Submission Year`,
    submission_date = `Submission Date`,
    region = stringr::str_squish(as.character(`Region / Contient`)),
    income_level = stringr::str_squish(as.character(`Income Level`)),
    language = as.character(Language),
    version = Version,
    metadata_filename = as.character(Filename)
  ) |>
  dplyr::distinct(country_key, .keep_all = TRUE)

pdf_files <- list.files(
  pdf_dir, pattern = "\\.pdf$", full.names = TRUE, ignore.case = TRUE
)
if (length(pdf_files) != EXPECTED_PDFS) {
  stop(
    "Expected ", EXPECTED_PDFS, " body-only PDFs, but found ",
    length(pdf_files), " in:\n", pdf_dir
  )
}

extract_country <- function(filename) {
  country <- stringr::str_match(basename(filename), "^\\[([^]]+)\\]")[, 2]
  ifelse(is.na(country), tools::file_path_sans_ext(basename(filename)), country)
}

document_index <- tibble::tibble(
  pdf_path = sort(pdf_files),
  pdf_file = basename(sort(pdf_files)),
  country_from_file = extract_country(sort(pdf_files))
) |>
  dplyr::mutate(country_key = normalize_country(country_from_file)) |>
  dplyr::left_join(metadata, by = "country_key") |>
  dplyr::mutate(country = dplyr::coalesce(country_metadata, country_from_file))

unmatched_metadata <- document_index |>
  dplyr::filter(is.na(country_metadata)) |>
  dplyr::select(country_from_file, pdf_file)
if (nrow(unmatched_metadata) > 0L && STRICT_CHECKS) {
  stop(
    "Metadata were not matched for: ",
    paste(unmatched_metadata$country_from_file, collapse = ", ")
  )
}

# Candidate matching and global longest-match resolution.
empty_candidates <- function() {
  tibble::tibble(
    keyword_id = integer(), canonical_term = character(), category = character(),
    level = character(), search_form = character(), case_sensitive = logical(),
    page = integer(), start = integer(), end = integer(), match_length = integer(),
    variant_characters = integer(), variant_words = integer(),
    matched_text = character(), context = character()
  )
}

add_context <- function(data, page_text, characters = 110L) {
  if (nrow(data) == 0L) {
    data$context <- character()
    return(data)
  }
  context_start <- pmax(1L, data$start - characters)
  context_end <- pmin(nchar(page_text), data$end + characters)
  data$context <- stringi::stri_sub(page_text, context_start, context_end)
  data
}

find_page_candidates <- function(page_text, page_number) {
  if (!nzchar(page_text)) return(empty_candidates())
  matches <- vector("list", nrow(variant_table))
  for (index in seq_len(nrow(variant_table))) {
    variant <- variant_table[index, ]
    locations <- stringi::stri_locate_all_regex(
      page_text,
      variant$pattern[[1]],
      opts_regex = stringi::stri_opts_regex(
        case_insensitive = !variant$case_sensitive[[1]]
      )
    )[[1]]
    if (nrow(locations) == 0L || all(is.na(locations))) next
    locations <- locations[stats::complete.cases(locations), , drop = FALSE]
    matches[[index]] <- tibble::tibble(
      keyword_id = variant$keyword_id[[1]],
      canonical_term = variant$canonical_term[[1]],
      category = variant$category[[1]],
      level = variant$level[[1]],
      search_form = variant$search_form[[1]],
      case_sensitive = variant$case_sensitive[[1]],
      page = as.integer(page_number),
      start = as.integer(locations[, "start"]),
      end = as.integer(locations[, "end"]),
      match_length = as.integer(locations[, "end"] - locations[, "start"] + 1L),
      variant_characters = variant$variant_characters[[1]],
      variant_words = variant$variant_words[[1]],
      matched_text = stringi::stri_sub(
        page_text, from = locations[, "start"], to = locations[, "end"]
      ),
      context = ""
    )
  }
  candidates <- dplyr::bind_rows(matches)
  if (nrow(candidates) == 0L) return(empty_candidates())
  candidates |>
    dplyr::arrange(
      dplyr::desc(match_length),
      dplyr::desc(variant_characters),
      start,
      canonical_term
    ) |>
    dplyr::distinct(
      start, end, canonical_term, .keep_all = TRUE
    )
}

resolve_longest_matches <- function(candidates, page_text) {
  if (nrow(candidates) == 0L) {
    rejected <- candidates |>
      dplyr::mutate(
        blocked_by_term = character(),
        blocked_by_search_form = character(),
        blocked_by_matched_text = character()
      )
    return(list(selected = candidates, rejected = rejected))
  }
  candidates <- candidates |>
    dplyr::arrange(
      dplyr::desc(match_length),
      dplyr::desc(variant_words),
      dplyr::desc(variant_characters),
      start,
      canonical_term
    ) |>
    dplyr::mutate(candidate_row = dplyr::row_number())
  keep <- rep(FALSE, nrow(candidates))
  blocked_by <- rep(NA_integer_, nrow(candidates))
  accepted_rows <- integer()
  for (index in seq_len(nrow(candidates))) {
    if (length(accepted_rows) == 0L) {
      keep[[index]] <- TRUE
      accepted_rows <- c(accepted_rows, index)
      next
    }
    overlaps <- accepted_rows[
      candidates$start[[index]] <= candidates$end[accepted_rows] &
        candidates$end[[index]] >= candidates$start[accepted_rows]
    ]
    if (length(overlaps) == 0L) {
      keep[[index]] <- TRUE
      accepted_rows <- c(accepted_rows, index)
    } else {
      blocked_by[[index]] <- overlaps[[1]]
    }
  }
  selected <- candidates[keep, , drop = FALSE] |>
    dplyr::select(-"candidate_row") |>
    dplyr::arrange(start, end) |>
    add_context(page_text)
  rejected <- candidates[!keep, , drop = FALSE]
  if (nrow(rejected) > 0L) {
    blocker_rows <- blocked_by[!keep]
    rejected <- rejected |>
      dplyr::mutate(
        blocked_by_term = candidates$canonical_term[blocker_rows],
        blocked_by_search_form = candidates$search_form[blocker_rows],
        blocked_by_matched_text = candidates$matched_text[blocker_rows]
      ) |>
      dplyr::select(-"candidate_row") |>
      dplyr::arrange(start, end) |>
      add_context(page_text)
  } else {
    rejected <- rejected |>
      dplyr::mutate(
        blocked_by_term = character(),
        blocked_by_search_form = character(),
        blocked_by_matched_text = character()
      ) |>
      dplyr::select(-"candidate_row")
  }
  list(selected = selected, rejected = rejected)
}

analyse_pdf <- function(path) {
  extraction_warnings <- character()
  tryCatch(
    {
      pages <- withCallingHandlers(
        pdftools::pdf_text(path),
        warning = function(warning_condition) {
          extraction_warnings <<- unique(c(
            extraction_warnings, conditionMessage(warning_condition)
          ))
          invokeRestart("muffleWarning")
        }
      )
      page_text <- vapply(pages, normalize_text, character(1))
      page_word_count <- stringi::stri_count_words(page_text)
      page_results <- vector("list", length(page_text))
      for (page_number in seq_along(page_text)) {
        candidates <- find_page_candidates(page_text[[page_number]], page_number)
        page_results[[page_number]] <- resolve_longest_matches(
          candidates, page_text[[page_number]]
        )
      }
      list(
        page_count = as.integer(length(page_text)),
        report_word_count = as.integer(sum(page_word_count, na.rm = TRUE)),
        extraction_status = "OK",
        extraction_warning = paste(extraction_warnings, collapse = " | "),
        extraction_error = "",
        selected = dplyr::bind_rows(purrr::map(page_results, "selected")),
        rejected = dplyr::bind_rows(purrr::map(page_results, "rejected"))
      )
    },
    error = function(error) {
      list(
        page_count = NA_integer_, report_word_count = NA_integer_,
        extraction_status = "ERROR",
        extraction_warning = paste(extraction_warnings, collapse = " | "),
        extraction_error = conditionMessage(error),
        selected = empty_candidates(),
        rejected = empty_candidates() |>
          dplyr::mutate(
            blocked_by_term = character(),
            blocked_by_search_form = character(),
            blocked_by_matched_text = character()
          )
      )
    }
  )
}

message(
  "Analysing ", length(pdf_files), " body-only PDFs with ", nrow(keywords),
  " canonical terms and ", nrow(variant_table), " approved search forms..."
)
analysis_results <- vector("list", nrow(document_index))
for (index in seq_len(nrow(document_index))) {
  message(sprintf(
    "[%03d/%03d] %s", index, nrow(document_index), document_index$pdf_file[[index]]
  ))
  analysis_results[[index]] <- analyse_pdf(document_index$pdf_path[[index]])
}

documents <- document_index |>
  dplyr::mutate(
    page_count = purrr::map_int(analysis_results, "page_count"),
    report_word_count = purrr::map_int(analysis_results, "report_word_count"),
    extraction_status = purrr::map_chr(analysis_results, "extraction_status"),
    extraction_warning = purrr::map_chr(analysis_results, "extraction_warning"),
    extraction_error = purrr::map_chr(analysis_results, "extraction_error"),
    pdf_md5 = unname(tools::md5sum(pdf_path))
  )
failed_extractions <- documents |>
  dplyr::filter(extraction_status != "OK")
if (nrow(failed_extractions) > 0L && STRICT_CHECKS) {
  stop(
    "PDF text extraction failed for: ", paste(failed_extractions$pdf_file, collapse = ", "),
    ". No formal results were saved."
  )
}

bind_document_matches <- function(result_name) {
  purrr::map_dfr(seq_len(nrow(documents)), function(index) {
    matches <- analysis_results[[index]][[result_name]]
    if (nrow(matches) == 0L) return(matches)
    matches |>
      dplyr::mutate(
        country = documents$country[[index]],
        country_key = documents$country_key[[index]],
        pdf_file = documents$pdf_file[[index]],
        region = documents$region[[index]],
        income_level = documents$income_level[[index]],
        report_word_count = documents$report_word_count[[index]],
        .before = 1
      )
  })
}

occurrences <- bind_document_matches("selected") |>
  dplyr::mutate(occurrence_id = dplyr::row_number(), .before = 1)
overlap_rejections <- bind_document_matches("rejected") |>
  dplyr::mutate(rejection_id = dplyr::row_number(), .before = 1)

retained_overlap_check <- occurrences |>
  dplyr::arrange(pdf_file, page, start, end) |>
  dplyr::group_by(pdf_file, page) |>
  dplyr::mutate(
    previous_max_end = dplyr::lag(cummax(end)),
    retained_overlap = !is.na(previous_max_end) &
      start <= previous_max_end
  ) |>
  dplyr::ungroup() |>
  dplyr::filter(retained_overlap)
if (nrow(retained_overlap_check) > 0L) {
  stop("Internal validation failed: retained matches still overlap.")
}

# Country, level, category, region, and income-group results.
occurrence_counts <- occurrences |>
  dplyr::count(country_key, canonical_term, name = "keyword_count")
country_term <- tidyr::crossing(
  documents |>
    dplyr::select(
      country, country_key, pdf_file,
      region, income_level, report_word_count
    ),
  keywords |>
    dplyr::select(
      keyword_id, canonical_term, category,
      level, search_aliases, case_sensitive
    )
) |>
  dplyr::left_join(occurrence_counts, by = c("country_key", "canonical_term")) |>
  dplyr::mutate(
    keyword_count = dplyr::coalesce(keyword_count, 0L),
    keyword_percentage = dplyr::if_else(
      report_word_count > 0L,
      100 * keyword_count / report_word_count,
      NA_real_
    )
  )

country_overall <- country_term |>
  dplyr::group_by(
    country, country_key, pdf_file,
    region, income_level, report_word_count
  ) |>
  dplyr::summarise(keyword_count = sum(keyword_count), .groups = "drop") |>
  dplyr::mutate(
    keyword_percentage = 100 * keyword_count / report_word_count
  ) |>
  dplyr::arrange(dplyr::desc(keyword_percentage))
country_level <- country_term |>
  dplyr::group_by(
    country, country_key, pdf_file,
    region, income_level, report_word_count, level
  ) |>
  dplyr::summarise(keyword_count = sum(keyword_count), .groups = "drop") |>
  dplyr::mutate(
    keyword_percentage = 100 * keyword_count / report_word_count
  )
country_category <- country_term |>
  dplyr::group_by(
    country, country_key, pdf_file,
    region, income_level, report_word_count,
    level, category
  ) |>
  dplyr::summarise(keyword_count = sum(keyword_count), .groups = "drop") |>
  dplyr::mutate(
    keyword_percentage = 100 * keyword_count / report_word_count
  )
if (sum(country_term$keyword_count) != nrow(occurrences)) {
  stop("Internal validation failed: counts do not reconcile to occurrences.")
}

summarise_group <- function(data, group_variable, extra_groups = character()) {
  data |>
    dplyr::group_by(dplyr::across(dplyr::all_of(c(group_variable, extra_groups)))) |>
    dplyr::summarise(
      n_countries = dplyr::n_distinct(country),
      mean_keyword_count = mean(keyword_count, na.rm = TRUE),
      mean_keyword_percentage = mean(keyword_percentage, na.rm = TRUE),
      .groups = "drop"
    )
}
region_overall <- summarise_group(country_overall, "region")
income_overall <- summarise_group(country_overall, "income_level")
region_level <- summarise_group(country_level, "region", "level")
income_level <- summarise_group(country_level, "income_level", "level")
region_category <- summarise_group(country_category, "region", c("level", "category"))
income_category <- summarise_group(country_category, "income_level", c("level", "category"))


# Explicit rankings: percentage and absolute count are different rankings.
rank_results <- function(data, groups = character()) {
  data |>
    dplyr::group_by(dplyr::across(dplyr::all_of(groups))) |>
    dplyr::mutate(
      rank_percentage = dplyr::min_rank(dplyr::desc(keyword_percentage)),
      rank_count = dplyr::min_rank(dplyr::desc(keyword_count))
    ) |>
    dplyr::ungroup() |>
    dplyr::arrange(dplyr::across(dplyr::all_of(groups)),
                   rank_percentage, country)
}
country_overall <- rank_results(country_overall)
country_level <- rank_results(country_level, "level")
country_category <- rank_results(country_category, c("level", "category"))
mental_health_ranking <- country_category |>
  dplyr::filter(stringr::str_detect(
    stringr::str_to_lower(category), "mental"
  )) |>
  dplyr::arrange(rank_percentage, country)

# New body-only texts require fresh contextual review.
# This table is a review queue, not automatic exclusion.
review_queue <- occurrences |>
  dplyr::filter(
    stringr::str_detect(stringr::str_to_lower(category), "mental") |
      stringr::str_to_lower(canonical_term) %in%
      c("stress", "depression", "mental", "psychosocial", "just", "life", "harm")
  ) |>
  dplyr::mutate(review_decision = "", review_note = "")
duplicate_files <- documents |>
  dplyr::group_by(pdf_md5) |>
  dplyr::filter(dplyr::n() > 1L) |>
  dplyr::ungroup() |>
  dplyr::select(country, pdf_file, pdf_md5)
# Equal hashes identify byte-identical files; differently encoded duplicates
# can still exist and must be checked separately.
global_summary <- tibble::tibble(
  n_country_records = nrow(country_overall),
  total_body_words = sum(country_overall$report_word_count),
  keyword_occurrences = sum(country_overall$keyword_count),
  pooled_percentage = 100 * sum(country_overall$keyword_count) /
    sum(country_overall$report_word_count),
  mean_country_percentage = mean(country_overall$keyword_percentage),
  median_country_percentage = median(country_overall$keyword_percentage)
)
level_summary <- country_level |>
  dplyr::group_by(level) |>
  dplyr::summarise(
    n_country_records = dplyr::n(),
    countries_with_matches = sum(keyword_count > 0),
    keyword_count = sum(keyword_count),
    body_words = sum(report_word_count),
    pooled_percentage = 100 * sum(keyword_count) /
      sum(report_word_count),
    mean_country_percentage = mean(keyword_percentage),
    .groups = "drop"
  )

# Map preparation.
country_aliases <- c(
  "Holy See" = "VAT", "Vatican" = "VAT", "Micronesia" = "FSM",
  "Federated States of Micronesia" = "FSM", "Republic of Korea" = "KOR",
  "Turkiye" = "TUR", "Cabo Verde" = "CPV",
  "Sao Tome and Principe" = "STP", "The Bahamas" = "BHS"
)
add_iso3 <- function(data) {
  data |>
    dplyr::mutate(
      iso3 = countrycode::countrycode(
        country, origin = "country.name", destination = "iso3c",
        custom_match = country_aliases, warn = FALSE
      )
    )
}
country_overall <- add_iso3(country_overall)
country_level <- add_iso3(country_level)
country_category <- add_iso3(country_category)
unmatched_map_countries <- country_overall |>
  dplyr::filter(is.na(iso3)) |>
  dplyr::select(country, pdf_file)
if (nrow(unmatched_map_countries) > 0L) {
  warning(
    "No map polygon code was found for: ",
    paste(unmatched_map_countries$country, collapse = ", "),
    ". Their tabular results are retained."
  )
}
world <- rnaturalearth::ne_countries(scale = "medium", returnclass = "sf") |>
  dplyr::filter(admin != "Antarctica") |>
  dplyr::mutate(
    iso3 = dplyr::if_else(iso_a3 == "-99", adm0_a3, iso_a3)
  )

ucl_purple <- "#500778"
ucl_navy <- "#002855"
ucl_blue <- "#005EB8"
ucl_cyan <- "#00A6D6"
ucl_red <- "#D50032"
no_data_colour <- "#D9D9D9"

journal_theme <- function(base_size = 10) {
  ggplot2::theme_classic(base_size = base_size) +
    ggplot2::theme(
      text = ggplot2::element_text(family = "Helvetica"),
      plot.title = ggplot2::element_text(face = "bold", size = base_size + 2),
      plot.subtitle = ggplot2::element_blank(),
      legend.position = "bottom",
      legend.title = ggplot2::element_text(face = "bold"),
      strip.background = ggplot2::element_blank(),
      strip.text = ggplot2::element_text(face = "bold"),
      axis.title = ggplot2::element_text(face = "bold")
    )
}
map_theme <- function(base_size = 9) {
  journal_theme(base_size) +
    ggplot2::theme(
      axis.title = ggplot2::element_blank(), axis.text = ggplot2::element_blank(),
      axis.ticks = ggplot2::element_blank(), axis.line = ggplot2::element_blank(),
      panel.border = ggplot2::element_blank(),
      plot.margin = ggplot2::margin(5, 8, 5, 8)
    )
}
format_metric <- function(x, metric) {
  if (metric == "keyword_percentage") {
    scales::number(x, accuracy = 0.01, suffix = "%")
  } else {
    scales::comma(x, accuracy = 1)
  }
}
metric_scale_labels <- function(metric) {
  if (metric == "keyword_percentage") {
    scales::label_number(accuracy = 0.1, suffix = "%")
  } else {
    scales::label_comma(accuracy = 1)
  }
}

make_label_points <- function(data, metric, group_column = NULL) {
  ranked <- if (is.null(group_column)) {
    data |>
      dplyr::filter(!is.na(iso3)) |>
      dplyr::slice_max(.data[[metric]], n = 5L, with_ties = FALSE)
  } else {
    data |>
      dplyr::filter(!is.na(iso3)) |>
      dplyr::group_by(.data[[group_column]]) |>
      dplyr::slice_max(.data[[metric]], n = 5L, with_ties = FALSE) |>
      dplyr::ungroup()
  }
  label_sf <- suppressWarnings(
    sf::st_point_on_surface(
      world |>
        dplyr::select(iso3, geometry) |>
        dplyr::inner_join(ranked, by = "iso3")
    )
  )
  coordinates <- sf::st_coordinates(label_sf)
  label_sf |>
    sf::st_drop_geometry() |>
    dplyr::mutate(
      label_x = coordinates[, 1], label_y = coordinates[, 2],
      map_label = paste0(
        country, "\n", format_metric(.data[[metric]], metric)
      )
    )
}

make_overall_map <- function(metric, title, legend_name, high_colour) {
  map_data <- world |> dplyr::left_join(country_overall, by = "iso3")
  labels <- make_label_points(country_overall, metric)
  ggplot2::ggplot(map_data) +
    ggplot2::geom_sf(
      ggplot2::aes(fill = .data[[metric]]), colour = "white", linewidth = 0.12
    ) +
    ggrepel::geom_text_repel(
      data = labels,
      ggplot2::aes(x = label_x, y = label_y, label = map_label),
      inherit.aes = FALSE, size = 3, fontface = "plain", colour = "#222222",
      box.padding = 0.22, point.padding = 0.1, min.segment.length = 0,
      segment.colour = "#666666", segment.size = 0.22,
      max.overlaps = Inf, seed = 27
    ) +
    ggplot2::scale_fill_gradient(
      low = "#EEE8F2", high = high_colour, na.value = no_data_colour,
      labels = metric_scale_labels(metric), name = legend_name
    ) +
    ggplot2::coord_sf(
      crs = "+proj=robin", default_crs = sf::st_crs(4326), expand = FALSE
    ) +
    ggplot2::labs(title = title) +
    map_theme()
}

make_level_map <- function(metric, title, legend_name, high_colour) {
  level_values <- sort(unique(country_level$level))
  map_data <- dplyr::bind_rows(lapply(level_values, function(level_value) {
    world |> dplyr::mutate(level = level_value)
  })) |>
    dplyr::left_join(country_level, by = c("iso3", "level"))
  labels <- make_label_points(country_level, metric, "level")
  ggplot2::ggplot(map_data) +
    ggplot2::geom_sf(
      ggplot2::aes(fill = .data[[metric]]), colour = "white", linewidth = 0.1
    ) +
    ggrepel::geom_text_repel(
      data = labels,
      ggplot2::aes(x = label_x, y = label_y, label = map_label),
      inherit.aes = FALSE, size = 2.6, fontface = "plain", colour = "#222222",
      box.padding = 0.18, point.padding = 0.08, min.segment.length = 0,
      segment.colour = "#777777", segment.size = 0.2,
      max.overlaps = Inf, seed = 27
    ) +
    ggplot2::facet_wrap(~level, ncol = 1) +
    ggplot2::scale_fill_gradient(
      low = "#E8F1F8", high = high_colour, na.value = no_data_colour,
      labels = metric_scale_labels(metric), name = legend_name
    ) +
    ggplot2::coord_sf(
      crs = "+proj=robin", default_crs = sf::st_crs(4326), expand = FALSE
    ) +
    ggplot2::labs(title = title) +
    map_theme()
}

make_category_map <- function(metric, title, legend_name, high_colour) {
  category_values <- unique(country_category$category)
  map_data <- dplyr::bind_rows(lapply(category_values, function(category_value) {
    world |> dplyr::mutate(category = category_value)
  })) |>
    dplyr::left_join(country_category, by = c("iso3", "category"))
  ggplot2::ggplot(map_data) +
    ggplot2::geom_sf(
      ggplot2::aes(fill = .data[[metric]]), colour = "white", linewidth = 0.07
    ) +
    ggplot2::facet_wrap(~category, ncol = 3) +
    ggplot2::scale_fill_gradient(
      low = "#F1ECF6", high = high_colour, na.value = no_data_colour,
      labels = metric_scale_labels(metric), name = legend_name
    ) +
    ggplot2::coord_sf(
      crs = "+proj=robin", default_crs = sf::st_crs(4326), expand = FALSE
    ) +
    ggplot2::labs(title = title) +
    map_theme(base_size = 8)
}

make_top20_plot <- function(metric, title, axis_title, fill_colour) {
  plot_data <- country_overall |>
    dplyr::slice_max(.data[[metric]], n = 20L, with_ties = FALSE) |>
    dplyr::mutate(country_plot = stats::reorder(country, .data[[metric]]))
  ggplot2::ggplot(
    plot_data, ggplot2::aes(x = .data[[metric]], y = country_plot)
  ) +
    ggplot2::geom_col(fill = fill_colour, width = 0.68) +
    ggplot2::geom_text(
      ggplot2::aes(label = format_metric(.data[[metric]], metric)),
      hjust = -0.12, size = 3
    ) +
    ggplot2::scale_x_continuous(
      labels = metric_scale_labels(metric),
      expand = ggplot2::expansion(mult = c(0, 0.18))
    ) +
    ggplot2::labs(title = title, x = axis_title, y = NULL) +
    journal_theme()
}

make_mean_plot <- function(data, group_column, metric, title, axis_title, fill_colour) {
  ggplot2::ggplot(
    data,
    ggplot2::aes(
      x = .data[[metric]],
      y = stats::reorder(.data[[group_column]], .data[[metric]])
    )
  ) +
    ggplot2::geom_col(fill = fill_colour, width = 0.65) +
    ggplot2::geom_text(
      ggplot2::aes(label = format_metric(.data[[metric]], metric)),
      hjust = -0.12, size = 3
    ) +
    ggplot2::scale_x_continuous(
      labels = metric_scale_labels(metric),
      expand = ggplot2::expansion(mult = c(0, 0.2))
    ) +
    ggplot2::labs(title = title, x = axis_title, y = NULL) +
    journal_theme()
}

plots <- list(
  overall_country_percentage_map = make_overall_map(
    "keyword_percentage", "Overall country-level keyword percentage",
    "Keyword percentage", ucl_purple
  ),
  overall_country_frequency_map = make_overall_map(
    "keyword_count", "Overall country-level keyword frequency",
    "Keyword frequency", ucl_navy
  ),
  country_percentage_by_level_maps = make_level_map(
    "keyword_percentage", "Country-level keyword percentage by search level",
    "Keyword percentage", ucl_purple
  ),
  country_frequency_by_level_maps = make_level_map(
    "keyword_count", "Country-level keyword frequency by search level",
    "Keyword frequency", ucl_blue
  ),
  country_percentage_by_category_maps = make_category_map(
    "keyword_percentage", "Country-level keyword percentage by category",
    "Keyword percentage", ucl_purple
  ),
  country_frequency_by_category_maps = make_category_map(
    "keyword_count", "Country-level keyword frequency by category",
    "Keyword frequency", ucl_navy
  ),
  top20_country_percentage = make_top20_plot(
    "keyword_percentage", "Top 20 countries by overall keyword percentage",
    "Keyword percentage", ucl_purple
  ),
  top20_country_frequency = make_top20_plot(
    "keyword_count", "Top 20 countries by overall keyword frequency",
    "Keyword frequency", ucl_navy
  ),
  mean_percentage_by_region = make_mean_plot(
    region_overall, "region", "mean_keyword_percentage",
    "Mean keyword percentage by region", "Mean country percentage", ucl_blue
  ),
  mean_percentage_by_income = make_mean_plot(
    income_overall, "income_level", "mean_keyword_percentage",
    "Mean keyword percentage by income group", "Mean country percentage", ucl_cyan
  ),
  mean_frequency_by_region = make_mean_plot(
    region_overall, "region", "mean_keyword_count",
    "Mean keyword frequency by region", "Mean country frequency", ucl_navy
  ),
  mean_frequency_by_income = make_mean_plot(
    income_overall, "income_level", "mean_keyword_count",
    "Mean keyword frequency by income group", "Mean country frequency", ucl_red
  )
)

extraction_log <- documents |>
  dplyr::select(
    country, pdf_file, page_count, report_word_count,
    extraction_status, extraction_warning, extraction_error,
    region, income_level, pdf_md5
  )
run_manifest <- tibble::tibble(
  item = c(
    "Run timestamp", "Keyword workbook", "Keyword workbook MD5",
    "Keyword sheet", "PDF folder", "PDF count", "Canonical term count",
    "Search form count", "Retained occurrence count",
    "Rejected overlapping candidate count", "Counting method",
    "Percentage denominator"
  ),
  value = c(
    format(Sys.time(), "%Y-%m-%d %H:%M:%S %Z"), basename(keyword_file),
    unname(tools::md5sum(keyword_file)), keyword_sheet, pdf_dir,
    as.character(length(pdf_files)), as.character(nrow(keywords)),
    as.character(nrow(variant_table)), as.character(nrow(occurrences)),
    as.character(nrow(overlap_rejections)),
    "Canonical-term counting with global longest-match priority and non-overlapping spans",
    "Keyword occurrences / body-only PDF word count x 100"
  )
)
qc_summary <- tibble::tibble(
  check = c(
    "Body-only PDF count", "Approved canonical term count",
    "Variant-to-canonical collisions", "PDF extraction failures",
    "Metadata matches", "Retained overlapping occurrences",
    "Occurrence-to-summary reconciliation"
  ),
  expected = c(
    as.character(EXPECTED_PDFS), as.character(EXPECTED_CANONICAL_TERMS),
    "0", "0", as.character(EXPECTED_PDFS), "0", "Exact"
  ),
  observed = c(
    as.character(length(pdf_files)), as.character(nrow(keywords)),
    as.character(nrow(variant_collisions)), as.character(nrow(failed_extractions)),
    as.character(sum(!is.na(document_index$country_metadata))),
    as.character(nrow(retained_overlap_check)),
    ifelse(sum(country_term$keyword_count) == nrow(occurrences), "Exact", "Mismatch")
  ),
  status = c(
    ifelse(length(pdf_files) == EXPECTED_PDFS, "PASS", "FAIL"),
    ifelse(nrow(keywords) == EXPECTED_CANONICAL_TERMS, "PASS", "FAIL"),
    ifelse(nrow(variant_collisions) == 0L, "PASS", "FAIL"),
    ifelse(nrow(failed_extractions) == 0L, "PASS", "FAIL"),
    ifelse(sum(!is.na(document_index$country_metadata)) == EXPECTED_PDFS, "PASS", "FAIL"),
    ifelse(nrow(retained_overlap_check) == 0L, "PASS", "FAIL"),
    ifelse(sum(country_term$keyword_count) == nrow(occurrences), "PASS", "FAIL")
  )
)

if (interactive() && OPEN_PREVIEW_TABLES) {
  View(qc_summary)
  View(extraction_log)
  View(global_summary)
  View(level_summary)
  View(mental_health_ranking)
  View(review_queue)
  View(duplicate_files)
  View(country_overall)
  View(country_level)
  View(country_category)
  View(occurrences)
  View(overlap_rejections)
}
invisible(lapply(plots, print))

save_results <- function() {
  if (any(qc_summary$status != "PASS") && STRICT_CHECKS) {
    stop("At least one quality-control check failed. Outputs were not saved.")
  }
  table_dir <- file.path(output_dir, "tables")
  figure_dir <- file.path(output_dir, "figures")
  dir.create(table_dir, recursive = TRUE, showWarnings = FALSE)
  dir.create(figure_dir, recursive = TRUE, showWarnings = FALSE)
  
  workbook_tables <- list(
    global_summary = global_summary,
    level_summary = level_summary,
    mental_health_ranking = mental_health_ranking,
    review_queue = review_queue,
    duplicate_files = duplicate_files,
    run_manifest = run_manifest,
    qc_summary = qc_summary,
    keyword_dictionary = keywords,
    approved_search_forms = variant_table |> dplyr::select(-pattern),
    extraction_log = extraction_log,
    retained_occurrences = occurrences,
    overlap_rejections = overlap_rejections,
    country_term = country_term,
    country_overall = country_overall,
    country_level = country_level,
    country_category = country_category,
    region_overall_mean = region_overall,
    income_overall_mean = income_overall,
    region_level_mean = region_level,
    income_level_mean = income_level,
    region_category_mean = region_category,
    income_category_mean = income_category,
    unmatched_map_countries = unmatched_map_countries
  )
  writexl::write_xlsx(
    workbook_tables,
    file.path(table_dir, "body_only_keyword_analysis_longest_match.xlsx")
  )
  
  csv_tables <- list(
    retained_occurrences = occurrences,
    overlap_rejections = overlap_rejections,
    country_term = country_term,
    country_overall = country_overall,
    country_level = country_level,
    country_category = country_category,
    region_overall_mean = region_overall,
    income_overall_mean = income_overall,
    region_level_mean = region_level,
    income_level_mean = income_level,
    region_category_mean = region_category,
    income_category_mean = income_category
  )
  purrr::iwalk(csv_tables, function(data, table_name) {
    readr::write_csv(data, file.path(table_dir, paste0(table_name, ".csv")))
  })
  
  figure_sizes <- tibble::tribble(
    ~figure_name, ~width, ~height,
    "overall_country_percentage_map", 14, 7.5,
    "overall_country_frequency_map", 14, 7.5,
    "country_percentage_by_level_maps", 13, 10,
    "country_frequency_by_level_maps", 13, 10,
    "country_percentage_by_category_maps", 16, 13,
    "country_frequency_by_category_maps", 16, 13,
    "top20_country_percentage", 8, 7,
    "top20_country_frequency", 8, 7,
    "mean_percentage_by_region", 9, 5.5,
    "mean_percentage_by_income", 8, 4.5,
    "mean_frequency_by_region", 9, 5.5,
    "mean_frequency_by_income", 8, 4.5
  )
  purrr::pwalk(figure_sizes, function(figure_name, width, height) {
    plot <- plots[[figure_name]]
    ggplot2::ggsave(
      file.path(figure_dir, paste0(figure_name, ".png")), plot,
      width = width, height = height, units = "in", dpi = 600, bg = "white"
    )
    ggplot2::ggsave(
      file.path(figure_dir, paste0(figure_name, ".pdf")), plot,
      width = width, height = height, units = "in",
      device = grDevices::pdf, bg = "white"
    )
  })
  message("Saved tables and figures to: ", normalizePath(output_dir))
  invisible(output_dir)
}

if (SAVE_OUTPUTS) {
  save_results()
} else {
  message(
    "Preview complete. Nothing was saved. Review qc_summary, extraction_log, ",
    "retained_occurrences, overlap_rejections, and the RStudio Plots pane."
  )
  message("After checking, run save_results() in the Console. No rerun is required.")
}
save_results()
output_dir
utils::browseURL(normalizePath(output_dir))