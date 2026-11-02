##NOTES FOR METHODS##
# For REGEX bits, I (NWB) used Claude Opus 5.5

# Libraries --------------------------------------------------------------

library(dplyr)
library(readr)
library(stringr)
library(here)
#library(terra)

# Assumptions ------------------------------------------------------------

# Report dates (whole years)
lately_window <- 5L # "lately", "recently"            -> article year minus 5, to article year
some_window <- 10L # "some years ago", "a few years"  -> article year minus 10, to article year
approx_window <- 2L # "about 1855", "?1884", "c. 1858" -> year plus or minus 2

# Quality classes (DRAFT: what are folksthoughts?)
newspaper_specimen_class <- "B" # newspaper "specimen" with no catalogue number: "A" or "B"
heard_only_class <- "C" # heard-only records (diagnostic call descriptions can be overridden to "B" in the sheet)

# I saved the dropbox sharing url to .Renviron to stop it leaking to github.
# quiet = TRUE stops download.file() printing the URL in the console.

tmp <- tempfile(fileext = ".xlsx")

download.file(Sys.getenv("WHEKAU_XLSX_URL"), tmp, mode = "wb", quiet = TRUE)

s1_raw <- readxl::read_excel(
    tmp,
    sheet = "Sheet1",
    col_types = "text",
    na = c("", "NA", "N/A", "na", "n/a")
) |>
    janitor::clean_names()

# Step 1: record IDs (permanent record_id column in the spreadsheet) -----

s1 <- s1_raw |>
    mutate(excel_row = row_number() + 1L) # handy for finding rows in the Excel spreadsheet

stopifnot(
    !anyNA(s1$record_id), # every row has an ID
    !anyDuplicated(s1$record_id) # no ID used twice
)

# Step 2: remove non-records ---------------------------------------------

s1 <- s1 |>
    filter(
        !str_detect(
            coalesce(duplicated, ""),
            regex("example", ignore_case = TRUE)
        )
    ) # removing the kakapo example

# Step 3: coordinates and location error as numbers ----------------------

s1 <- s1 |>
    mutate(
        error_km_raw = error_km,
        x_nztm = suppressWarnings(as.numeric(nztm_east)),
        y_nztm = suppressWarnings(as.numeric(nztm_north)),
        error_km = suppressWarnings(as.numeric(error_km_raw)),
        has_coords = !is.na(x_nztm) & !is.na(y_nztm),
        coords_in_nz = has_coords &
            between(x_nztm, 1.05e6, 2.10e6) & # NZTM Easting range for mainland NZ
            between(y_nztm, 4.70e6, 6.25e6) # NZTM Northing range
    )

count(s1, has_coords, coords_in_nz)

# Step 4: standardise categories -----------------------------------------

s1 <- s1 |>
    mutate(
        island_std = case_when(
            is.na(island) ~ NA_character_,
            str_detect(str_to_lower(island), "^south") ~ "South",
            str_detect(str_to_lower(island), "^north") ~ "North",
            str_detect(str_to_lower(island), "stewart") ~ "Stewart",
            str_detect(str_to_lower(island), "chatham") ~ "Chatham",
            TRUE ~ "CHECK"
        ),

        relevance_std = str_to_upper(str_trim(relevance)),
        is_duplicate = str_to_upper(str_trim(coalesce(duplicated, ""))) %in%
            c("Y", "YES"),

        has_catalogue_no = !is.na(specimen_id), # catalogued, traceable museum specimen

        obs_type_std = {
            t <- str_to_lower(str_trim(observation_type))
            case_when(
                is.na(t) ~ NA_character_,
                str_detect(t, "midden") ~ "midden_bone",
                str_detect(t, "pellet") ~ "pellets",
                str_detect(t, "bone|fossil") ~ "subfossil_bone",
                str_detect(t, "egg") & !str_detect(t, "specimen") ~ "egg",
                str_detect(t, "specimen|skin|mounted") ~ "specimen",
                str_detect(
                    t,
                    "killed|captured|dead|remains"
                ) ~ "killed_or_dead",
                str_detect(t, "nest") ~ "nest",
                str_detect(t, "seen|sighting") ~ "seen",
                str_detect(t, "heard|call") ~ "heard",
                str_detect(
                    t,
                    "newspaper|report|obituary"
                ) ~ "report_unspecified",
                str_detect(t, "commentary") ~ "commentary",
                str_detect(t, "search") ~ "search_no_find",
                TRUE ~ "other"
            )
        },
        # split specimens: catalogued in a museum vs reported in print
        obs_type_std = case_when(
            obs_type_std == "specimen" &
                has_catalogue_no ~ "specimen_catalogued",
            obs_type_std == "specimen" &
                !has_catalogue_no ~ "specimen_reported",
            TRUE ~ obs_type_std
        ),

        obs_uncertain = str_detect(
            coalesce(observation_type, ""),
            "\\?|potential"
        )
    )

# rows whose Location names a museum, and whether they carry coordinates
museum_loc <- s1 |>
    filter(str_detect(
        coalesce(location, ""),
        regex("museum", ignore_case = TRUE)
    ))

museum_loc |> select(record_id, location, has_coords, x_nztm, y_nztm)

# do any other rows share exactly those coordinates? (a sign they were placed at the museum too)
s1 |>
    filter(has_coords) |>
    semi_join(filter(museum_loc, has_coords), by = c("x_nztm", "y_nztm")) |>
    select(record_id, location, x_nztm, y_nztm)

s1 |> filter(has_coords) |> count(x_nztm, y_nztm, sort = TRUE) |> head(15)


# Step 5a: article year and tidy dates ------------------------

# I used claude to do the regex: \\d{4} means four digits. ^ and $ mean the start and end of the cell,
# so "^\\d{4}$" matches a cell that holds only a year. str_extract() returns NA when nothing matches,
# so cells like "N.D." become NA without errors. case_when() evaluates every right-hand side for every
# row, so each one uses str_extract() (never a bare as.integer()) to avoid "NAs introduced by coercion".

s1 <- s1 |>
    mutate(
        article_year = as.integer(str_extract(article_date, "\\d{4}")), # first 4-digit number; "N.D." gives NA

        tidy = str_trim(observation_date_tidy),
        tidy_min = case_when(
            str_detect(tidy, "^\\d{4}$") ~ as.integer(str_extract(
                tidy,
                "\\d{4}"
            )), # 1903
            str_detect(
                tidy,
                "^\\d{4}\\s*-\\s*\\d{4}$"
            ) ~ as.integer(str_extract(tidy, "^\\d{4}")), # 1903-1904
            str_detect(tidy, "^>\\s*\\d{4}$") ~ as.integer(str_extract(
                tidy,
                "\\d{4}"
            )), # >1854 = after 1854
            TRUE ~ NA_integer_
        ),
        tidy_max = case_when(
            str_detect(tidy, "^\\d{4}$") ~ as.integer(str_extract(
                tidy,
                "\\d{4}"
            )),
            str_detect(
                tidy,
                "^\\d{4}\\s*-\\s*\\d{4}$"
            ) ~ as.integer(str_extract(tidy, "\\d{4}$")),
            str_detect(tidy, "^<\\s*\\d{4}$") ~ as.integer(str_extract(
                tidy,
                "\\d{4}"
            )), # <1870 = before 1870
            TRUE ~ NA_integer_
        )
    )


# Step 5b: free-text observation dates -----------------------------------

# Stage 1: label which pattern each free-text date matches

s1 <- s1 |>
    mutate(
        t = str_to_lower(str_squish(observation_date)),
        y4 = as.integer(str_extract(t, "\\d{4}")), # first 4-digit year
        n_yrs = str_count(coalesce(t, ""), "\\d{4}"), # how many 4-digit years
        ago_a = as.integer(str_match(
            t,
            "(\\d+)\\s*(?:-|to)?\\s*(\\d+)?\\s*years?\\s*(?:ago|since)"
        )[, 2]),
        ago_b = as.integer(str_match(
            t,
            "(\\d+)\\s*(?:-|to)?\\s*(\\d+)?\\s*years?\\s*(?:ago|since)"
        )[, 3]),
        decade = as.integer(str_match(t, "(\\d{3})0s")[, 2]) * 10L,

        date_rule = case_when(
            is.na(t) ~ NA_character_,
            str_detect(t, "^\\d{4}\\s*-\\s*\\d{1,2}$") ~ "short_range", # 1903-4
            n_yrs >= 2 ~ "multi_year", # 1896 ... 1898
            str_detect(t, "late\\s+\\d{3}0s") ~ "late_decade", # late 1880s
            str_detect(t, "early\\s+\\d{3}0s") ~ "early_decade",
            !is.na(decade) ~ "decade", # 1880s
            str_detect(t, "pre|before") & !is.na(y4) ~ "before", # Pre-1870
            str_detect(t, "post|after") & !is.na(y4) ~ "after", # post-1854
            !is.na(y4) &
                str_detect(t, "\\?|about|around|circa|c\\.") ~ "approx_year",
            !is.na(y4) ~ "exact_year",
            !is.na(ago_a) ~ "years_ago", # 18 - 20 years ago
            str_detect(t, "lately|recent") ~ "lately",
            str_detect(t, "some years|few years") ~ "some_years",
            str_detect(t, "pre-european") ~ "pre_european",
            TRUE ~ "vague"
        ) # "many years ago", "?"
    )

# Stage 2: convert each pattern to an earliest and latest year

s1 <- s1 |>
    mutate(
        multi_min = sapply(str_extract_all(coalesce(t, ""), "\\d{4}"), \(v) {
            if (length(v)) min(as.integer(v)) else NA_integer_
        }),
        multi_max = sapply(str_extract_all(coalesce(t, ""), "\\d{4}"), \(v) {
            if (length(v)) max(as.integer(v)) else NA_integer_
        }),
        end_yr = as.integer(str_extract(t, "\\d{1,2}$")),
        about = str_detect(coalesce(t, ""), "about|some|around"), # widen "about 30 years ago" by approx_window

        text_min = case_when(
            date_rule == "short_range" ~ y4,
            date_rule == "multi_year" ~ multi_min,
            date_rule == "late_decade" ~ decade + 5L,
            date_rule == "early_decade" ~ decade,
            date_rule == "decade" ~ decade,
            date_rule == "after" ~ y4,
            date_rule == "approx_year" ~ y4 - approx_window,
            date_rule == "exact_year" ~ y4,
            date_rule == "years_ago" ~ article_year -
                pmax(ago_a, ago_b, na.rm = TRUE) -
                if_else(about, approx_window, 0L),
            date_rule == "lately" ~ article_year - lately_window,
            date_rule == "some_years" ~ article_year - some_window,
            TRUE ~ NA_integer_
        ),
        text_max = case_when(
            date_rule == "short_range" ~ y4 -
                y4 %% as.integer(10^nchar(end_yr)) +
                end_yr, # 1903-4 -> 1904 (^ gives a decimal, so convert back to a whole number)
            date_rule == "multi_year" ~ multi_max,
            date_rule == "late_decade" ~ decade + 9L,
            date_rule == "early_decade" ~ decade + 4L,
            date_rule == "decade" ~ decade + 9L,
            date_rule == "before" ~ y4,
            date_rule == "approx_year" ~ y4 + approx_window,
            date_rule == "exact_year" ~ y4,
            date_rule == "years_ago" ~ article_year -
                pmin(ago_a, ago_b, na.rm = TRUE) +
                if_else(about, approx_window, 0L),
            date_rule %in%
                c("lately", "some_years", "vague", "after") ~ article_year,
            TRUE ~ NA_integer_
        ),

        # final dates: the tidy column wins, then free text, then "no later than the article"
        obs_year_min = coalesce(
            tidy_min,
            if_else(is.na(tidy_max), text_min, NA_integer_)
        ),
        obs_year_max = coalesce(
            tidy_max,
            if_else(is.na(tidy_min), text_max, NA_integer_),
            article_year
        ),
        date_source = case_when(
            !is.na(tidy_min) | !is.na(tidy_max) ~ "tidy column",
            !is.na(date_rule) ~ paste("free text:", date_rule),
            !is.na(article_year) ~ "article year only",
            TRUE ~ "none"
        )
    )

# Bone records (subfossil and midden) get no calendar year: any year attached to them is when
# they were found or reported, not when the bird lived. article_year is kept as the report date.

s1 <- s1 |>
    mutate(
        bone_record = obs_type_std %in% c("subfossil_bone", "midden_bone"),
        obs_year_min = if_else(bone_record, NA, obs_year_min),
        obs_year_max = if_else(bone_record, NA, obs_year_max),
        date_source = if_else(
            bone_record,
            "bone record: no calendar date",
            date_source
        )
    )

# Step 6: evidence period, captive flag and quality class ----------------

if (!"quality_class" %in% names(s1)) {
    s1$quality_class <- NA_character_
} # works before the column exists

s1 <- s1 |>
    mutate(
        evidence_period = case_when(
            bone_record ~ "pre_european",
            obs_type_std %in%
                c(
                    "seen",
                    "heard",
                    "nest",
                    "killed_or_dead",
                    "egg",
                    "pellets",
                    "specimen_catalogued",
                    "specimen_reported",
                    "report_unspecified"
                ) ~ "historical",
            TRUE ~ NA_character_ # commentary, searches, other: not observations
        ),

        captive = str_detect(
            coalesce(notes, ""),
            regex("captiv", ignore_case = TRUE)
        ),

        quality_auto = case_when(
            relevance_std != "Y" | is.na(relevance_std) ~ "D",
            is.na(evidence_period) ~ "D",
            has_catalogue_no ~ "A", # catalogued skins, eggs, bones
            obs_type_std == "specimen_reported" ~ newspaper_specimen_class,
            obs_uncertain ~ "C", # "Seen?", "potential sighting"
            obs_type_std == "heard" ~ heard_only_class,
            obs_type_std == "report_unspecified" ~ "C",
            TRUE ~ "B" # seen, nest, killed, eggs, pellets, uncatalogued bones
        ),

        quality = coalesce(str_to_upper(str_trim(quality_class)), quality_auto)
    )

# Step 7: write the cleaned table for modelling --------------------------

s1_clean <- s1 |>
    select(
        record_id,
        excel_row,
        source,
        location,
        region,
        island_std,
        x_nztm,
        y_nztm,
        error_km,
        geographic_certainty,
        has_coords,
        coords_in_nz,
        obs_type_std,
        obs_uncertain,
        has_catalogue_no,
        specimen_id,
        museum,
        article_year,
        obs_year_min,
        obs_year_max,
        date_source,
        relevance_std,
        is_duplicate,
        duplicate_link,
        captive,
        evidence_period,
        quality_auto,
        quality_class,
        quality,
        notes,
        url
    )

write_csv(s1_clean, here::here("data", "whekau_records_clean.csv"), na = "")
