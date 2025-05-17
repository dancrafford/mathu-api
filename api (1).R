rm(list=ls())

# Install required packages if not already installed
if (!requireNamespace("readxl", quietly = TRUE)) install.packages("readxl")
if (!requireNamespace("dplyr", quietly = TRUE)) install.packages("dplyr")
if (!requireNamespace("tidyr", quietly = TRUE)) install.packages("tidyr")
if (!requireNamespace("writexl", quietly = TRUE)) install.packages("writexl")
if (!requireNamespace("openxlsx", quietly = TRUE)) install.packages("openxlsx")
if (!requireNamespace("plumber", quietly = TRUE)) install.packages("plumber")

# Load libraries
library(readxl)
library(dplyr)
library(tidyr)
library(writexl)
library(openxlsx)  # Using openxlsx for better control over formatting
library(plumber)

r <- plumb("api.R")
r$run(port = 8000)





# Function to process a single workbook
process_workbook <- function(file_path) {
  # Extract workbook name (without extension) to use for naming
  workbook_name <- tools::file_path_sans_ext(basename(file_path))
  
  # Extract Grade and Subject from workbook name
  # Assuming format like "Grade 10 Maths" or "Grade 9 Science"
  grade <- NA
  subject <- NA
  
  # Try to extract grade (looking for patterns like "Grade 10" or "Grade9")
  grade_match <- regexpr("Grade\\s*\\d+", workbook_name, ignore.case = TRUE)
  if (grade_match > 0) {
    grade_text <- regmatches(workbook_name, grade_match)
    # Extract just the number
    grade <- gsub("Grade\\s*", "", grade_text, ignore.case = TRUE)
  }
  
  # Try to extract subject (after finding grade)
  if (grade_match > 0) {
    # Get everything after the grade
    subject_text <- substr(workbook_name, 
                           grade_match + attr(grade_match, "match.length"), 
                           nchar(workbook_name))
    # Clean up leading/trailing spaces and common words
    subject <- gsub("^\\s+|\\s+$", "", subject_text)
    if (subject == "") {
      subject <- NA
    }
  }
  
  cat("Extracted Grade:", grade, "and Subject:", subject, "from filename\n")
  
  # Get all sheet names in the workbook
  sheet_names <- excel_sheets(file_path)
  
  # Initialize an empty list to store processed sheets
  processed_sheets <- list()
  
  # Read and process each sheet
  for (i in seq_along(sheet_names)) {
    sheet_name <- sheet_names[i]
    
    # Read the sheet data
    sheet_data <- read_excel(file_path, sheet = sheet_name)
    
    # Print column names to debug
    cat("Sheet:", sheet_name, "- Columns:", paste(colnames(sheet_data), collapse=", "), "\n")
    
    # Check if we have the expected columns
    if (!"Name" %in% colnames(sheet_data) || !"Surname" %in% colnames(sheet_data)) {
      cat("Error: Expected columns 'Name' and 'Surname' not found in sheet", sheet_name, "\n")
      next
    }
    
    # Remove summary rows (Totals, Averages, Homework Completion %)
    sheet_data <- sheet_data %>%
      filter(!grepl("^(Totals:|Averages:|Homework Completion %:)", sheet_data[[1]]))
    
    # Combine Name and Surname into a single column
    sheet_data <- sheet_data %>%
      mutate(Student = paste(Name, Surname)) %>%
      # Add Grade and Subject columns
      mutate(Grade = grade, Subject = subject) %>%
      # Reorder columns: Student, Grade, Subject, then everything else except Name/Surname
      select(Student, Grade, Subject, everything(), -Name, -Surname)
    
    # Store the processed sheet
    processed_sheets[[i]] <- sheet_data
  }
  
  # If no sheets were successfully processed, return NULL
  if (length(processed_sheets) == 0) {
    cat("No sheets were successfully processed for", workbook_name, "\n")
    return(NULL)
  }
  
  # Combine all sheets horizontally with Student column as the key
  master_data <- processed_sheets[[1]]
  
  if (length(processed_sheets) > 1) {
    for (i in 2:length(processed_sheets)) {
      # Get the current sheet
      current_sheet <- processed_sheets[[i]]
      
      # For merging, we need to handle duplicate column names
      # Get all column names except Student, Grade, and Subject
      current_cols <- setdiff(colnames(current_sheet), c("Student", "Grade", "Subject"))
      
      # Rename columns in current_sheet to avoid duplicates
      # We'll add sheet number as suffix only if there are duplicate columns
      for (col in current_cols) {
        if (col %in% colnames(master_data)) {
          # If duplicate, rename with sheet index
          new_col <- paste0(col, "_Sheet", i)
          colnames(current_sheet)[colnames(current_sheet) == col] <- new_col
        }
      }
      
      # Add a blank separator column to master_data
      separator_name <- paste0("Separator_", i-1)
      master_data[[separator_name]] <- NA
      
      # Merge with the master data
      master_data <- master_data %>%
        left_join(current_sheet %>% select(Student), by = "Student") %>%
        bind_cols(current_sheet %>% select(-Student, -Grade, -Subject))
    }
  }
  
  # Write the master sheet to a new Excel file
  output_file <- paste0(workbook_name, "_Master.xlsx")
  write_xlsx(master_data, output_file)
  
  cat("Created master sheet for", workbook_name, "at", output_file, "\n")
  
  return(master_data)
}

# Function to scale percentage values if needed
scale_percentage <- function(x) {
  if (is.numeric(x)) {
    # Check if values are in decimal format (0-1) or percentage (0-100)
    if (all(x <= 1, na.rm = TRUE) && any(x > 0, na.rm = TRUE)) {
      # If all values are <= 1 and at least one is > 0, assume decimal format (0-1)
      return(x * 100)
    } else {
      # Otherwise assume it's already in percentage format (0-100)
      return(x)
    }
  }
  return(x)
}

# Function to create a master of masters sheet by aggregating metrics
create_master_of_masters <- function(all_masters) {
  # Install required package if needed
  if (!requireNamespace("dplyr", quietly = TRUE) || 
      packageVersion("dplyr") < "1.1.0") {
    cat("Note: Using summarize() for aggregation\n")
    use_reframe <- FALSE
  } else {
    cat("Note: Using reframe() for aggregation\n")
    use_reframe <- TRUE
  }
  
  # Initialize an empty dataframe to store the master of masters
  master_of_masters <- data.frame()
  
  # Process each master sheet
  for (master_name in names(all_masters)) {
    master_data <- all_masters[[master_name]]
    
    # Skip if NULL
    if (is.null(master_data)) {
      next
    }
    
    # Get key columns (Student, Grade, Subject)
    key_cols <- master_data %>%
      select(Student, Grade, Subject)
    
    # Get metric columns (skip key columns and separator columns)
    metric_cols <- master_data %>%
      select(-Student, -Grade, -Subject, -matches("^Separator_"))
    
    # Group metric columns by their base names (without sheet numbers)
    # e.g., "Homework.Progress" and "Homework.Progress_Sheet2" should be grouped
    col_names <- colnames(metric_cols)
    base_names <- gsub("_Sheet\\d+$", "", col_names)
    
    # Create a list to store aggregated metrics
    aggregated_metrics <- list()
    
    # Process each unique base column name
    for (base_name in unique(base_names)) {
      # Get all columns with this base name
      related_cols <- col_names[base_names == base_name]
      
      # Convert percentage columns to numeric if needed
      for (col in related_cols) {
        if (is.character(metric_cols[[col]])) {
          # If it's a percentage, convert to numeric
          if (any(grepl("%", metric_cols[[col]], fixed = TRUE))) {
            metric_cols[[col]] <- as.numeric(gsub("%", "", metric_cols[[col]], fixed = TRUE))
          }
        }
      }
      
      # Check if this is a percentage-based column
      is_percentage <- grepl("Progress|[Ss]core", base_name, ignore.case = TRUE)
      
      # Handle numeric columns
      if (all(sapply(metric_cols[related_cols], is.numeric))) {
        if (is_percentage) {
          # For percentage columns, calculate mean (ignoring NA values)
          aggregated_col <- rowMeans(metric_cols[related_cols], na.rm = TRUE)
        } else {
          # For other numeric columns, calculate the sum
          aggregated_col <- rowSums(metric_cols[related_cols], na.rm = TRUE)
        }
      } else {
        # For non-numeric columns, use the first non-NA value
        aggregated_col <- apply(metric_cols[related_cols], 1, function(x) {
          non_na <- x[!is.na(x)]
          if (length(non_na) > 0) non_na[1] else NA
        })
      }
      
      # Store the aggregated values
      aggregated_metrics[[base_name]] <- aggregated_col
    }
    
    # Combine key columns with aggregated metrics
    student_aggregated <- key_cols %>%
      bind_cols(as.data.frame(aggregated_metrics))
    
    # Add to master of masters
    if (nrow(master_of_masters) == 0) {
      master_of_masters <- student_aggregated
    } else {
      # Ensure columns match by adding missing columns with NA values
      for (col in setdiff(colnames(student_aggregated), colnames(master_of_masters))) {
        master_of_masters[[col]] <- NA
      }
      for (col in setdiff(colnames(master_of_masters), colnames(student_aggregated))) {
        student_aggregated[[col]] <- NA
      }
      
      # Bind rows
      master_of_masters <- bind_rows(master_of_masters, student_aggregated)
    }
  }
  
  # Print column names to help with debugging
  cat("Columns in master_of_masters:\n")
  for (col in colnames(master_of_masters)) {
    cat("  -", col, "\n")
  }
  
  # Remove Homework Progress column if it exists
  if ("Homework.Progress" %in% colnames(master_of_masters)) {
    cat("Removing Homework.Progress column...\n")
    master_of_masters <- master_of_masters %>% select(-Homework.Progress)
  }
  
  # Use the actual column names from the data with dots
  concepts_completed_col <- "Concepts.Completed"
  concepts_assigned_col <- "Concepts.Assigned"
  exercises_completed_col <- "Exercises.Completed"
  exercises_assigned_col <- "Exercises.Assigned"
  assessments_completed_col <- "Assessments.Completed"
  assessments_assigned_col <- "Assessments.Assigned"
  assessment_score_col <- "Assessment.Score"
  
  # Check if these columns exist in the data
  required_cols <- c(concepts_completed_col, concepts_assigned_col,
                     exercises_completed_col, exercises_assigned_col,
                     assessments_completed_col, assessments_assigned_col,
                     assessment_score_col)
  
  missing_cols <- setdiff(required_cols, colnames(master_of_masters))
  
  if (length(missing_cols) > 0) {
    cat("WARNING: The following columns are missing:", paste(missing_cols, collapse=", "), "\n")
    cat("Looking for alternative column names...\n")
    
    # Try finding similar columns using pattern matching
    if (concepts_completed_col %in% missing_cols) {
      alt_col <- grep("Concepts.*Completed", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", concepts_completed_col, "\n")
        concepts_completed_col <- alt_col[1]
      }
    }
    
    if (concepts_assigned_col %in% missing_cols) {
      alt_col <- grep("Concepts.*Assigned", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", concepts_assigned_col, "\n")
        concepts_assigned_col <- alt_col[1]
      }
    }
    
    if (exercises_completed_col %in% missing_cols) {
      alt_col <- grep("Exercises.*Completed", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", exercises_completed_col, "\n")
        exercises_completed_col <- alt_col[1]
      }
    }
    
    if (exercises_assigned_col %in% missing_cols) {
      alt_col <- grep("Exercises.*Assigned", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", exercises_assigned_col, "\n")
        exercises_assigned_col <- alt_col[1]
      }
    }
    
    if (assessments_completed_col %in% missing_cols) {
      alt_col <- grep("Assessments.*Completed", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", assessments_completed_col, "\n")
        assessments_completed_col <- alt_col[1]
      }
    }
    
    if (assessments_assigned_col %in% missing_cols) {
      alt_col <- grep("Assessments.*Assigned", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", assessments_assigned_col, "\n")
        assessments_assigned_col <- alt_col[1]
      }
    }
    
    if (assessment_score_col %in% missing_cols) {
      alt_col <- grep("[Aa]ssessment.*[Ss]core", colnames(master_of_masters), value = TRUE)
      if (length(alt_col) > 0) {
        cat("Using", alt_col[1], "instead of", assessment_score_col, "\n")
        assessment_score_col <- alt_col[1]
      }
    }
  }
  
  # Check if we have the required columns after trying alternatives
  have_all_columns <- all(c(concepts_completed_col, concepts_assigned_col,
                            exercises_completed_col, exercises_assigned_col,
                            assessments_completed_col, assessments_assigned_col,
                            assessment_score_col) %in% 
                            colnames(master_of_masters))
  
  if (have_all_columns) {
    cat("All required columns found. Proceeding with calculations...\n")
    
    # Add completion percentage calculations
    master_of_masters$`Concept.Completion.%` <- ifelse(
      master_of_masters[[concepts_assigned_col]] > 0,
      100 * master_of_masters[[concepts_completed_col]] / master_of_masters[[concepts_assigned_col]],
      NA
    )
    
    master_of_masters$`Exercise.Completion.%` <- ifelse(
      master_of_masters[[exercises_assigned_col]] > 0,
      100 * master_of_masters[[exercises_completed_col]] / master_of_masters[[exercises_assigned_col]],
      NA
    )
    
    master_of_masters$`Assessment.Completion.%` <- ifelse(
      master_of_masters[[assessments_assigned_col]] > 0,
      100 * master_of_masters[[assessments_completed_col]] / master_of_masters[[assessments_assigned_col]],
      NA
    )
    
    # Calculate overall homework completion
    master_of_masters$`Overall.Homework.Completion.%` <- rowMeans(
      master_of_masters[, c("Concept.Completion.%", "Exercise.Completion.%", "Assessment.Completion.%")],
      na.rm = TRUE
    )
    
    # Round all percentage columns to whole numbers
    master_of_masters$`Concept.Completion.%` <- round(master_of_masters$`Concept.Completion.%`)
    master_of_masters$`Exercise.Completion.%` <- round(master_of_masters$`Exercise.Completion.%`)
    master_of_masters$`Assessment.Completion.%` <- round(master_of_masters$`Assessment.Completion.%`)
    master_of_masters$`Overall.Homework.Completion.%` <- round(master_of_masters$`Overall.Homework.Completion.%`)
    
    # Scale and round Assessment.Score if it exists
    if (assessment_score_col %in% colnames(master_of_masters)) {
      cat("Scaling Assessment Score if needed...\n")
      # Print range to debug
      cat("Assessment Score range:", min(master_of_masters[[assessment_score_col]], na.rm = TRUE), 
          "to", max(master_of_masters[[assessment_score_col]], na.rm = TRUE), "\n")
      
      master_of_masters[[assessment_score_col]] <- scale_percentage(master_of_masters[[assessment_score_col]])
      master_of_masters[[assessment_score_col]] <- round(master_of_masters[[assessment_score_col]])
    }
    
    # Function to create grade summaries - using sums for metrics
    create_grade_summary <- function(data) {
      if (use_reframe) {
        # Use reframe for dplyr >= 1.1.0
        grade_summaries <- data %>%
          group_by(Grade) %>%
          reframe(
            Student = paste0("Grade ", Grade, " Average"),
            Subject = "All",
            # Sum the count columns
            !!sym(concepts_completed_col) := sum(!!sym(concepts_completed_col), na.rm = TRUE),
            !!sym(concepts_assigned_col) := sum(!!sym(concepts_assigned_col), na.rm = TRUE),
            !!sym(exercises_completed_col) := sum(!!sym(exercises_completed_col), na.rm = TRUE),
            !!sym(exercises_assigned_col) := sum(!!sym(exercises_assigned_col), na.rm = TRUE),
            !!sym(assessments_completed_col) := sum(!!sym(assessments_completed_col), na.rm = TRUE),
            !!sym(assessments_assigned_col) := sum(!!sym(assessments_assigned_col), na.rm = TRUE),
            # Average the assessment score - with scaling for 0-1 values
            !!sym(assessment_score_col) := {
              scores <- scale_percentage(!!sym(assessment_score_col))
              round(mean(scores, na.rm = TRUE))
            }
          ) %>%
          distinct()
      } else {
        # Use summarize for older dplyr versions
        grade_summaries <- data %>%
          group_by(Grade) %>%
          summarize(
            Student = first(paste0("Grade ", Grade, " Average")),
            Subject = "All",
            # Sum the count columns
            !!sym(concepts_completed_col) := sum(!!sym(concepts_completed_col), na.rm = TRUE),
            !!sym(concepts_assigned_col) := sum(!!sym(concepts_assigned_col), na.rm = TRUE),
            !!sym(exercises_completed_col) := sum(!!sym(exercises_completed_col), na.rm = TRUE),
            !!sym(exercises_assigned_col) := sum(!!sym(exercises_assigned_col), na.rm = TRUE),
            !!sym(assessments_completed_col) := sum(!!sym(assessments_completed_col), na.rm = TRUE),
            !!sym(assessments_assigned_col) := sum(!!sym(assessments_assigned_col), na.rm = TRUE),
            # Average the assessment score - with scaling for 0-1 values
            !!sym(assessment_score_col) := {
              scores <- scale_percentage(!!sym(assessment_score_col))
              round(mean(scores, na.rm = TRUE))
            },
            .groups = "drop"
          ) %>%
          distinct()
      }
      
      # Calculate completion percentages based on the sums
      grade_summaries$`Concept.Completion.%` <- ifelse(
        grade_summaries[[concepts_assigned_col]] > 0,
        round(100 * grade_summaries[[concepts_completed_col]] / grade_summaries[[concepts_assigned_col]]),
        NA
      )
      
      grade_summaries$`Exercise.Completion.%` <- ifelse(
        grade_summaries[[exercises_assigned_col]] > 0,
        round(100 * grade_summaries[[exercises_completed_col]] / grade_summaries[[exercises_assigned_col]]),
        NA
      )
      
      grade_summaries$`Assessment.Completion.%` <- ifelse(
        grade_summaries[[assessments_assigned_col]] > 0,
        round(100 * grade_summaries[[assessments_completed_col]] / grade_summaries[[assessments_assigned_col]]),
        NA
      )
      
      # Calculate overall homework completion
      grade_summaries$`Overall.Homework.Completion.%` <- round(rowMeans(
        grade_summaries[, c("Concept.Completion.%", "Exercise.Completion.%", "Assessment.Completion.%")],
        na.rm = TRUE
      ))
      
      return(grade_summaries)
    }
    
    # Function to create subject summaries - using sums for metrics
    create_subject_summary <- function(data) {
      if (use_reframe) {
        # Use reframe for dplyr >= 1.1.0
        subject_summaries <- data %>%
          filter(Grade %in% c("10", "11", "12")) %>%
          group_by(Grade, Subject) %>%
          reframe(
            Student = paste0("Grade ", Grade, " ", Subject, " Average"),
            # Sum the count columns
            !!sym(concepts_completed_col) := sum(!!sym(concepts_completed_col), na.rm = TRUE),
            !!sym(concepts_assigned_col) := sum(!!sym(concepts_assigned_col), na.rm = TRUE),
            !!sym(exercises_completed_col) := sum(!!sym(exercises_completed_col), na.rm = TRUE),
            !!sym(exercises_assigned_col) := sum(!!sym(exercises_assigned_col), na.rm = TRUE),
            !!sym(assessments_completed_col) := sum(!!sym(assessments_completed_col), na.rm = TRUE),
            !!sym(assessments_assigned_col) := sum(!!sym(assessments_assigned_col), na.rm = TRUE),
            # Average the assessment score - with scaling for 0-1 values
            !!sym(assessment_score_col) := {
              scores <- scale_percentage(!!sym(assessment_score_col))
              round(mean(scores, na.rm = TRUE))
            }
          ) %>%
          distinct()
      } else {
        # Use summarize for older dplyr versions
        subject_summaries <- data %>%
          filter(Grade %in% c("10", "11", "12")) %>%
          group_by(Grade, Subject) %>%
          summarize(
            Student = first(paste0("Grade ", Grade, " ", Subject, " Average")),
            # Sum the count columns
            !!sym(concepts_completed_col) := sum(!!sym(concepts_completed_col), na.rm = TRUE),
            !!sym(concepts_assigned_col) := sum(!!sym(concepts_assigned_col), na.rm = TRUE),
            !!sym(exercises_completed_col) := sum(!!sym(exercises_completed_col), na.rm = TRUE),
            !!sym(exercises_assigned_col) := sum(!!sym(exercises_assigned_col), na.rm = TRUE),
            !!sym(assessments_completed_col) := sum(!!sym(assessments_completed_col), na.rm = TRUE),
            !!sym(assessments_assigned_col) := sum(!!sym(assessments_assigned_col), na.rm = TRUE),
            # Average the assessment score - with scaling for 0-1 values
            !!sym(assessment_score_col) := {
              scores <- scale_percentage(!!sym(assessment_score_col))
              round(mean(scores, na.rm = TRUE))
            },
            .groups = "drop"
          ) %>%
          distinct()
      }
      
      # Calculate completion percentages based on the sums
      subject_summaries$`Concept.Completion.%` <- ifelse(
        subject_summaries[[concepts_assigned_col]] > 0,
        round(100 * subject_summaries[[concepts_completed_col]] / subject_summaries[[concepts_assigned_col]]),
        NA
      )
      
      subject_summaries$`Exercise.Completion.%` <- ifelse(
        subject_summaries[[exercises_assigned_col]] > 0,
        round(100 * subject_summaries[[exercises_completed_col]] / subject_summaries[[exercises_assigned_col]]),
        NA
      )
      
      subject_summaries$`Assessment.Completion.%` <- ifelse(
        subject_summaries[[assessments_assigned_col]] > 0,
        round(100 * subject_summaries[[assessments_completed_col]] / subject_summaries[[assessments_assigned_col]]),
        NA
      )
      
      # Calculate overall homework completion
      subject_summaries$`Overall.Homework.Completion.%` <- round(rowMeans(
        subject_summaries[, c("Concept.Completion.%", "Exercise.Completion.%", "Assessment.Completion.%")],
        na.rm = TRUE
      ))
      
      return(subject_summaries)
    }
    
    # Create summary rows
    cat("Creating grade summary statistics...\n")
    grade_summary <- create_grade_summary(master_of_masters)
    
    cat("Creating grade-subject summary statistics for grades 10-12...\n")
    subject_summary <- create_subject_summary(master_of_masters)
    
    # Add empty row as a separator
    cat("Adding separator row...\n")
    empty_row <- master_of_masters[1, ]
    empty_row[] <- NA
    empty_row$Student <- "--- SUMMARY STATISTICS ---"
    
    # Combine everything, ensuring we don't have duplicate summary rows
    cat("Combining student data with summaries...\n")
    final_master <- bind_rows(
      master_of_masters,
      empty_row,
      grade_summary,
      subject_summary
    )
    
    return(final_master)
  } else {
    cat("WARNING: Some required columns are still missing. Cannot create all summary statistics.\n")
    # Return the data with any partial calculations we were able to perform
    return(master_of_masters)
  }
}

# Main function to process all workbooks in the current working directory
process_all_workbooks <- function() {
  # Get the current working directory
  current_dir <- getwd()
  cat("Using current working directory:", current_dir, "\n")
  
  # Get all Excel files in the current directory
  excel_files <- list.files(path = current_dir, 
                            pattern = "\\.(xlsx|xls)$", 
                            full.names = TRUE)
  
  # Remove any existing master files from the list
  excel_files <- excel_files[!grepl("_Master\\.", excel_files)]
  
  if (length(excel_files) == 0) {
    cat("No Excel files found in the current directory.\n")
    return(NULL)
  }
  
  cat("Found", length(excel_files), "Excel file(s) to process.\n")
  
  # Process each workbook
  results <- list()
  for (file in excel_files) {
    cat("Processing", basename(file), "...\n")
    results[[basename(file)]] <- process_workbook(file)
  }
  
  cat("All workbooks processed successfully!\n")
  
  # Create master of masters sheet
  cat("Creating master of masters sheet with additional metrics...\n")
  master_of_masters <- create_master_of_masters(results)
  
  # Create a new workbook with all master sheets and the master of masters
  all_masters_workbook <- list()
  all_masters_workbook[["Master of Masters"]] <- master_of_masters
  
  # Add each individual master sheet
  for (name in names(results)) {
    if (!is.null(results[[name]])) {
      sheet_name <- tools::file_path_sans_ext(name)
      # Truncate sheet name if too long (Excel has a 31 character limit for sheet names)
      if (nchar(sheet_name) > 30) {
        sheet_name <- substr(sheet_name, 1, 30)
      }
      all_masters_workbook[[sheet_name]] <- results[[name]]
    }
  }
  
  # Write to a single Excel file
  write_xlsx(all_masters_workbook, "All_Masters.xlsx")
  cat("Created consolidated workbook with all masters at All_Masters.xlsx\n")
  
  return(all_masters_workbook)
}

# Function to create a transposed master report
# Function to create a transposed master report
create_transposed_master_report <- function(input_file = "All_Masters.xlsx", 
                                            output_file = "Transposed_Report.xlsx",
                                            date_range = "Term 1 2025",  
                                            export_individual_sheets = TRUE) {
  
  # Read the Master of Masters sheet
  cat("Reading Master of Masters sheet from", input_file, "...\n")
  master_data <- read_excel(input_file, sheet = "Master of Masters")
  
  # Debug: Print unique grades and subjects
  cat("Unique grades in data:", paste(unique(master_data$Grade), collapse=", "), "\n")
  cat("Unique subjects in data:", paste(unique(master_data$Subject), collapse=", "), "\n")
  
  # Filter out summary rows (those starting with "---" or "Grade")
  student_data <- master_data %>%
    filter(!grepl("^---", Student) & !grepl("^Grade", Student))
  
  # Handle grades with no subjects separately (Grade 8, Grade 9, etc.)
  # First, get all grades in the data
  all_grades <- unique(student_data$Grade)
  
  # Initialize Excel workbook using openxlsx
  wb <- createWorkbook()
  
  # Calculate total UNIQUE students per grade - MODIFIED
  total_by_grade <- student_data %>%
    select(Grade, Student) %>%  # Select only Grade and Student columns
    distinct() %>%  # Keep only unique Grade-Student combinations
    group_by(Grade) %>%
    summarise(Count = n()) %>%  # Count unique students per grade
    arrange(Grade)
  
  # Calculate total students per grade and subject combination
  total_by_grade_subject <- student_data %>%
    group_by(Grade, Subject) %>%
    summarise(Count = n()) %>%
    arrange(Grade, Subject)
  
  # Create summary sheet
  summary_sheet <- data.frame(
    Category = c("Date Range", 
                 paste0("Total Students Grade ", total_by_grade$Grade),
                 "",
                 paste0("Total Students Grade ", total_by_grade_subject$Grade, " ", total_by_grade_subject$Subject)),
    Count = c(date_range,
              as.character(total_by_grade$Count),
              "",
              as.character(total_by_grade_subject$Count))
  )
  
  addWorksheet(wb, "Summary")
  writeData(wb, "Summary", summary_sheet)
  
  # Calculate ranks for each grade - MODIFIED to use highest points per student
  # First create a temporary dataframe with unique Student-Grade combinations and their MAXIMUM points
  grade_ranks <- student_data %>%
    # For each Student and Grade, get their MAXIMUM points across all subjects
    group_by(Student, Grade) %>%
    summarize(`Points..out.of.600.` = max(`Points..out.of.600.`, na.rm = TRUE), .groups = "drop") %>%
    # Now calculate rank within each grade based on these maximum points
    group_by(Grade) %>%
    mutate(`Overall Department Or Grade Rank` = rank(desc(`Points..out.of.600.`), ties.method = "min")) %>%
    ungroup() %>%
    # Keep only Student, Grade and the calculated rank
    select(Student, Grade, `Overall Department Or Grade Rank`)
  
  # Join this back to the original student_data using both Student and Grade as keys
  student_data <- student_data %>%
    left_join(grade_ranks, by = c("Student", "Grade"))
  
  
  # Calculate ranks for each grade and subject
  student_data <- student_data %>%
    group_by(Grade, Subject) %>%
    mutate(`Subject Rank` = rank(desc(`Points..out.of.600.`), ties.method = "min"),
           `Total Students Grade & Subject` = n()) %>%
    ungroup()
  
  # Create Position Subject Suffix
  student_data <- student_data %>%
    mutate(`Position Subject Suffix` = case_when(
      `Subject Rank` == 1 ~ "ST",
      `Subject Rank` == 2 ~ "ND",
      `Subject Rank` == 3 ~ "RD",
      TRUE ~ "TH"
    ))
  
  # Create Rating Status based on Overall.Homework.Completion.%
  student_data <- student_data %>%
    mutate(RatingStatus = case_when(
      `Assessment.Score` <= 30 ~ "BelowAverage",  
      `Assessment.Score` <= 60 ~ "Average",       
      TRUE ~ "Good"                               
    ))
  
  
  
  # Create Progress Circle calculations
  student_data <- student_data %>%
    mutate(
      ProgressCircleHomework = paste0("PCH", ceiling(`Overall.Homework.Completion.%` / 10)),
      ProgressCircleExercise = paste0("PCE", ceiling(`Exercise.Completion.%` / 10)),
      ProgressCircleAssessment = paste0("PCA", ceiling(`Assessment.Completion.%` / 10)),
      ProgressCircleConcepts = paste0("PCC", ceiling(`Concept.Completion.%` / 10))
    )
  
  # Adjust Progress Circle values to max of 10
  student_data <- student_data %>%
    mutate(
      ProgressCircleHomework = gsub("PCH11", "PCH10", ProgressCircleHomework),
      ProgressCircleExercise = gsub("PCE11", "PCE10", ProgressCircleExercise),
      ProgressCircleAssessment = gsub("PCA11", "PCA10", ProgressCircleAssessment),
      ProgressCircleConcepts = gsub("PCC11", "PCC10", ProgressCircleConcepts)
    )
  
  # Define row names for transposed structure (used multiple times)
  row_names <- c(
    "Date Range",
    "Full Name",
    "Grade",
    "Subject",
    "Total Term Points",
    "Average Assessment Score",
    "Overall Homework Completion",
    "Concept Completion",
    "Assessment Completion",
    "Exercise Completion",
    "RatingStatus",
    "ProgressCircleHomework",
    "ProgressCircleExercise",
    "ProgressCircleAssessment",
    "ProgressCircleConcepts",
    "Overall Department Or Grade Rank",
    "Subject Rank",
    "Total Students Grade",
    "Position Subject Suffix"
  )
  
  # Define which rows should be numeric in the output
  numeric_rows <- c(
    "Total Term Points",
    "Average Assessment Score",
    "Overall Homework Completion",
    "Concept Completion",
    "Assessment Completion", 
    "Exercise Completion",
    "Overall Department Or Grade Rank",
    "Subject Rank",
    "Total Students Grade"
  )
  
  # First, handle each grade as a whole (for grades 8, 9 or any grade without subjects)
  for (grade in all_grades) {
    # For each grade, check if there are lower grades (like 8 or 9) with no subjects
    # or with NA/blank subjects
    grade_data <- student_data %>% 
      filter(Grade == grade)
    
    # Check if this grade has blank/NA subjects or a mix
    has_blank_subjects <- any(is.na(grade_data$Subject) | grade_data$Subject == "")
    subject_count <- length(unique(grade_data$Subject[!is.na(grade_data$Subject) & grade_data$Subject != ""]))
    
    # If grade is 8 or 9 OR it has no defined subjects, create a whole-grade sheet
    if (grade %in% c("8", "9", 8, 9) || (has_blank_subjects && subject_count == 0)) {
      cat("Creating consolidated sheet for Grade", grade, "with", nrow(grade_data), "students\n")
      
      # Create sheet name
      sheet_name <- paste0("Grade ", grade)
      
      # Add worksheet to workbook
      addWorksheet(wb, sheet_name)
      
      # Create the initial dataframe with row names
      transposed_df <- data.frame(RowName = row_names)
      
      # Write the row names to the first column
      writeData(wb, sheet_name, transposed_df, startCol = 1, startRow = 1)
      
      # For each student, write the data directly to the Excel sheet
      for (j in 1:nrow(grade_data)) {
        student <- grade_data$Student[j]
        col_position <- j + 1  # Column position (1st column is row names)
        
        # Write the student name as the column header
        writeData(wb, sheet_name, student, startCol = col_position, startRow = 1)
        
        # Values to be written
        values <- list(
          date_range,
          student,
          grade,
          ifelse(is.na(grade_data$Subject[j]) || grade_data$Subject[j] == "", "N/A", grade_data$Subject[j]),
          grade_data$`Points..out.of.600.`[j],
          grade_data$`Assessment.Score`[j],
          grade_data$`Overall.Homework.Completion.%`[j],
          grade_data$`Concept.Completion.%`[j],
          grade_data$`Assessment.Completion.%`[j],
          grade_data$`Exercise.Completion.%`[j],
          grade_data$RatingStatus[j],
          grade_data$ProgressCircleHomework[j],
          grade_data$ProgressCircleExercise[j],
          grade_data$ProgressCircleAssessment[j],
          grade_data$ProgressCircleConcepts[j],
          grade_data$`Overall Department Or Grade Rank`[j],
          grade_data$`Subject Rank`[j],
          grade_data$`Total Students Grade & Subject`[j],
          grade_data$`Position Subject Suffix`[j]
        )
        
        # Write each value with appropriate type
        for (k in 1:length(values)) {
          # Get the row name for current value
          current_row_name <- row_names[k]
          
          # Check if this should be a numeric field
          if (current_row_name %in% numeric_rows) {
            # Write as numeric
            writeData(wb, sheet_name, as.numeric(values[[k]]), 
                      startCol = col_position, startRow = k + 1)
          } else {
            # Write as text
            writeData(wb, sheet_name, values[[k]], 
                      startCol = col_position, startRow = k + 1)
          }
        }
      }
    }
  }
  
  # Now handle grade/subject combinations
  # Get unique grade and subject combinations
  grade_subjects <- student_data %>%
    filter(!is.na(Subject) & Subject != "") %>%  # Exclude blank/NA subjects
    select(Grade, Subject) %>%
    distinct() %>%
    arrange(Grade, Subject)
  
  # Loop through each grade/subject combination and create a sheet
  for (i in 1:nrow(grade_subjects)) {
    current_grade <- grade_subjects$Grade[i]
    current_subject <- grade_subjects$Subject[i]
    
    # Skip if current_subject is blank or NA
    if (is.na(current_subject) || current_subject == "") {
      next
    }
    
    # Filter data for current grade/subject
    current_data <- student_data %>%
      filter(Grade == current_grade & Subject == current_subject)
    
    # Skip if no data
    if (nrow(current_data) == 0) {
      next
    }
    
    # Debug info
    cat("Creating sheet for Grade", current_grade, "Subject:", current_subject, 
        "with", nrow(current_data), "students\n")
    
    # Create sheet name
    sheet_name <- paste0("Grade ", current_grade, " ", current_subject)
    
    # Truncate if too long (Excel limit is 31 chars)
    if (nchar(sheet_name) > 31) {
      sheet_name <- substr(sheet_name, 1, 31)
    }
    
    # Add worksheet to workbook
    addWorksheet(wb, sheet_name)
    
    # Create the initial dataframe with row names
    transposed_df <- data.frame(RowName = row_names)
    
    # Write the row names to the first column
    writeData(wb, sheet_name, transposed_df, startCol = 1, startRow = 1)
    
    # For each student, write the data directly to the Excel sheet
    for (j in 1:nrow(current_data)) {
      student <- current_data$Student[j]
      col_position <- j + 1  # Column position (1st column is row names)
      
      # Write the student name as the column header
      writeData(wb, sheet_name, student, startCol = col_position, startRow = 1)
      
      # Values to be written
      values <- list(
        date_range,
        student,
        current_grade,
        current_subject,
        current_data$`Points..out.of.600.`[j],
        current_data$`Assessment.Score`[j],
        current_data$`Overall.Homework.Completion.%`[j],
        current_data$`Concept.Completion.%`[j],
        current_data$`Assessment.Completion.%`[j],
        current_data$`Exercise.Completion.%`[j],
        current_data$RatingStatus[j],
        current_data$ProgressCircleHomework[j],
        current_data$ProgressCircleExercise[j],
        current_data$ProgressCircleAssessment[j],
        current_data$ProgressCircleConcepts[j],
        current_data$`Overall Department Or Grade Rank`[j],
        current_data$`Subject Rank`[j],
        current_data$`Total Students Grade & Subject`[j],
        current_data$`Position Subject Suffix`[j]
      )
      
      # Write each value with appropriate type
      for (k in 1:length(values)) {
        # Get the row name for current value
        current_row_name <- row_names[k]
        
        # Check if this should be a numeric field
        if (current_row_name %in% numeric_rows) {
          # Write as numeric
          writeData(wb, sheet_name, as.numeric(values[[k]]), 
                    startCol = col_position, startRow = k + 1)
          
          # After writing all student data to the sheet, add a "Total Students" row
          total_students <- nrow(current_data)
          writeData(wb, sheet_name, "Total Student Subject", startCol = 1, startRow = length(row_names) + 2)
          writeData(wb, sheet_name, total_students, startCol = col_position, startRow = length(row_names) + 2)
          
        } else {
          # Write as text
          writeData(wb, sheet_name, values[[k]], 
                    startCol = col_position, startRow = k + 1)
        }
      }
    }
  }
  
  # Save the workbook
  cat("Writing transposed report to", output_file, "...\n")
  saveWorkbook(wb, output_file, overwrite = TRUE)
  
  cat("Transposed report created successfully!\n")
  return(output_file)
}

#Create individual report export
export_sheets_as_individual_files <- function(main_file = "Transposed_Report.xlsx") {
  # Make sure the required packages are loaded
  library(openxlsx)
  
  # Check if googlesheets4 and googledrive packages are installed
  if (!requireNamespace("googlesheets4", quietly = TRUE)) {
    install.packages("googlesheets4")
  }
  if (!requireNamespace("googledrive", quietly = TRUE)) {
    install.packages("googledrive")
  }
  
  library(googlesheets4)
  library(googledrive)
  
  # Authenticate with Google (user will need to authorize in browser first time)
  googlesheets4::gs4_auth()
  googledrive::drive_auth()
  
  # Prompt the user for the folder name after Google authentication
  cat("Please enter the name of the Google Drive folder to use:\n")
  google_folder_name <- readline(prompt = "> ")
  
  # Check if the user provided a folder name, use default if not
  if (google_folder_name == "") {
    google_folder_name <- "Individual Reporting"
    cat("No folder name provided. Using default:", google_folder_name, "\n")
  }
  
  # Check if the specified folder exists, if not create it
  folder_exists <- googledrive::drive_find(
    type = "folder",
    pattern = google_folder_name
  )
  
  if (nrow(folder_exists) == 0) {
    # Folder doesn't exist, create it
    folder <- googledrive::drive_mkdir(google_folder_name)
    folder_id <- folder$id
    cat("Created new Google Drive folder:", google_folder_name, "\n")
  } else {
    # Folder exists, use the first matching folder
    folder_id <- folder_exists$id[1]
    cat("Using existing Google Drive folder:", google_folder_name, "\n")
  }
  
  # Get the sheet names from the main file
  sheet_names <- getSheetNames(main_file)
  
  # Skip the Summary sheet
  sheet_names <- sheet_names[sheet_names != "Summary"]
  
  cat("Exporting", length(sheet_names), "sheets as Google Sheets to folder:", google_folder_name, "\n")
  
  # Process each sheet
  for (sheet_name in sheet_names) {
    cat("  Exporting sheet:", sheet_name, "\n")
    
    # Read the data from the sheet
    sheet_data <- read.xlsx(main_file, sheet = sheet_name)
    
    # Create a safe filename
    safe_name <- gsub("[^a-zA-Z0-9]", "_", sheet_name)
    sheet_title <- paste0("Transposed_", safe_name)
    
    # Create a new Google Sheet
    tryCatch({
      # Create the sheet (initially in the root of My Drive)
      new_sheet <- googlesheets4::gs4_create(sheet_title, sheets = sheet_data)
      
      # Move the sheet to the specified folder
      googledrive::drive_mv(
        file = as_id(new_sheet),
        path = as_id(folder_id)
      )
      
      cat("    Created Google Sheet:", sheet_title, "in folder:", google_folder_name, "\n")
    }, error = function(e) {
      cat("    Error creating/moving sheet:", e$message, "\n")
    })
  }
  
  cat("All Google Sheets created successfully in folder:", google_folder_name, "!\n")
}

#---------------------------------------------------------------------------------------------------------

# Function to add a total points summary sheet to the All Masters workbook
# Function to add a total points summary sheet to the All Masters workbook
add_total_points_summary <- function(input_file = "All_Masters.xlsx") {
  # Read the Master of Masters sheet
  cat("Reading Master of Masters sheet from", input_file, "...\n")
  master_data <- read_excel(input_file, sheet = "Master of Masters")
  
  # Filter out summary rows (those starting with "---" or "Grade")
  student_data <- master_data %>%
    filter(!grepl("^---", Student) & !grepl("^Grade", Student))
  
  # Get all sheets in the workbook to examine each subject
  all_sheets <- excel_sheets(input_file)
  # Remove "Master of Masters" and any other non-subject sheets
  subject_sheets <- all_sheets[!all_sheets %in% c("Master of Masters", "Total Points Summary", "Points Summary Stats")]
  
  cat("Found", length(subject_sheets), "subject sheets to analyze\n")
  
  # Initialize a list to store results for each subject/grade
  subject_grade_points <- list()
  
  # Process each subject sheet to determine max points
  for(sheet_name in subject_sheets) {
    cat("Analyzing sheet:", sheet_name, "\n")
    
    # Read the sheet data
    sheet_data <- read_excel(input_file, sheet = sheet_name)
    
    # Find all "Points..out.of.600." columns
    points_columns <- grep("Points..out.of.600.", names(sheet_data), value = TRUE)
    
    # Extract grade and subject from the data
    # Get the first row that has valid Grade values
    first_valid_row <- which(!is.na(sheet_data$Grade))[1]
    if(is.na(first_valid_row) || length(first_valid_row) == 0) {
      cat("  Warning: Could not determine Grade for sheet", sheet_name, "\n")
      next
    }
    
    grade <- sheet_data$Grade[first_valid_row]
    
    # Handle case where subject might be NA or empty (like Grade 9)
    subject <- NA
    if("Subject" %in% names(sheet_data) && !is.na(sheet_data$Subject[first_valid_row]) && 
       sheet_data$Subject[first_valid_row] != "") {
      subject <- sheet_data$Subject[first_valid_row]
    } else {
      # Use sheet name as subject if it's not in the data
      subject <- sheet_name
    }
    
    cat("  Grade:", grade, ", Subject:", subject, "\n")
    cat("  Found", length(points_columns), "points columns\n")
    
    # Count how many columns are 300 vs 600
    cols_600 <- 0
    cols_300 <- 0
    
    # Analyze each points column to determine if it's out of 300 or 600
    total_max_points <- 0
    for(col in points_columns) {
      # Get max value in this column
      max_value <- max(sheet_data[[col]], na.rm = TRUE)
      
      # Determine if out of 300 or 600
      max_points <- ifelse(max_value > 300, 600, 300)
      
      # Update counters
      if(max_points == 600) {
        cols_600 <- cols_600 + 1
      } else {
        cols_300 <- cols_300 + 1
      }
      
      total_max_points <- total_max_points + max_points
      
      cat("    Column:", col, "- Max value:", max_value, "- Inferred max:", max_points, "\n")
    }
    
    cat("  Columns out of 600:", cols_600, ", Columns out of 300:", cols_300, "\n")
    cat("  Total maximum points for", grade, subject, ":", total_max_points, "\n")
    
    # Store the results
    subject_grade_points[[paste(grade, subject)]] <- list(
      Grade = grade,
      Subject = subject,
      Total_Max_Points = total_max_points,
      Columns_600 = cols_600,
      Columns_300 = cols_300
    )
  }
  
  # Convert list to dataframe
  subject_grade_summary <- do.call(rbind, lapply(subject_grade_points, function(x) {
    data.frame(
      Grade = x$Grade,
      Subject = x$Subject,
      Total_Max_Points = x$Total_Max_Points,
      Columns_600 = x$Columns_600,
      Columns_300 = x$Columns_300,
      stringsAsFactors = FALSE
    )
  }))
  subject_grade_summary <- as.data.frame(subject_grade_summary)
  
  # Extract unique student-grade-subject combinations with their points
  student_points <- student_data %>%
    # Select only the relevant columns
    select(Student, Grade, Subject, `Points..out.of.600.`) %>%
    # Keep unique student-grade-subject combinations
    distinct() %>%
    # Remove any NA or empty subjects
    filter(!is.na(Subject) & Subject != "")
  
  # Special handling for cases where Subject is NA (like Grade 9)
  # Find students with no subject who aren't already in student_points
  students_no_subject <- student_data %>%
    filter(is.na(Subject) | Subject == "") %>%
    select(Student, Grade, `Points..out.of.600.`) %>%
    distinct() %>%
    # Add "General" as the subject
    mutate(Subject = "General")
  
  # Combine regular students and those with no subject
  student_points <- bind_rows(student_points, students_no_subject)
  
  # Calculate highest points for each student (across subjects)
  highest_points <- student_points %>%
    group_by(Student, Grade) %>%
    summarize(
      `Highest Points` = max(`Points..out.of.600.`, na.rm = TRUE),
      `Best Subject` = Subject[which.max(`Points..out.of.600.`)]
    ) %>%
    ungroup()
  
  # Rename column for student data and join with highest points
  student_points <- student_points %>%
    rename(Points = `Points..out.of.600.`) %>%
    # Join with highest points data
    left_join(highest_points, by = c("Student", "Grade"))
  
  # Create the grade/subject summary with descriptive column names
  grade_subject_summary <- subject_grade_summary %>%
    arrange(Grade, Subject) %>%
    rename(
      `Max Available Points` = Total_Max_Points,
      `600 Columns` = Columns_600,  
      `300 Columns` = Columns_300  
    )
  
  # Load the existing workbook
  wb <- loadWorkbook(input_file)
  
  # 1. Create the Total Points Summary sheet (students only)
  if ("Total Points Summary" %in% names(wb)) {
    removeWorksheet(wb, "Total Points Summary")
  }
  addWorksheet(wb, "Total Points Summary")
  
  # Arrange student data
  arranged_student_points <- student_points %>%
    # Keep only the desired columns (no Max Available Points)
    select(Student, Grade, Subject, Points, `Highest Points`, `Best Subject`) %>%
    arrange(Grade, Student, Subject)
  
  # Write the student data only to the Total Points Summary sheet
  writeData(wb, "Total Points Summary", arranged_student_points)
  
  # 2. Create a separate Points Summary Stats sheet
  if ("Points Summary Stats" %in% names(wb)) {
    removeWorksheet(wb, "Points Summary Stats")
  }
  addWorksheet(wb, "Points Summary Stats")
  
  # Write the grade/subject summary directly to the Points Summary Stats sheet
  writeData(wb, "Points Summary Stats", grade_subject_summary %>% 
              select(Grade, Subject, `Max Available Points`, 
                     `600 Columns`, `300 Columns`))
  
  # Save the workbook
  saveWorkbook(wb, input_file, overwrite = TRUE)
  
  cat("Added Total Points Summary and Points Summary Stats sheets to", input_file, "\n")
  return(input_file)
}

# Run the main functions
results <- process_all_workbooks()
add_total_points_summary("All_Masters.xlsx")
create_transposed_master_report("All_Masters.xlsx", "Transposed_Report.xlsx", "15 FEB - 28 MAR (Term 1)")
export_sheets_as_individual_files("Transposed_Report.xlsx")