# author: "Andrei-Emil Constantinescu"
# date: "2023-04-16"

#### Functions ####
# Function to make a flextable object fit a Word page
FitFlextableToPage <- function(ft, pgwidth = 6){
  
  ft_out <- ft %>% autofit()
  
  ft_out <- width(ft_out, width = dim(ft_out)$widths*pgwidth /(flextable_dim(ft_out)$widths))
  return(ft_out)
}

# Load packages function. If not found, install, then load.
process_packages <- function(package_names) {
  for (package_name in package_names) {
    if (!requireNamespace(package_name, quietly = TRUE)) {
      message("Package '", package_name, "' is not installed. Installing now...")
      install.packages(package_name)
    }
    library(package_name, character.only = TRUE)
  }
}

# Function to run logistic regression
# Maybe next time I could use a random effects model for the cs
analyze_c <- function(data) {
  # Perform logistic regression
  model <- glm(screening ~ demographic_variables,
               data = data, family = binomial(link = "logit"))
  
  # Calculate odds ratios and confidence intervals
  model_summary <- tidy(model, conf.int = TRUE, exponentiate = TRUE)
  
  return(list(model_summary = model_summary))
}

#### Install packages and load them ####
packageVec <- c("data.table", "tidyverse", "ggpubr", "xlsx", 
                "gtsummary", "flextable", "officer", "naniar",
                "broom", "caret", "pROC", "mice", 
                "metafor", "polycor", "corrplot")
process_packages(packageVec)

# Set working directory to location of the file
setwd(dirname(rstudioapi::getSourceEditorContext()$path))

# Set seed for replicability
set.seed(20230416)

#### Preparing the datasets ####
cs_data <- list.files(path = "../00_data", pattern = "*.xlsx", full.names = TRUE) %>% ##
  setNames(nm = .) %>% # 
  lapply(read.xlsx, 1) %>%
  rbindlist(use.names=TRUE, fill=TRUE, idcol="ID") %>%
  mutate(cID = str_extract(cID, "test\\d+")) %>%
  select(!contains("NA."))

# Replace all NA characters and values with "missing" string
cs_data[is.na(cs_data) | cs_data == "NA"] <- NA

# Check the type of the variables in the dataframe
str(cs_data)

# Height and Alcohol are character, should be numeric
change_cols <- c("Height", "Alcohol")
cs_data[, (change_cols) := lapply(.SD, as.numeric), .SDcols = change_cols]

# Create dataframe which contains only rows with at least one missing value
cs_data_missing <- cs_data %>% 
  select(where(~sum(is.na(.x)) > 0)) %>% 
  filter(if_any(everything(), is.na))

# Plot the missing values in an informative way
missing_plot <- gg_miss_upset(cs_data_missing, sets.x.label = "Nr. missing")
png(file="../03_figures/figure_1.png", res = 300, units = "mm",
    width = 180, height = 110) 
missing_plot
dev.off()

# Use the 'mice' package to impute the missing data
# First, change the cID, Deps and ES variables to factors
change_cols <- c("cID", "Deps", "ES")
cs_data[, (change_cols) := lapply(.SD, as.factor), .SDcols = change_cols]
cs_data[, screening := factor(screening, levels = c("N", "Y"), 
                                   labels = c(0, 1))]

# Imputation methods for each variable
# See '?mice' for info on the imputation methods
imputation_methods <- make.method(cs_data)
imputation_methods[c("Alcohol", "Height")] <- "norm.nob"
imputation_methods["ES"] <- "polr"

# Perform multiple imputation
mice_output <- mice(cs_data, m = 5, maxit = 50, method = imputation_methods, seed = 42)
completed_data <- complete(mice_output, 1)

# Create descriptive statistics table with the imputed data
# The mean is presented for continuous variables
# For categorical variables, the number of patients who correspond to a category
#  is displayed (n) along with the total number of patients in the categorical
#  variable (N).
completed_data_table <- completed_data %>% 
  select(cID, VARS) %>% 
  tbl_summary(     
    by = cID,                                               # stratify entire table by cID
    statistic = list(all_continuous() ~ "{mean} \n({sd})",        # stats and format for continuous columns
                     all_categorical() ~ "{n} / {N} \n({p})"),   # stats and format for categorical columns
    digits = all_continuous() ~ 1,                              # rounding for continuous columns
    type   = all_categorical() ~ "categorical",                 # force all categorical levels to display
    label  = list(                                              # display labels for column names
      cID ~ "c ID",
      screening ~ "Screening (Yes/No)",
      ...
    missing_text = "Missing",
    missing = "no"
  ) 

# Save as flextable to then output as .docx Word table
completed_data_ft <- as_flex_table(completed_data_table) %>% 
  padding(padding = 0, part = "all") %>% 
  FitFlextableToPage(11.7)

# Change the layout to landscape
sect_properties <- prop_section(
  page_size = page_size(
    orient = "landscape",
    width = 11.7, height = 8.3
  ),
  type = "continuous",
  page_margins = page_mar(0.5, 0.5, 0.5, 0.5)
)
save_as_docx(completed_data_ft, path = "../04_tables/table_1.docx", pr_section = sect_properties)

# Analyze the combined data i.e. run logistic regression with the characteristics
# Perform logistic regression
analysis_result <- analyze_c(completed_data)

# Extract and combine model summaries
model_summaries <- bind_rows(analysis_result) %>%
  filter(term != "(Intercept)") %>%
  as.data.table()

#### Sensitivity analysis ####
# Studying the performance of the model by performing
#  a cross-validation analysis to generate an average AUC

# Prepare k-fold cross-validation
folds <- createFolds(completed_data$screening, k = 10)
auc_results <- c()

# Perform cross-validation
for (i in 1:length(folds)) {
  # Split the data into training and testing sets
  train_data <- completed_data[-folds[[i]], ]
  test_data <- completed_data[folds[[i]], ]
  
  # Fit the logistic regression model on the training data
  model <- glm(screening ~ VARS, 
               data = train_data, family = binomial(link = "logit"))
  
  # Compute the AUC on the test data
  roc_obj <- roc(test_data$screening, predict(model, type = "response", newdata = test_data))
  auc <- auc(roc_obj)
  auc_results <- c(auc_results, auc)
}

# Calculate the average AUC across all iterations
mean_auc <- mean(auc_results)
print(paste("Average AUC:", mean_auc)) # pretty good, but not great

# Compute the pairwise correlation matrix including categorical variables and p-values
mixed_cor_results <- hetcor(completed_data[, c("variables_of_interest")], ML = TRUE)
print(mixed_cor_results) # nothing notable

# Create forest plot to show summary of logitreg results
png(file="../03_figures/figure_2.png", res = 300, units = "mm",
    width = 180, height = 110) 
forest_plot <- metafor::forest(x = model_summaries$estimate, ci.lb = model_summaries$conf.low, ci.ub = model_summaries$conf.high,
                               slab = model_summaries$term, annotate=TRUE, header = "Variable",
                               cex = 1, xlim = c(-2,4.5), alim=c(0,4), steps = 9,
                               refline = 1, xlab = "Odds Ratio")
### add text with AUC
text(-1.3, -0.24, pos=3, cex=0.75, bquote(paste("AUC-ROC = ",
                                                .(formatC(signif(mean_auc))))))
dev.off()





