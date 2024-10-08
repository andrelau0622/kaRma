% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/karma.R
\name{karma}
\alias{karma}
\title{Causal Inference Analysis with Multiple Methods and Excel Output}
\usage{
karma(
  data,
  outcome_var,
  method,
  treatment_var = NULL,
  covariates = NULL,
  threshold = NULL,
  use_mice = FALSE,
  m = 5,
  mice_methods = NULL,
  save_excel = FALSE,
  file_name = "causal_results.xlsx"
)
}
\arguments{
\item{data}{The dataset to be analyzed.}

\item{outcome_var}{The outcome variable of interest.}

\item{method}{The causal inference method to use. Choose from 'OLS', 'IV', 'GLM', 'PSM', 'DID', 'RDD', 'CausalGraph', 'IPW', 'CausalForest', 'Bayesian', 'SCM'.}

\item{treatment_var}{The treatment variable (where applicable).}

\item{covariates}{A vector of covariates to adjust for.}

\item{threshold}{The threshold for RDD (if applicable).}

\item{use_mice}{Whether to perform MICE imputation (TRUE/FALSE).}

\item{m}{The number of imputations for MICE.}

\item{mice_methods}{A named list specifying the imputation methods for each variable.}

\item{save_excel}{Boolean indicating whether to save the results to Excel.}

\item{file_name}{The name of the Excel file to save results.}
}
\description{
This function performs causal inference analysis using various methods
and supports missing data imputation with MICE. Results can be saved
to an Excel file and relevant plots can be generated.
}
