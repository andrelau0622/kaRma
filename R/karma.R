#' Causal Inference Analysis with Multiple Methods and Excel Output
#'
#' This function performs causal inference analysis using various methods
#' and supports missing data imputation with MICE. Results can be saved
#' to an Excel file and relevant plots can be generated.
#'
#' @param data The dataset to be analyzed.
#' @param outcome_var The outcome variable of interest.
#' @param method The causal inference method to use. Choose from 'OLS', 'IV', 'GLM', 'PSM', 'DID', 'RDD', 'CausalGraph', 'IPW', 'CausalForest', 'Bayesian', 'SCM'.
#' @param treatment_var The treatment variable (where applicable).
#' @param covariates A vector of covariates to adjust for.
#' @param threshold The threshold for RDD (if applicable).
#' @param use_mice Whether to perform MICE imputation (TRUE/FALSE).
#' @param m The number of imputations for MICE.
#' @param mice_methods A named list specifying the imputation methods for each variable.
#' @param save_excel Boolean indicating whether to save the results to Excel.
#' @param file_name The name of the Excel file to save results.
#'
#' @import MatchIt
#' @import fixest
#' @import rdrobust
#' @import dagitty
#' @import ipw
#' @import grf
#' @import AER
#' @import Synth
#' @import brms
#' @import mice
#' @import openxlsx
#' @export
karma <- function(data, outcome_var, method, treatment_var = NULL, covariates = NULL, 
                  threshold = NULL, use_mice = FALSE, m = 5, mice_methods = NULL, 
                  save_excel = FALSE, file_name = "causal_results.xlsx") {
  
  # Step 1: MICE 插补缺失值
  if (use_mice) {
    print("开始进行MICE插补...")
    
    # 使用默认的 MICE 方法，或者传入的 mice_methods
    if (is.null(mice_methods)) {
      mice_methods <- "pmm"  # 使用默认的 Predictive Mean Matching (pmm) 方法
    }
    
    mice_data <- mice(data, m = m, method = mice_methods, maxit = 5, seed = 500)
    data <- complete(mice_data)  # 使用插补后的数据
    print("MICE 插补完成。")
  }
  
  # Step 2: 根据指定方法执行相应的因果推断分析
  result <- switch(method,
                   
                   # OLS 回归分析
                   "OLS" = {
                     formula <- as.formula(paste(outcome_var, "~", paste(covariates, collapse = " + ")))
                     model <- lm(formula, data = data)
                     summary(model)
                   },
                   
                   # 工具变量回归分析
                   "IV" = {
                     formula <- as.formula(paste(outcome_var, "~", paste(covariates, collapse = " + "), "|", treatment_var))
                     model <- ivreg(formula, data = data)
                     summary(model)
                   },
                   
                   # GLM 广义线性模型
                   "GLM" = {
                     formula <- as.formula(paste(outcome_var, "~", paste(covariates, collapse = " + ")))
                     model <- glm(formula, family = binomial(link = "logit"), data = data)
                     summary(model)
                   },
                   
                   # 倾向得分匹配 PSM
                   "PSM" = {
                     formula <- as.formula(paste(treatment_var, "~", paste(covariates, collapse = " + ")))
                     psm_model <- matchit(formula, data = data, method = "nearest")
                     matched_data <- match.data(psm_model)
                     formula_outcome <- as.formula(paste(outcome_var, "~", treatment_var, "+", paste(covariates, collapse = " + ")))
                     model_psm <- lm(formula_outcome, data = matched_data)
                     summary(model_psm)
                   },
                   
                   # 双重差分 DID
                   "DID" = {
                     formula <- as.formula(paste(outcome_var, "~", treatment_var, "* post | id + time"))
                     model <- feols(formula, data = data)
                     summary(model)
                   },
                   
                   # 断点回归 RDD
                   "RDD" = {
                     if (is.null(threshold)) stop("Please specify a threshold for RDD.")
                     rdd_model <- rdrobust(data[[outcome_var]], data[[treatment_var]], c = threshold)
                     summary(rdd_model)
                   },
                   
                   # 因果图模型 Causal Graphical Models
                   "CausalGraph" = {
                     dag <- dagitty(paste("dag {", paste(paste(covariates, "->", outcome_var), collapse = " "), "}"))
                     plot(dag)
                     adjustmentSets(dag, exposure = treatment_var, outcome = outcome_var)
                   },
                   
                   # 逆概率加权 IPW
                   "IPW" = {
                     ipw_model <- ipwpoint(exposure = data[[treatment_var]], family = "binomial", link = "logit",
                                           numerator = ~ 1, denominator = as.formula(paste("~", paste(covariates, collapse = " + "))),
                                           data = data)
                     data$weights <- ipw_model$ipw.weights
                     formula_outcome <- as.formula(paste(outcome_var, "~", treatment_var))
                     model_ipw <- glm(formula_outcome, weights = data$weights, data = data)
                     summary(model_ipw)
                   },
                   
                   # 因果森林 Causal Forest
                   "CausalForest" = {
                     X <- as.matrix(data[, covariates])
                     W <- data[[treatment_var]]
                     Y <- data[[outcome_var]]
                     causal_forest_model <- causal_forest(X, Y, W)
                     predict(causal_forest_model)
                   },
                   
                   # 贝叶斯因果推断 Bayesian Causal Inference
                   "Bayesian" = {
                     formula <- as.formula(paste(outcome_var, "~", paste(covariates, collapse = " + ")))
                     model_bayesian <- brm(formula, data = data, family = bernoulli())
                     summary(model_bayesian)
                   },
                   
                   # 合成控制法 SCM
                   "SCM" = {
                     dataprep_out <- dataprep(foo = data, predictors = covariates,
                                              dependent = outcome_var, unit.variable = "unit", time.variable = "time",
                                              treatment.identifier = 1, controls.identifier = c(2,3),
                                              time.predictors.prior = 1:5, time.optimize.ssr = 1:5)
                     synth_out <- synth(dataprep_out)
                     synth_out
                   },
                   
                   stop("Unknown method")
  )
  
  # Step 3: Save results to Excel (if needed)
  if (save_excel) {
    print(paste("Saving results to", file_name, "..."))
    openxlsx::write.xlsx(result, file_name)
    print("Results saved to Excel.")
  }
  
  return(result)
}
