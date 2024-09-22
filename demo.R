# 安装kaRma包
devtools::install_github('andrelau0622/kaRma')

# 加载所需的包
library(kaRma)
library(dplyr)

# 使用kaRma包进行因果推断分析
res <- karma(
  data = my_data,               # 要分析的数据集
  outcome_var = "outcome",       # 指定结果变量
  method = "OLS",                # 使用双重差分（DID）进行因果推断
  treatment_var = "treatment",   # 指定干预变量
  covariates = c("age", "gender"), # 指定要调整的协变量
  
  use_mice = TRUE,               # 启用 MICE 插补
  m = 5,                        # 指定 MICE 插补次数为 5
  mice_methods = "pmm",          # 为所有变量指定预测均值匹配 (PMM) 插补方法
  
  save_excel = TRUE,             # 保存结果
  file_name = "res.xlsx" # 保存结果的 Excel 文件名
)

# 输出结果
print(res)