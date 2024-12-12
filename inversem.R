# 安装并加载必要的包
if (!require("readxl")) install.packages("readxl", version = "4.5.0")
if (!require("mediation")) install.packages("mediation", version = "4.5.0")
if (!require("openxlsx")) install.packages("openxlsx")

library(readxl)
library(mediation)
library(openxlsx)

# 读取Excel数据
# 假设文件名为"data.xlsx"，每个自变量、中介变量和因变量分别在不同的工作表中
file_path <- "data.xlsx"
X <- read_excel(file_path, sheet = "Sheet1")
M <- read_excel(file_path, sheet = "Sheet2")
Y <- read_excel(file_path, sheet = "Sheet3")

# 初始化结果存储列表
results_list <- list()

# 遍历每个X、M、Y的组合，进行中介分析
for (i in 1:nrow(X)) {
  for (j in 1:nrow(M)) {
    for (k in 1:nrow(Y)) {
      X_value <- X[i, -1]  # 假设第一列是OTU ID，其余列是样本数据
      M_value <- M[j, -1]
      Y_value <- Y[k, -1]
      
      # 创建数据框用于分析
      data <- data.frame(X = as.numeric(X_value), M = as.numeric(M_value), Y = as.numeric(Y_value))
      
      # 定义中介模型
      model_m <- lm(M ~ X, data = data)  # 自变量对中介变量的影响
      model_y <- lm(Y ~ X + M, data = data)  # 自变量和中介变量对因变量的影响
      
      # 进行中介分析
      mediation_result <- mediate(model_m, model_y, treat = "X", mediator = "M", boot = TRUE, sims = 1000)
      
      # 提取所需的结果值
      beta_a <- coef(model_m)["X"]  # 自变量对中介变量的系数 (a)
      p_value_a <- summary(model_m)$coefficients["X", 4]  # p值
      beta_b <- coef(model_y)["M"]  # 中介变量对因变量的系数 (b)
      p_value_b <- summary(model_y)$coefficients["M", 4]  # p值
      beta_c <- coef(lm(Y ~ X, data = data))["X"]  # 自变量对因变量的总效应 (c)
      p_value_c <- summary(lm(Y ~ X, data = data))$coefficients["X", 4]  # p值
      
      # 提取直接效应 c'
      direct_effect_c_prime <- mediation_result$z0  # 直接效应 (c')
      direct_p_value_c_prime <- mediation_result$z0.p  # 直接效应的 p 值
      
      indirect_effect <- mediation_result$d0  # 间接效应 (a * b)
      indirect_p_value <- mediation_result$d0.p  # 中介效应的 p 值
      mediation_percentage <- (indirect_effect / (indirect_effect + direct_effect_c_prime)) * 100  # 中介效应的百分比
      
      # 进行逆中介分析
      # 逆中介指的是交换因变量和中介变量的角色，测试中介关系的反向情况
      model_y_inverse <- lm(Y ~ X, data = data)  # 自变量对因变量的影响
      model_m_inverse <- lm(M ~ X + Y, data = data)  # 自变量和因变量对中介变量的影响
      
      # 进行逆中介分析
      inverse_mediation_result <- mediate(model_y_inverse, model_m_inverse, treat = "X", mediator = "Y", boot = TRUE, sims = 1000)
      
      # 提取逆中介分析的结果值
      inverse_indirect_effect <- inverse_mediation_result$d0  # 逆中介效应
      inverse_indirect_p_value <- inverse_mediation_result$d0.p  # 逆中介效应的 p 值
      
      # 将结果存储到列表中
      results_list[[length(results_list) + 1]] <- data.frame(
        X_OTU_ID = X[i, 1],
        M_OTU_ID = M[j, 1],
        Y_OTU_ID = Y[k, 1],
        beta_a = beta_a,
        p_value_a = p_value_a,
        beta_b = beta_b,
        p_value_b = p_value_b,
        beta_c = beta_c,
        p_value_c = p_value_c,
        direct_effect_c_prime = direct_effect_c_prime,
        direct_p_value_c_prime = direct_p_value_c_prime,
        indirect_effect = indirect_effect,
        indirect_p_value = indirect_p_value,
        mediation_percentage = mediation_percentage,
        inverse_indirect_effect = inverse_indirect_effect,
        inverse_indirect_p_value = inverse_indirect_p_value
      )
    }
  }
}

# 合并所有结果为一个数据框
final_results <- do.call(rbind, results_list)

# 将结果写入Excel文件
write.xlsx(final_results, file = "mediation_results.xlsx", rowNames = FALSE)

print("中介分析和逆中介分析已完成，结果保存在mediation_results.xlsx中。")
