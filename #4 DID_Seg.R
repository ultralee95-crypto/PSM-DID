# ==============================================================================
# 소재(seg=1) / 부품(seg=2) / 장비(seg=3)별 DID 분석
# RQ1 (패널 FE ATT) + Simple DID (2019 vs 2024) — log 변환 버전
# OLS
# 종업원수삭제 (결측치 다수), 수출, 개발비용계 추가. 20260308
# ==============================================================================

packages <- c("readxl", "dplyr", "tidyr", "ggplot2", "broom", "sandwich",
              "lmtest", "car", "writexl", "knitr", "scales", "plm", "MASS")

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}

select <- dplyr::select
lag    <- dplyr::lag

setwd("/Users/ultra/PSM-DID")
getwd()

matched <- read_excel("matched_dataset_segPSM.xlsx", col_types = "text")

cat("=== 데이터 확인 ===\n")
cat("행수:", nrow(matched), "\n")
cat("seg 분포:\n"); print(table(matched$seg))
cat("treat 분포:\n"); print(table(matched$treat))

to_num <- function(x) as.numeric(x)

# ------------------------------------------------------------------------------
# [핵심] 부호 보존 log 변환 함수 (음수값 처리)
# sign(x) * log1p(|x|) : 음수 → 음수 방향 유지, 0 → 0
# ------------------------------------------------------------------------------
log_signed <- function(x) {
  ifelse(is.na(x), NA, sign(x) * log1p(abs(x)))
}

# ==============================================================================
# 1. Wide → Long 패널 변환 (2019-2024)
# ==============================================================================

years <- 2019:2024

panel_list <- lapply(1:nrow(matched), function(i) {
  row <- matched[i, ]
  data.frame(
    firm_id     = i,
    seg         = to_num(row$seg),
    분류        = row$분류,
    treat       = to_num(row$treat),
    region      = row$region,
    subclass    = row$subclass,
    year        = years,
    #emp       = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S05000.종업원수")]])),
    asset     = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S15000.자산총계")]])),
    debt      = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18000.부채총계")]])),
    equity    = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18900.자본총계")]])),
    capital   = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18100.자본금")]])),
    sales     = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S21100.총매출액")]])),
    op_profit = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S25000.영업이익(손실)")]])),
    patent    = sapply(years, function(y) to_num(row[[paste0("p", y)]])),
    # ↓ 추가
    export    = sapply(years, function(y) to_num(row[[paste0("exportamt", y)]])),
    rdcost    = sapply(years, function(y) to_num(row[[paste0("rdcost", y)]])),
    stringsAsFactors = FALSE
  )
})

panel <- bind_rows(panel_list) %>% arrange(firm_id, year)

cat("\n패널:", nrow(panel), "obs,", length(unique(panel$firm_id)), "firms\n")

# ==============================================================================
# 2. 원시 수준 변수 → Log 변환
# ==============================================================================

panel <- panel %>%
  mutate(
    log_patent    = log1p(pmax(patent,  0, na.rm = TRUE)),   # ln(특허+1)
    #log_emp_raw   = ifelse(!is.na(emp)   & emp   > 0, log(emp),   NA),  # ln(종업원수)
    log_asset_raw = ifelse(!is.na(asset) & asset > 0, log(asset), NA),  # ln(자산총계)
    log_debt      = log1p(pmax(debt,    0, na.rm = TRUE)),   # ln(부채총계+1)
    log_equity    = log_signed(equity),                       # ±ln(자본총계+1)
    log_capital   = log1p(pmax(capital, 0, na.rm = TRUE)),   # ln(자본금+1)
    log_sales     = ifelse(!is.na(sales) & sales > 0, log(sales), NA),  # ln(매출액)
    log_opprofit  = log_signed(op_profit),                     # ±ln(영업이익+1)
    log_export = log1p(pmax(export, 0, na.rm = TRUE)), # 수출금액
    log_rdcost = log1p(pmax(rdcost, 0, na.rm = TRUE)) # 연구개발비 
  )

# ==============================================================================
# 3. 종속변수 생성
# ==============================================================================

panel <- panel %>%
  mutate(
    post = ifelse(year >= 2020, 1, 0),
    did  = treat * post
  )

# ── [혁신성] ln(특허출원수+1) ──
panel$innov_patent <- panel$log_patent

# ── [수익성] log(비율) ──
panel$profit_opm <- ifelse(
  panel$sales > 0,
  log_signed(panel$op_profit / panel$sales),
  NA
)
panel$profit_roa <- ifelse(
  panel$asset > 0,
  log_signed(panel$op_profit / panel$asset),
  NA
)
panel$profit_roe <- ifelse(
  panel$equity != 0,
  log_signed(panel$op_profit / panel$equity),
  NA
)

# ── [안정성] log(비율) ──
panel$stab_debtratio <- ifelse(
  panel$equity != 0,
  panel$log_debt - log_signed(panel$equity),
  NA
)
panel$stab_equitydep <- ifelse(
  panel$asset > 0,
  log_signed(panel$equity) - panel$log_asset_raw,
  NA
)

# ── [활동성] log(비율) ──
panel$act_assetturn <- ifelse(
  panel$asset > 0,
  panel$log_sales - panel$log_asset_raw,
  NA
)
panel$act_equityturn <- ifelse(
  panel$equity != 0,
  panel$log_sales - log_signed(panel$equity),
  NA
)

# ── Winsorize (log 변환 후 이상치 처리) ──
winsorize <- function(x, probs = c(0.01, 0.99)) {
  q <- quantile(x, probs, na.rm = TRUE)
  x[x < q[1] & !is.na(x)] <- q[1]
  x[x > q[2] & !is.na(x)] <- q[2]
  return(x)
}

ratio_vars <- c("profit_opm", "profit_roa", "profit_roe",
                "stab_debtratio", "stab_equitydep",
                "act_assetturn", "act_equityturn")

for (v in ratio_vars) {
  panel[[v]] <- winsorize(panel[[v]])
}

# ── 통제변수 ──
#panel$log_emp   <- panel$log_emp_raw
panel$log_asset <- panel$log_asset_raw
panel$leverage  <- panel$debt / ifelse(panel$asset == 0, NA, panel$asset)

cat("\n=== 변수 생성 완료 ===\n")

# ==============================================================================
# 4. DV 정의 — RQ1용
#    성장성: grow_* 대신 log 수준 변수 사용 (두-way FE가 기저 수준 흡수)
# ==============================================================================

dv_info <- data.frame(
  dv     = c("innov_patent",
             "log_asset_raw", "log_sales", "log_opprofit",
             "profit_opm", "profit_roa", "profit_roe",
             "stab_debtratio", "stab_equitydep",
             "act_assetturn", "act_equityturn"),
  label  = c("ln(특허+1)",
             "ln자산", "ln매출", "ln영업이익",
             "ln영업이익률", "ln(ROA)", "ln(ROE)",
             "ln부채비율", "ln자기자본의존도",
             "ln총자산회전율", "ln자기자본회전율"),
  cat    = c("혁신성",
             "성장성", "성장성", "성장성",
             "수익성", "수익성", "수익성",
             "안정성", "안정성",
             "활동성", "활동성"),
  method = rep("fe", 11),
  stringsAsFactors = FALSE
)

sig_mark <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "***",
                ifelse(p < 0.01, "**",
                       ifelse(p < 0.05, "*",
                              ifelse(p < 0.1, ".", "ns")))))
}

# ==============================================================================
# 5. RQ1 분석 함수 — Two-way FE (firm + year), 군집 표준오차
# ==============================================================================

run_rq1 <- function(sub, dv, method) {
  df <- sub %>%
    filter(!is.na(.data[[dv]]), #!is.na(log_emp), 
           !is.na(log_asset), !is.na(leverage))
  tryCatch({
    if (method == "ols") {
      fml <- as.formula(paste0(dv, " ~ did + treat + log_export + log_rdcost + 
                               log_asset + leverage + factor(year)"))
      m   <- lm(fml, data = df)
      rob <- coeftest(m, vcov = vcovCL(m, cluster = df$firm_id))
      data.frame(coef = rob["did",1], se = rob["did",2], stat = rob["did",3], pval = rob["did",4])
    } else {
      pdf <- pdata.frame(df, index = c("firm_id", "year"), drop.index = FALSE)
      fml <- as.formula(paste0(dv, " ~ did + log_emp + log_asset + leverage"))
      m   <- plm(fml, data = pdf, model = "within", effect = "twoways")
      rob <- coeftest(m, vcov = vcovHC(m, type = "HC1", cluster = "group"))
      data.frame(coef = rob["did",1], se = rob["did",2], stat = rob["did",3], pval = rob["did",4])
    }
  }, error = function(e) data.frame(coef = NA, se = NA, stat = NA, pval = NA))
}

# ==============================================================================
# 6. 부문별 RQ1 실행
# ==============================================================================

seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")

all_rq1 <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub <- panel %>% filter(seg == s)
  
  n_t <- length(unique(sub$firm_id[sub$treat == 1]))
  n_c <- length(unique(sub$firm_id[sub$treat == 0]))
  
  cat("\n", paste(rep("=", 80), collapse = ""), "\n")
  cat("  ", seg_name, " 부문 (처치:", n_t, "/ 통제:", n_c, ")\n")
  cat(paste(rep("=", 80), collapse = ""), "\n")
  
  cat(sprintf("  %-18s %8s %10s %10s %8s %5s\n",
              "카테고리/지표", "ATT", "SE", "stat", "p", "Sig"))
  cat("  ", paste(rep("-", 65), collapse = ""), "\n")
  
  rq1_rows <- list()
  for (i in 1:nrow(dv_info)) {
    r <- run_rq1(sub, dv_info$dv[i], dv_info$method[i])
    cat(sprintf("  [%s] %-12s %8.4f %10.4f %10.3f %8.4f %5s\n",
                dv_info$cat[i], dv_info$label[i],
                r$coef, r$se, r$stat, r$pval, sig_mark(r$pval)))
    rq1_rows[[i]] <- data.frame(
      부문 = seg_name, 카테고리 = dv_info$cat[i], 지표 = dv_info$label[i],
      ATT = r$coef, SE = r$se, t_z = r$stat, p_value = r$pval,
      유의성 = sig_mark(r$pval), stringsAsFactors = FALSE
    )
  }
  all_rq1[[seg_name]] <- bind_rows(rq1_rows)
}

# ==============================================================================
# 7. RQ1 결과 통합
# ==============================================================================

rq1_all <- bind_rows(all_rq1)

rq1_wide <- rq1_all %>%
  select(부문, 지표, ATT, p_value, 유의성) %>%
  pivot_wider(
    names_from  = 부문,
    values_from = c(ATT, p_value, 유의성),
    names_glue  = "{부문}_{.value}"
  ) %>%
  select(지표,
         소재_ATT, 소재_p_value, 소재_유의성,
         부품_ATT, 부품_p_value, 부품_유의성,
         장비_ATT, 장비_p_value, 장비_유의성)

cat("\n=== RQ1 부문별 비교 (Wide) ===\n")
print(rq1_wide, n = Inf)

# ==============================================================================
# 8. 시각화
# ==============================================================================

if (Sys.info()["sysname"] == "Darwin") {
  kfont <- "AppleGothic"
} else if (Sys.info()["sysname"] == "Windows") {
  kfont <- "Malgun Gothic"
} else {
  kfont <- "NanumGothic"
}

# ── (a) RQ1: 부문별 ATT dot plot ──
plot_rq1 <- rq1_all %>%
  mutate(
    지표 = factor(지표, levels = rev(dv_info$label)),
    부문 = factor(부문, levels = c("소재", "부품", "장비")),
    유의 = ifelse(p_value < 0.05, "유의 (p<0.05)", "비유의")
  )

p1 <- ggplot(plot_rq1, aes(x = ATT, y = 지표, color = 부문, shape = 유의)) +
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray40") +
  geom_point(size = 3, position = position_dodge(width = 0.6)) +
  geom_errorbarh(aes(xmin = ATT - 1.96 * SE, xmax = ATT + 1.96 * SE),
                 height = 0.2, position = position_dodge(width = 0.6)) +
  scale_color_manual(values = c("소재" = "#8E44AD", "부품" = "#2980B9", "장비" = "#27AE60")) +
  scale_shape_manual(values = c("유의 (p<0.05)" = 16, "비유의" = 1)) +
  labs(title = "RQ1: 부문별 평균 처치효과 (ATT)",
       subtitle = "PSM-DID Two-way FE | 소재·부품·장비 × 11개 종속변수",
       x = "ATT (DID Coefficient)", y = "", color = "부문", shape = "유의성") +
  theme_minimal(base_family = kfont) +
  theme(plot.title = element_text(face = "bold", size = 14), legend.position = "top")

ggsave("RQ1_ATT_by_segment.png", p1, width = 12, height = 8, dpi = 300)
cat("✓ RQ1_ATT_by_segment.png\n")

# ── (b) RQ1: 카테고리별 Facet bar ──
p1f <- ggplot(plot_rq1, aes(x = 지표, y = ATT, fill = 부문)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_errorbar(aes(ymin = ATT - 1.96 * SE, ymax = ATT + 1.96 * SE),
                position = position_dodge(width = 0.8), width = 0.2) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
  facet_wrap(~ 카테고리, scales = "free", ncol = 2) +
  scale_fill_manual(values = c("소재" = "#8E44AD", "부품" = "#2980B9", "장비" = "#27AE60")) +
  labs(title = "부문별 DID 추정량 (카테고리별)",
       x = "", y = "ATT", fill = "부문") +
  theme_minimal(base_family = kfont) +
  theme(plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1, size = 9),
        strip.text  = element_text(face = "bold", size = 12),
        legend.position = "top")

ggsave("RQ1_ATT_by_category_facet.png", p1f, width = 14, height = 10, dpi = 300)
cat("✓ RQ1_ATT_by_category_facet.png\n")

# ── (c) 평행추세: 부문별 주요 지표 ──
key_dvs    <- c("innov_patent", "log_asset_raw", "log_sales", "profit_roa",
                "stab_debtratio", "act_assetturn")
key_labels <- c("ln(특허+1)", "ln자산", "ln매출", "ln(ROA)",
                "ln부채비율", "ln총자산회전율")

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub <- panel %>% filter(seg == s)
  
  plots <- list()
  for (j in seq_along(key_dvs)) {
    trend <- sub %>%
      group_by(year, treat) %>%
      summarise(m  = mean(.data[[key_dvs[j]]], na.rm = TRUE),
                se = sd(.data[[key_dvs[j]]], na.rm = TRUE) / sqrt(n()),
                .groups = "drop") %>%
      mutate(grp = ifelse(treat == 1, "처치", "통제"))
    
    plots[[j]] <- ggplot(trend, aes(x = year, y = m, color = grp)) +
      geom_ribbon(aes(ymin = m - 1.96*se, ymax = m + 1.96*se, fill = grp),
                  alpha = 0.12, color = NA) +
      geom_line(linewidth = 1) + geom_point(size = 2) +
      geom_vline(xintercept = 2019.5, linetype = "dashed", color = "gray40") +
      scale_color_manual(values = c("처치" = "#e74c3c", "통제" = "#3498db")) +
      scale_fill_manual(values  = c("처치" = "#e74c3c", "통제" = "#3498db")) +
      labs(title = key_labels[j], x = "", y = "", color = "", fill = "") +
      scale_x_continuous(breaks = 2019:2024) +
      theme_minimal(base_family = kfont) +
      theme(legend.position = ifelse(j == 1, "top", "none"),
            plot.title = element_text(face = "bold", size = 10))
  }
  
  if (require(gridExtra, quietly = TRUE)) {
    g <- gridExtra::grid.arrange(
      grobs = plots, ncol = 3,
      top = grid::textGrob(paste0(seg_name, " 부문 평행추세"),
                           gp = grid::gpar(fontsize = 14, fontface = "bold",
                                           fontfamily = kfont))
    )
    ggsave(paste0("parallel_trends_", seg_name, ".png"),
           g, width = 14, height = 8, dpi = 300)
  } else {
    for (j in seq_along(plots)) {
      ggsave(paste0("trend_", seg_name, "_", key_dvs[j], ".png"),
             plots[[j]], width = 6, height = 4, dpi = 300)
    }
  }
  cat("✓ 평행추세:", seg_name, "\n")
}

# ==============================================================================
# 9. Simple DID (2019 vs 2024, log 수준 변수 기준)
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  Simple DID (2019 vs 2024 비교) — log 변환 변수\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

simple_vars <- data.frame(
  label = c("ln자산", "ln매출", "ln부채", "ln자본금",
            #"ln종업원수", 
            "ln영업이익", "ln특허","ln수출금액", "ln연구개발비"), # ← 추가
  var   = c("log_asset_raw", "log_sales", "log_debt", "log_capital",
            #"log_emp_raw", 
            "log_opprofit", "log_patent","log_export", "log_rdcost"), # ← 추가
  stringsAsFactors = FALSE
)

simple_all <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub_pre  <- panel %>% filter(seg == s, year == 2019)
  sub_post <- panel %>% filter(seg == s, year == 2024)
  
  sub_wide <- sub_pre %>%
    select(firm_id, treat) %>%
    left_join(
      sub_pre  %>% select(firm_id, all_of(simple_vars$var)) %>%
        rename_with(~paste0(.x, "_pre"),  -firm_id), by = "firm_id") %>%
    left_join(
      sub_post %>% select(firm_id, all_of(simple_vars$var)) %>%
        rename_with(~paste0(.x, "_post"), -firm_id), by = "firm_id")
  
  cat("\n───", seg_name, "───\n")
  cat(sprintf("  %-12s %12s %12s %12s %12s %12s %8s %5s\n",
              "변수", "처치_pre", "처치_post", "통제_pre", "통제_post", "DID", "t", "sig"))
  cat("  ", paste(rep("-", 90), collapse = ""), "\n")
  
  seg_rows <- list()
  for (v in 1:nrow(simple_vars)) {
    col_pre  <- paste0(simple_vars$var[v], "_pre")
    col_post <- paste0(simple_vars$var[v], "_post")
    sub_wide$diff <- sub_wide[[col_post]] - sub_wide[[col_pre]]
    
    t_mean_pre  <- mean(sub_wide[[col_pre]] [sub_wide$treat == 1], na.rm = TRUE)
    t_mean_post <- mean(sub_wide[[col_post]][sub_wide$treat == 1], na.rm = TRUE)
    c_mean_pre  <- mean(sub_wide[[col_pre]] [sub_wide$treat == 0], na.rm = TRUE)
    c_mean_post <- mean(sub_wide[[col_post]][sub_wide$treat == 0], na.rm = TRUE)
    
    reg     <- lm(diff ~ treat, data = sub_wide)
    did_est <- coef(reg)["treat"]
    t_val   <- summary(reg)$coefficients["treat", "t value"]
    p_val   <- summary(reg)$coefficients["treat", "Pr(>|t|)"]
    
    cat(sprintf("  %-12s %12.4f %12.4f %12.4f %12.4f %12.4f %8.3f %5s\n",
                simple_vars$label[v],
                t_mean_pre, t_mean_post, c_mean_pre, c_mean_post,
                did_est, t_val, sig_mark(p_val)))
    
    seg_rows[[v]] <- data.frame(
      부문 = seg_name, 변수 = simple_vars$label[v],
      처치_pre = t_mean_pre, 처치_post = t_mean_post,
      통제_pre = c_mean_pre, 통제_post = c_mean_post,
      DID = did_est, t_value = t_val, p_value = p_val,
      유의성 = sig_mark(p_val), stringsAsFactors = FALSE
    )
  }
  simple_all[[seg_name]] <- bind_rows(seg_rows)
}

simple_result <- bind_rows(simple_all)

# Wide format
simple_wide <- simple_result %>%
  select(부문, 변수, DID, p_value, 유의성) %>%
  pivot_wider(
    names_from  = 부문,
    values_from = c(DID, p_value, 유의성),
    names_glue  = "{부문}_{.value}"
  ) %>%
  select(변수,
         소재_DID, 소재_p_value, 소재_유의성,
         부품_DID, 부품_p_value, 부품_유의성,
         장비_DID, 장비_p_value, 장비_유의성)

cat("\n=== Simple DID 부문별 비교 (Wide) ===\n")
print(simple_wide)

# ==============================================================================
# 10. Excel 저장
# ==============================================================================

write_xlsx(list(
  RQ1_ATT         = rq1_all,
  RQ1_Wide        = rq1_wide,
  Simple_DID      = simple_result,
  Simple_DID_Wide = simple_wide
), "DID_RQ1_SimpleDID.xlsx")

cat("\n✓ 저장 완료: DID_RQ1_SimpleDID.xlsx\n")
cat("\n=== 분석 완료 ===\n")