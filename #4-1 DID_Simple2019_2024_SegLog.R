# ==============================================================================
# 소재(seg=1) / 부품(seg=2) / 장비(seg=3)별 DID 분석
# ① Simple DID (2019 vs 2024)  ② 변수별 평행추세 검정 + 시각화
# 수출금액, 개발비, 특허 포함 | 20260310 수정
# ==============================================================================

packages <- c("readxl", "dplyr", "tidyr", "ggplot2", "broom", "sandwich",
              "lmtest", "car", "writexl", "knitr", "scales",
              "gridExtra", "grid")

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
cat("seg 분포:\n");   print(table(matched$seg))
cat("treat 분포:\n"); print(table(matched$treat))

to_num <- function(x) as.numeric(x)

# ------------------------------------------------------------------------------
# 부호 보존 log 변환 (음수값 처리)
# ------------------------------------------------------------------------------
log_signed <- function(x) {
  ifelse(is.na(x), NA, sign(x) * log1p(abs(x)))
}

# NA/NaN 숫자값을 sprintf에 안전하게 전달
safe_num <- function(x) ifelse(is.na(x) | is.nan(x), 0, x)
safe_str <- function(x) ifelse(is.na(x) | x == "NA", "ns", x)

sig_mark <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "***",
                ifelse(p < 0.01,  "**",
                       ifelse(p < 0.05,  "*",
                              ifelse(p < 0.1,   ".",  "ns")))))
}

# ==============================================================================
# 1. Wide → Long 패널 변환 (2018–2024)
# ==============================================================================

safe_col <- function(row, col_name) {
  if (col_name %in% names(row) && !is.null(row[[col_name]])) {
    to_num(row[[col_name]])
  } else {
    NA_real_
  }
}


years <- 2018:2024

panel_list <- lapply(1:nrow(matched), function(i) {
  row <- matched[i, ]
  data.frame(
    firm_id  = i,
    seg      = to_num(row$seg),
    seg_name = row$seg_name,
    treat    = to_num(row$treat),
    n_funded = nchar(gsub("0", "", row$fundedpattern)),
    year     = years,
    asset    = sapply(years, function(y) safe_col(row, paste0(y, "/Annual S15000.자산총계"))),
    debt     = sapply(years, function(y) safe_col(row, paste0(y, "/Annual S18000.부채총계"))),
    equity   = sapply(years, function(y) safe_col(row, paste0(y, "/Annual S18900.자본총계"))),
    capital  = sapply(years, function(y) safe_col(row, paste0(y, "/Annual S18100.자본금"))),
    sales    = sapply(years, function(y) safe_col(row, paste0(y, "/Annual S21100.총수익"))),
    op_profit= sapply(years, function(y) safe_col(row, paste0(y, "/Annual S25000.영업이익(손실)"))),
    patent   = sapply(years, function(y) safe_col(row, paste0("p", y))),
    export   = sapply(years, function(y) safe_col(row, paste0("exportamt", y))),
    rdcost   = sapply(years, function(y) safe_col(row, paste0("rdcost", y))),
    lbcost   = sapply(years, function(y) safe_col(row, paste0("lbcost", y))),
    mflbcost = sapply(years, function(y) safe_col(row, paste0("mflbcost", y))),
    stringsAsFactors = FALSE
  )
})

panel <- bind_rows(panel_list) %>% arrange(firm_id, year)
cat("패널:", nrow(panel), "obs /", length(unique(panel$firm_id)), "firms\n")

# 어느 연도·컬럼에서 NA가 발생했는지 확인
panel %>%
  filter(year %in% c(2018, 2019)) %>%
  summarise(
    na_lbcost   = sum(is.na(lbcost)),
    na_mflbcost = sum(is.na(mflbcost)),
    na_rdcost   = sum(is.na(rdcost)),
    na_export   = sum(is.na(export))
  ) %>%
  print()

# ==============================================================================
# 2. Log 변환
# ==============================================================================

panel <- panel %>%
  mutate(
    log_patent  = log1p(pmax(patent,   0, na.rm = TRUE)),
    log_asset   = ifelse(!is.na(asset)  & asset  > 0, log(asset),  NA),
    log_debt    = log1p(pmax(debt,     0, na.rm = TRUE)),
    log_capital = log1p(pmax(capital,  0, na.rm = TRUE)),
    log_sales   = ifelse(!is.na(sales)  & sales  > 0, log(sales),  NA),
    log_opprofit= log_signed(op_profit),
    log_export  = log1p(pmax(export,   0, na.rm = TRUE)),
    log_rdcost  = log1p(pmax(rdcost,   0, na.rm = TRUE)),
    log_lbcost  = log1p(pmax(lbcost, 0, na.rm = TRUE)),
    log_mflbcost= log1p(pmax(mflbcost, 0, na.rm = TRUE)),
    post = ifelse(year >= 2020, 1, 0),
    did  = treat * post
  )

seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")

# ==============================================================================
# 3. Simple DID (2019 vs 2024)
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  Simple DID (Pre=2019 vs Post=2024)\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

simple_vars <- data.frame(
  label = c("ln자산",    "ln매출",    "ln부채",     "ln자본금",
            "ln영업이익", "ln(특허+1)", "ln(수출+1)", "ln(개발비+1)",
            "ln(인건비+1)", "ln(노무비+1)"),   # ← 2개 추가
  var   = c("log_asset", "log_sales", "log_debt",    "log_capital",
            "log_opprofit", "log_patent", "log_export", "log_rdcost",
            "log_lbcost", "log_mflbcost"),        # ← 기존 그대로
  cat   = c("성장성", "성장성", "안정성", "성장성",
            "수익성", "혁신성", "활동성", "혁신성",
            "활동성", "활동성"),                   # ← 2개 추가
  stringsAsFactors = FALSE
)

simple_all <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub_pre  <- panel %>% filter(seg == s, year == 2019)
  sub_post <- panel %>% filter(seg == s, year == 2024)
  
  # pre/post wide 결합
  sub_wide <- sub_pre %>%
    select(firm_id, treat) %>%
    left_join(
      sub_pre  %>% select(firm_id, all_of(simple_vars$var)) %>%
        rename_with(~ paste0(.x, "_pre"),  -firm_id),
      by = "firm_id") %>%
    left_join(
      sub_post %>% select(firm_id, all_of(simple_vars$var)) %>%
        rename_with(~ paste0(.x, "_post"), -firm_id),
      by = "firm_id")
  
  cat("\n───", seg_name, "───\n")
  cat(sprintf("  %-16s %10s %10s %10s %10s %10s %8s %5s\n",
              "변수", "처치_pre", "처치_post", "통제_pre", "통제_post",
              "DID", "t값", "유의"))
  cat("  ", paste(rep("-", 85), collapse = ""), "\n")
  
  seg_rows <- list()
  for (v in 1:nrow(simple_vars)) {
    col_pre  <- paste0(simple_vars$var[v], "_pre")
    col_post <- paste0(simple_vars$var[v], "_post")
    sub_wide$diff <- sub_wide[[col_post]] - sub_wide[[col_pre]]
    
    t_pre  <- mean(sub_wide[[col_pre]] [sub_wide$treat == 1], na.rm = TRUE)
    t_post <- mean(sub_wide[[col_post]][sub_wide$treat == 1], na.rm = TRUE)
    c_pre  <- mean(sub_wide[[col_pre]] [sub_wide$treat == 0], na.rm = TRUE)
    c_post <- mean(sub_wide[[col_post]][sub_wide$treat == 0], na.rm = TRUE)
    
    reg    <- lm(diff ~ treat, data = sub_wide)
    did_est<- coef(reg)["treat"]
    t_val  <- summary(reg)$coefficients["treat", "t value"]
    p_val  <- summary(reg)$coefficients["treat", "Pr(>|t|)"]
    
    cat(sprintf("  %-16s %10.4f %10.4f %10.4f %10.4f %10.4f %8.3f %5s\n",
                simple_vars$label[v],
                safe_num(t_pre), safe_num(t_post), safe_num(c_pre), safe_num(c_post),
                safe_num(did_est), safe_num(t_val), safe_str(sig_mark(p_val))))
    
    seg_rows[[v]] <- data.frame(
      부문 = seg_name, 카테고리 = simple_vars$cat[v],
      변수 = simple_vars$label[v],
      처치_pre = t_pre, 처치_post = t_post,
      통제_pre = c_pre, 통제_post = c_post,
      DID = did_est, t_value = t_val, p_value = p_val,
      유의성 = sig_mark(p_val),
      stringsAsFactors = FALSE
    )
  }
  simple_all[[seg_name]] <- bind_rows(seg_rows)
}

simple_result <- bind_rows(simple_all) %>% as.data.frame()

# Wide format 출력
simple_wide <- simple_result %>%
  select(변수, 부문, DID, p_value, 유의성) %>%
  pivot_wider(
    names_from  = 부문,
    values_from = c(DID, p_value, 유의성),
    names_glue  = "{부문}_{.value}"
  ) %>%
  select(변수,
         소재_DID, 소재_p_value, 소재_유의성,
         부품_DID, 부품_p_value, 부품_유의성,
         장비_DID, 장비_p_value, 장비_유의성) %>%
  as.data.frame()

cat("\n=== Simple DID Wide ===\n")
print(as.data.frame(simple_wide))

# ==============================================================================
# 4. 변수별 평행추세 검정 (pre-trend t-test: 2018→2019 기울기 차이)
# 2018->2019로 해야 함.
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  평행추세 검정 (Pre-trend Test: 2018→2019 slope difference)\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

pt_all <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub <- panel %>% filter(seg == s)
  
  cat(sprintf("\n─── %s ───\n", seg_name))
  cat(sprintf("  %-16s %8s %8s %6s\n", "변수", "t값", "p값", "판정"))
  cat("  ", paste(rep("-", 44), collapse = ""), "\n")
  
  pt_rows <- list()
  for (v in 1:nrow(simple_vars)) {
    varname <- simple_vars$var[v]
    label   <- simple_vars$label[v]
    
    # 2019→2020 차분
    d19 <- sub %>% filter(year == 2018) %>% select(firm_id, treat, val19 = all_of(varname))
    d20 <- sub %>% filter(year == 2019) %>% select(firm_id, val20 = all_of(varname))
    d   <- left_join(d19, d20, by = "firm_id") %>%
      mutate(slope = val20 - val19) %>%
      filter(!is.na(slope))
    
    treat_slope <- d$slope[d$treat == 1]
    ctrl_slope  <- d$slope[d$treat == 0]
    
    if (length(treat_slope) < 3 || length(ctrl_slope) < 3) {
      pt_rows[[v]] <- data.frame(부문=seg_name, 변수=label, t값=NA, p값=NA, 판정="n/a")
      next
    }
    
    tt   <- t.test(treat_slope, ctrl_slope, var.equal = FALSE)
    t_v  <- tt$statistic
    p_v  <- tt$p.value
    judg <- ifelse(p_v > 0.1, "✅ 충족",
                   ifelse(p_v > 0.05, "⚠ 경계", "❌ 위반"))
    
    cat(sprintf("  %-16s %8.3f %8.4f  %s\n", label, safe_num(t_v), safe_num(p_v), judg))
    
    pt_rows[[v]] <- data.frame(
      부문 = seg_name, 변수 = label,
      t값  = round(t_v, 4), p값 = round(p_v, 4),
      판정 = judg,
      stringsAsFactors = FALSE
    )
  }
  pt_all[[seg_name]] <- bind_rows(pt_rows)
}

pt_result <- bind_rows(pt_all) %>% as.data.frame()

cat("\n=== 평행추세 검정 결과 요약 ===\n")
print(as.data.frame(pt_result))

# ==============================================================================
# 5. 시각화 설정 (한글 폰트)
# ==============================================================================

if (Sys.info()["sysname"] == "Darwin") {
  kfont <- "AppleGothic"
} else if (Sys.info()["sysname"] == "Windows") {
  kfont <- "Malgun Gothic"
} else {
  kfont <- "NanumGothic"
}

col_treat <- "#E05555"
col_ctrl  <- "#4A90D9"

# ==============================================================================
# 6. Simple DID 시각화 — 부문별 DID 계수 dot plot (변수 × 부문)
# ==============================================================================

plot_simple <- simple_result %>%
  mutate(
    변수 = factor(변수, levels = rev(simple_vars$label)),
    부문 = factor(부문, levels = c("소재", "부품", "장비")),
    유의 = ifelse(p_value < 0.05, "유의(p<0.05)", "비유의")
  )

p_did <- ggplot(plot_simple, aes(x = DID, y = 변수, color = 부문, shape = 유의)) +
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray40") +
  geom_point(size = 3.5, position = position_dodge(width = 0.6)) +
  scale_color_manual(values = c("소재" = "#8E44AD", "부품" = "#2980B9", "장비" = "#27AE60")) +
  scale_shape_manual(values = c("유의(p<0.05)" = 16, "비유의" = 1)) +
  facet_wrap(~ 카테고리, scales = "free_x", ncol = 2) +
  labs(
    title    = "Simple DID 추정량 (Pre=2019, Post=2024)",
    subtitle = "변수별 · 부문별 비교 | ln 변환 기준",
    x = "DID 계수", y = NULL, color = "부문", shape = "유의성"
  ) +
  theme_minimal(base_family = kfont) +
  theme(
    plot.title   = element_text(face = "bold", size = 14),
    legend.position = "top",
    strip.text   = element_text(face = "bold", size = 11),
    axis.text.y  = element_text(size = 9)
  )

ggsave("SimpleDID_dotplot.png", p_did, width = 13, height = 9, dpi = 300)
cat("\n✓ SimpleDID_dotplot.png 저장\n")

# ==============================================================================
# 7. 변수별 평행추세 시각화 — 부문 × 변수 전체 그리드 (3 × 8)
# ==============================================================================

# 처치/통제 집단 연도별 평균 계산
trend_data <- panel %>%
  group_by(seg, year, treat) %>%
  summarise(
    across(all_of(simple_vars$var),
           ~ mean(.x, na.rm = TRUE), .names = "{.col}"),
    .groups = "drop"
  ) %>%
  mutate(
    seg_name = seg_labels[as.character(seg)],
    grp      = ifelse(treat == 1, "처치", "통제")
  )

# 부문별로 각 변수의 평행추세 플롯 저장
for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub_trend <- trend_data %>% filter(seg == s)
  
  # 각 변수별 플롯 생성
  plot_list <- lapply(1:nrow(simple_vars), function(v) {
    varname <- simple_vars$var[v]
    label   <- simple_vars$label[v]
    
    # pre-trend p값 가져오기
    pt_row <- pt_result %>% filter(부문 == seg_name, 변수 == label)
    p_lab  <- if (nrow(pt_row) > 0 && !is.na(pt_row$p값))
      sprintf("pre-trend p=%.3f (%s)", pt_row$p값, sig_mark(pt_row$p값))
    else "pre-trend: n/a"
    
    df_v <- sub_trend %>%
      select(year, grp, val = all_of(varname)) %>%
      filter(!is.na(val))
    
    # 95% CI를 위한 SE 재계산 (전체 패널에서)
    se_df <- panel %>%
      filter(seg == s, !is.na(.data[[varname]])) %>%
      group_by(year, treat) %>%
      summarise(
        m  = mean(.data[[varname]], na.rm = TRUE),
        se = sd(.data[[varname]],   na.rm = TRUE) / sqrt(n()),
        .groups = "drop"
      ) %>%
      mutate(grp = ifelse(treat == 1, "처치", "통제"))
    
    ggplot(se_df, aes(x = year, y = m, color = grp, fill = grp)) +
      # 신뢰구간 리본
      geom_ribbon(aes(ymin = m - 1.96*se, ymax = m + 1.96*se),
                  alpha = 0.12, color = NA) +
      geom_line(linewidth = 1.1) +
      geom_point(size = 2.2) +
      # 지원 기간 음영
      annotate("rect", xmin = 2019.5, xmax = 2022.5,
               ymin = -Inf, ymax = Inf, alpha = 0.06, fill = "gold") +
      # 처치 시작선
      geom_vline(xintercept = 2019.5, linetype = "dashed",
                 color = "gray40", linewidth = 0.6) +
      # pre-trend 판정 레이블
      annotate("text", x = 2019.6, y = Inf,
               label = p_lab, hjust = 0, vjust = 1.5,
               size = 2.8, color = "gray30") +
      scale_color_manual(values = c("처치" = col_treat, "통제" = col_ctrl)) +
      scale_fill_manual( values = c("처치" = col_treat, "통제" = col_ctrl)) +
      scale_x_continuous(breaks = 2019:2024) +
      labs(
        title = label,
        x = NULL, y = NULL, color = NULL, fill = NULL
      ) +
      theme_minimal(base_family = kfont) +
      theme(
        plot.title      = element_text(face = "bold", size = 9.5),
        axis.text.x     = element_text(size = 7.5, angle = 30, hjust = 1),
        axis.text.y     = element_text(size = 7.5),
        legend.position = if (v == 1) "top" else "none",
        legend.text     = element_text(size = 8),
        panel.grid.minor= element_blank()
      )
  })
  
  # 8개 변수 2행 4열 배치
  g <- gridExtra::grid.arrange(
    grobs = plot_list,
    ncol  = 4,
    top   = grid::textGrob(
      paste0("평행추세 검정: ", seg_name, " 부문 (처치 vs 통제)"),
      gp = grid::gpar(fontsize = 13, fontface = "bold", fontfamily = kfont)
    )
  )
  
  fname <- paste0("ParallelTrends_", seg_name, ".png")
  ggsave(fname, g, width = 18, height = 9, dpi = 300)
  cat("✓", fname, "저장\n")
}

# ==============================================================================
# 8. 평행추세 히트맵 — 3부문 × 8변수 판정 요약
# ==============================================================================

pt_heat <- pt_result %>%
  mutate(
    판정색 = case_when(
      p값 > 0.1  ~ "충족",
      p값 > 0.05 ~ "경계",
      TRUE       ~ "위반"
    ),
    부문 = factor(부문, levels = c("소재", "부품", "장비")),
    변수 = factor(변수, levels = rev(simple_vars$label))
  )

p_heat <- ggplot(pt_heat, aes(x = 부문, y = 변수, fill = 판정색)) +
  geom_tile(color = "white", linewidth = 0.8) +
  geom_text(aes(label = sprintf("p=%.3f", p값)), size = 3.2, color = "gray20") +
  scale_fill_manual(
    values = c("충족" = "#A8D5A2", "경계" = "#FFE08A", "위반" = "#F4A6A0"),
    name   = "평행추세"
  ) +
  labs(
    title    = "평행추세 검정 결과 요약 (Pre-trend Test)",
    subtitle = "2019→2020 기울기 차이 t-test | 초록=충족, 노랑=경계(p<0.1), 빨강=위반(p<0.05)",
    x = "부문", y = "변수"
  ) +
  theme_minimal(base_family = kfont) +
  theme(
    plot.title    = element_text(face = "bold", size = 13),
    plot.subtitle = element_text(size = 9, color = "gray40"),
    axis.text     = element_text(size = 10),
    legend.position = "top"
  )

ggsave("ParallelTrends_Heatmap.png", p_heat, width = 8, height = 7, dpi = 300)
cat("✓ ParallelTrends_Heatmap.png 저장\n")

# ==============================================================================
# 9. Excel 저장
# ==============================================================================

write_xlsx(
  list(
    Simple_DID        = simple_result,
    Simple_DID_Wide   = simple_wide,
    ParallelTrends    = pt_result
  ),
  "DID_SimpleDID_PT.xlsx"
)

cat("\n✓ 저장 완료: DID_SimpleDID_PT.xlsx\n")
cat("\n=== 분석 완료 ===\n")

