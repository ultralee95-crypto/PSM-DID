# ==============================================================================
# 투자 횟수별(1회/2회/3회) × 부문별(소재/부품/장비) Simple DID 분석
# - 처치집단: 2020·2021·2022 중 지원받은 횟수(n_funded)로 세분
# - 통제집단: 동일 부문 내 PSM 매칭된 미지원 기업(n_funded == 0)
# - Pre=2019, Post=2024, 변수: 자산·매출·부채·자본금·영업이익·특허·수출·개발비
# - 출력: ① 콘솔 수치표  ② Wide 비교표  ③ 평행추세 검정
#         ④ 시각화 4종  ⑤ Excel 저장
# 20260310
# ==============================================================================

packages <- c("readxl", "dplyr", "tidyr", "ggplot2",
              "sandwich", "lmtest", "writexl",
              "gridExtra", "grid", "scales")

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}

select <- dplyr::select
lag    <- dplyr::lag

setwd("/Users/ultra/PSM-DID")

matched <- read_excel("matched_dataset_segPSM.xlsx", col_types = "text")

cat("=== 데이터 확인 ===\n")
cat("행수:", nrow(matched), "\n")
cat("fundedpattern 분포:\n"); print(table(matched$fundedpattern))

# ==============================================================================
# 0. 공통 함수
# ==============================================================================

to_num    <- function(x) as.numeric(x)
log_signed <- function(x) ifelse(is.na(x), NA, sign(x) * log1p(abs(x)))
safe_num  <- function(x) ifelse(is.na(x) | is.nan(x), 0, x)
safe_str  <- function(x) ifelse(is.na(x) | x == "NA", "ns", x)

sig_mark <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "***",
                ifelse(p < 0.01, "**",
                       ifelse(p < 0.05, "*",
                              ifelse(p < 0.1, ".", "ns")))))
}

# ==============================================================================
# 1. Wide → Long 패널 변환 (2019–2024)
# ==============================================================================

years <- 2019:2024

panel_list <- lapply(1:nrow(matched), function(i) {
  row <- matched[i, ]
  # 투자 횟수: fundedpattern에서 "1"의 개수
  pattern   <- as.character(row$fundedpattern)
  n_funded  <- nchar(gsub("0", "", pattern))
  
  data.frame(
    firm_id    = i,
    seg        = to_num(row$seg),
    seg_name   = as.character(row$seg_name),
    treat      = to_num(row$treat),
    n_funded   = n_funded,               # 0 / 1 / 2 / 3
    fundedpattern = pattern,
    year       = years,
    asset      = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S15000.자산총계")]])),
    debt       = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18000.부채총계")]])),
    equity     = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18900.자본총계")]])),
    capital    = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18100.자본금")]])),
    sales      = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S21100.총매출액")]])),
    op_profit  = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S25000.영업이익(손실)")]])),
    patent     = sapply(years, function(y) to_num(row[[paste0("p", y)]])),
    export     = sapply(years, function(y) to_num(row[[paste0("exportamt", y)]])),
    rdcost     = sapply(years, function(y) to_num(row[[paste0("rdcost", y)]])),
    stringsAsFactors = FALSE
  )
})

panel <- bind_rows(panel_list) %>% arrange(firm_id, year)
cat("\n패널:", nrow(panel), "obs /", length(unique(panel$firm_id)), "firms\n")

# ==============================================================================
# 2. Log 변환
# ==============================================================================

panel <- panel %>%
  mutate(
    log_asset    = ifelse(!is.na(asset)  & asset  > 0, log(asset),  NA),
    log_sales    = ifelse(!is.na(sales)  & sales  > 0, log(sales),  NA),
    log_debt     = log1p(pmax(debt,    0, na.rm = TRUE)),
    log_capital  = log1p(pmax(capital, 0, na.rm = TRUE)),
    log_opprofit = log_signed(op_profit),
    log_patent   = log1p(pmax(patent,  0, na.rm = TRUE)),
    log_export   = log1p(pmax(export,  0, na.rm = TRUE)),
    log_rdcost   = log1p(pmax(rdcost,  0, na.rm = TRUE))
  )

# 부문 라벨
seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")

# n_funded 분포 확인
cat("\n=== n_funded × seg 분포 (처치 기업) ===\n")
print(
  panel %>%
    filter(year == 2019) %>%
    count(seg_name, n_funded) %>%
    pivot_wider(names_from = n_funded, values_from = n, values_fill = 0) %>%
    as.data.frame()
)

# ==============================================================================
# 3. 분석 변수 정의
# ==============================================================================

ana_vars <- data.frame(
  label = c("ln자산",   "ln매출",    "ln부채",   "ln자본금",
            "ln영업이익", "ln(특허+1)", "ln(수출+1)", "ln(개발비+1)"),
  var   = c("log_asset", "log_sales", "log_debt", "log_capital",
            "log_opprofit", "log_patent", "log_export", "log_rdcost"),
  cat   = c("성장성", "성장성", "안정성", "성장성",
            "수익성", "혁신성", "활동성", "혁신성"),
  stringsAsFactors = FALSE
)

n_funded_levels <- c(1, 2, 3)
n_funded_labels <- c("1" = "1회", "2" = "2회", "3" = "3회")

# ==============================================================================
# 4. Simple DID — 투자 횟수별 × 부문별
#    처치(n회 지원) vs 통제(동일 부문 내 n_funded==0)
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  투자 횟수별 Simple DID (Pre=2019, Post=2024)\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

did_all <- list()  # 전체 결과 누적

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  
  # ── 통제집단 (해당 부문, n_funded == 0, 2019/2024) ──
  ctrl_pre  <- panel %>% filter(seg == s, n_funded == 0, year == 2019)
  ctrl_post <- panel %>% filter(seg == s, n_funded == 0, year == 2024)
  
  ctrl_wide <- ctrl_pre %>%
    select(firm_id) %>%
    left_join(
      ctrl_pre  %>% select(firm_id, all_of(ana_vars$var)) %>%
        rename_with(~ paste0(.x, "_pre"),  -firm_id), by = "firm_id") %>%
    left_join(
      ctrl_post %>% select(firm_id, all_of(ana_vars$var)) %>%
        rename_with(~ paste0(.x, "_post"), -firm_id), by = "firm_id") %>%
    mutate(treat_group = "통제", n_funded = 0)
  
  for (n in n_funded_levels) {
    n_label <- n_funded_labels[as.character(n)]
    
    # ── 처치집단 (해당 부문, n_funded == n) ──
    tr_pre  <- panel %>% filter(seg == s, n_funded == n, year == 2019)
    tr_post <- panel %>% filter(seg == s, n_funded == n, year == 2024)
    
    n_treat <- nrow(tr_pre)
    n_ctrl  <- nrow(ctrl_pre)
    
    if (n_treat < 5) {
      cat(sprintf("\n  [%s - %s] 처치 표본 부족 (N=%d), 건너뜀\n",
                  seg_name, n_label, n_treat))
      next
    }
    
    tr_wide <- tr_pre %>%
      select(firm_id) %>%
      left_join(
        tr_pre  %>% select(firm_id, all_of(ana_vars$var)) %>%
          rename_with(~ paste0(.x, "_pre"),  -firm_id), by = "firm_id") %>%
      left_join(
        tr_post %>% select(firm_id, all_of(ana_vars$var)) %>%
          rename_with(~ paste0(.x, "_post"), -firm_id), by = "firm_id") %>%
      mutate(treat_group = "처치", n_funded = n)
    
    # 처치 + 통제 합치기 (treat 더미: 처치=1, 통제=0)
    combined <- bind_rows(
      tr_wide   %>% mutate(treat_dummy = 1L),
      ctrl_wide %>% mutate(treat_dummy = 0L)
    )
    
    # ── 헤더 출력 ──
    cat(sprintf(
      "\n─── %s  |  %s  (처치 N=%d, 통제 N=%d) ───\n",
      seg_name, n_label, n_treat, n_ctrl))
    cat(sprintf(
      "  %-16s %10s %10s %10s %10s %10s %8s %5s\n",
      "변수", "처치_pre", "처치_post", "통제_pre", "통제_post", "DID", "t값", "유의"))
    cat("  ", paste(rep("-", 87), collapse = ""), "\n")
    
    seg_n_rows <- list()
    
    for (v in 1:nrow(ana_vars)) {
      col_pre  <- paste0(ana_vars$var[v], "_pre")
      col_post <- paste0(ana_vars$var[v], "_post")
      
      combined$diff <- combined[[col_post]] - combined[[col_pre]]
      
      t_pre  <- mean(combined[[col_pre]] [combined$treat_dummy == 1], na.rm = TRUE)
      t_post <- mean(combined[[col_post]][combined$treat_dummy == 1], na.rm = TRUE)
      c_pre  <- mean(combined[[col_pre]] [combined$treat_dummy == 0], na.rm = TRUE)
      c_post <- mean(combined[[col_post]][combined$treat_dummy == 0], na.rm = TRUE)
      
      # DID = lm(diff ~ treat_dummy)
      reg     <- lm(diff ~ treat_dummy, data = combined)
      did_est <- coef(reg)["treat_dummy"]
      t_val   <- summary(reg)$coefficients["treat_dummy", "t value"]
      p_val   <- summary(reg)$coefficients["treat_dummy", "Pr(>|t|)"]
      
      cat(sprintf(
        "  %-16s %10.4f %10.4f %10.4f %10.4f %10.4f %8.3f %5s\n",
        ana_vars$label[v],
        safe_num(t_pre), safe_num(t_post),
        safe_num(c_pre), safe_num(c_post),
        safe_num(did_est), safe_num(t_val),
        safe_str(sig_mark(p_val))))
      
      seg_n_rows[[v]] <- data.frame(
        부문      = seg_name,
        투자횟수  = n_label,
        n_funded  = n,
        카테고리  = ana_vars$cat[v],
        변수      = ana_vars$label[v],
        N_처치    = n_treat,
        N_통제    = n_ctrl,
        처치_pre  = safe_num(t_pre),
        처치_post = safe_num(t_post),
        통제_pre  = safe_num(c_pre),
        통제_post = safe_num(c_post),
        DID       = safe_num(did_est),
        t_value   = safe_num(t_val),
        p_value   = safe_num(p_val),
        유의성    = safe_str(sig_mark(p_val)),
        stringsAsFactors = FALSE
      )
    }
    
    key <- paste0(seg_name, "_", n_label)
    did_all[[key]] <- bind_rows(seg_n_rows)
  }
}

did_result <- bind_rows(did_all) %>% as.data.frame()

# ==============================================================================
# 5. Wide 비교표 — 투자횟수 × 변수  (부문별 파일로 분리)
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  Wide 비교표 (DID 추정량 | 부문 × 투자횟수 × 변수)\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

wide_list <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  
  wide <- did_result %>%
    filter(부문 == seg_name) %>%
    select(변수, 투자횟수, DID, p_value, 유의성) %>%
    pivot_wider(
      names_from  = 투자횟수,
      values_from = c(DID, p_value, 유의성),
      names_glue  = "{투자횟수}_{.value}"
    ) %>%
    select(변수,
           `1회_DID`, `1회_p_value`, `1회_유의성`,
           `2회_DID`, `2회_p_value`, `2회_유의성`,
           `3회_DID`, `3회_p_value`, `3회_유의성`) %>%
    as.data.frame()
  
  cat(sprintf("\n=== %s 부문 ===\n", seg_name))
  print(wide)
  
  wide_list[[seg_name]] <- wide
}

# ==============================================================================
# 6. 평행추세 검정 — 투자횟수별 × 부문별 (2019→2020 기울기 차이)
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  평행추세 검정 (Pre-trend: 2019→2020 slope t-test)\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

pt_all <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  
  # 통제집단 2019→2020 차분
  ctrl_sub <- panel %>% filter(seg == s, n_funded == 0)
  ctrl_d19 <- ctrl_sub %>% filter(year == 2019) %>%
    select(firm_id, all_of(ana_vars$var))
  ctrl_d20 <- ctrl_sub %>% filter(year == 2020) %>%
    select(firm_id, all_of(ana_vars$var))
  ctrl_diff <- left_join(ctrl_d19, ctrl_d20, by = "firm_id",
                         suffix = c("_19", "_20"))
  
  for (n in n_funded_levels) {
    n_label <- n_funded_labels[as.character(n)]
    tr_sub  <- panel %>% filter(seg == s, n_funded == n)
    
    if (nrow(tr_sub %>% filter(year == 2019)) < 5) next
    
    tr_d19  <- tr_sub %>% filter(year == 2019) %>%
      select(firm_id, all_of(ana_vars$var))
    tr_d20  <- tr_sub %>% filter(year == 2020) %>%
      select(firm_id, all_of(ana_vars$var))
    tr_diff <- left_join(tr_d19, tr_d20, by = "firm_id",
                         suffix = c("_19", "_20"))
    
    cat(sprintf("\n─── %s | %s ───\n", seg_name, n_label))
    cat(sprintf("  %-16s %8s %8s  %s\n", "변수", "t값", "p값", "판정"))
    cat("  ", paste(rep("-", 46), collapse = ""), "\n")
    
    for (v in 1:nrow(ana_vars)) {
      vname <- ana_vars$var[v]
      label <- ana_vars$label[v]
      
      tr_slope <- tr_diff[[paste0(vname, "_20")]] -
        tr_diff[[paste0(vname, "_19")]]
      ct_slope <- ctrl_diff[[paste0(vname, "_20")]] -
        ctrl_diff[[paste0(vname, "_19")]]
      
      tr_slope <- tr_slope[!is.na(tr_slope)]
      ct_slope <- ct_slope[!is.na(ct_slope)]
      
      if (length(tr_slope) < 3 || length(ct_slope) < 3) {
        pt_all[[paste(seg_name, n_label, label)]] <-
          data.frame(부문=seg_name, 투자횟수=n_label, 변수=label,
                     t값=NA_real_, p값=NA_real_, 판정="n/a",
                     stringsAsFactors=FALSE)
        next
      }
      
      tt   <- t.test(tr_slope, ct_slope, var.equal = FALSE)
      t_v  <- as.numeric(tt$statistic)
      p_v  <- tt$p.value
      judg <- ifelse(p_v > 0.1,  "✅ 충족",
                     ifelse(p_v > 0.05, "⚠ 경계", "❌ 위반"))
      
      cat(sprintf("  %-16s %8.3f %8.4f  %s\n",
                  label, safe_num(t_v), safe_num(p_v), judg))
      
      pt_all[[paste(seg_name, n_label, label)]] <-
        data.frame(부문=seg_name, 투자횟수=n_label, n_funded=n,
                   변수=label,
                   t값=round(t_v, 4), p값=round(p_v, 4),
                   판정=judg, stringsAsFactors=FALSE)
    }
  }
}

pt_result <- bind_rows(pt_all) %>% as.data.frame()

cat("\n=== 평행추세 결과 요약 ===\n")
print(pt_result)

# ==============================================================================
# 7. 시각화
# ==============================================================================

if (Sys.info()["sysname"] == "Darwin") {
  kfont <- "AppleGothic"
} else if (Sys.info()["sysname"] == "Windows") {
  kfont <- "Malgun Gothic"
} else {
  kfont <- "NanumGothic"
}

# 색상 팔레트
col_ctrl    <- "#4A90D9"   # 통제
col_1time   <- "#91C46C"   # 1회
col_2times  <- "#F5A623"   # 2회
col_3times  <- "#E05555"   # 3회

nf_colors <- c("통제" = col_ctrl,
               "1회"  = col_1time,
               "2회"  = col_2times,
               "3회"  = col_3times)

# ── (A) DID 계수 dot-plot: 부문 × 투자횟수 × 변수 ──────────────────────────
plot_did <- did_result %>%
  mutate(
    변수     = factor(변수, levels = rev(ana_vars$label)),
    투자횟수 = factor(투자횟수, levels = c("1회", "2회", "3회")),
    부문     = factor(부문, levels = c("소재", "부품", "장비")),
    유의     = ifelse(p_value < 0.05, "유의(p<0.05)", "비유의")
  )

p_dot <- ggplot(plot_did,
                aes(x = DID, y = 변수, color = 투자횟수, shape = 유의)) +
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray40") +
  geom_point(size = 3.2, position = position_dodge(width = 0.7)) +
  scale_color_manual(values = c("1회" = col_1time,
                                "2회" = col_2times,
                                "3회" = col_3times)) +
  scale_shape_manual(values = c("유의(p<0.05)" = 16, "비유의" = 1)) +
  facet_wrap(~ 부문, ncol = 3, scales = "free_x") +
  labs(
    title    = "투자 횟수별 Simple DID 추정량",
    subtitle = "Pre=2019, Post=2024 | 부문 × 투자횟수 × 변수",
    x = "DID 계수", y = NULL, color = "투자횟수", shape = "유의성"
  ) +
  theme_minimal(base_family = kfont) +
  theme(
    plot.title      = element_text(face = "bold", size = 14),
    strip.text      = element_text(face = "bold", size = 12),
    legend.position = "top",
    axis.text.y     = element_text(size = 9)
  )

ggsave("DID_Nfunded_dotplot.png", p_dot, width = 15, height = 9, dpi = 300)
cat("\n✓ DID_Nfunded_dotplot.png 저장\n")

# ── (B) 카테고리별 Facet bar: 부문 × 투자횟수 ───────────────────────────────
p_bar <- ggplot(plot_did,
                aes(x = 변수, y = DID, fill = 투자횟수, alpha = 유의)) +
  geom_col(position = position_dodge(width = 0.8), width = 0.7) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray30") +
  scale_fill_manual(values = c("1회" = col_1time,
                               "2회" = col_2times,
                               "3회" = col_3times)) +
  scale_alpha_manual(values = c("유의(p<0.05)" = 1.0, "비유의" = 0.45),
                     guide  = "none") +
  facet_grid(부문 ~ 카테고리, scales = "free_x") +
  labs(
    title    = "투자 횟수별 DID 추정량 — 카테고리 × 부문",
    subtitle = "진한 색 = 유의(p<0.05), 연한 색 = 비유의",
    x = NULL, y = "DID 계수", fill = "투자횟수"
  ) +
  theme_minimal(base_family = kfont) +
  theme(
    plot.title      = element_text(face = "bold", size = 13),
    strip.text      = element_text(face = "bold", size = 10),
    axis.text.x     = element_text(angle = 40, hjust = 1, size = 8),
    legend.position = "top"
  )

ggsave("DID_Nfunded_barplot.png", p_bar, width = 16, height = 12, dpi = 300)
cat("✓ DID_Nfunded_barplot.png 저장\n")

# ── (C) 평행추세 추이 — 부문별, 대표 변수(자산) × 투자횟수 ──────────────────
# 연도별 평균 (n_funded 그룹별)
trend_nf <- panel %>%
  filter(year %in% 2019:2024) %>%
  mutate(
    그룹 = case_when(
      n_funded == 0 ~ "통제",
      n_funded == 1 ~ "1회",
      n_funded == 2 ~ "2회",
      n_funded == 3 ~ "3회"
    ),
    그룹 = factor(그룹, levels = c("통제", "1회", "2회", "3회"))
  )

# 대표 변수 4개 × 부문 3개 그리드
key_vars   <- c("log_asset", "log_sales", "log_patent", "log_rdcost")
key_labels <- c("ln자산", "ln매출", "ln(특허+1)", "ln(개발비+1)")

trend_plots <- list()
idx <- 1

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub_nf   <- trend_nf %>% filter(seg == s)
  
  for (vi in seq_along(key_vars)) {
    varname <- key_vars[vi]
    label   <- key_labels[vi]
    
    se_df <- sub_nf %>%
      filter(!is.na(.data[[varname]])) %>%
      group_by(year, 그룹) %>%
      summarise(
        m  = mean(.data[[varname]], na.rm = TRUE),
        se = sd(.data[[varname]],   na.rm = TRUE) / sqrt(n()),
        .groups = "drop"
      )
    
    trend_plots[[idx]] <- ggplot(se_df, aes(x = year, y = m,
                                            color = 그룹, fill = 그룹)) +
      geom_ribbon(aes(ymin = m - 1.96*se, ymax = m + 1.96*se),
                  alpha = 0.10, color = NA) +
      geom_line(linewidth = 1.0) +
      geom_point(size = 1.8) +
      annotate("rect", xmin = 2019.5, xmax = 2022.5,
               ymin = -Inf, ymax = Inf, alpha = 0.05, fill = "gold") +
      geom_vline(xintercept = 2019.5, linetype = "dashed",
                 color = "gray50", linewidth = 0.5) +
      scale_color_manual(values = nf_colors) +
      scale_fill_manual( values = nf_colors) +
      scale_x_continuous(breaks = 2019:2024) +
      labs(title = paste0(seg_name, " | ", label),
           x = NULL, y = NULL, color = NULL, fill = NULL) +
      theme_minimal(base_family = kfont) +
      theme(
        plot.title      = element_text(face = "bold", size = 9),
        axis.text.x     = element_text(size = 7, angle = 30, hjust = 1),
        axis.text.y     = element_text(size = 7),
        legend.position = if (idx == 1) "top" else "none",
        legend.text     = element_text(size = 8),
        panel.grid.minor= element_blank()
      )
    
    idx <- idx + 1
  }
}

g_trend <- gridExtra::grid.arrange(
  grobs = trend_plots,
  ncol  = 4,
  top   = grid::textGrob(
    "평행추세: 투자 횟수별 추이 (통제·1회·2회·3회)",
    gp = grid::gpar(fontsize = 13, fontface = "bold", fontfamily = kfont)
  )
)

ggsave("ParallelTrends_Nfunded.png", g_trend,
       width = 18, height = 12, dpi = 300)
cat("✓ ParallelTrends_Nfunded.png 저장\n")

# ── (D) 평행추세 히트맵 — 부문 × 투자횟수 × 변수 ────────────────────────────
pt_heat <- pt_result %>%
  filter(!is.na(p값)) %>%
  mutate(
    판정색   = case_when(
      p값 > 0.1  ~ "충족",
      p값 > 0.05 ~ "경계",
      TRUE       ~ "위반"
    ),
    부문     = factor(부문,     levels = c("소재", "부품", "장비")),
    투자횟수 = factor(투자횟수, levels = c("1회", "2회", "3회")),
    변수     = factor(변수,     levels = rev(ana_vars$label)),
    x_label  = paste0(부문, "\n", 투자횟수)
  )

p_heat <- ggplot(pt_heat,
                 aes(x = interaction(투자횟수, 부문, sep = "\n"),
                     y = 변수, fill = 판정색)) +
  geom_tile(color = "white", linewidth = 0.7) +
  geom_text(aes(label = sprintf("%.3f", p값)),
            size = 2.8, color = "gray20") +
  scale_fill_manual(
    values = c("충족" = "#A8D5A2", "경계" = "#FFE08A", "위반" = "#F4A6A0"),
    name   = "평행추세 판정"
  ) +
  labs(
    title    = "평행추세 검정 히트맵 — 투자 횟수별 × 부문별",
    subtitle = "초록=충족(p>0.1)  노랑=경계(p<0.1)  빨강=위반(p<0.05)",
    x = "부문 × 투자횟수", y = NULL
  ) +
  theme_minimal(base_family = kfont) +
  theme(
    plot.title    = element_text(face = "bold", size = 13),
    plot.subtitle = element_text(size = 9, color = "gray40"),
    axis.text.x   = element_text(size = 9),
    axis.text.y   = element_text(size = 9),
    legend.position = "top"
  )

ggsave("ParallelTrends_Nfunded_Heatmap.png", p_heat,
       width = 12, height = 8, dpi = 300)
cat("✓ ParallelTrends_Nfunded_Heatmap.png 저장\n")

# ==============================================================================
# 8. Excel 저장
# ==============================================================================

# Wide 시트 3개 (부문별)
excel_sheets <- list(
  DID_결과_전체    = did_result,
  평행추세_결과    = pt_result,
  Wide_소재        = wide_list[["소재"]],
  Wide_부품        = wide_list[["부품"]],
  Wide_장비        = wide_list[["장비"]]
)

write_xlsx(excel_sheets, "DID_Nfunded_Seg.xlsx")

cat("\n✓ 저장 완료: DID_Nfunded_Seg.xlsx\n")
cat("  시트: DID_결과_전체 / 평행추세_결과 / Wide_소재 / Wide_부품 / Wide_장비\n")
cat("\n=== 분석 완료 ===\n")