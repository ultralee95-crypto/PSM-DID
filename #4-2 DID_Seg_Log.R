# ==============================================================================
# 소재(seg=1) / 부품(seg=2) / 장비(seg=3)별 DID 분석
# 5개 카테고리 × 11개 종속변수 체계 (전체 log 변환 버전)
# ※ 모든 측정 변수 log 변환 적용
#    - 수준 변수   : log1p(pmax(x, 0))
#    - 음수 허용   : sign(x) * log1p(|x|)   [op_profit, equity]
#    - 성장률      : log 차분 (log_t - log_{t-1})
#    - 비율 변수   : log(분자) - log(분모)
#    - innov_patent: Poisson → FE (연속형으로 전환)
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
    fund_count  = to_num(row$fund2020) + to_num(row$fund2021) + to_num(row$fund2022),
    total_gfund = to_num(row$gfundvol2020) + to_num(row$gfundvol2021) + to_num(row$gfundvol2022),
    year        = years,
    emp       = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S05000.종업원수")]])),
    asset     = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S15000.자산총계")]])),
    debt      = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18000.부채총계")]])),
    equity    = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18900.자본총계")]])),
    capital   = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S18100.자본금")]])),
    sales     = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S21100.총매출액")]])),
    op_profit = sapply(years, function(y) to_num(row[[paste0(y, "/Annual S25000.영업이익(손실)")]])),
    patent    = sapply(years, function(y) to_num(row[[paste0("p", y)]])),
    stringsAsFactors = FALSE
  )
})

panel <- bind_rows(panel_list) %>% arrange(firm_id, year)

cat("\n패널:", nrow(panel), "obs,", length(unique(panel$firm_id)), "firms\n")

# ==============================================================================
# 2. 원시 수준 변수 → Log 변환
# ==============================================================================
# ● 항상 양수(≥0) 변수: log1p(pmax(x, 0))
# ● 음수 가능 변수   : log_signed(x)

panel <- panel %>%
  mutate(
    log_patent   = log1p(pmax(patent,   0, na.rm = TRUE)),   # ln(특허+1)
    log_emp_raw   = ifelse(!is.na(emp)   & emp   > 0, log(emp),   NA),   # ln(종업원+1)
    log_asset_raw = ifelse(!is.na(asset) & asset > 0, log(asset), NA),   # ln(자산총계+1)
    log_debt     = log1p(pmax(debt,     0, na.rm = TRUE)),   # ln(부채총계+1)
    log_equity   = log_signed(equity),                        # ±ln(자본총계+1)
    log_capital  = log1p(pmax(capital,  0, na.rm = TRUE)),   # ln(자본금+1)
    log_sales     = ifelse(!is.na(sales) & sales > 0, log(sales), NA), # ln(매출액+1)
    log_opprofit = log_signed(op_profit)                      # ±ln(영업이익+1)
  )

# ==============================================================================
# 3. 종속변수 생성 (모두 log 변환 기반)
# ==============================================================================

panel <- panel %>%
  mutate(
    post = ifelse(year >= 2020, 1, 0),
    did  = treat * post
  )

# ── [혁신성] ln(특허출원수 + 1) ──
panel$innov_patent <- panel$log_patent

# ── [성장성] log 차분 (= 로그 증가율 ≈ 연속 복리 성장률) ──
panel <- panel %>%
  group_by(firm_id) %>%
  mutate(
    grow_asset     = log_asset_raw   - lag(log_asset_raw),    # Δln(자산)
    grow_sales     = log_sales       - lag(log_sales),         # Δln(매출)
    grow_netincome = log_opprofit    - lag(log_opprofit)       # Δln(영업이익)
  ) %>%
  ungroup()

# ── [수익성] log(비율) = log(분자) - log(분모) ──
panel$profit_opm <- ifelse(
  panel$sales > 0,
  log_signed(panel$op_profit / panel$sales),    # ln(영업이익률)
  NA
)
panel$profit_roa <- ifelse(
  panel$asset > 0,
  log_signed(panel$op_profit / panel$asset),    # ln(ROA)
  NA
)
panel$profit_roe <- ifelse(
  panel$equity != 0,
  log_signed(panel$op_profit / panel$equity),   # ln(ROE)
  NA
)

# ── [안정성] log(비율) ──
panel$stab_debtratio <- ifelse(
  panel$equity != 0,
  panel$log_debt - log_signed(panel$equity),    # ln(부채/자본) = ln부채비율
  NA
)
panel$stab_equitydep <- ifelse(
  panel$asset > 0,
  log_signed(panel$equity) - panel$log_asset_raw,  # ln(자본/자산) = ln자기자본의존도
  NA
)

# ── [활동성] log(비율) ──
panel$act_assetturn <- ifelse(
  panel$asset > 0,
  panel$log_sales - panel$log_asset_raw,        # ln(매출/자산) = ln총자산회전율
  NA
)
panel$act_equityturn <- ifelse(
  panel$equity != 0,
  panel$log_sales - log_signed(panel$equity),   # ln(매출/자본) = ln자기자본회전율
  NA
)

# ── Winsorize (log 변환 후 이상치 처리) ──
winsorize <- function(x, probs = c(0.01, 0.99)) {
  q <- quantile(x, probs, na.rm = TRUE)
  x[x < q[1] & !is.na(x)] <- q[1]
  x[x > q[2] & !is.na(x)] <- q[2]
  return(x)
}

ratio_vars <- c("grow_asset", "grow_sales", "grow_netincome",
                "profit_opm", "profit_roa", "profit_roe",
                "stab_debtratio", "stab_equitydep",
                "act_assetturn", "act_equityturn")

for (v in ratio_vars) {
  panel[[v]] <- winsorize(panel[[v]])
}

# ── 통제변수 (log 변환) ──
panel$log_emp   <- panel$log_emp_raw    # ln(종업원+1)
panel$log_asset <- panel$log_asset_raw  # ln(자산총계+1)
panel$leverage  <- panel$debt / ifelse(panel$asset == 0, NA, panel$asset)  # 부채비율 (통제)

# ── Dosage ──
panel$dose1 <- as.integer(panel$treat == 1 & panel$fund_count == 1)
panel$dose2 <- as.integer(panel$treat == 1 & panel$fund_count == 2)
panel$dose3 <- as.integer(panel$treat == 1 & panel$fund_count == 3)
panel$dose1_post <- panel$dose1 * panel$post
panel$dose2_post <- panel$dose2 * panel$post
panel$dose3_post <- panel$dose3 * panel$post

# ── 투자금액 log 변환 ──
panel$log_gfund      <- log1p(panel$total_gfund)    # ln(투자금액+1)
panel$log_gfund_post <- panel$log_gfund * panel$post

# ── 특허 시차 (log 변환, 매개분석용) ──
panel <- panel %>%
  group_by(firm_id) %>%
  mutate(
    patent_lag1     = lag(patent),
    log_patent_lag1 = lag(log_patent)   # ln(특허+1)의 시차
  ) %>%
  ungroup()

cat("\n=== 변수 생성 완료 (전체 log 변환) ===\n")

# ==============================================================================
# 4. DV 정의
#    ※ innov_patent: Poisson 제거 → FE (log 변환으로 연속형 변수)
# ==============================================================================

dv_info <- data.frame(
  dv     = c("innov_patent", "grow_asset", "grow_sales", "grow_netincome",
             "profit_opm", "profit_roa", "profit_roe",
             "stab_debtratio", "stab_equitydep",
             "act_assetturn", "act_equityturn"),
  label  = c("ln(특허+1)", "Δln자산", "Δln매출", "Δln영업이익",
             "ln영업이익률", "ln(ROA)", "ln(ROE)",
             "ln부채비율", "ln자기자본의존도",
             "ln총자산회전율", "ln자기자본회전율"),
  cat    = c("혁신성", "성장성", "성장성", "성장성",
             "수익성", "수익성", "수익성",
             "안정성", "안정성",
             "활동성", "활동성"),
  method = c("fe",  "fe", "fe", "fe",   # ← innov_patent: poisson → fe
             "fe",  "fe",  "fe",
             "fe",  "fe",
             "fe",  "fe"),
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
# 5. 분석 함수 (RQ1 ~ RQ5)
# ==============================================================================

# ── RQ1: 평균 처치효과 ──
run_rq1 <- function(sub, dv, method) {
  df <- sub %>% filter(!is.na(.data[[dv]]), !is.na(log_emp), !is.na(log_asset), !is.na(leverage))
  tryCatch({
    if (method == "ols") {
      fml <- as.formula(paste0(dv, " ~ did + treat + log_emp + log_asset + leverage + factor(year)"))
      m   <- lm(fml, data = df)
      rob <- coeftest(m, vcov = vcovCL(m, cluster = df$firm_id))
      data.frame(coef = rob["did",1], se = rob["did",2], stat = rob["did",3], pval = rob["did",4], irr = NA)
    } else {
      pdf <- pdata.frame(df, index = c("firm_id", "year"), drop.index = FALSE)
      fml <- as.formula(paste0(dv, " ~ did + log_emp + log_asset + leverage"))
      m   <- plm(fml, data = pdf, model = "within", effect = "twoways")
      rob <- coeftest(m, vcov = vcovHC(m, type = "HC1", cluster = "group"))
      data.frame(coef = rob["did",1], se = rob["did",2], stat = rob["did",3], pval = rob["did",4], irr = NA)
    }
  }, error = function(e) data.frame(coef=NA, se=NA, stat=NA, pval=NA, irr=NA))
}

# ── RQ2: 투자 횟수별 효과 ──
run_rq2 <- function(sub, dv, method) {
  df <- sub %>% filter(!is.na(.data[[dv]]), !is.na(log_emp), !is.na(log_asset), !is.na(leverage))
  tryCatch({
    if (method == "ols") {
      fml <- as.formula(paste0(dv, " ~ dose1_post + dose2_post + dose3_post + dose1 + dose2 + dose3 + log_emp + log_asset + leverage + factor(year)"))
      m   <- lm(fml, data = df)
      rob <- coeftest(m, vcov = vcovCL(m, cluster = df$firm_id))
    } else {
      pdf <- pdata.frame(df, index = c("firm_id", "year"), drop.index = FALSE)
      fml <- as.formula(paste0(dv, " ~ dose1_post + dose2_post + dose3_post + log_emp + log_asset + leverage"))
      m   <- plm(fml, data = pdf, model = "within", effect = "twoways")
      rob <- coeftest(m, vcov = vcovHC(m, type = "HC1", cluster = "group"))
    }
    data.frame(
      d1_coef = rob["dose1_post",1], d1_p = rob["dose1_post",4],
      d2_coef = rob["dose2_post",1], d2_p = rob["dose2_post",4],
      d3_coef = rob["dose3_post",1], d3_p = rob["dose3_post",4]
    )
  }, error = function(e) data.frame(d1_coef=NA,d1_p=NA,d2_coef=NA,d2_p=NA,d3_coef=NA,d3_p=NA))
}

# ── RQ3: 투자 금액 효과 ──
run_rq3 <- function(sub, dv, method) {
  df <- sub %>% filter(!is.na(.data[[dv]]), !is.na(log_emp), !is.na(log_asset), !is.na(leverage))
  tryCatch({
    if (method == "ols") {
      fml <- as.formula(paste0(dv, " ~ log_gfund_post + log_gfund + log_emp + log_asset + leverage + factor(year)"))
      m   <- lm(fml, data = df)
      rob <- coeftest(m, vcov = vcovCL(m, cluster = df$firm_id))
    } else {
      pdf <- pdata.frame(df, index = c("firm_id", "year"), drop.index = FALSE)
      fml <- as.formula(paste0(dv, " ~ log_gfund_post + log_emp + log_asset + leverage"))
      m   <- plm(fml, data = pdf, model = "within", effect = "twoways")
      rob <- coeftest(m, vcov = vcovHC(m, type = "HC1", cluster = "group"))
    }
    data.frame(coef = rob["log_gfund_post",1], se = rob["log_gfund_post",2], pval = rob["log_gfund_post",4])
  }, error = function(e) data.frame(coef=NA, se=NA, pval=NA))
}

# ── RQ5: 매개효과 (특허 → 재무) ──
run_rq5 <- function(sub, dv, method) {
  df <- sub %>% filter(!is.na(.data[[dv]]), !is.na(log_emp), !is.na(log_asset),
                       !is.na(leverage), !is.na(log_patent_lag1))
  tryCatch({
    # a-path: did → ln(특허+1) [FE, 연속형으로 변경]
    pdf_a <- pdata.frame(df, index = c("firm_id", "year"), drop.index = FALSE)
    m_a   <- plm(innov_patent ~ did + log_emp + log_asset + leverage,
                 data = pdf_a, model = "within", effect = "twoways")
    rob_a  <- coeftest(m_a, vcov = vcovHC(m_a, type = "HC1", cluster = "group"))
    a_coef <- rob_a["did",1]
    a_se   <- rob_a["did",2]
    a_pval <- rob_a["did",4]
    
    # b-path: ln(특허_lag) → Y
    if (method == "ols") {
      fml <- as.formula(paste0(dv, " ~ did + treat + log_patent_lag1 + log_emp + log_asset + leverage + factor(year)"))
      m_cb  <- lm(fml, data = df)
      rob_cb <- coeftest(m_cb, vcov = vcovCL(m_cb, cluster = df$firm_id))
    } else {
      pdf_b  <- pdata.frame(df, index = c("firm_id", "year"), drop.index = FALSE)
      fml    <- as.formula(paste0(dv, " ~ did + log_patent_lag1 + log_emp + log_asset + leverage"))
      m_cb   <- plm(fml, data = pdf_b, model = "within", effect = "twoways")
      rob_cb <- coeftest(m_cb, vcov = vcovHC(m_cb, type = "HC1", cluster = "group"))
    }
    b_coef <- rob_cb["log_patent_lag1",1]
    b_se   <- rob_cb["log_patent_lag1",2]
    b_pval <- rob_cb["log_patent_lag1",4]
    
    # Sobel test
    indirect <- a_coef * b_coef
    sobel_se <- sqrt(a_coef^2 * b_se^2 + b_coef^2 * a_se^2)
    sobel_z  <- indirect / sobel_se
    sobel_p  <- 2 * pnorm(-abs(sobel_z))
    
    # ※ log 변환이므로 IRR 대신 beta 계수 보고
    data.frame(a_coef=a_coef, a_pval=a_pval,
               b_coef=b_coef, b_pval=b_pval,
               indirect=indirect, sobel_z=sobel_z, sobel_p=sobel_p)
  }, error = function(e) data.frame(a_coef=NA,a_pval=NA,b_coef=NA,b_pval=NA,
                                    indirect=NA,sobel_z=NA,sobel_p=NA))
}

# ==============================================================================
# 6. 부문별 분석 실행
# ==============================================================================

seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")

all_rq1 <- list(); all_rq2 <- list(); all_rq3 <- list(); all_rq5 <- list()

for (s in c(1, 2, 3)) {
  seg_name <- seg_labels[as.character(s)]
  sub <- panel %>% filter(seg == s)
  
  n_t <- length(unique(sub$firm_id[sub$treat == 1]))
  n_c <- length(unique(sub$firm_id[sub$treat == 0]))
  
  cat("\n", paste(rep("=", 80), collapse = ""), "\n")
  cat("  ", seg_name, " 부문 분석 (처치:", n_t, " / 통제:", n_c, ")\n")
  cat(paste(rep("=", 80), collapse = ""), "\n")
  
  # ── RQ1 ──
  cat("\n─── RQ1: 평균 처치효과 (ATT) ───\n")
  cat(sprintf("  %-18s %8s %10s %10s %8s %5s\n",
              "카테고리/지표", "ATT", "SE", "stat", "p", "Sig"))
  cat("  ", paste(rep("-", 65), collapse = ""), "\n")
  
  rq1_rows <- list()
  for (i in 1:nrow(dv_info)) {
    r <- run_rq1(sub, dv_info$dv[i], dv_info$method[i])
    cat(sprintf("  [%s] %-12s %8.4f %10.4f %10.3f %8.4f %5s\n",
                dv_info$cat[i], dv_info$label[i],
                r$coef, r$se, r$stat, r$pval, sig_mark(r$pval)))
    rq1_rows[[i]] <- data.frame(부문=seg_name, 카테고리=dv_info$cat[i], 지표=dv_info$label[i],
                                ATT=r$coef, SE=r$se, t_z=r$stat, p_value=r$pval,
                                유의성=sig_mark(r$pval), stringsAsFactors=FALSE)
  }
  all_rq1[[seg_name]] <- bind_rows(rq1_rows)
  
  # ── RQ2 ──
  cat("\n─── RQ2: 투자 횟수별 효과 ───\n")
  cat(sprintf("  %-18s %10s %6s %10s %6s %10s %6s\n",
              "지표", "1회(β)", "p", "2회(β)", "p", "3회(β)", "p"))
  cat("  ", paste(rep("-", 65), collapse = ""), "\n")
  
  rq2_rows <- list()
  for (i in 1:nrow(dv_info)) {
    r <- run_rq2(sub, dv_info$dv[i], dv_info$method[i])
    cat(sprintf("  %-18s %10.4f %5.3f%s %10.4f %5.3f%s %10.4f %5.3f%s\n",
                dv_info$label[i],
                r$d1_coef, r$d1_p, sig_mark(r$d1_p),
                r$d2_coef, r$d2_p, sig_mark(r$d2_p),
                r$d3_coef, r$d3_p, sig_mark(r$d3_p)))
    rq2_rows[[i]] <- data.frame(부문=seg_name, 지표=dv_info$label[i],
                                `1회`=r$d1_coef, `1회_p`=r$d1_p,
                                `2회`=r$d2_coef, `2회_p`=r$d2_p,
                                `3회`=r$d3_coef, `3회_p`=r$d3_p,
                                stringsAsFactors=FALSE, check.names=FALSE)
  }
  all_rq2[[seg_name]] <- bind_rows(rq2_rows)
  
  # ── RQ3 ──
  cat("\n─── RQ3: 투자 금액 효과 ───\n")
  cat(sprintf("  %-18s %10s %8s %5s\n", "지표", "β(lnFund)", "p", "Sig"))
  cat("  ", paste(rep("-", 45), collapse = ""), "\n")
  
  rq3_rows <- list()
  for (i in 1:nrow(dv_info)) {
    r <- run_rq3(sub, dv_info$dv[i], dv_info$method[i])
    cat(sprintf("  %-18s %10.4f %8.4f %5s\n",
                dv_info$label[i], r$coef, r$pval, sig_mark(r$pval)))
    rq3_rows[[i]] <- data.frame(부문=seg_name, 지표=dv_info$label[i],
                                β=r$coef, SE=r$se, p_value=r$pval,
                                유의성=sig_mark(r$pval), stringsAsFactors=FALSE)
  }
  all_rq3[[seg_name]] <- bind_rows(rq3_rows)
  
  # ── RQ5 ──
  cat("\n─── RQ5: 매개효과 (ln특허 → 재무) ───\n")
  cat(sprintf("  %-18s %8s %6s %8s %6s %9s %8s %6s\n",
              "지표(Y)", "a(β)", "p(a)", "b(β)", "p(b)", "a×b", "Sobel z", "p"))
  cat("  ", paste(rep("-", 75), collapse = ""), "\n")
  
  rq5_dvs  <- dv_info %>% filter(dv != "innov_patent")
  rq5_rows <- list()
  for (i in 1:nrow(rq5_dvs)) {
    r <- run_rq5(sub, rq5_dvs$dv[i], rq5_dvs$method[i])
    cat(sprintf("  %-18s %8.4f %5.3f%s %8.4f %5.3f%s %9.5f %8.3f %5.3f%s\n",
                rq5_dvs$label[i],
                r$a_coef, r$a_pval, sig_mark(r$a_pval),
                r$b_coef, r$b_pval, sig_mark(r$b_pval),
                r$indirect, r$sobel_z, r$sobel_p, sig_mark(r$sobel_p)))
    rq5_rows[[i]] <- data.frame(부문=seg_name, 지표=rq5_dvs$label[i],
                                a_β=r$a_coef, a_p=r$a_pval,
                                b_β=r$b_coef, b_p=r$b_pval,
                                `a×b`=r$indirect, Sobel_z=r$sobel_z, Sobel_p=r$sobel_p,
                                유의성=sig_mark(r$sobel_p),
                                stringsAsFactors=FALSE, check.names=FALSE)
  }
  all_rq5[[seg_name]] <- bind_rows(rq5_rows)
}

# ==============================================================================
# 7. 결과 통합
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  결과 통합 및 부문별 비교\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

rq1_all <- bind_rows(all_rq1)
rq2_all <- bind_rows(all_rq2)
rq3_all <- bind_rows(all_rq3)
rq5_all <- bind_rows(all_rq5)

# RQ1 Wide: 부문별 비교 테이블
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
# 8. Excel 저장
# ==============================================================================

write_xlsx(list(
  RQ1_ATT       = rq1_all,
  RQ1_Wide      = rq1_wide,
  RQ2_Dosage    = rq2_all,
  RQ3_Amount    = rq3_all,
  RQ5_Mediation = rq5_all
), "DID_segmented_RQ1_RQ5_log.xlsx")

cat("\n✓ 저장: DID_segmented_RQ1_RQ5_log.xlsx\n")

# ==============================================================================
# 9. 시각화
# ==============================================================================

cat("\n=== 시각화 생성 ===\n")

if (Sys.info()["sysname"] == "Darwin") {
  kfont <- "AppleGothic"
} else if (Sys.info()["sysname"] == "Windows") {
  kfont <- "Malgun Gothic"
} else {
  kfont <- "NanumGothic"
}

# ── (a) RQ1: 부문별 ATT 비교 ──
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
  labs(title = "RQ1: 부문별 평균 처치효과 (ATT) — Log 변환",
       subtitle = "PSM-DID | 소재·부품·장비 × 11개 종속변수 (log 변환)",
       x = "ATT (DID Coefficient)", y = "", color = "부문", shape = "유의성") +
  theme_minimal(base_family = kfont) +
  theme(plot.title = element_text(face = "bold", size = 14), legend.position = "top")

ggsave("RQ1_ATT_by_segment_log.png", p1, width = 12, height = 8, dpi = 300)
cat("✓ RQ1_ATT_by_segment_log.png\n")

# ── (b) RQ1: 카테고리별 Facet ──
p1f <- ggplot(plot_rq1, aes(x = 지표, y = ATT, fill = 부문)) +
  geom_bar(stat = "identity", position = position_dodge(width = 0.8), width = 0.7) +
  geom_errorbar(aes(ymin = ATT - 1.96 * SE, ymax = ATT + 1.96 * SE),
                position = position_dodge(width = 0.8), width = 0.2) +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red") +
  facet_wrap(~ 카테고리, scales = "free", ncol = 2) +
  scale_fill_manual(values = c("소재" = "#8E44AD", "부품" = "#2980B9", "장비" = "#27AE60")) +
  labs(title = "부문별 DID 추정량 — Log 변환 (카테고리별)",
       x = "", y = "ATT", fill = "부문") +
  theme_minimal(base_family = kfont) +
  theme(plot.title = element_text(face = "bold", size = 14),
        axis.text.x = element_text(angle = 45, hjust = 1, size = 9),
        strip.text  = element_text(face = "bold", size = 12),
        legend.position = "top")

ggsave("RQ1_ATT_by_category_facet_log.png", p1f, width = 14, height = 10, dpi = 300)
cat("✓ RQ1_ATT_by_category_facet_log.png\n")

# ── (c) 평행추세: 부문별 주요 지표 (log 변환 변수) ──
key_dvs    <- c("innov_patent", "grow_asset", "grow_sales", "profit_roa",
                "stab_debtratio", "act_assetturn")
key_labels <- c("ln(특허+1)", "Δln자산", "Δln매출", "ln(ROA)",
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
      top = grid::textGrob(paste0(seg_name, " 부문 평행추세 (log 변환)"),
                           gp = grid::gpar(fontsize = 14, fontface = "bold",
                                           fontfamily = kfont))
    )
    ggsave(paste0("parallel_trends_log_", seg_name, ".png"),
           g, width = 14, height = 8, dpi = 300)
  } else {
    for (j in seq_along(plots)) {
      ggsave(paste0("trend_log_", seg_name, "_", key_dvs[j], ".png"),
             plots[[j]], width = 6, height = 4, dpi = 300)
    }
  }
  cat("✓ 평행추세 (log):", seg_name, "\n")
}

# ==============================================================================
# 10. Simple DID (log 변환 변수 기준, 2019 vs 2024)
# ==============================================================================

cat("\n", paste(rep("=", 80), collapse = ""), "\n")
cat("  Simple DID (2019 vs 2024 비교) — log 변환 변수\n")
cat(paste(rep("=", 80), collapse = ""), "\n")

# Simple DID용 변수: log 변환된 수준 변수 사용
simple_vars <- data.frame(
  label  = c("ln자산", "ln매출", "ln부채", "ln자본금", "ln종업원수", "ln영업이익", "ln특허"),
  var    = c("log_asset_raw", "log_sales", "log_debt", "log_capital",
             "log_emp_raw", "log_opprofit", "log_patent"),
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
    
    t_mean_pre  <- mean(sub_wide[[col_pre]] [sub_wide$treat==1], na.rm=TRUE)
    t_mean_post <- mean(sub_wide[[col_post]][sub_wide$treat==1], na.rm=TRUE)
    c_mean_pre  <- mean(sub_wide[[col_pre]] [sub_wide$treat==0], na.rm=TRUE)
    c_mean_post <- mean(sub_wide[[col_post]][sub_wide$treat==0], na.rm=TRUE)
    
    reg     <- lm(diff ~ treat, data = sub_wide)
    did_est <- coef(reg)["treat"]
    t_val   <- summary(reg)$coefficients["treat", "t value"]
    p_val   <- summary(reg)$coefficients["treat", "Pr(>|t|)"]
    
    cat(sprintf("  %-12s %12.4f %12.4f %12.4f %12.4f %12.4f %8.3f %5s\n",
                simple_vars$label[v],
                t_mean_pre, t_mean_post, c_mean_pre, c_mean_post,
                did_est, t_val, sig_mark(p_val)))
    
    seg_rows[[v]] <- data.frame(부문=seg_name, 변수=simple_vars$label[v],
                                처치_pre=t_mean_pre, 처치_post=t_mean_post,
                                통제_pre=c_mean_pre, 통제_post=c_mean_post,
                                DID=did_est, t_value=t_val, p_value=p_val,
                                유의성=sig_mark(p_val), stringsAsFactors=FALSE)
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

cat("\n=== Simple DID 부문별 비교 (Wide, log 변환) ===\n")
print(simple_wide)

# 최종 통합 저장
write_xlsx(list(
  Simple_DID       = simple_result,
  Simple_DID_Wide  = simple_wide,
  RQ1_ATT          = rq1_all,
  RQ1_Wide         = rq1_wide,
  RQ2_Dosage       = rq2_all,
  RQ3_Amount       = rq3_all,
  RQ5_Mediation    = rq5_all
), "DID_segmented_all_results_log.xlsx")

cat("\n✓ 최종 저장: DID_segmented_all_results_log.xlsx\n")
cat("\n=== 전체 분석 완료 (log 변환 버전) ===\n")