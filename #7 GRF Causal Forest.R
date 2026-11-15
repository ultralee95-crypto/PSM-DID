# ==============================================================================
# #7 GRF Causal Forest — PSM-DID 강건성 검증
#
# 주 분석(DR-DID)과 동일한 3가지 비교를 Causal Forest로 재추정하여 비교
#   (1) DR-DID  : 부문별 처치군 전체 vs 통제군
#   (2) 횟수별  : 1회 · 2회 · 3회 수혜 처치군 vs 동일 통제군   (#4-2 구조)
#   (3) 금액별  : 금액 하위 · 중위 · 상위 처치군 vs 동일 통제군 (#4-3 구조)
#
# 설계 (강건성 검증용)
#   · 표본      : matched_dataset_segPSM.xlsx (1:1 부문별 PSM)
#   · 결과변수  : ΔY = Y_post − Y_2019 (#4-1 정의), post = 2024 / 2025
#   · 공변량 X  : PSM 공변량 8개 + 산업 · 지역 + 사전추세(Y_2019 − Y_2018)
#   · 적합      : 부문 × 결과 × 연도마다 숲 1회 적합(부문 전체) → 비교별 ATT는 subset 으로 산출
#   · W.hat     : 처치 비율로 고정 (매칭 설계상 X와 독립)
#   · clusters  : 매칭 쌍(subclass) 단위 군집 표준오차
#   · seed      : 10회 반복 → ATT 중앙값 · 범위 보고
#   · 판정      : CF ATT 부호 = DR-DID 부호 & CF 95% CI가 DR-DID 추정치 포함 → "강건"
#
# 작성 : 2026-11-15
# ==============================================================================

packages <- c("readxl", "dplyr", "tidyr", "writexl", "grf")
for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}
select <- dplyr::select

# ------------------------------------------------------------------------------
# 0. 설정
# ------------------------------------------------------------------------------
NUM_TREES  <- 4000
SEEDS      <- 1:10
POST_YEARS <- c(2024, 2025)
seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")

to_num   <- function(x) as.numeric(gsub(",", "", as.character(x)))
sig_mark <- function(p) ifelse(is.na(p), "", ifelse(p < .001, "***", ifelse(p < .01, "**",
                         ifelse(p < .05, "*", ifelse(p < .1, ".", "ns")))))

# ------------------------------------------------------------------------------
# 1. 데이터
# ------------------------------------------------------------------------------
dat <- read_excel("matched_dataset_segPSM.xlsx", col_types = "text")
dat <- dat %>%
  mutate(
    treat_n   = to_num(treat),
    n_funded  = nchar(gsub("0", "", ifelse(is.na(fundedpattern), "", fundedpattern))),
    fund_tot  = rowSums(cbind(to_num(gfundvol2020), to_num(gfundvol2021), to_num(gfundvol2022)), na.rm = TRUE)
  )
cat("행수:", nrow(dat), "\n"); print(table(dat$seg, dat$treat))

# ------------------------------------------------------------------------------
# 2. 결과변수 (#4-1 정의, 부채 · 인건비 결측은 NA 유지)
# ------------------------------------------------------------------------------
l1_zero <- function(x) log1p(pmax(ifelse(is.na(x), 0, x), 0))
l1_na   <- function(x) ifelse(is.na(x), NA, log1p(pmax(x, 0)))
ycol    <- function(df, nm) if (nm %in% names(df)) to_num(df[[nm]]) else rep(NA_real_, nrow(df))

year_vars <- function(df, y) {
  asset <- ycol(df, paste0(y, "/Annual S15000.자산총계"))
  sales <- ycol(df, paste0(y, "/Annual S21100.총수익"))
  tibble(
    log_asset  = ifelse(!is.na(asset) & asset > 0, log(asset), NA),
    log_sales  = ifelse(!is.na(sales) & sales > 0, log(sales), NA),
    log_debt   = l1_na(ycol(df, paste0(y, "/Annual S18000.부채총계"))),
    op_profit  = ycol(df, paste0("opm", y)),
    log_export = l1_zero(ycol(df, paste0("exportamt", y))),
    log_rdcost = l1_zero(ycol(df, paste0("rdcost", y))),
    log_lbcost = l1_na(ycol(df, paste0("lbcost", y))),
    log_patent = l1_zero(ycol(df, paste0("p", y)))
  )
}
outcomes <- data.frame(
  var   = c("log_asset", "log_sales", "log_debt", "op_profit", "log_export", "log_rdcost", "log_lbcost"),
  label = c("ln자산", "ln매출", "ln부채", "영업이익", "ln(수출+1)", "ln(개발비+1)", "ln(인건비+1)"),
  stringsAsFactors = FALSE)
all_y  <- c(outcomes$var, "log_patent")
x_vars <- c("log_자산", "log_매출", "log_부채", "log_업력", "opm",
            "log_연구개발비", "log_수출", "log_인건비", "산업", "지역")

# ------------------------------------------------------------------------------
# 3. DR-DID (#4-1 방식) : 비교 표본(idx)에서 diff ~ treat + 나머지 7개 사전변수
# ------------------------------------------------------------------------------
dr_did <- function(idx, Y, pre19, W_all, v) {
  ok <- idx & !is.na(Y)
  df <- data.frame(diff = Y[ok], treat = W_all[ok], as.data.frame(pre19[ok, setdiff(all_y, v)]))
  f  <- summary(lm(diff ~ ., data = df))$coefficients
  c(b = f["treat", 1], se = f["treat", 2], p = f["treat", 4])
}

# ------------------------------------------------------------------------------
# 4. 부문 × 연도 × 변수마다 Causal Forest를 1회 적합 (부문 전체 표본)
#    → 비교(전체 / 횟수 / 금액)별 ATT는 average_treatment_effect(subset=)로 산출
#      (동일 통제군 공유 : #4-2 · #4-3 DID 구조와 동일)
# ------------------------------------------------------------------------------
results <- list()
for (s in names(seg_labels)) {
  seg_name <- seg_labels[s]
  sub <- dat %>% filter(seg == s)

  # 금액 3분위 (#4-3과 동일 : 처치군 중 금액 > 0 기준 33.3% · 66.7%)
  pos <- sub$fund_tot[sub$treat_n == 1 & sub$fund_tot > 0]
  q33 <- quantile(pos, 0.333, na.rm = TRUE); q67 <- quantile(pos, 0.667, na.rm = TRUE)
  amt <- ifelse(sub$treat_n == 0, "ctrl", ifelse(sub$fund_tot <= q33, "하위", ifelse(sub$fund_tot <= q67, "중위", "상위")))
  cnt <- ifelse(sub$treat_n == 0, "ctrl", paste0(sub$n_funded, "회"))
  comparisons <- list(
    list(분석 = "DR-DID", 그룹 = "전체", idx = rep(TRUE, nrow(sub))),
    list(분석 = "횟수별", 그룹 = "1회", idx = cnt %in% c("ctrl", "1회")),
    list(분석 = "횟수별", 그룹 = "2회", idx = cnt %in% c("ctrl", "2회")),
    list(분석 = "횟수별", 그룹 = "3회", idx = cnt %in% c("ctrl", "3회")),
    list(분석 = "금액별", 그룹 = "하위", idx = amt %in% c("ctrl", "하위")),
    list(분석 = "금액별", 그룹 = "중위", idx = amt %in% c("ctrl", "중위")),
    list(분석 = "금액별", 그룹 = "상위", idx = amt %in% c("ctrl", "상위")))

  X_all  <- sapply(x_vars, function(v) to_num(sub[[v]]))
  W_all  <- sub$treat_n
  cl_all <- as.integer(factor(sub$subclass))
  p18 <- year_vars(sub, 2018); p19 <- year_vars(sub, 2019)

  for (py in POST_YEARS) {
    post <- year_vars(sub, py)
    for (k in seq_len(nrow(outcomes))) {
      v  <- outcomes$var[k]
      Y  <- post[[v]] - p19[[v]]
      pt <- p19[[v]] - p18[[v]]
      ok <- !is.na(Y) & complete.cases(X_all)
      X_k <- cbind(X_all[ok, , drop = FALSE], pretrend = ifelse(is.na(pt[ok]), 0, pt[ok]))
      Y_k <- Y[ok]; W_k <- W_all[ok]
      w_hat <- rep(mean(W_k), length(W_k))       # 매칭 설계상 처치 비율로 고정

      # seed 반복 적합
      forests <- lapply(SEEDS, function(sd)
        causal_forest(X_k, Y_k, W_k, W.hat = w_hat, clusters = cl_all[ok],
                      num.trees = NUM_TREES, seed = sd))

      for (cmp in comparisons) {
        sub_idx <- cmp$idx[ok]                   # 해당 그룹 처치군 + 전체 통제군
        est <- sapply(forests, function(cf) {
          a <- average_treatment_effect(cf, target.sample = "treated", subset = sub_idx)
          c(a["estimate"], a["std.err"]) })
        b <- median(est[1, ]); se <- median(est[2, ])
        p <- 2 * pnorm(-abs(b / se)); lo <- b - 1.96 * se; hi <- b + 1.96 * se
        dr <- dr_did(cmp$idx, Y, p19, W_all, v)

        results[[length(results) + 1]] <- data.frame(
          분석 = cmp$분석, 부문 = seg_name, 그룹 = cmp$그룹, post_year = py, 변수 = outcomes$label[k],
          N처치 = sum(W_k[sub_idx] == 1), N통제 = sum(W_k[sub_idx] == 0),
          DR_DID = dr["b"], DR_SE = dr["se"], DR_p = dr["p"], DR_sig = sig_mark(dr["p"]),
          CF_ATT = b, CF_SE = se, CF_p = p, CF_sig = sig_mark(p), CF_CI_low = lo, CF_CI_high = hi,
          CF_seed_min = min(est[1, ]), CF_seed_max = max(est[1, ]), CF_DR_비율 = b / dr["b"],
          판정 = ifelse(sign(b) == sign(dr["b"]) & lo <= dr["b"] & dr["b"] <= hi, "강건", "불일치"),
          row.names = NULL)
        cat(sprintf("%-6s %s %-4s %d %-12s DR %+.3f%-4s | CF %+.3f%-4s [%s]\n",
                    cmp$분석, seg_name, cmp$그룹, py, outcomes$label[k],
                    dr["b"], sig_mark(dr["p"]), b, sig_mark(p),
                    ifelse(sign(b) == sign(dr["b"]) & lo <= dr["b"] & dr["b"] <= hi, "강건", "불일치")))
      }
    }
  }
}
res_df <- bind_rows(results)

# ------------------------------------------------------------------------------
# 5. 요약 · 저장
# ------------------------------------------------------------------------------
fmt <- function(b, p) sprintf("%+.3f%s", b, ifelse(sig_mark(p) == "ns", "", sig_mark(p)))
wide_tbl <- function(an) {
  res_df %>% filter(분석 == an) %>%
    mutate(DR = fmt(DR_DID, DR_p), CF = fmt(CF_ATT, CF_p)) %>%
    select(부문, 그룹, 변수, post_year, DR, CF, 판정) %>%
    pivot_wider(names_from = post_year, values_from = c(DR, CF, 판정))
}
summary_tbl <- res_df %>%
  group_by(분석) %>%
  summarise(비교수 = n(),
            강건 = sum(판정 == "강건"),
            DR유의_CF유의 = sum(DR_p < .05 & CF_p < .05),
            DR유의_CF비유의 = sum(DR_p < .05 & CF_p >= .05),
            부호일치 = sum(sign(DR_DID) == sign(CF_ATT)),
            CF_DR_비율_중앙값 = median(CF_DR_비율[DR_p < .05], na.rm = TRUE))
print(summary_tbl)

write_xlsx(list(
  `00_설정`    = data.frame(항목 = c("표본", "트리 수", "seed", "W.hat", "clusters", "공변량", "판정 기준"),
                            값 = c("1:1 부문별 PSM 매칭 표본", NUM_TREES, paste(range(SEEDS), collapse = "~"),
                                   "비교 표본의 처치 비율로 고정", "매칭 쌍(subclass)",
                                   paste(c(x_vars, "사전추세"), collapse = ", "),
                                   "CF 부호 = DR 부호 & CF 95% CI가 DR 추정치 포함 → 강건")),
  `01_요약`    = summary_tbl,
  `02_DR-DID`  = wide_tbl("DR-DID"),
  `03_횟수별`  = wide_tbl("횟수별"),
  `04_금액별`  = wide_tbl("금액별"),
  `05_전체결과` = res_df
), "GRF_CausalForest_강건성_DR_횟수_금액.xlsx")
cat("\n✓ 저장 : GRF_CausalForest_강건성_DR_횟수_금액.xlsx\n")
