# ==============================================================================
# #8 Rosenbaum 민감도 분석 — 미관측 교란(숨은 선정 편의)에 대한 강건성
#
# 질문 : "관측되지 않은 요인(기술력 · 경영 역량 등)이 지원 선정과 성장에 함께
#         영향을 주었다면, 그 요인이 얼마나 강해야 결론이 뒤집히는가?"
#
# Γ (감마) : 관측 공변량이 같은 매칭 쌍 안에서, 숨은 요인 때문에 한 기업이
#            지원받을 오즈가 다른 기업보다 최대 Γ배 높을 수 있다고 가정
#            Γ = 1 → 숨은 교란 없음 (현재 PSM의 가정)
# 임계 Γ  : 단측 p값 상한이 .05를 넘기 시작하는 Γ  → 클수록 강건
#
# 설계
#   · 표본     : matched_dataset_segPSM.xlsx (1:1 부문별 PSM, 매칭 쌍 = subclass)
#   · 결과     : 쌍별 차이 d = ΔY_처치 − ΔY_통제,  ΔY = Y_post − Y_2019 (#4-1 정의)
#   · 검정     : (1) Wilcoxon 부호순위 Rosenbaum 상한 (주, rbounds::psens 와 동일 원리)
#                (2) Huber M-검정 상한 (sensitivitymw::senmw, 극단값에 덜 민감 — 보조)
#   · 방향     : 쌍 차이 평균의 부호 방향으로 단측 검정 (음(−)의 효과는 부호를 바꿔 검정)
#   · 비교     : DR-DID 전체 / 횟수별(1·2·3회) / 금액 3분위별(하위·중위·상위)
#                — 하위집단은 "해당 처치기업 + 그 매칭 짝" 쌍만 사용 (쌍 구조 유지)
#   · 벤치마크 : 매칭 전 표본의 성향점수 로짓에서 관측 공변량 1 SD가 지원 오즈를
#                바꾸는 배수 exp(|β|·SD) → 임계 Γ와 비교해 크기를 해석
#
# 주의
#   · 쌍 차이만 사용하므로 DR-DID의 회귀 보정은 반영되지 않음 → Γ=1 p값은 DR-DID p값과 다를 수 있음
#   · 숨은 요인이 결과를 거의 완벽히 예측한다는 최악의 경우를 가정한 보수적 상한
#
# 실행 위치 : matched_dataset_segPSM.xlsx, sobujang_integrated_with_valuesearch+Patent.xlsx 가 있는 폴더
# 작성 : 2026-11-15
# ==============================================================================

packages <- c("readxl", "dplyr", "tidyr", "writexl", "sensitivitymw")
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
POST_YEARS <- c(2024, 2025)
ALPHA      <- 0.05
GAMMA_MAX  <- 10                                  # 임계 Γ 탐색 상한
GAMMA_GRID <- c(1, 1.1, 1.2, 1.3, 1.4, 1.5, 1.75, 2, 2.5, 3)   # Γ별 p값 표
seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")

to_num   <- function(x) as.numeric(gsub(",", "", as.character(x)))
sig_mark <- function(p) ifelse(is.na(p), "", ifelse(p < .001, "***", ifelse(p < .01, "**",
                         ifelse(p < .05, "*", ifelse(p < .1, ".", "ns")))))

# ------------------------------------------------------------------------------
# 1. 데이터 (#7과 동일)
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
    log_lbcost = l1_na(ycol(df, paste0("lbcost", y)))
  )
}
outcomes <- data.frame(
  var   = c("log_asset", "log_sales", "log_debt", "op_profit", "log_export", "log_rdcost", "log_lbcost"),
  label = c("ln자산", "ln매출", "ln부채", "영업이익", "ln(수출+1)", "ln(개발비+1)", "ln(인건비+1)"),
  stringsAsFactors = FALSE)

# ------------------------------------------------------------------------------
# 3. Rosenbaum 상한 함수
# ------------------------------------------------------------------------------
# (1) Wilcoxon 부호순위 검정의 Γ 상한 p값 (정규근사, 0 차이 제외 · 동순위 평균순위)
#     T = 양(+)의 차이 순위합,  p = Γ/(1+Γ)
#     E(T) = p·Σr,  Var(T) = p(1−p)·Σr²   → 단측 상한 p = 1 − Φ((T − E)/√Var)
wilcox_upper_p <- function(d, gamma) {
  d <- d[!is.na(d) & d != 0]
  if (length(d) < 5) return(NA_real_)
  r  <- rank(abs(d))
  T  <- sum(r[d > 0])
  pg <- gamma / (1 + gamma)
  E  <- pg * sum(r); V <- pg * (1 - pg) * sum(r^2)
  1 - pnorm((T - E) / sqrt(V))
}

# (2) Huber M-검정의 Γ 상한 p값 (Rosenbaum 2007, sensitivitymw)
huber_upper_p <- function(yt, yc, gamma) {
  ok <- !is.na(yt) & !is.na(yc)
  if (sum(ok) < 5 || all(yt[ok] == yc[ok])) return(NA_real_)
  tryCatch(senmw(cbind(yt[ok], yc[ok]), gamma = gamma, method = "h")$pval,
           error = function(e) NA_real_)
}

# 임계 Γ : 상한 p값이 ALPHA가 되는 Γ (Γ=1에서 이미 비유의면 NA, GAMMA_MAX 초과면 Inf)
critical_gamma <- function(pfun) {
  p1 <- pfun(1)
  if (is.na(p1) || p1 >= ALPHA) return(NA_real_)
  pmax_ <- pfun(GAMMA_MAX)
  if (is.na(pmax_)) return(NA_real_)
  if (pmax_ < ALPHA) return(Inf)
  uniroot(function(g) pfun(g) - ALPHA, c(1, GAMMA_MAX), tol = 1e-3)$root
}

gamma_grade <- function(g) ifelse(is.na(g), "Γ=1에서 비유의",
                       ifelse(g < 1.1, "매우 취약 (<1.1)",
                       ifelse(g < 1.5, "보통 (1.1~1.5)",
                       ifelse(g < 2.0, "꽤 강건 (1.5~2.0)", "매우 강건 (≥2.0)"))))

# ------------------------------------------------------------------------------
# 4. 부문 × 연도 × 변수 × 비교별 계산
# ------------------------------------------------------------------------------
results <- list(); curves <- list()
for (s in names(seg_labels)) {
  seg_name <- seg_labels[s]
  sub <- dat %>% filter(seg == s)

  # 금액 3분위 (#4-3 · #7과 동일 : 처치군 중 금액 > 0 기준 33.3% · 66.7%)
  pos <- sub$fund_tot[sub$treat_n == 1 & sub$fund_tot > 0]
  q33 <- quantile(pos, 0.333, na.rm = TRUE); q67 <- quantile(pos, 0.667, na.rm = TRUE)

  p19 <- year_vars(sub, 2019)
  for (py in POST_YEARS) {
    post <- year_vars(sub, py)
    for (k in seq_len(nrow(outcomes))) {
      v  <- outcomes$var[k]
      dY <- post[[v]] - p19[[v]]

      # 매칭 쌍 구성 : subclass별 처치 1 · 통제 1
      long <- data.frame(sc = sub$subclass, tr = sub$treat_n, dY = dY,
                         nf = sub$n_funded, ft = sub$fund_tot)
      trt <- long %>% filter(tr == 1) %>% select(sc, yt = dY, nf, ft)
      ctl <- long %>% filter(tr == 0) %>% select(sc, yc = dY)
      pr  <- inner_join(trt, ctl, by = "sc") %>% filter(!is.na(yt), !is.na(yc)) %>%
        mutate(d   = yt - yc,
               cnt = paste0(nf, "회"),
               amt = ifelse(ft <= q33, "하위", ifelse(ft <= q67, "중위", "상위")))

      comparisons <- list(
        list(분석 = "DR-DID", 그룹 = "전체", idx = rep(TRUE, nrow(pr))),
        list(분석 = "횟수별", 그룹 = "1회", idx = pr$cnt == "1회"),
        list(분석 = "횟수별", 그룹 = "2회", idx = pr$cnt == "2회"),
        list(분석 = "횟수별", 그룹 = "3회", idx = pr$cnt == "3회"),
        list(분석 = "금액별", 그룹 = "하위", idx = pr$amt == "하위"),
        list(분석 = "금액별", 그룹 = "중위", idx = pr$amt == "중위"),
        list(분석 = "금액별", 그룹 = "상위", idx = pr$amt == "상위"))

      for (cmp in comparisons) {
        q <- pr[cmp$idx, ]
        if (nrow(q) < 5) next
        dir <- ifelse(mean(q$d) < 0, -1, 1)          # 효과 방향으로 단측 검정
        dd  <- dir * q$d; yt <- dir * q$yt; yc <- dir * q$yc

        pw <- function(g) wilcox_upper_p(dd, g)
        ph <- function(g) huber_upper_p(yt, yc, g)
        gw <- critical_gamma(pw)
        gh <- tryCatch(critical_gamma(ph), error = function(e) NA_real_)

        results[[length(results) + 1]] <- data.frame(
          분석 = cmp$분석, 부문 = seg_name, 그룹 = cmp$그룹, post_year = py, 변수 = outcomes$label[k],
          쌍수 = nrow(q), 쌍차이_평균 = mean(q$d), 쌍차이_중앙값 = median(q$d),
          방향 = ifelse(dir > 0, "+", "−"),
          W_p_Γ1 = pw(1), W_sig = sig_mark(pw(1)), W_임계Γ = gw,
          M_p_Γ1 = ph(1), M_sig = sig_mark(ph(1)), M_임계Γ = gh,
          판정 = gamma_grade(gw), row.names = NULL)

        curves[[length(curves) + 1]] <- data.frame(
          분석 = cmp$분석, 부문 = seg_name, 그룹 = cmp$그룹, post_year = py, 변수 = outcomes$label[k],
          t(setNames(sapply(GAMMA_GRID, pw), paste0("Γ=", GAMMA_GRID))), check.names = FALSE)

        cat(sprintf("%-6s %s %-4s %d %-12s 쌍 %3d | p(Γ=1) %.3f%-4s 임계Γ W %5s  M %5s\n",
                    cmp$분석, seg_name, cmp$그룹, py, outcomes$label[k], nrow(q),
                    pw(1), sig_mark(pw(1)),
                    ifelse(is.na(gw), "—", sprintf("%.2f", gw)),
                    ifelse(is.na(gh), "—", sprintf("%.2f", gh))))
      }
    }
  }
}
res_df   <- bind_rows(results)
curve_df <- bind_rows(curves)

# ------------------------------------------------------------------------------
# 5. 벤치마크 : 관측 공변량 1 SD가 지원 오즈를 바꾸는 배수 (매칭 전 표본, 부문별 로짓)
#    → "결과를 뒤집으려면 숨은 요인이 ○○보다 강해야 한다"로 해석
# ------------------------------------------------------------------------------
raw <- read_excel("sobujang_integrated_with_valuesearch+Patent.xlsx", col_types = "text")
pre <- raw %>%
  filter(group %in% c("Treatgroup", "Controlgroup")) %>%
  transmute(seg = as.character(seg), treat = as.integer(group == "Treatgroup"),
            자산 = to_num(`2019/Annual S15000.자산총계`), 매출 = to_num(`2019/Annual S21100.총수익`),
            부채 = to_num(`2019/Annual S18000.부채총계`), 업력 = to_num(age), opm = to_num(opm2019),
            연구개발비 = to_num(rdcost2019), 수출 = to_num(exportamt2019), 인건비 = to_num(lbcost2019),
            산업 = factor(KSIC_mid_code), 지역 = factor(region)) %>%
  filter(!is.na(자산), !is.na(매출), !is.na(부채), !is.na(업력), !is.na(opm)) %>%
  mutate(log_자산 = log(자산 + 1), log_매출 = log(매출 + 1), log_부채 = log(부채 + 1),
         log_업력 = log(업력 + 1), log_연구개발비 = log(연구개발비 + 1),
         log_수출 = log(수출 + 1), log_인건비 = log(인건비 + 1))
base <- c("log_자산", "log_매출", "log_부채", "log_업력", "log_연구개발비", "log_수출", "log_인건비", "opm")

bench <- list()
for (s in names(seg_labels)) {
  b <- pre %>% filter(seg == s) %>% droplevels()
  rhs <- base
  for (f in c("산업", "지역")) if (nlevels(b[[f]]) > 1) rhs <- c(rhs, f)
  fit <- suppressWarnings(glm(reformulate(rhs, "treat"), data = b, family = binomial()))
  co  <- coef(fit)[base]
  sds <- sapply(b[base], sd, na.rm = TRUE)
  bench[[s]] <- data.frame(부문 = seg_labels[s], 공변량 = base,
                           계수 = co, SD = sds, `1SD당_오즈배수` = exp(abs(co) * sds),
                           check.names = FALSE, row.names = NULL)
}
bench_df <- bind_rows(bench) %>% arrange(부문, desc(`1SD당_오즈배수`))
bench_max <- bench_df %>% group_by(부문) %>% slice_max(`1SD당_오즈배수`, n = 1) %>%
  select(부문, 최강_관측공변량 = 공변량, 최강_오즈배수 = `1SD당_오즈배수`)

res_df <- res_df %>% left_join(bench_max, by = "부문") %>%
  mutate(벤치마크_비교 = ifelse(is.na(W_임계Γ), "",
                         ifelse(W_임계Γ > 최강_오즈배수, "최강 관측공변량보다 강해야 뒤집힘",
                                "최강 관측공변량 수준 이하에서 뒤집힘")))

# ------------------------------------------------------------------------------
# 6. 요약 · 저장
# ------------------------------------------------------------------------------
fmtg <- function(g) ifelse(is.na(g), "—", ifelse(is.infinite(g), paste0(">", GAMMA_MAX), sprintf("%.2f", g)))
wide_tbl <- function(an) {
  res_df %>% filter(분석 == an) %>%
    mutate(값 = paste0(fmtg(W_임계Γ), " (p=", sprintf("%.3f", W_p_Γ1), ")")) %>%
    select(부문, 그룹, 변수, post_year, 값) %>%
    pivot_wider(names_from = post_year, values_from = 값, names_prefix = "임계Γ_")
}
summary_tbl <- res_df %>%
  group_by(분석, 부문) %>%
  summarise(비교수 = n(),
            `Γ=1 유의(p<.05)` = sum(W_p_Γ1 < ALPHA, na.rm = TRUE),
            `임계Γ ≥1.2` = sum(W_임계Γ >= 1.2, na.rm = TRUE),
            `임계Γ ≥1.5` = sum(W_임계Γ >= 1.5, na.rm = TRUE),
            `임계Γ ≥2.0` = sum(W_임계Γ >= 2.0, na.rm = TRUE),
            임계Γ_최대 = suppressWarnings(max(W_임계Γ, na.rm = TRUE)), .groups = "drop") %>%
  mutate(임계Γ_최대 = ifelse(is.infinite(임계Γ_최대) & 임계Γ_최대 < 0, NA, 임계Γ_최대))
print(summary_tbl)

write_xlsx(list(
  `00_설정`     = data.frame(
    항목 = c("표본", "결과", "주 검정", "보조 검정", "방향", "임계 Γ", "판정 구간", "벤치마크", "주의"),
    값 = c("1:1 부문별 PSM 매칭 쌍 (subclass)",
           "쌍 차이 d = ΔY_처치 − ΔY_통제, ΔY = Y_post − Y_2019",
           "Wilcoxon 부호순위 Rosenbaum 상한 (정규근사, 0 차이 제외)",
           "Huber M-검정 상한 (sensitivitymw::senmw, method = 'h')",
           "쌍 차이 평균의 부호 방향 단측 검정",
           paste0("상한 p값이 ", ALPHA, "가 되는 Γ (— = Γ=1에서 이미 비유의)"),
           "<1.1 매우 취약 / 1.1~1.5 보통 / 1.5~2.0 꽤 강건 / ≥2.0 매우 강건",
           "매칭 전 부문별 로짓에서 관측 공변량 1 SD당 오즈 배수 exp(|β|·SD)",
           "쌍 차이만 사용(회귀 보정 없음) · 최악의 경우를 가정한 보수적 상한")),
  `01_요약`     = summary_tbl,
  `02_DR-DID`   = wide_tbl("DR-DID"),
  `03_횟수별`   = wide_tbl("횟수별"),
  `04_금액별`   = wide_tbl("금액별"),
  `05_벤치마크` = bench_df,
  `06_Γ별_p값`  = curve_df,
  `07_전체결과` = res_df
), "Rosenbaum_민감도분석_DR_횟수_금액.xlsx")
cat("\n✓ 저장 : Rosenbaum_민감도분석_DR_횟수_금액.xlsx\n")
