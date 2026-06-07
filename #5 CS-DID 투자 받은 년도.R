################################################################################
# 소부장 R&D 지원 정책 — Callaway-Sant'Anna Staggered DID
#
# 분석 내용:
#   1. CS-DID: ATT(g, t) — 코호트×시점 처치 효과 추정
#   2. 집계: θ_simple (전체 평균), θ_group (코호트별), θ_dynamic (동적 효과)
#   3. Event Study 그래프 (처치 후 경과 연도별 효과 시각화)
#   4. 부문별(소재·부품·장비) 독립 분석
#   5. 현행 PSM-DID(횟수별 3회) 결과와 직접 비교
#   6. Excel 저장
#
# 데이터: matched_dataset_segPSM.xlsx (PSM 매칭 완료, 2018~2025 패널)
#
# CS-DID 핵심 설계:
#   - first.treat = 최초 정부 R&D 지원 연도 (2020/2021/2022)
#   - never-treated = 통제군 538개사 (first.treat = 0)
#   - Pre 기간: 2018·2019 (2개년) → 평행추세 검정 가능
#   - Post 기간: 2020~2025 (6개년) → 동적 효과 e=0~+5
#   - 공변량: PSM 매칭 공변량 10개 (xformla 인수)
#
# 코호트 구조:
#   g=2020: 307개사 (부품176, 소재89, 장비42)
#   g=2021: 114개사 (부품 54, 소재36, 장비24)
#   g=2022: 117개사 (부품 59, 소재30, 장비28)
#
# 참고: _4-3_DID_투자금액_효과.R 설계 패턴 계승
# 작성: 20260607
################################################################################

# ==============================================================================
# 0. 패키지 설치 및 로드
# ==============================================================================
packages <- c(
  "readxl", "dplyr", "tidyr", "ggplot2",
  "writexl", "gridExtra", "grid", "scales",
  "did",        # Callaway & Sant'Anna (2021) — CS-DID 핵심
  "BMisc"       # did 패키지 의존성
)

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}

select <- dplyr::select
lag    <- dplyr::lag

setwd("/Users/ultra/PSM-DID")   # ← 경로 수정

# ==============================================================================
# 1. 데이터 로드 및 수치 변환 (_4-3 코드 동일)
# ==============================================================================
raw <- read_excel("matched_dataset_segPSM.xlsx", col_types = "text")

numeric_cols <- c(
  "fund2020", "fund2021", "fund2022",
  "gfundvol2020", "gfundvol2021", "gfundvol2022",
  "fundedpattern", "funded", "seg", "treat",
  grep("S15000|S18000|S18100|S21100|S25000|S05000|S21190|S21195|692080|692081|692082|692083|692084|692085|692086|692087|692088|124100|152000",
       names(raw), value = TRUE),
  paste0("exportamt", 2018:2025),
  paste0("rdcost",    2018:2025),
  paste0("lbcost",    2018:2025),
  paste0("opm",       2018:2025),
  paste0("p",         2019:2025),
  "age", "subclass", "weights", "distance",
  grep("^log_|^labor_prod", names(raw), value = TRUE)
)
numeric_cols <- intersect(numeric_cols, names(raw))
for (col in numeric_cols) raw[[col]] <- suppressWarnings(as.numeric(raw[[col]]))

# 횟수 변수 생성 (이진 funded → 횟수 n_funded)
raw$n_funded       <- raw$fund2020 + raw$fund2021 + raw$fund2022
raw$total_gfundvol <- raw$gfundvol2020 + raw$gfundvol2021 + raw$gfundvol2022

# ==============================================================================
# 2. CS-DID 핵심 변수 생성
# ==============================================================================

# ── first.treat: 최초 수혜 연도 (never-treated = 0) ──────────────────────────
# 2020 최초: fund2020=1 (n_funded=1,2,3 모두 포함)
# 2021 최초: fund2020=0, fund2021=1
# 2022 최초: fund2020=0, fund2021=0, fund2022=1
raw$first_treat <- ifelse(raw$treat == 0, 0,
                          ifelse(raw$fund2020 == 1, 2020,
                                 ifelse(raw$fund2021 == 1, 2021, 2022)))

# ── 기업 ID (정수형 필수) ────────────────────────────────────────────────────
raw$firm_id <- as.integer(factor(raw$사업자번호_clean))

cat("=== CS-DID 코호트 구성 ===\n")
print(table(raw$first_treat, raw$seg_name))

# ==============================================================================
# 3. Long Panel 변환 (wide → long, 연도 = 2018~2025)
# ==============================================================================

# ── 분석 변수 정의 (_4-3 make_var_info 참조, 특허 제외) ──────────────────────
var_info_cs <- list(
  list(vn="ln_asset", kr="ln(자산)",      raw_col_prefix="Annual S15000.자산총계",
       transform="log"),
  list(vn="ln_rev",   kr="ln(매출)",      raw_col_prefix="Annual S21100.총수익",
       transform="log"),
  list(vn="ln_debt",  kr="ln(부채)",      raw_col_prefix="Annual S18000.부채총계",
       transform="log"),
  list(vn="opm",      kr="OPM",           raw_col_prefix="opm",
       transform="raw"),
  list(vn="ln_exp1",  kr="ln(수출+1)",    raw_col_prefix="exportamt",
       transform="log1p"),
  list(vn="ln_rd1",   kr="ln(개발비+1)",  raw_col_prefix="rdcost",
       transform="log1p"),
  list(vn="ln_lb1",   kr="ln(인건비+1)",  raw_col_prefix="lbcost",
       transform="log1p")
)

YEARS <- 2018:2025

# ── 변환 함수 (_4-3 ln_transform 동일) ────────────────────────────────────────
apply_transform <- function(x, type) {
  x <- suppressWarnings(as.numeric(x))
  switch(type,
         "log"   = log(pmax(x, 1, na.rm = TRUE)),
         "log1p" = log(pmax(x, 0, na.rm = TRUE) + 1),
         "raw"   = x
  )
}

# ── Wide → Long 변환 ──────────────────────────────────────────────────────────
make_long_panel <- function(df, vi) {
  # 연도별 해당 변수값 추출
  year_vals <- sapply(YEARS, function(yr) {
    if (vi$transform == "raw") {
      col <- paste0(vi$raw_col_prefix, yr)
    } else {
      col <- paste0(yr, "/", vi$raw_col_prefix)
    }
    if (col %in% names(df)) apply_transform(df[[col]], vi$transform)
    else rep(NA_real_, nrow(df))
  })
  colnames(year_vals) <- as.character(YEARS)
  
  # Long format 조립
  long <- lapply(seq_along(YEARS), function(j) {
    yr <- YEARS[j]
    data.frame(
      firm_id      = df$firm_id,
      year         = yr,
      seg          = df$seg,
      seg_name     = df$seg_name,
      treat        = df$treat,
      first_treat  = df$first_treat,
      n_funded     = df$n_funded,
      Y            = year_vals[, j],
      # PSM 공변량 (사전값 2019 고정 — xformla용)
      log_asset_19 = apply_transform(df[["2019/Annual S15000.자산총계"]], "log"),
      log_rev_19   = apply_transform(df[["2019/Annual S21100.총수익"]],   "log"),
      log_debt_19  = apply_transform(df[["2019/Annual S18000.부채총계"]], "log"),
      opm_19       = apply_transform(df[["opm2019"]], "raw"),
      log_exp_19   = apply_transform(df[["exportamt2019"]], "log1p"),
      log_rd_19    = apply_transform(df[["rdcost2019"]],    "log1p"),
      log_lb_19    = apply_transform(df[["lbcost2019"]],    "log1p"),
      log_age_19   = log(pmax(as.numeric(df[["age"]]), 1, na.rm=TRUE)),
      stringsAsFactors = FALSE
    )
  })
  bind_rows(long)
}

# ==============================================================================
# 4. 유틸리티 함수
# ==============================================================================
sig_label <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "***",
                ifelse(p < 0.01,  "**",
                       ifelse(p < 0.05,  "*",
                              ifelse(p < 0.1,   ".", "ns")))))
}

# ATT(g,t) 결과 → 깔끔한 data.frame
tidy_att_gt <- function(att_obj, seg_name, var_name) {
  df <- data.frame(
    seg      = seg_name,
    var      = var_name,
    cohort   = att_obj$group,
    year     = att_obj$t,
    ATT      = round(att_obj$att,    4),
    SE       = round(att_obj$se,     4),
    ci_lower = round(att_obj$att - 1.96 * att_obj$se, 4),
    ci_upper = round(att_obj$att + 1.96 * att_obj$se, 4),
    p_val    = round(2 * (1 - pnorm(abs(att_obj$att / att_obj$se))), 4),
    stringsAsFactors = FALSE
  )
  df$sig <- sig_label(df$p_val)
  df$event_time <- df$year - df$cohort   # e = t - g
  df
}

# 집계 결과 → data.frame
tidy_aggte <- function(agg_obj, seg_name, var_name, type) {
  n <- length(agg_obj$overall.att)
  data.frame(
    seg      = seg_name,
    var      = var_name,
    type     = type,
    term     = if (type == "simple") "theta_simple"
    else if (type == "group") paste0("g=", agg_obj$egt)
    else paste0("e=", agg_obj$egt),
    ATT      = round(agg_obj$overall.att, 4),
    SE       = round(agg_obj$overall.se,  4),
    ci_lower = round(agg_obj$overall.att - 1.96 * agg_obj$overall.se, 4),
    ci_upper = round(agg_obj$overall.att + 1.96 * agg_obj$overall.se, 4),
    p_val    = round(2*(1-pnorm(abs(agg_obj$overall.att/agg_obj$overall.se))),4),
    stringsAsFactors = FALSE
  )
}

# ==============================================================================
# 5. CS-DID 메인 분석 — 부문별 × 결과변수별
# ==============================================================================

segments <- c("소재", "부품", "장비")

# 결과 컨테이너
res_att_gt  <- list()   # ATT(g,t) 전체
res_simple  <- list()   # θ_simple (전체 평균)
res_group   <- list()   # θ_group  (코호트별)
res_dynamic <- list()   # θ_dynamic (이벤트 시간별)
res_compare <- list()   # 현행 PSM-DID vs CS-DID 비교

cat("\n", paste(rep("=", 70), collapse=""), "\n")
cat("  CS-DID (Callaway & Sant'Anna, 2021) 분석 시작\n")
cat(paste(rep("=", 70), collapse=""), "\n")

for (seg in segments) {
  seg_code <- switch(seg, "소재"=1, "부품"=2, "장비"=3)
  df_seg   <- raw[raw$seg == seg_code, ]
  
  cat(sprintf("\n%s %s %s\n",
              paste(rep("─", 25), collapse=""),
              seg,
              paste(rep("─", 25), collapse="")))
  cat(sprintf("  N=%d | 처치=%d | 통제=%d\n",
              nrow(df_seg), sum(df_seg$treat), sum(df_seg$treat==0)))
  cat(sprintf("  코호트: g2020=%d, g2021=%d, g2022=%d\n",
              sum(df_seg$first_treat==2020),
              sum(df_seg$first_treat==2021),
              sum(df_seg$first_treat==2022)))
  
  for (vi in var_info_cs) {
    vn <- vi$vn
    kr <- vi$kr
    cat(sprintf("\n  [%s - %s]\n", seg, kr))
    
    # ── Long panel 생성 ─────────────────────────────────────────────────────
    panel_long <- make_long_panel(df_seg, vi)
    panel_long <- panel_long[!is.na(panel_long$Y), ]
    
    # ── CS-DID 실행 (att_gt) ────────────────────────────────────────────────
    # xformla: PSM 사전 공변량 7개 (산업·지역은 PSM에서 이미 처리됨)
    # control_group = "nevertreated": 통제군 538개사만 사용
    # base_period = "universal": 처치 직전 연도를 기준(pre-trend용)
    # clusterse = TRUE: 기업 수준 클러스터 표준오차
    
    att_result <- tryCatch({
      att_gt(
        yname         = "Y",
        tname         = "year",
        idname        = "firm_id",
        gname         = "first_treat",
        data          = panel_long,
        xformla       = ~ log_asset_19 + log_rev_19 + log_debt_19 +
          opm_19 + log_exp_19 + log_rd_19 +
          log_lb_19 + log_age_19,
        control_group = "nevertreated",   # never-treated 538개사만 통제군
        base_period   = "universal",      # 전체 Pre 기간 공통 기준
        est_method    = "dr",             # Doubly Robust 추정
        clustervars   = "firm_id",        # 기업 클러스터 SE
        bstrap        = TRUE,             # 부트스트랩 표준오차
        biters        = 999,              # 부트스트랩 반복
        print_details = FALSE
      )
    }, error = function(e) {
      cat(sprintf("    ⚠ att_gt 오류: %s\n", e$message))
      NULL
    })
    
    if (is.null(att_result)) next
    
    # ── ATT(g,t) 저장 ────────────────────────────────────────────────────────
    df_att <- tidy_att_gt(att_result, seg, kr)
    res_att_gt[[length(res_att_gt)+1]] <- df_att
    
    # ── 집계 1: θ_simple (전체 ATT 단순평균) ─────────────────────────────────
    agg_simple <- tryCatch(
      aggte(att_result, type = "simple"),
      error = function(e) NULL
    )
    if (!is.null(agg_simple)) {
      df_s <- data.frame(
        seg=seg, var=kr, type="simple",
        term="θ_simple",
        ATT      = round(agg_simple$overall.att, 4),
        SE       = round(agg_simple$overall.se,  4),
        ci_lower = round(agg_simple$overall.att - 1.96*agg_simple$overall.se, 4),
        ci_upper = round(agg_simple$overall.att + 1.96*agg_simple$overall.se, 4),
        p_val    = round(2*(1-pnorm(abs(agg_simple$overall.att/agg_simple$overall.se))),4)
      )
      df_s$sig <- sig_label(df_s$p_val)
      res_simple[[length(res_simple)+1]] <- df_s
      cat(sprintf("    θ_simple = %+.4f (SE=%.4f) %s\n",
                  df_s$ATT, df_s$SE, df_s$sig))
    }
    
    # ── 집계 2: θ_group (코호트별 평균) ──────────────────────────────────────
    agg_group <- tryCatch(
      aggte(att_result, type = "group"),
      error = function(e) NULL
    )
    if (!is.null(agg_group)) {
      df_g <- data.frame(
        seg      = seg,
        var      = kr,
        type     = "group",
        term     = paste0("g=", agg_group$egt),
        ATT      = round(agg_group$att.egt, 4),
        SE       = round(agg_group$se.egt,  4),
        ci_lower = round(agg_group$att.egt - 1.96*agg_group$se.egt, 4),
        ci_upper = round(agg_group$att.egt + 1.96*agg_group$se.egt, 4),
        p_val    = round(2*(1-pnorm(abs(agg_group$att.egt/agg_group$se.egt))),4)
      )
      df_g$sig <- sig_label(df_g$p_val)
      res_group[[length(res_group)+1]] <- df_g
      for (i in seq_len(nrow(df_g))) {
        cat(sprintf("    %s: ATT=%+.4f %s\n",
                    df_g$term[i], df_g$ATT[i], df_g$sig[i]))
      }
    }
    
    # ── 집계 3: θ_dynamic (이벤트 시간별) ────────────────────────────────────
    agg_dynamic <- tryCatch(
      aggte(att_result, type = "dynamic"),
      error = function(e) NULL
    )
    if (!is.null(agg_dynamic)) {
      df_d <- data.frame(
        seg       = seg,
        var       = kr,
        type      = "dynamic",
        event_time = agg_dynamic$egt,
        ATT       = round(agg_dynamic$att.egt, 4),
        SE        = round(agg_dynamic$se.egt,  4),
        ci_lower  = round(agg_dynamic$att.egt - 1.96*agg_dynamic$se.egt, 4),
        ci_upper  = round(agg_dynamic$att.egt + 1.96*agg_dynamic$se.egt, 4),
        p_val     = round(2*(1-pnorm(abs(agg_dynamic$att.egt/agg_dynamic$se.egt))),4)
      )
      df_d$sig <- sig_label(df_d$p_val)
      res_dynamic[[length(res_dynamic)+1]] <- df_d
      
      # Pre-trend 판정 (e < 0)
      pre_rows  <- df_d[df_d$event_time < 0, ]
      post_rows <- df_d[df_d$event_time >= 0, ]
      pre_pass  <- all(pre_rows$p_val > 0.1, na.rm=TRUE)
      cat(sprintf("    Pre-trend: %s | Post 유의: e=%s\n",
                  ifelse(pre_pass, "✅ 통과", "⚠ 주의"),
                  paste(post_rows$event_time[post_rows$p_val < 0.05],
                        collapse=",")))
    }
    
    # ── 현행 PSM-DID와의 비교 (부품 3회 수혜) ────────────────────────────────
    # 부품 부문에서만 3회 수혜 코호트(g=2020, n=185) 직접 비교
    if (seg == "부품" && !is.null(agg_group)) {
      df_g_sub <- df_g[df_g$term == "g=2020", ]
      if (nrow(df_g_sub) > 0) {
        res_compare[[length(res_compare)+1]] <- data.frame(
          seg          = seg,
          var          = kr,
          method_PSMDID = NA_real_,   # 현행 결과 수동 기입 (아래 참조)
          method_CSDID  = df_g_sub$ATT,
          sig_CSDID     = df_g_sub$sig,
          note          = "g=2020 코호트 (3회 수혜 포함, N=307)"
        )
      }
    }
  }
}

# 현행 PSM-DID 3회 수혜 결과 (표17 기준) 수동 입력
psm_ref <- data.frame(
  var            = c("ln(자산)","ln(매출)","ln(부채)","OPM",
                     "ln(수출+1)","ln(개발비+1)","ln(인건비+1)"),
  PSM_DID_2024   = c(0.249,  0.184,  0.354, -0.011,  0.698, 2.395, 0.264),
  PSM_DID_sig24  = c("***",  "*",    "*",   "ns",    "ns",  "***", "*"),
  PSM_DID_2025   = c(0.281,  0.241,  0.381, -0.006,  0.712, 2.987, 0.278),
  PSM_DID_sig25  = c("***",  "**",   "*",   "ns",    "ns",  "***", "*"),
  stringsAsFactors = FALSE
)

# ==============================================================================
# 6. Event Study 그래프 — 부문 × 변수
# ==============================================================================

cat("\n=== Event Study 그래프 생성 ===\n")

# 핵심 변수 4개만 그래프 (자산/매출/개발비/OPM)
key_vars <- c("ln(자산)", "ln(매출)", "ln(개발비+1)", "OPM")

plot_list <- list()

if (length(res_dynamic) > 0) {
  df_dyn_all <- bind_rows(res_dynamic)
  
  for (seg in segments) {
    for (kr in key_vars) {
      df_plot <- df_dyn_all[df_dyn_all$seg == seg &
                              df_dyn_all$var == kr, ]
      if (nrow(df_plot) == 0) next
      
      p <- ggplot(df_plot, aes(x = event_time, y = ATT)) +
        # 신뢰구간
        geom_ribbon(aes(ymin = ci_lower, ymax = ci_upper),
                    alpha = 0.15, fill = "#2C5F8A") +
        # 0선 (처치 전 기준선)
        geom_hline(yintercept = 0, linetype = "dashed",
                   color = "gray50", linewidth = 0.7) +
        # 처치 시점 표시
        geom_vline(xintercept = -0.5, linetype = "dotted",
                   color = "red3", linewidth = 0.8) +
        # 계수선
        geom_line(color = "#2C5F8A", linewidth = 1.0) +
        geom_point(aes(shape = ifelse(p_val < 0.05, "sig", "ns")),
                   size = 3, color = "#1A3A5C") +
        scale_shape_manual(values = c("sig"=16, "ns"=1),
                           name = NULL,
                           labels = c("sig"="p<0.05", "ns"="ns")) +
        # 유의 포인트 레이블
        geom_text(data = df_plot[df_plot$p_val < 0.05, ],
                  aes(label = sig),
                  vjust = -1.2, size = 3.5, color = "#C00000") +
        # 연도 레이블 (x축 하단)
        scale_x_continuous(
          breaks = df_plot$event_time,
          labels = paste0("e=", df_plot$event_time,
                          "\n(", df_plot$event_time + 2021, ")")
        ) +
        labs(
          title    = sprintf("[%s] %s — 동적 처치 효과 (CS-DID)", seg, kr),
          subtitle = "처치 후 경과 연도별 ATT | 음영=95% CI | 점선=처치 시점",
          x        = "처치 후 경과 연도 (event time)",
          y        = "ATT (처치 효과)"
        ) +
        theme_bw(base_size = 11) +
        theme(
          plot.title    = element_text(face = "bold", size = 12),
          plot.subtitle = element_text(color = "gray40", size = 9),
          panel.grid.minor = element_blank(),
          legend.position  = "bottom"
        )
      
      plot_key <- paste0(seg, "_", gsub("[\\(\\)\\+]", "", kr))
      plot_list[[plot_key]] <- p
    }
  }
  
  # ── 그래프 저장 ────────────────────────────────────────────────────────────
  if (length(plot_list) > 0) {
    n_plots <- length(plot_list)
    n_cols  <- 2
    n_rows  <- ceiling(n_plots / n_cols)
    
    pdf("CS_DID_EventStudy.pdf",
        width = 14, height = 5 * n_rows, onefile = TRUE)
    grid.arrange(grobs = plot_list, ncol = n_cols,
                 top = textGrob("소부장 CS-DID Event Study — 부문 × 변수",
                                gp = gpar(fontsize=14, fontface="bold")))
    dev.off()
    cat("  -> CS_DID_EventStudy.pdf 저장 완료\n")
    
    # 부품 핵심 변수 단독 저장 (논문용)
    if ("부품_ln자산" %in% names(plot_list)) {
      p_parts <- plot_list[grep("^부품", names(plot_list))]
      pdf("CS_DID_부품_핵심.pdf", width = 14, height = 10)
      grid.arrange(grobs = p_parts, ncol = 2)
      dev.off()
      cat("  -> CS_DID_부품_핵심.pdf 저장 완료\n")
    }
  }
}

# ==============================================================================
# 7. 현행 PSM-DID vs CS-DID 비교표 생성
# ==============================================================================

cat("\n=== PSM-DID vs CS-DID 비교 ===\n")

# CS-DID θ_simple 집계
if (length(res_simple) > 0) {
  df_simple_all <- bind_rows(res_simple)
  
  # 부품 부문 핵심 결과 출력
  df_s_part <- df_simple_all[df_simple_all$seg == "부품", ]
  cat("\n[부품] θ_simple (전체 코호트 평균):\n")
  cat(sprintf("  %-14s %8s %6s %6s\n", "변수", "ATT", "SE", "sig"))
  cat("  ", paste(rep("-", 38), collapse=""), "\n")
  for (i in seq_len(nrow(df_s_part))) {
    cat(sprintf("  %-14s %+8.4f %6.4f %6s\n",
                df_s_part$var[i], df_s_part$ATT[i],
                df_s_part$SE[i],  df_s_part$sig[i]))
  }
}

# 비교표: PSM-DID(3회, 2024) vs CS-DID(θ_simple 부품)
compare_tbl <- tryCatch({
  df_s_part <- bind_rows(res_simple)[bind_rows(res_simple)$seg == "부품", ]
  merged <- merge(psm_ref,
                  df_s_part[, c("var","ATT","SE","sig")],
                  by="var", all.x=TRUE)
  names(merged)[names(merged)=="ATT"] <- "CSDID_simple"
  names(merged)[names(merged)=="SE"]  <- "CSDID_SE"
  names(merged)[names(merged)=="sig"] <- "CSDID_sig"
  merged$consistency <- ifelse(
    !is.na(merged$CSDID_simple) & !is.na(merged$PSM_DID_2024),
    ifelse(sign(merged$CSDID_simple) == sign(merged$PSM_DID_2024),
           "✅ 방향 일치", "⚠ 방향 상이"),
    "N/A"
  )
  merged
}, error = function(e) NULL)

# ==============================================================================
# 8. Excel 저장
# ==============================================================================

cat("\n=== Excel 저장 ===\n")

excel_sheets <- list()

# ATT(g,t) — 코호트×시점 전체
if (length(res_att_gt) > 0) {
  df_att_all <- bind_rows(res_att_gt)
  excel_sheets[["ATT_gt_전체"]] <- df_att_all
  
  # 부문별 분리
  for (seg in segments) {
    key <- paste0("ATT_gt_", seg)
    excel_sheets[[key]] <- df_att_all[df_att_all$seg == seg, ]
  }
}

# θ_simple
if (length(res_simple) > 0) {
  df_s_wide <- bind_rows(res_simple) %>%
    select(seg, var, ATT, SE, ci_lower, ci_upper, p_val, sig) %>%
    pivot_wider(names_from = seg,
                values_from = c(ATT, SE, p_val, sig),
                names_glue = "{seg}_{.value}")
  excel_sheets[["집계_θsimple"]]      <- bind_rows(res_simple)
  excel_sheets[["집계_θsimple_wide"]] <- df_s_wide
}

# θ_group (코호트별)
if (length(res_group) > 0) {
  excel_sheets[["집계_θgroup"]] <- bind_rows(res_group)
}

# θ_dynamic (이벤트 시간별)
if (length(res_dynamic) > 0) {
  df_dyn <- bind_rows(res_dynamic)
  excel_sheets[["집계_θdynamic"]] <- df_dyn
  
  # 부문별 분리
  for (seg in segments) {
    key <- paste0("Dynamic_", seg)
    excel_sheets[[key]] <- df_dyn[df_dyn$seg == seg, ]
  }
}

# PSM-DID vs CS-DID 비교
if (!is.null(compare_tbl)) {
  excel_sheets[["PSM_vs_CS_비교_부품"]] <- compare_tbl
}

# 메타 정보
excel_sheets[["분석_메타"]] <- data.frame(
  항목   = c("방법론", "패키지", "control_group", "est_method",
           "base_period", "bootstrap", "clustervars",
           "공변량", "코호트g2020", "코호트g2021", "코호트g2022",
           "never_treated", "Pre기간", "관측연도"),
  내용   = c("Callaway & Sant'Anna (2021) Staggered DID",
           "did (R CRAN)",
           "nevertreated (통제군 538개사)",
           "Doubly Robust (DR)",
           "universal",
           "Bootstrap B=999",
           "firm_id (기업 클러스터)",
           "log_자산·매출·부채·R&D·수출·인건비·OPM·업력 (2019년 기준, 8개)",
           "307개사 (부품176, 소재89, 장비42)",
           "114개사 (부품54, 소재36, 장비24)",
           "117개사 (부품59, 소재30, 장비28)",
           "538개사",
           "2018·2019 (2개년)",
           "2018~2025 (8개년)"),
  stringsAsFactors = FALSE
)

write_xlsx(excel_sheets, path = "CS_DID_Staggered_Results.xlsx")
cat(sprintf("  -> CS_DID_Staggered_Results.xlsx 저장 완료 (%d 시트)\n",
            length(excel_sheets)))

# ==============================================================================
# 9. 주요 결과 콘솔 요약
# ==============================================================================

cat("\n", paste(rep("=", 70), collapse=""), "\n")
cat("  CS-DID 결과 요약\n")
cat(paste(rep("=", 70), collapse=""), "\n")

if (length(res_simple) > 0) {
  df_smry <- bind_rows(res_simple)
  for (seg in segments) {
    cat(sprintf("\n【%s】 θ_simple (전체 평균 처치 효과)\n", seg))
    sub <- df_smry[df_smry$seg == seg, c("var","ATT","SE","p_val","sig")]
    cat(sprintf("  %-14s %8s %6s %8s %6s\n",
                "변수", "ATT", "SE", "p값", "sig"))
    cat("  ", paste(rep("-", 48), collapse=""), "\n")
    for (i in seq_len(nrow(sub))) {
      cat(sprintf("  %-14s %+8.4f %6.4f %8.4f %6s\n",
                  sub$var[i], sub$ATT[i], sub$SE[i],
                  sub$p_val[i], sub$sig[i]))
    }
  }
}

cat("\n", paste(rep("=", 70), collapse=""), "\n")
cat("  해석 가이드\n")
cat(paste(rep("=", 70), collapse=""), "\n")
cat("
  ① θ_simple: 전체 코호트×시점 ATT의 가중평균
      → PSM-DID 결과와 방향·유의성 일치 = 강건성 확인 ✅

  ② θ_group(g=2020): 2020년 최초 수혜 기업 평균 효과
      → 3회 수혜 185개사 포함. PSM-DID 3회 결과와 가장 유사

  ③ θ_dynamic(e): 처치 후 e년 경과 시점 효과
      e=-1: Pre-trend 검정 (≈0이면 평행추세 통과)
      e= 0: 처치 즉각 효과
      e=+1~+2: 중단기 효과 (장비 시차 포착)
      e=+3~+5: 장기 누적 효과 (부품 심화 여부)

  ④ PSM-DID vs CS-DID 결과 일치 여부:
      일치 → 처치 시점 이질성이 추정에 큰 편향 없음을 의미
      불일치 → 스태거드 처치 보정이 중요함을 의미
      → 어느 쪽이든 논문의 강건성 논의에 기여
")

cat("====== CS-DID 분석 완료 ======\n")