################################################################################
# 소부장 R&D 지원 정책 — Callaway-Sant'Anna Staggered DID  [v2]
#
# [v2 수정 사항]
# BUG-1 수정: make_long_panel() 컬럼명 불일치
#   - exportamt·rdcost·lbcost 계열은 paste0(prefix, yr) 형태
#   - Annual S##### 계열은 paste0(yr, "/Annual ", prefix) 형태
#   - opm·export 계열은 paste0(prefix, yr) 형태
#   → vi$col_type 필드 추가하여 컬럼 생성 방식 명시적으로 구분
#
# BUG-2 수정: base_period = "universal" → "varying"
#   - "universal": 전체 코호트가 동일한 t*=처치 직전 연도를 기준으로 정규화
#     → g=2022의 e=-1 기준이 2021년이 되어
#       2018년=e=-4, 2019년=e=-3으로 event_time 왜곡
#     → "e=-3에서 Pre-trend 위반"처럼 보이지만 실제로는
#       g=2022 코호트의 2019년(처치 2년 전) 대비 2021년(처치 전)의 차이
#   - "varying": 각 코호트가 자신의 처치 직전 연도를 기준으로 독립 정규화
#     → g=2020: 기준=2019년 / g=2021: 기준=2020년 / g=2022: 기준=2021년
#     → 각 코호트의 Pre-trend를 올바르게 검정 가능
#
# 데이터: matched_dataset_segPSM.xlsx
# 코호트: g=2020(307개), g=2021(114개), g=2022(117개)
# 관측연도: 2018~2025 (8개년, 2023 포함 확인)
# Pre 기간: g=2020 → 2018·2019 / g=2021 → 2018~2020 / g=2022 → 2018~2021
################################################################################

# ==============================================================================
# 0. 패키지
# ==============================================================================
packages <- c(
  "readxl", "dplyr", "tidyr", "ggplot2",
  "writexl", "gridExtra", "grid", "scales",
  "did", "BMisc"
)
for (pkg in packages) {
  if (!require(pkg, character.only=TRUE, quietly=TRUE)) {
    install.packages(pkg, repos="https://cloud.r-project.org/")
    library(pkg, character.only=TRUE)
  }
}
select <- dplyr::select

setwd("/Users/ultra/PSM-DID")   # ← 경로 수정

# ==============================================================================
# 1. 데이터 로드 및 수치 변환
# ==============================================================================
raw <- read_excel("matched_dataset_segPSM.xlsx", col_types="text")

numeric_cols <- c(
  "fund2020","fund2021","fund2022",
  "gfundvol2020","gfundvol2021","gfundvol2022",
  "fundedpattern","funded","seg","treat",
  grep("S15000|S18000|S18100|S21100|S25000|S05000|S21190|S21195|692084|692087|124100",
       names(raw), value=TRUE),
  paste0("exportamt", 2018:2025),
  paste0("rdcost",    2018:2025),
  paste0("lbcost",    2018:2025),
  paste0("opm",       2018:2025),
  "age","subclass","weights","distance",
  grep("^log_", names(raw), value=TRUE)
)
numeric_cols <- intersect(numeric_cols, names(raw))
for (col in numeric_cols) raw[[col]] <- suppressWarnings(as.numeric(raw[[col]]))

raw$n_funded       <- raw$fund2020 + raw$fund2021 + raw$fund2022
raw$total_gfundvol <- raw$gfundvol2020 + raw$gfundvol2021 + raw$gfundvol2022

# ==============================================================================
# 2. CS-DID 핵심 변수
# ==============================================================================

# first_treat: 최초 수혜 연도 (never-treated = 0)
raw$first_treat <- ifelse(raw$treat == 0, 0,
                          ifelse(raw$fund2020 == 1, 2020,
                                 ifelse(raw$fund2021 == 1, 2021, 2022)))

# 기업 ID (정수형 필수)
raw$firm_id <- as.integer(factor(raw$사업자번호_clean))

cat("=== 코호트 구성 ===\n")
print(table(raw$first_treat, raw$seg_name))

cat("\n=== 연도별 자산 가용률 (데이터 존재 확인) ===\n")
for (yr in 2018:2025) {
  col <- paste0(yr, "/Annual S15000.자산총계")
  if (col %in% names(raw)) {
    n <- sum(!is.na(suppressWarnings(as.numeric(raw[[col]]))))
    cat(sprintf("  %d: %d/%d (%.1f%%)\n", yr, n, nrow(raw), n/nrow(raw)*100))
  }
}

# ==============================================================================
# 3. 결과변수 정의
# ==============================================================================
# [BUG-1 수정] col_type 필드로 컬럼명 생성 방식 명시적 구분
#   col_type = "annual"  → paste0(yr, "/Annual ", col_key)
#   col_type = "direct"  → paste0(col_key, yr)

var_info_cs <- list(
  list(vn="ln_asset", kr="ln(자산)",      cat="성장성",
       col_key="S15000.자산총계",  col_type="annual",  transform="log"),
  list(vn="ln_rev",   kr="ln(매출)",      cat="성장성",
       col_key="S21100.총수익",    col_type="annual",  transform="log"),
  list(vn="ln_debt",  kr="ln(부채)",      cat="안정성",
       col_key="S18000.부채총계",  col_type="annual",  transform="log"),
  list(vn="opm",      kr="OPM",           cat="수익성",
       col_key="opm",              col_type="direct",  transform="raw"),
  list(vn="ln_exp1",  kr="ln(수출+1)",    cat="활동성",
       col_key="exportamt",        col_type="direct",  transform="log1p"),
  list(vn="ln_rd1",   kr="ln(개발비+1)",  cat="혁신성",
       col_key="rdcost",           col_type="direct",  transform="log1p"),
  list(vn="ln_lb1",   kr="ln(인건비+1)",  cat="고용",
       col_key="lbcost",           col_type="direct",  transform="log1p")
)

YEARS <- 2018:2025

# ==============================================================================
# 4. 변환 및 패널 생성 함수
# ==============================================================================

apply_transform <- function(x, type) {
  x <- suppressWarnings(as.numeric(x))
  switch(type,
         "log"   = log(pmax(x, 1, na.rm=TRUE)),
         "log1p" = log(pmax(x, 0, na.rm=TRUE) + 1),
         "raw"   = x,
         x
  )
}

# [BUG-1 수정] col_type에 따라 올바른 컬럼명 생성
get_col_name <- function(vi, yr) {
  if (vi$col_type == "annual") {
    paste0(yr, "/Annual ", vi$col_key)   # "2020/Annual S15000.자산총계"
  } else {
    paste0(vi$col_key, yr)               # "exportamt2020", "rdcost2020"
  }
}

make_long_panel <- function(df, vi) {
  # 연도별 Y값 추출
  year_vals <- sapply(YEARS, function(yr) {
    col <- get_col_name(vi, yr)
    if (col %in% names(df)) {
      apply_transform(df[[col]], vi$transform)
    } else {
      cat(sprintf("    ⚠ 컬럼 없음: %s\n", col))
      rep(NA_real_, nrow(df))
    }
  })
  colnames(year_vals) <- as.character(YEARS)
  
  # 사전 공변량 (2019 고정)
  cov_asset_19 <- apply_transform(df[["2019/Annual S15000.자산총계"]], "log")
  cov_rev_19   <- apply_transform(df[["2019/Annual S21100.총수익"]],   "log")
  cov_debt_19  <- apply_transform(df[["2019/Annual S18000.부채총계"]], "log")
  cov_opm_19   <- apply_transform(df[["opm2019"]], "raw")
  cov_exp_19   <- apply_transform(df[["exportamt2019"]], "log1p")
  cov_rd_19    <- apply_transform(df[["rdcost2019"]],    "log1p")
  cov_lb_19    <- apply_transform(df[["lbcost2019"]],    "log1p")
  cov_age_19   <- log(pmax(suppressWarnings(as.numeric(df[["age"]])), 1, na.rm=TRUE))
  
  long <- lapply(seq_along(YEARS), function(j) {
    data.frame(
      firm_id     = df$firm_id,
      year        = YEARS[j],
      seg         = df$seg,
      seg_name    = df$seg_name,
      treat       = df$treat,
      first_treat = df$first_treat,
      n_funded    = df$n_funded,
      Y           = year_vals[, j],
      # xformla 공변량 (2019 기준 고정값)
      cov_asset   = cov_asset_19,
      cov_rev     = cov_rev_19,
      cov_debt    = cov_debt_19,
      cov_opm     = cov_opm_19,
      cov_exp     = cov_exp_19,
      cov_rd      = cov_rd_19,
      cov_lb      = cov_lb_19,
      cov_age     = cov_age_19,
      stringsAsFactors = FALSE
    )
  })
  bind_rows(long)
}

sig_label <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "***",
                ifelse(p < 0.01,  "**",
                       ifelse(p < 0.05,  "*",
                              ifelse(p < 0.1,   ".", "ns")))))
}

tidy_att_gt <- function(att_obj, seg_nm, var_nm) {
  data.frame(
    seg        = seg_nm,
    var        = var_nm,
    cohort     = att_obj$group,
    year       = att_obj$t,
    event_time = att_obj$t - att_obj$group,   # e = t - g
    ATT        = round(att_obj$att,    4),
    SE         = round(att_obj$se,     4),
    ci_lower   = round(att_obj$att - 1.96 * att_obj$se, 4),
    ci_upper   = round(att_obj$att + 1.96 * att_obj$se, 4),
    p_val      = round(2*(1-pnorm(abs(att_obj$att/att_obj$se))), 4),
    sig        = sig_label(2*(1-pnorm(abs(att_obj$att/att_obj$se)))),
    stringsAsFactors = FALSE
  )
}

# ==============================================================================
# 5. [BUG-2 수정] base_period = "varying" 사용 이유
# ==============================================================================
# "universal": 모든 코호트가 동일한 기준 시점 사용
#   → g=2022: e=-1 기준 = 2021년, e=-3 = 2019년(Pre-trend 2개 기간 전)
#   → e=-3 ATT 는 2019년 vs 2021년 차이 → 이질적 코호트간 비교 왜곡
#
# "varying": 각 코호트가 자신의 처치 직전 연도를 기준으로 독립 정규화
#   → g=2020: 기준=2019 / g=2021: 기준=2020 / g=2022: 기준=2021
#   → event_time e=-1은 각 코호트의 처치 1년 전 → 코호트간 비교 의미있음
#   → Pre-trend는 e=-2(g=2020: 2018), e=-2(g=2021: 2019), e=-2(g=2022: 2020)
#   → 2018 이전 데이터 없어도 e=-2(=2018년)까지만 추정 → 안전

# ==============================================================================
# 6. 메인 분석
# ==============================================================================
segments <- c("소재","부품","장비")
seg_codes <- c("소재"=1, "부품"=2, "장비"=3)

res_att_gt  <- list()
res_simple  <- list()
res_group   <- list()
res_dynamic <- list()

cat("\n", paste(rep("=",70), collapse=""), "\n")
cat("  CS-DID v2 분석 시작 (base_period='varying')\n")
cat(paste(rep("=",70), collapse=""), "\n")

for (seg in segments) {
  
  df_seg <- raw[raw$seg == seg_codes[seg], ]
  cat(sprintf("\n%s [%s] N=%d (처치=%d, 통제=%d, g2020=%d, g2021=%d, g2022=%d)\n",
              paste(rep("─",20),collapse=""), seg, nrow(df_seg),
              sum(df_seg$treat), sum(df_seg$treat==0),
              sum(df_seg$first_treat==2020),
              sum(df_seg$first_treat==2021),
              sum(df_seg$first_treat==2022)))
  
  for (vi in var_info_cs) {
    
    cat(sprintf("  %-15s ... ", vi$kr))
    
    panel_long <- make_long_panel(df_seg, vi)
    panel_long <- panel_long[!is.na(panel_long$Y), ]
    
    n_obs <- nrow(panel_long)
    n_firms <- length(unique(panel_long$firm_id))
    cat(sprintf("(기업%d, 행%d) ", n_firms, n_obs))
    
    # ── att_gt 추정 ──────────────────────────────────────────────────────────
    att_res <- tryCatch({
      att_gt(
        yname         = "Y",
        tname         = "year",
        idname        = "firm_id",
        gname         = "first_treat",
        data          = panel_long,
        # [BUG-1 수정] xformla 변수명을 make_long_panel의 cov_* 와 일치
        xformla       = ~ cov_asset + cov_rev + cov_debt +
          cov_opm + cov_exp + cov_rd +
          cov_lb + cov_age,
        control_group = "nevertreated",
        # [BUG-2 수정] varying: 코호트별 처치 직전 연도를 각자의 기준으로
        base_period   = "varying",
        est_method    = "dr",
        clustervars   = "firm_id",
        bstrap        = TRUE,
        biters        = 1999,    # v1에서 999 → 1999로 증가 (소표본 SE 안정화)
        print_details = FALSE
      )
    }, error = function(e) {
      cat(sprintf("\n    ⚠ 오류: %s\n", e$message)); NULL
    })
    
    if (is.null(att_res)) next
    
    # ATT(g,t)
    df_att <- tidy_att_gt(att_res, seg, vi$kr)
    res_att_gt[[length(res_att_gt)+1]] <- df_att
    
    # θ_simple
    agg_s <- tryCatch(aggte(att_res, type="simple"), error=function(e) NULL)
    if (!is.null(agg_s)) {
      res_simple[[length(res_simple)+1]] <- data.frame(
        seg=seg, var=vi$kr, cat=vi$cat,
        ATT     = round(agg_s$overall.att, 4),
        SE      = round(agg_s$overall.se,  4),
        ci_lower= round(agg_s$overall.att - 1.96*agg_s$overall.se, 4),
        ci_upper= round(agg_s$overall.att + 1.96*agg_s$overall.se, 4),
        p_val   = round(2*(1-pnorm(abs(agg_s$overall.att/agg_s$overall.se))),4),
        sig     = sig_label(2*(1-pnorm(abs(agg_s$overall.att/agg_s$overall.se)))),
        stringsAsFactors=FALSE
      )
      cat(sprintf("θ_s=%+.3f%s\n", agg_s$overall.att,
                  sig_label(2*(1-pnorm(abs(agg_s$overall.att/agg_s$overall.se))))))
    } else cat("\n")
    
    # θ_group
    agg_g <- tryCatch(aggte(att_res, type="group"), error=function(e) NULL)
    if (!is.null(agg_g)) {
      df_g <- data.frame(
        seg=seg, var=vi$kr, cat=vi$cat,
        cohort  = agg_g$egt,
        ATT     = round(agg_g$att.egt, 4),
        SE      = round(agg_g$se.egt,  4),
        ci_lower= round(agg_g$att.egt - 1.96*agg_g$se.egt, 4),
        ci_upper= round(agg_g$att.egt + 1.96*agg_g$se.egt, 4),
        p_val   = round(2*(1-pnorm(abs(agg_g$att.egt/agg_g$se.egt))),4),
        sig     = sig_label(2*(1-pnorm(abs(agg_g$att.egt/agg_g$se.egt)))),
        stringsAsFactors=FALSE
      )
      res_group[[length(res_group)+1]] <- df_g
    }
    
    # θ_dynamic
    agg_d <- tryCatch(aggte(att_res, type="dynamic"), error=function(e) NULL)
    if (!is.null(agg_d)) {
      df_d <- data.frame(
        seg        = seg,
        var        = vi$kr,
        cat        = vi$cat,
        event_time = agg_d$egt,
        ATT        = round(agg_d$att.egt, 4),
        SE         = round(agg_d$se.egt,  4),
        ci_lower   = round(agg_d$att.egt - 1.96*agg_d$se.egt, 4),
        ci_upper   = round(agg_d$att.egt + 1.96*agg_d$se.egt, 4),
        p_val      = round(2*(1-pnorm(abs(agg_d$att.egt/agg_d$se.egt))),4),
        sig        = sig_label(2*(1-pnorm(abs(agg_d$att.egt/agg_d$se.egt)))),
        stringsAsFactors=FALSE
      )
      res_dynamic[[length(res_dynamic)+1]] <- df_d
    }
  }
}

# ==============================================================================
# 7. Pre-trend 검정 요약 (varying 기준)
# ==============================================================================
cat("\n", paste(rep("=",70),collapse=""), "\n")
cat("  Pre-trend 검정 요약 (e < 0, base_period='varying')\n")
cat(paste(rep("=",70),collapse=""), "\n")
cat("  해석: varying 기준에서 e=-1은 각 코호트의 처치 직전 연도\n")
cat("        g=2020: e=-1=2019 / g=2021: e=-1=2020 / g=2022: e=-1=2021\n")
cat("        e=-2: g=2020→2018, g=2021→2019, g=2022→2020\n\n")

if (length(res_dynamic) > 0) {
  df_dyn_all <- bind_rows(res_dynamic)
  pre_df <- df_dyn_all[df_dyn_all$event_time < 0, ]
  pre_viol <- pre_df[!is.na(pre_df$p_val) & pre_df$p_val < 0.10, ]
  
  if (nrow(pre_viol) > 0) {
    cat("  ⚠ Pre-trend 위반 가능 (p<0.10):\n")
    print(pre_viol[,c("seg","var","event_time","ATT","SE","p_val","sig")])
  } else {
    cat("  ✅ 모든 변수에서 Pre-trend 가정 충족 (e<0 에서 p>0.10)\n")
  }
}

# ==============================================================================
# 8. Event Study 그래프
# ==============================================================================
cat("\n=== Event Study 그래프 생성 ===\n")

# event_time 레이블 생성 (varying 기준 → 코호트 평균 연도 계산)
# g=2020: e=-2=2018, e=0=2020, e=+4=2024
# g=2021: e=-2=2019, e=0=2021, e=+4=2025
# g=2022: e=-2=2020, e=0=2022, e=+3=2025
# → Dynamic plot에서는 event_time만 표시 (연도는 코호트별 상이)

make_event_plot <- function(df_d, seg_nm, var_nm) {
  df_plot <- df_d[df_d$seg==seg_nm & df_d$var==var_nm, ]
  if (nrow(df_plot) == 0) return(NULL)
  
  # Pre/Post 구분
  df_plot$period <- ifelse(df_plot$event_time < 0, "Pre", "Post")
  
  ggplot(df_plot, aes(x=event_time, y=ATT)) +
    # 신뢰구간
    geom_ribbon(aes(ymin=ci_lower, ymax=ci_upper, fill=period),
                alpha=0.15, show.legend=FALSE) +
    scale_fill_manual(values=c("Pre"="gray60","Post"="#2C5F8A")) +
    # 기준선
    geom_hline(yintercept=0, linetype="dashed", color="gray40", linewidth=0.6) +
    # 처치 시점 구분선
    geom_vline(xintercept=-0.5, linetype="dotted", color="red3", linewidth=0.8) +
    # 계수선 (Pre: 회색, Post: 파랑)
    geom_line(aes(color=period), linewidth=1.0) +
    geom_point(aes(color=period,
                   shape=ifelse(p_val < 0.05, "sig","ns")),
               size=3) +
    scale_color_manual(values=c("Pre"="gray50","Post"="#1A3A5C"),
                       name="기간") +
    scale_shape_manual(values=c("sig"=16,"ns"=1), name=NULL,
                       labels=c("sig"="p<0.05","ns"="ns")) +
    # 유의 레이블
    geom_text(data=df_plot[!is.na(df_plot$p_val) & df_plot$p_val < 0.05,],
              aes(label=sig), vjust=-1.3, size=3.5, color="#C00000") +
    # x축: event_time
    scale_x_continuous(breaks=df_plot$event_time,
                       labels=paste0("e=",df_plot$event_time)) +
    labs(
      title    = sprintf("[%s] %s", seg_nm, var_nm),
      subtitle = sprintf("Pre: e<0 (각 코호트 처치 전) | Post: e≥0 | 점선=처치 시점\n기준: base_period='varying' (코호트별 독립 정규화)"),
      x        = "Event Time (e = 처치 후 경과 연도)",
      y        = "ATT"
    ) +
    theme_bw(base_size=10) +
    theme(
      plot.title    = element_text(face="bold", size=11),
      plot.subtitle = element_text(color="gray40", size=8),
      panel.grid.minor = element_blank(),
      legend.position  = "right"
    )
}

key_vars <- c("ln(자산)","ln(매출)","ln(개발비+1)","OPM")
plot_list <- list()

if (length(res_dynamic) > 0) {
  df_dyn_all <- bind_rows(res_dynamic)
  
  for (seg in segments) {
    for (vr in key_vars) {
      p <- make_event_plot(df_dyn_all, seg, vr)
      if (!is.null(p)) {
        key <- paste0(seg,"_",gsub("[\\(\\)\\+/]","",vr))
        plot_list[[key]] <- p
      }
    }
  }
  
  if (length(plot_list) > 0) {
    # 전체 그래프 PDF
    n_rows <- ceiling(length(plot_list) / 2)
    pdf("CS_DID_v2_EventStudy.pdf", width=14, height=5*n_rows, onefile=TRUE)
    grid.arrange(grobs=plot_list, ncol=2,
                 top=textGrob(
                   "소부장 CS-DID Event Study [v2: base_period=varying]",
                   gp=gpar(fontsize=13, fontface="bold")))
    dev.off()
    cat("  -> CS_DID_v2_EventStudy.pdf\n")
    
    # 부품 핵심 4변수 단독
    p_part <- plot_list[grep("^부품", names(plot_list))]
    if (length(p_part) > 0) {
      pdf("CS_DID_v2_부품핵심.pdf", width=14, height=10)
      grid.arrange(grobs=p_part, ncol=2,
                   top=textGrob("[부품] CS-DID Event Study",
                                gp=gpar(fontsize=12, fontface="bold")))
      dev.off()
      cat("  -> CS_DID_v2_부품핵심.pdf\n")
    }
  }
}

# ==============================================================================
# 9. PSM-DID vs CS-DID 비교표
# ==============================================================================

# 현행 PSM-DID 결과 (표17: 부품 3회 DR-DID)
psm_ref <- data.frame(
  var          = c("ln(자산)","ln(매출)","ln(부채)","OPM",
                   "ln(수출+1)","ln(개발비+1)","ln(인건비+1)"),
  PSM_3회_2024 = c(0.249, 0.184, 0.354,-0.011, 0.698, 2.395, 0.264),
  sig_PSM_2024 = c("***",  "*",  "*",   "ns",  "ns",  "***",  "*"),
  PSM_3회_2025 = c(0.281, 0.241, 0.381,-0.006, 0.712, 2.987, 0.278),
  sig_PSM_2025 = c("***", "**",  "*",   "ns",  "ns",  "***",  "*"),
  stringsAsFactors=FALSE
)

# CS-DID θ_group g=2020 (부품, 3회 수혜 185개사 포함)
compare_tbl <- NULL
if (length(res_group) > 0) {
  df_grp_all <- bind_rows(res_group)
  df_g_2020  <- df_grp_all[df_grp_all$seg=="부품" &
                             df_grp_all$cohort==2020,
                           c("var","ATT","SE","p_val","sig")]
  names(df_g_2020)[2:5] <- c("CS_g2020_ATT","CS_g2020_SE",
                             "CS_g2020_p","CS_g2020_sig")
  
  compare_tbl <- merge(psm_ref, df_g_2020, by="var", all.x=TRUE)
  compare_tbl$방향일치 <- ifelse(
    !is.na(compare_tbl$CS_g2020_ATT),
    ifelse(sign(compare_tbl$CS_g2020_ATT)==sign(compare_tbl$PSM_3회_2024),
           "✅ 일치","⚠ 불일치"),
    "N/A"
  )
  compare_tbl$해석 <- ifelse(
    compare_tbl$방향일치=="✅ 일치" & compare_tbl$CS_g2020_p < 0.10,
    "★강건",
    ifelse(compare_tbl$방향일치=="✅ 일치",
           "방향강건(유의성약화)", "검토필요")
  )
}

# ==============================================================================
# 10. Excel 저장
# ==============================================================================
cat("\n=== Excel 저장 ===\n")

xl <- list()

if (length(res_att_gt) > 0) {
  df_att_all <- bind_rows(res_att_gt)
  xl[["ATT_gt_전체"]] <- df_att_all
  for (seg in segments)
    xl[[paste0("ATT_gt_",seg)]] <- df_att_all[df_att_all$seg==seg,]
}

if (length(res_simple) > 0) {
  df_s <- bind_rows(res_simple)
  xl[["θsimple"]] <- df_s
  # 부문별 wide
  df_s_wide <- df_s %>%
    select(var, seg, ATT, SE, p_val, sig) %>%
    pivot_wider(names_from=seg,
                values_from=c(ATT,SE,p_val,sig),
                names_glue="{seg}_{.value}")
  xl[["θsimple_wide"]] <- df_s_wide
}

if (length(res_group) > 0) {
  df_grp <- bind_rows(res_group)
  xl[["θgroup"]] <- df_grp
  # 부품 g=2020 별도 시트
  xl[["θgroup_부품_g2020"]] <- df_grp[df_grp$seg=="부품" &
                                      df_grp$cohort==2020,]
}

if (length(res_dynamic) > 0) {
  df_dyn <- bind_rows(res_dynamic)
  xl[["θdynamic"]] <- df_dyn
  for (seg in segments)
    xl[[paste0("Dynamic_",seg)]] <- df_dyn[df_dyn$seg==seg,]
  
  # Pre-trend 검정 결과만 별도
  pre_only <- df_dyn[df_dyn$event_time < 0 & !is.na(df_dyn$p_val),]
  xl[["PreTrend_검정"]] <- pre_only
}

if (!is.null(compare_tbl))
  xl[["PSM_vs_CS_비교_부품"]] <- compare_tbl

# 버그 수정 요약
xl[["버그수정_요약"]] <- data.frame(
  번호 = c("BUG-1","BUG-2","개선"),
  항목 = c("컬럼명 불일치 → ln수출+1·개발비+1·인건비+1 NA",
         "base_period='universal' → event_time 해석 왜곡",
         "biters 999 → 1999 (소표본 SE 안정화)"),
  원인 = c(
    "exportamt·rdcost·lbcost 계열은 paste0(prefix,yr) 형식이나
     Annual 계열 처리 로직(paste0(yr,'/Annual',prefix))이 잘못 적용됨",
    "universal: 모든 코호트가 동일 기준시점 사용
     → g=2022의 e=-3이 2019년을 가리켜 Pre-trend 위반처럼 보임
     실제로 2017년 데이터 없음에도 e=-4 추정 시도",
    "g=2021(N=54~36), g=2022(N=28~59) 소표본에서 부트스트랩 분산 불안정"),
  수정 = c(
    "col_type 필드로 'annual'/'direct' 구분 → get_col_name() 함수 추가",
    "base_period='varying': 각 코호트가 자신의 처치 직전연도를 독립 기준으로 사용
     g=2020→기준2019 / g=2021→기준2020 / g=2022→기준2021",
    "biters=1999로 증가"),
  stringsAsFactors=FALSE
)

xl[["분석_메타"]] <- data.frame(
  항목 = c("버전","방법론","패키지","control_group","est_method",
         "base_period","bootstrap_biters","clustervars",
         "공변량","관측연도","코호트g2020","코호트g2021","코호트g2022"),
  내용 = c("v2 (2026-06-07)",
         "Callaway & Sant'Anna (2021) Staggered DID",
         "did (R CRAN)",
         "nevertreated (통제군 538개사)",
         "Doubly Robust (DR)",
         "varying [BUG-2 수정: universal→varying]",
         "1999 [개선: 999→1999]",
         "firm_id (기업 클러스터)",
         "cov_asset·rev·debt·opm·exp·rd·lb·age (2019년 기준, 8개) [BUG-1 수정]",
         "2018~2025 (8개년, 2023 포함)",
         "307개사 (부품176, 소재89, 장비42)",
         "114개사 (부품54, 소재36, 장비24)",
         "117개사 (부품59, 소재30, 장비28)"),
  stringsAsFactors=FALSE
)

write_xlsx(xl, path="CS_DID_v2_Results.xlsx")
cat(sprintf("  -> CS_DID_v2_Results.xlsx (%d 시트)\n", length(xl)))

# ==============================================================================
# 11. 콘솔 최종 요약
# ==============================================================================
cat("\n", paste(rep("=",70),collapse=""), "\n")
cat("  최종 결과 요약\n")
cat(paste(rep("=",70),collapse=""), "\n")

if (length(res_simple) > 0) {
  df_s <- bind_rows(res_simple)
  for (seg in segments) {
    sub <- df_s[df_s$seg==seg, c("var","cat","ATT","SE","p_val","sig")]
    cat(sprintf("\n【%s】θ_simple\n", seg))
    cat(sprintf("  %-14s %-5s %8s %6s %8s %5s\n","변수","구분","ATT","SE","p값","sig"))
    cat("  ", paste(rep("-",50),collapse=""),"\n")
    for (i in seq_len(nrow(sub)))
      cat(sprintf("  %-14s %-5s %+8.4f %6.4f %8.4f %5s\n",
                  sub$var[i],sub$cat[i],sub$ATT[i],
                  sub$SE[i],sub$p_val[i],sub$sig[i]))
  }
}

if (!is.null(compare_tbl)) {
  cat("\n【부품】PSM-DID(3회) vs CS-DID(g=2020) 비교\n")
  cat(sprintf("  %-14s %10s %-4s %10s %-5s %8s\n",
              "변수","PSM2024","sig","CS_g2020","sig","판정"))
  cat("  ", paste(rep("-",60),collapse=""),"\n")
  for (i in seq_len(nrow(compare_tbl)))
    cat(sprintf("  %-14s %10.3f %-4s %10.4f %-5s %s\n",
                compare_tbl$var[i],
                compare_tbl$PSM_3회_2024[i], compare_tbl$sig_PSM_2024[i],
                ifelse(is.na(compare_tbl$CS_g2020_ATT[i]),NA,
                       compare_tbl$CS_g2020_ATT[i]),
                ifelse(is.na(compare_tbl$CS_g2020_sig[i]),"",
                       compare_tbl$CS_g2020_sig[i]),
                compare_tbl$해석[i]))
}

cat("\n====== CS-DID v2 완료 ======\n")