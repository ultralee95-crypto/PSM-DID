################################################################################
# 소부장 R&D 지원 정책 — NTIS 변수별 이질성 분석
#
# 분석 변수 (4개):
#   ① rd_stage      : 연구개발단계 (응용=2 vs 개발=3, 2그룹 비교)
#   ② support_type  : 지원유형 (자유공모=1 vs 품목지정=2, 2그룹 비교)
#   ③ research_dur  : 연구기간(년) — Dose-Response + 그룹 비교
#   ④ priv_match_rate: 민간매칭률 — Dose-Response (보완/구축 효과)
#
# 분석 구조 (TLC 코드와 동일):
#   STEP 1. 기술통계 — 변수별 사전(2019) 공변량 균형 확인
#   STEP 2. OLS 이진 비교 — rd_stage, support_type 2그룹 Interaction DID
#   STEP 3. ANOVA + Tukey — 그룹 간 성과 차이 검정
#   STEP 4. Dose-Response — research_dur, priv_match_rate 연속형 분석
#   STEP 5. Random Forest — 4개 변수 중요도 비교
#   STEP 6. 시각화
#   STEP 7. Excel 저장
#
# 데이터: matched_dataset_segPSM.xlsx
#   처치군(treat=1) 538개 / 통제군(treat=0) 538개
#   ① rd_stage      : 처치군만 유효 (통제군 NA → 더미=0으로 처리)
#   ② support_type  : 처치군만 유효 (동일)
#   ③ research_dur  : 처치군만 유효 (통제군=0으로 처리)
#   ④ priv_match_rate: 처치군만 유효 (통제군=0으로 처리)
#
# 참고 문헌:
#   Arrow(1962) 지식공공재 / Lach(2002) crowding-in·out /
#   González et al.(2005) 민간매칭 / Wallsten(2000) SBIR
#
# 수정 이력:
#   v1: 최초 작성 (TLC 코드 _6_TLC_분석.R 구조 기반)
#       BUG 예방: coeftest 위치인덱스, 폰트 속성 완전 제거
################################################################################

# ==============================================================================
# 0. 패키지
# ==============================================================================
pkgs <- c("readxl","dplyr","tidyr","ggplot2","writexl",
          "sandwich","lmtest","car",
          "ranger","gridExtra","grid","scales")

for (p in pkgs) {
  if (!require(p, character.only = TRUE, quietly = TRUE)) {
    install.packages(p, repos = "https://cloud.r-project.org/")
    library(p, character.only = TRUE)
  }
}

select <- dplyr::select
set.seed(2025)
setwd("/Users/ultra/PSM-DID")

# ==============================================================================
# 1. 데이터 로드
# ==============================================================================
cat("── 데이터 로드 중...\n")
raw <- read_excel("matched_dataset_segPSM.xlsx", col_types = "text")

num_cols <- c(
  "fund2020","fund2021","fund2022",
  "gfundvol2020","gfundvol2021","gfundvol2022",
  "funded","seg","treat","age",
  "subclass","weights","distance",
  "rd_stage","support_type","research_dur","priv_match_rate",
  grep("S15000|S18000|S21100|S25000|S05000",
       names(raw), value = TRUE),
  paste0("exportamt", 2018:2025),
  paste0("rdcost",    2018:2025),
  paste0("lbcost",    2018:2025),
  paste0("opm",       2018:2025),
  grep("^log_", names(raw), value = TRUE)
)
num_cols <- intersect(num_cols, names(raw))
for (col in num_cols) {
  raw[[col]] <- suppressWarnings(as.numeric(raw[[col]]))
}
cat(sprintf("  로드 완료: %d행 x %d열\n", nrow(raw), ncol(raw)))

# ==============================================================================
# 2. 기본 전처리
# ==============================================================================
df_all <- raw %>%
  mutate(
    seg_f  = factor(seg, levels=c(1,2,3),
                    labels=c("소재","부품","장비")),
    seg_nm = as.character(seg_f),
    total_vol = gfundvol2020 + gfundvol2021 + gfundvol2022,
    n_funded  = fund2020 + fund2021 + fund2022,
    ln_vol    = log(pmax(total_vol, 0) + 1),
    
    # ── ① rd_stage 더미 (개발연구=3 기준, 응용연구=2 비교)
    # 기초연구(1)는 N=5로 제외 → 응용 vs 개발 2그룹
    rd_stage_v = case_when(
      rd_stage == 2 ~ "응용연구",
      rd_stage == 3 ~ "개발연구",
      TRUE          ~ NA_character_
    ),
    rd_응용더미 = ifelse(!is.na(rd_stage) & rd_stage == 2, 1L, 0L),
    
    # ── ② support_type 더미 (자유공모=1 기준, 품목지정=2 비교)
    # 하향식(3)은 N=7로 제외 → 자유공모 vs 품목지정 2그룹
    sp_stage_v = case_when(
      support_type == 1 ~ "자유공모형",
      support_type == 2 ~ "품목지정형",
      TRUE              ~ NA_character_
    ),
    sp_품목더미 = ifelse(!is.na(support_type) & support_type == 2, 1L, 0L),
    
    # ── ③ research_dur 그룹화 (단기≤2 / 중기=3 / 장기≥4)
    dur_grp = case_when(
      is.na(research_dur)  ~ NA_character_,
      research_dur <= 2    ~ "단기(1-2년)",
      research_dur <= 3.5  ~ "중기(3년)",
      TRUE                 ~ "장기(4년+)"
    ),
    dur_grp_f = factor(dur_grp,
                       levels=c("단기(1-2년)","중기(3년)","장기(4년+)")),
    
    # ── ④ priv_match_rate 그룹화 (3분위)
    pmt_grp = case_when(
      is.na(priv_match_rate)       ~ NA_character_,
      priv_match_rate < 0.18       ~ "저매칭(<18%)",
      priv_match_rate < 0.24       ~ "중매칭(18-24%)",
      TRUE                         ~ "고매칭(24%+)"
    ),
    pmt_grp_f = factor(pmt_grp,
                       levels=c("저매칭(<18%)","중매칭(18-24%)","고매칭(24%+)")),
    
    # 통제군 연속형 변수 → 0 처리 (Interaction DID 구조 유지)
    research_dur_fill    = ifelse(is.na(research_dur),    0, research_dur),
    priv_match_rate_fill = ifelse(is.na(priv_match_rate), 0, priv_match_rate)
  )

cat(sprintf("  전체 표본: %d개 (처치군 %d / 통제군 %d)\n",
            nrow(df_all), sum(df_all$treat==1), sum(df_all$treat==0)))

# 처치군 변수 분포 확인
cat("\n=== 처치군 변수 분포 ===\n")
t_df <- df_all %>% filter(treat==1)
cat("[rd_stage] "); print(table(t_df$rd_stage_v, useNA="ifany"))
cat("[support_type] "); print(table(t_df$sp_stage_v, useNA="ifany"))
cat("[research_dur 그룹] "); print(table(t_df$dur_grp_f, useNA="ifany"))
cat("[priv_match_rate 그룹] "); print(table(t_df$pmt_grp_f, useNA="ifany"))

# ==============================================================================
# 3. 결과변수·공변량 공통 정의 (TLC 코드와 동일)
# ==============================================================================
make_diff_Y <- function(df, post_yr) {
  pre <- "2019"; yr <- as.character(post_yr)
  df %>% mutate(
    Y_asset = log(pmax(.data[[paste0(yr,"/Annual S15000.자산총계")]],1)) -
      log(pmax(.data[[paste0(pre,"/Annual S15000.자산총계")]],1)),
    Y_rev   = log(pmax(.data[[paste0(yr,"/Annual S21100.총수익")]],1)) -
      log(pmax(.data[[paste0(pre,"/Annual S21100.총수익")]],1)),
    Y_debt  = log(pmax(.data[[paste0(yr,"/Annual S18000.부채총계")]],1)) -
      log(pmax(.data[[paste0(pre,"/Annual S18000.부채총계")]],1)),
    Y_rd    = log(pmax(.data[[paste0("rdcost",yr)]],0)+1) -
      log(pmax(.data[[paste0("rdcost",pre)]],0)+1),
    Y_lb    = log(pmax(.data[[paste0("lbcost",yr)]],0)+1) -
      log(pmax(.data[[paste0("lbcost",pre)]],0)+1),
    Y_opm   = .data[[paste0("opm",yr)]] - .data[[paste0("opm",pre)]],
    Y_exp   = log(pmax(.data[[paste0("exportamt",yr)]],0)+1) -
      log(pmax(.data[[paste0("exportamt",pre)]],0)+1)
  )
}

var_info <- list(
  list(col="Y_asset", kr="ln(자산)"),
  list(col="Y_rev",   kr="ln(매출)"),
  list(col="Y_debt",  kr="ln(부채)"),
  list(col="Y_rd",    kr="ln(개발비+1)"),
  list(col="Y_lb",    kr="ln(인건비+1)"),
  list(col="Y_opm",   kr="OPM"),
  list(col="Y_exp",   kr="ln(수출+1)")
)

cov_fmla <- paste(
  "log_자산 + log_매출 + log_부채 + log_업력 +",
  "opm + log_연구개발비 + log_수출 + log_인건비 +",
  "ln_vol + factor(KSIC_mid_code) + factor(region)"
)

sig_star <- function(p) {
  dplyr::case_when(
    p < 0.001 ~ "***", p < 0.01 ~ "**",
    p < 0.05  ~ "*",   p < 0.10 ~ ".",
    TRUE ~ "ns"
  )
}

# coeftest 안전 추출 함수 (BUG-1 방지: 위치 인덱스 사용)
safe_coef <- function(mod, pattern) {
  tryCatch({
    ct     <- coeftest(mod, vcov = vcovHC(mod, type="HC1"))
    ct_sub <- ct[grep(pattern, rownames(ct)), , drop=FALSE]
    if (nrow(ct_sub) == 0) return(NULL)
    data.frame(
      term  = rownames(ct_sub),
      coef  = as.numeric(ct_sub[,1]),
      se    = as.numeric(ct_sub[,2]),
      t_val = as.numeric(ct_sub[,3]),
      p_val = as.numeric(ct_sub[,4]),
      stringsAsFactors = FALSE
    )
  }, error = function(e) NULL)
}

# ==============================================================================
# STEP 1. 기술통계 — 변수별 처치군 사전 공변량 균형
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 1. 기술통계\n")
cat(strrep("=",60), "\n")

cov_stat_cols <- c("log_자산","log_매출","log_부채","log_업력",
                   "opm","log_연구개발비","log_수출","log_인건비",
                   "age","total_vol","n_funded")

res_stat <- list()

# 분석 변수별 처치군 서브셋 기술통계
ntis_vars <- list(
  list(grp_col="rd_stage_v",  grp_nm="rd_stage",
       valid=c("응용연구","개발연구")),
  list(grp_col="sp_stage_v",  grp_nm="support_type",
       valid=c("자유공모형","품목지정형")),
  list(grp_col="dur_grp_f",   grp_nm="research_dur",
       valid=c("단기(1-2년)","중기(3년)","장기(4년+)")),
  list(grp_col="pmt_grp_f",   grp_nm="priv_match_rate",
       valid=c("저매칭(<18%)","중매칭(18-24%)","고매칭(24%+)"))
)

for (nv in ntis_vars) {
  sub_t <- df_all %>%
    filter(treat==1, !is.na(.data[[nv$grp_col]]),
           .data[[nv$grp_col]] %in% nv$valid)
  for (cov in cov_stat_cols) {
    if (!cov %in% names(sub_t)) next
    by_grp <- sub_t %>%
      group_by(.data[[nv$grp_col]]) %>%
      summarise(mean=mean(.data[[cov]],na.rm=T),
                sd  =sd(.data[[cov]],  na.rm=T),
                n   =n(), .groups="drop")
    frm <- as.formula(paste(cov, "~", nv$grp_col))
    f_p <- tryCatch(
      summary(aov(frm, data=sub_t))[[1]]$`Pr(>F)`[1],
      error=function(e) NA_real_
    )
    for (i in seq_len(nrow(by_grp))) {
      res_stat[[length(res_stat)+1]] <- data.frame(
        ntis_var  = nv$grp_nm,
        grp_label = as.character(by_grp[[nv$grp_col]][i]),
        covariate = cov,
        n         = by_grp$n[i],
        mean      = round(by_grp$mean[i], 4),
        sd        = round(by_grp$sd[i],   4),
        F_p       = round(f_p, 4),
        F_sig     = sig_star(f_p)
      )
    }
  }
}
df_stat <- bind_rows(res_stat)
cat("  -> 완료\n")

# ==============================================================================
# STEP 2. OLS Interaction DID
#   ① rd_stage:     TR × 응용연구더미  (개발연구 기준)
#   ② support_type: TR × 품목지정형더미 (자유공모 기준)
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 2. OLS Interaction DID (rd_stage, support_type)\n")
cat(strrep("=",60), "\n")

res_ols_bin <- list()

# 이진 비교 변수 정의
bin_vars <- list(
  list(
    nm       = "rd_stage",
    kr       = "연구개발단계",
    dummy    = "rd_응용더미",
    ref_lbl  = "개발연구(기준)",
    cmp_lbl  = "응용연구",
    pattern  = "^rd_응용더미",
    filter_f = quote(!is.na(rd_stage_v) & rd_stage_v %in% c("응용연구","개발연구"))
  ),
  list(
    nm       = "support_type",
    kr       = "지원유형",
    dummy    = "sp_품목더미",
    ref_lbl  = "자유공모형(기준)",
    cmp_lbl  = "품목지정형",
    pattern  = "^sp_품목더미",
    filter_f = quote(!is.na(sp_stage_v) & sp_stage_v %in% c("자유공모형","품목지정형"))
  )
)

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_all, yr)
  
  for (bv in bin_vars) {
    for (vi in var_info) {
      df_sub <- df_yr %>%
        filter(!is.na(.data[[vi$col]])) %>%
        filter(treat == 0 | eval(bv$filter_f))
      
      for (seg_nm in c("전체","소재","부품","장비")) {
        sub <- if (seg_nm=="전체") df_sub else
          df_sub %>% filter(seg_nm == !!seg_nm)
        if (nrow(sub) < 30) next
        
        # Interaction DID:
        # diff_Y ~ TR + TR×더미 + 더미 + 공변량
        # TR×더미 계수 = 처치 + (품목지정/응용) 추가 효과
        fmla_str <- paste(
          vi$col, "~ treat +",
          sprintf("treat:%s + %s +", bv$dummy, bv$dummy),
          if (seg_nm=="전체") "factor(seg_f) +" else "",
          cov_fmla
        )
        fmla <- as.formula(fmla_str)
        mod  <- tryCatch(lm(fmla, data=sub), error=function(e) NULL)
        if (is.null(mod)) next
        
        # 1. TR 주효과 (기준 그룹의 처치 효과)
        cr_tr <- safe_coef(mod, "^treat$")
        # 2. TR×더미 (비교 그룹의 추가 효과)
        cr_ia <- safe_coef(mod, bv$pattern)
        
        for (cr in list(cr_tr, cr_ia)) {
          if (is.null(cr)) next
          for (i in seq_len(nrow(cr))) {
            res_ols_bin[[length(res_ols_bin)+1]] <- data.frame(
              ntis_var = bv$nm,
              ntis_kr  = bv$kr,
              seg      = seg_nm,
              var      = vi$kr,
              year     = yr,
              N        = nrow(sub),
              term     = cr$term[i],
              coef     = round(cr$coef[i],  4),
              se       = round(cr$se[i],    4),
              t_val    = round(cr$t_val[i], 3),
              p_val    = round(cr$p_val[i], 4),
              sig      = sig_star(cr$p_val[i]),
              ci_lo    = round(cr$coef[i] - 1.96*cr$se[i], 4),
              ci_hi    = round(cr$coef[i] + 1.96*cr$se[i], 4)
            )
          }
        }
      }
    }
  }
}
df_ols_bin <- bind_rows(res_ols_bin)
cat("  -> 완료\n")

# ==============================================================================
# STEP 3. ANOVA + Tukey — rd_stage, support_type, research_dur 그룹 비교
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 3. ANOVA + Tukey\n")
cat(strrep("=",60), "\n")

res_anova <- list()
res_tukey <- list()

anova_vars <- list(
  list(nm="rd_stage",       grp="rd_stage_v",
       filter_f=quote(!is.na(rd_stage_v) & rd_stage_v %in% c("응용연구","개발연구"))),
  list(nm="support_type",   grp="sp_stage_v",
       filter_f=quote(!is.na(sp_stage_v) & sp_stage_v %in% c("자유공모형","품목지정형"))),
  list(nm="research_dur",   grp="dur_grp_f",
       filter_f=quote(!is.na(dur_grp_f)))
)

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_all %>% filter(treat==1), yr)
  
  for (av in anova_vars) {
    df_sub <- df_yr %>% filter(eval(av$filter_f))
    
    for (vi in var_info) {
      sub_all <- df_sub %>% filter(!is.na(.data[[vi$col]]))
      
      for (seg_nm in c("전체","소재","부품","장비")) {
        sub <- if (seg_nm=="전체") sub_all else
          sub_all %>% filter(seg_nm == !!seg_nm)
        if (nrow(sub) < 20) next
        
        fmla    <- as.formula(paste(vi$col, "~", av$grp))
        aov_mod <- tryCatch(aov(fmla, data=sub), error=function(e) NULL)
        if (is.null(aov_mod)) next
        
        aov_sum <- summary(aov_mod)[[1]]
        grp_means <- sub %>%
          group_by(.data[[av$grp]]) %>%
          summarise(mean_Y = round(mean(.data[[vi$col]], na.rm=T), 4),
                    n=n(), .groups="drop")
        
        res_anova[[length(res_anova)+1]] <- data.frame(
          ntis_var = av$nm,
          seg      = seg_nm,
          var      = vi$kr,
          year     = yr,
          N        = nrow(sub),
          F_stat   = round(aov_sum$`F value`[1], 3),
          F_p      = round(aov_sum$`Pr(>F)`[1],  4),
          F_sig    = sig_star(aov_sum$`Pr(>F)`[1]),
          그룹분포  = paste(
            paste(grp_means[[av$grp]], round(grp_means$mean_Y,3),
                  sep="="), collapse=" / ")
        )
        
        tukey <- tryCatch({
          tk <- TukeyHSD(aov_mod, av$grp)[[av$grp]]
          as.data.frame(tk) %>%
            mutate(comparison=rownames(tk),
                   ntis_var=av$nm, seg=seg_nm,
                   var=vi$kr, year=yr,
                   sig=sig_star(`p adj`))
        }, error=function(e) NULL)
        if (!is.null(tukey)) res_tukey[[length(res_tukey)+1]] <- tukey
      }
    }
  }
}

df_anova <- bind_rows(res_anova)
df_tukey <- if (length(res_tukey)>0) {
  bind_rows(res_tukey) %>%
    rename(diff_mean=diff, ci_lo=lwr, ci_hi=upr, p_adj=`p adj`) %>%
    select(ntis_var, seg, var, year, comparison,
           diff_mean, ci_lo, ci_hi, p_adj, sig)
} else data.frame()
cat("  -> 완료\n")

# ==============================================================================
# STEP 4. Dose-Response — research_dur, priv_match_rate
#   모형: diff_Y ~ TR + TR×X + TR×X² + X + X² + 공변량
#   처치군: 실제값 / 통제군: 0 (fill 변수 사용)
#   β₃(TR×X²) < 0 → 역U형 (최적 존재)
#   β₃ > 0         → U형
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 4. Dose-Response (research_dur, priv_match_rate)\n")
cat(strrep("=",60), "\n")

res_dr <- list()

dr_vars <- list(
  list(nm="research_dur",    fill="research_dur_fill",
       kr="연구기간(년)",    scale=1),
  list(nm="priv_match_rate", fill="priv_match_rate_fill",
       kr="민간매칭률",      scale=100)
)

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_all, yr)
  
  for (dv in dr_vars) {
    df_sub <- df_yr %>%
      mutate(
        X     = .data[[dv$fill]],
        X2    = X^2,
        TR_X  = treat * X,
        TR_X2 = treat * X^2
      )
    
    for (vi in var_info) {
      sub_all <- df_sub %>% filter(!is.na(.data[[vi$col]]))
      
      for (seg_nm in c("전체","소재","부품","장비")) {
        sub <- if (seg_nm=="전체") sub_all else
          sub_all %>% filter(seg_nm == !!seg_nm)
        if (nrow(sub) < 30) next
        
        seg_term <- if (seg_nm=="전체") "factor(seg_f) +" else ""
        fmla <- as.formula(paste(
          vi$col, "~ treat + TR_X + TR_X2 + X + X2 +",
          seg_term, cov_fmla
        ))
        
        mod <- tryCatch(lm(fmla, data=sub), error=function(e) NULL)
        if (is.null(mod)) next
        
        cr <- safe_coef(mod, "^treat$|^TR_X$|^TR_X2$")
        if (is.null(cr)) next
        
        b1 <- cr$coef[cr$term=="treat"]
        b2 <- cr$coef[cr$term=="TR_X"]
        b3 <- cr$coef[cr$term=="TR_X2"]
        
        # 최적값 계산 (역U형: β₃<0)
        opt_raw <- if (!is.na(b3) && b3 < 0 && !is.na(b2))
          -b2 / (2*b3) else NA_real_
        opt_disp <- opt_raw * dv$scale  # 단위 변환 (매칭률 → %)
        
        # F검정 (TR_X, TR_X2 동시 유의성)
        f_test <- tryCatch({
          linearHypothesis(mod, c("TR_X=0","TR_X2=0"),
                           vcov.=vcovHC(mod, type="HC1"))$`Pr(>F)`[2]
        }, error=function(e) NA_real_)
        
        shape_nm <- if (is.na(b3)) "계산불가"
        else if (b3 < 0) "역U형(최적존재)"
        else "U형(최저점)"
        
        res_dr[[length(res_dr)+1]] <- data.frame(
          ntis_var  = dv$nm,
          ntis_kr   = dv$kr,
          seg       = seg_nm,
          var       = vi$kr,
          year      = yr,
          N         = nrow(sub),
          b_TR      = round(b1, 5),
          b_TR_X    = round(b2, 5),
          b_TR_X2   = round(b3, 5),
          F_p       = round(f_test, 4),
          F_sig     = sig_star(f_test),
          shape     = shape_nm,
          opt_value = round(opt_disp, 3),
          opt_unit  = ifelse(dv$nm=="priv_match_rate","(%)", "(년)")
        )
      }
    }
  }
}
df_dr <- bind_rows(res_dr)
cat("  -> 완료\n")

# ==============================================================================
# STEP 5. Random Forest — 4개 NTIS 변수 중요도
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 5. Random Forest 변수 중요도\n")
cat(strrep("=",60), "\n")

rf_covs <- c("rd_stage","support_type","research_dur_fill","priv_match_rate_fill",
             "log_자산","log_매출","log_부채","log_업력",
             "opm","log_연구개발비","log_수출","log_인건비",
             "age","ln_vol","n_funded","seg")

res_vi <- list()

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_all %>% filter(treat==1), yr)
  
  for (vi in var_info) {
    df_sub <- df_yr %>%
      filter(!is.na(.data[[vi$col]])) %>%
      select(all_of(c(vi$col, rf_covs, "seg_nm"))) %>%
      mutate(across(where(is.character), ~as.numeric(factor(.))))
    
    for (seg_nm in c("전체","소재","부품","장비")) {
      sub <- if (seg_nm=="전체") df_sub else
        df_sub %>% filter(seg_nm == !!seg_nm)
      if (nrow(sub) < 30) next
      
      sub_rf <- sub %>% select(-seg_nm, -any_of("seg_f"))
      avail_covs <- intersect(rf_covs, names(sub_rf))
      if (length(avail_covs) < 3) next
      
      fmla_rf <- as.formula(paste(vi$col, "~",
                                  paste(avail_covs, collapse=" + ")))
      rf_mod <- tryCatch(
        ranger(fmla_rf, data=sub_rf, num.trees=1000,
               importance="impurity", num.threads=1, seed=2025),
        error=function(e) NULL
      )
      if (is.null(rf_mod)) next
      
      vi_df <- data.frame(
        variable   = names(rf_mod$variable.importance),
        importance = rf_mod$variable.importance,
        r2_oob     = rf_mod$r.squared
      ) %>% arrange(desc(importance))
      vi_df$seg  <- seg_nm
      vi_df$var  <- vi$kr
      vi_df$year <- yr
      res_vi[[length(res_vi)+1]] <- vi_df
    }
  }
}
df_vi <- bind_rows(res_vi)
cat("  -> 완료\n")

# ==============================================================================
# STEP 6. 시각화
# (BUG 예방: 모든 폰트 속성 제거 — face=, fontface=, base_family=)
# ==============================================================================
cat("\n-- 시각화 생성 중...\n")

# ── 6-1. rd_stage × support_type Coefficient Plot ─────────────────
grp_cols <- list(
  rd_stage     = c("응용연구"="#D85A30", "개발연구"="#0F6E56"),
  support_type = c("품목지정형"="#1A3A5C", "자유공모형"="#D98E2B")
)

pdf("NTIS변수_CoefPlot.pdf", width=13, height=8)
for (nv_nm in c("rd_stage","support_type")) {
  nv_kr <- if (nv_nm=="rd_stage") "연구개발단계" else "지원유형"
  ref_lbl <- if (nv_nm=="rd_stage") "개발연구(기준)" else "자유공모형(기준)"
  ia_pat  <- if (nv_nm=="rd_stage") "rd_응용더미" else "sp_품목더미"
  
  for (yr in c(2024, 2025)) {
    for (seg_nm in c("전체","소재","부품","장비")) {
      sub_ols <- df_ols_bin %>%
        filter(ntis_var==nv_nm, seg==seg_nm, year==yr)
      if (nrow(sub_ols)==0) next
      
      # TR 주효과 + TR×더미 모두 표시
      sub_plot <- sub_ols %>%
        mutate(
          효과구분 = case_when(
            term == "treat" ~ ref_lbl,
            grepl(ia_pat, term) ~
              if (nv_nm=="rd_stage") "응용연구 추가효과" else "품목지정 추가효과",
            TRUE ~ term
          ),
          var = factor(var, levels=c("ln(자산)","ln(매출)","ln(부채)",
                                     "ln(개발비+1)","ln(인건비+1)",
                                     "OPM","ln(수출+1)"))
        )
      
      p <- ggplot(sub_plot,
                  aes(x=var, y=coef, color=효과구분, group=효과구분)) +
        geom_hline(yintercept=0, linetype="dashed", color="gray60") +
        geom_pointrange(aes(ymin=ci_lo, ymax=ci_hi),
                        position=position_dodge(width=0.45),
                        size=0.7, linewidth=0.9) +
        geom_text(aes(label=sig, y=ci_hi+0.02),
                  position=position_dodge(width=0.45), size=3.5) +
        scale_color_manual(
          values = c("#0F6E56","#D98E2B","#1A3A5C","#D85A30"),
          name   = nv_kr) +
        labs(
          title    = sprintf("[%s] %s 처치 효과 (Post=%d)", seg_nm, nv_kr, yr),
          subtitle = paste("Interaction DID: diff_Y ~ TR + TR×더미 + 더미 + 공변량",
                           "HC1 강건 SE, 95% CI", sep="\n"),
          x=NULL, y="계수 (처치 효과)",
          caption = "*** p<0.001  ** p<0.01  * p<0.05  . p<0.10"
        ) +
        theme_minimal() +
        theme(legend.position="top", axis.text.x=element_text(size=10),
              plot.title=element_text(size=13),
              plot.subtitle=element_text(size=9, color="gray40"),
              panel.grid.minor=element_blank())
      print(p)
    }
  }
}
dev.off()
cat("  -> NTIS변수_CoefPlot.pdf 저장\n")

# ── 6-2. research_dur Box Plot (처치군, 그룹별 성과) ─────────────
dur_cols <- c("단기(1-2년)"="#D85A30","중기(3년)"="#1A3A5C","장기(4년+)"="#0F6E56")

pdf("NTIS변수_ResearchDur_BoxPlot.pdf", width=14, height=9)
for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_all %>% filter(treat==1, !is.na(dur_grp_f)), yr)
  for (seg_nm in c("부품","소재","장비")) {
    sub <- df_yr %>% filter(seg_nm == !!seg_nm)
    if (nrow(sub) < 10) next
    
    plot_list <- lapply(var_info, function(vi) {
      sub_v <- sub %>% filter(!is.na(.data[[vi$col]]))
      f_p_val <- df_anova %>%
        filter(ntis_var=="research_dur", seg==seg_nm,
               var==vi$kr, year==yr) %>% pull(F_p)
      subtitle_str <- if (length(f_p_val)>0)
        sprintf("ANOVA p=%.4f %s", f_p_val[1], sig_star(f_p_val[1])) else ""
      
      ggplot(sub_v, aes(x=dur_grp_f, y=.data[[vi$col]], fill=dur_grp_f)) +
        geom_boxplot(alpha=0.75, outlier.size=1, outlier.alpha=0.4, width=0.55) +
        geom_hline(yintercept=0, linetype="dashed", color="red", linewidth=0.6) +
        stat_summary(fun=mean, geom="point", shape=23,
                     size=2.5, fill="white", color="black") +
        scale_fill_manual(values=dur_cols) +
        labs(title=vi$kr, subtitle=subtitle_str, x=NULL, y="diff_Y") +
        theme_minimal() +
        theme(legend.position="none",
              plot.title=element_text(size=11),
              plot.subtitle=element_text(size=9, color="gray40"),
              axis.text.x=element_text(size=8))
    })
    
    gridExtra::grid.arrange(
      grobs=plot_list, ncol=4,
      top=grid::textGrob(
        sprintf("[%s] 연구기간별 성과 (처치군, Post=%d) N=%d",
                seg_nm, yr, nrow(sub)),
        gp=grid::gpar(fontsize=12))
    )
  }
}
dev.off()
cat("  -> NTIS변수_ResearchDur_BoxPlot.pdf 저장\n")

# ── 6-3. Dose-Response 요약 Heatmap (F_sig 기준) ─────────────────
pdf("NTIS변수_DoseResponse_요약.pdf", width=12, height=7)
for (dv_nm in c("research_dur","priv_match_rate")) {
  dv_kr <- if (dv_nm=="research_dur") "연구기간" else "민간매칭률"
  
  sub_dr <- df_dr %>%
    filter(ntis_var==dv_nm, seg != "전체") %>%
    mutate(
      sig_order = case_when(
        F_sig=="***" ~ 4, F_sig=="**" ~ 3,
        F_sig=="*"  ~ 2, F_sig=="."  ~ 1, TRUE ~ 0
      ),
      var = factor(var, levels=c("ln(자산)","ln(매출)","ln(부채)",
                                 "ln(개발비+1)","ln(인건비+1)",
                                 "OPM","ln(수출+1)"))
    )
  if (nrow(sub_dr)==0) next
  
  p <- sub_dr %>%
    mutate(label = paste0(shape, "\n",
                          ifelse(!is.na(opt_value),
                                 paste0("최적=",round(opt_value,2),opt_unit),
                                 "")) ) %>%
    ggplot(aes(x=var, y=interaction(seg,year), fill=sig_order)) +
    geom_tile(color="white", linewidth=0.5) +
    geom_text(aes(label=F_sig), size=3.5) +
    scale_fill_gradient2(
      low="#FFFFFF", mid="#B5D4F4", high="#185FA5",
      midpoint=2, limits=c(0,4),
      breaks=c(0,1,2,3,4),
      labels=c("ns",".","*","**","***"),
      name="유의성"
    ) +
    labs(
      title    = sprintf("[%s] Dose-Response F검정 유의성 히트맵", dv_kr),
      subtitle = "역U형/U형 유의 여부 (F검정 p < 0.10 = 곡선 관계 존재)",
      x=NULL, y="부문×연도"
    ) +
    theme_minimal() +
    theme(axis.text.x=element_text(size=9, angle=30, hjust=1),
          axis.text.y=element_text(size=9),
          plot.title=element_text(size=13),
          plot.subtitle=element_text(size=9, color="gray40"))
  print(p)
}
dev.off()
cat("  -> NTIS변수_DoseResponse_요약.pdf 저장\n")

# ── 6-4. RF 변수 중요도 — NTIS 4개 변수 순위 강조 ────────────────
vi_label <- c(
  "rd_stage"             = "연구개발단계",
  "support_type"         = "지원유형",
  "research_dur_fill"    = "연구기간",
  "priv_match_rate_fill" = "민간매칭률",
  "log_자산"             = "log(자산)",
  "log_매출"             = "log(매출)",
  "log_부채"             = "log(부채)",
  "log_업력"             = "log(업력)",
  "opm"                  = "OPM",
  "log_연구개발비"        = "log(R&D)",
  "log_수출"             = "log(수출)",
  "log_인건비"           = "log(인건비)",
  "age"                  = "업력",
  "ln_vol"               = "ln(투자금액)",
  "n_funded"             = "투자횟수",
  "seg"                  = "부문"
)
ntis_vars_nm <- c("rd_stage","support_type",
                  "research_dur_fill","priv_match_rate_fill")

pdf("NTIS변수_RF중요도.pdf", width=10, height=6)
for (yr in c(2024, 2025)) {
  for (seg_nm in c("전체","부품","소재","장비")) {
    sub_vi <- df_vi %>%
      filter(seg==seg_nm, year==yr) %>%
      group_by(variable) %>%
      summarise(importance=mean(importance), .groups="drop") %>%
      mutate(
        var_lbl  = ifelse(variable %in% names(vi_label),
                          vi_label[variable], variable),
        is_ntis  = variable %in% ntis_vars_nm
      ) %>%
      arrange(desc(importance))
    
    if (nrow(sub_vi)==0) next
    
    p <- sub_vi %>%
      ggplot(aes(x=reorder(var_lbl, importance),
                 y=importance, fill=is_ntis)) +
      geom_col(width=0.7, alpha=0.85) +
      scale_fill_manual(
        values=c("TRUE"="#D98E2B","FALSE"="#5B7494"),
        labels=c("기타 공변량","NTIS 신규 변수"),
        name=""
      ) +
      coord_flip() +
      labs(
        title    = sprintf("[%s] RF 변수 중요도 — Post %d", seg_nm, yr),
        subtitle = "NTIS 신규 4개 변수(주황) 순위 확인",
        x=NULL, y="Impurity 중요도"
      ) +
      theme_minimal() +
      theme(legend.position="top",
            axis.text.y=element_text(size=9),
            plot.title=element_text(size=12),
            plot.subtitle=element_text(size=9, color="gray40"))
    print(p)
  }
}
dev.off()
cat("  -> NTIS변수_RF중요도.pdf 저장\n")

# ==============================================================================
# STEP 7. Excel 저장
# ==============================================================================
cat("\n-- Excel 저장 중...\n")

# Dose-Response Wide (F_sig만 추출)
df_dr_wide <- df_dr %>%
  mutate(결과 = paste0(F_sig, " / ", shape,
                     ifelse(!is.na(opt_value),
                            paste0(" / 최적=",round(opt_value,2),opt_unit), ""))) %>%
  select(ntis_var, ntis_kr, seg, var, year, 결과) %>%
  pivot_wider(names_from=year, values_from=결과,
              names_prefix="Post") %>%
  rename(변수=ntis_var, 변수설명=ntis_kr, 부문=seg, 성과변수=var)

sheets <- list(
  
  메타정보 = data.frame(
    항목 = c("분석 변수","표본 구조","STEP2 모형","STEP4 모형",
           "기준연도","성과연도","OLS 표준오차","사후 비교","RF 나무 수"),
    내용 = c(
      "rd_stage(응용vs개발), support_type(자유공모vs품목지정), research_dur, priv_match_rate",
      "전체 1076 (처치538+통제538) / NTIS 변수는 처치군만 유효, 통제군=0",
      "diff_Y ~ TR + TR×더미 + 더미 + 공변량 (Interaction DID)",
      "diff_Y ~ TR + TR×X + TR×X² + X + X² + 공변량 (Dose-Response)",
      "2019년",
      "2024년, 2025년",
      "HC1 이분산 강건 표준오차",
      "Tukey HSD",
      "1,000")
  ),
  
  기술통계 = df_stat %>%
    rename(NTIS변수=ntis_var, 그룹=grp_label, 공변량=covariate,
           N=n, 평균=mean, 표준편차=sd, ANOVA_p=F_p, ANOVA_sig=F_sig),
  
  OLS_Interaction_상세 = df_ols_bin %>%
    rename(NTIS변수=ntis_var, NTIS설명=ntis_kr, 부문=seg,
           성과변수=var, 연도=year, 표본수=N,
           계수항=term, 계수=coef, 표준오차=se,
           t값=t_val, p값=p_val, 유의성=sig,
           CI하한=ci_lo, CI상한=ci_hi),
  
  ANOVA_그룹비교 = df_anova %>%
    rename(NTIS변수=ntis_var, 부문=seg, 성과변수=var,
           연도=year, 표본수=N, F통계량=F_stat,
           p값=F_p, 유의성=F_sig, 그룹평균=그룹분포),
  
  Tukey_사후비교 = df_tukey %>%
    rename(NTIS변수=ntis_var, 부문=seg, 성과변수=var,
           연도=year, 비교쌍=comparison, 평균차이=diff_mean,
           CI하한=ci_lo, CI상한=ci_hi, 조정p값=p_adj, 유의성=sig),
  
  DoseResponse_상세 = df_dr %>%
    rename(NTIS변수=ntis_var, 변수설명=ntis_kr, 부문=seg,
           성과변수=var, 연도=year, 표본수=N,
           b_TR=b_TR, b_TRxX=b_TR_X, b_TRxX2=b_TR_X2,
           F_p=F_p, 유의성=F_sig, 곡선형태=shape,
           최적값=opt_value, 단위=opt_unit),
  
  DoseResponse_Wide = df_dr_wide,
  
  RF_변수중요도 = df_vi %>%
    rename(부문=seg, 성과변수=var, 연도=year,
           공변량=variable, 중요도=importance, OOB_R2=r2_oob)
)

write_xlsx(sheets, "NTIS변수_이질성분석_결과.xlsx")
cat("  -> NTIS변수_이질성분석_결과.xlsx 저장\n")

# ==============================================================================
# 콘솔 최종 요약
# ==============================================================================
cat("\n", strrep("=",65), "\n")
cat("  NTIS 변수 이질성 분석 — 결과 요약\n")
cat(strrep("=",65), "\n")

cat("\n[1] OLS Interaction DID — 유의한 계수 (부품 부문)\n")
tmp1 <- df_ols_bin %>%
  filter(seg=="부품", sig %in% c("***","**","*",".")) %>%
  select(ntis_var, var, year, term, coef, se, p_val, sig) %>%
  as.data.frame()
if (nrow(tmp1) > 0) print(tmp1) else cat("  (유의 없음)\n")

cat("\n[2] ANOVA — 유의한 그룹 차이\n")
tmp2 <- df_anova %>%
  filter(F_sig %in% c("***","**","*",".")) %>%
  select(ntis_var, seg, var, year, F_stat, F_p, F_sig, 그룹분포) %>%
  as.data.frame()
if (nrow(tmp2) > 0) print(head(tmp2, 20)) else cat("  (유의 없음)\n")

cat("\n[3] Dose-Response — 유의한 곡선 관계\n")
tmp3 <- df_dr %>%
  filter(F_sig %in% c("***","**","*",".")) %>%
  select(ntis_var, seg, var, year, F_sig, shape, opt_value, opt_unit) %>%
  as.data.frame()
if (nrow(tmp3) > 0) print(head(tmp3, 20)) else cat("  (유의 없음)\n")

cat("\n[4] RF 중요도 — NTIS 변수 순위 (부품, 2024)\n")
tmp4 <- df_vi %>%
  filter(seg=="부품", year==2024) %>%
  group_by(variable) %>%
  summarise(imp=mean(importance), .groups="drop") %>%
  mutate(rank=rank(-imp), lbl=vi_label[variable]) %>%
  filter(variable %in% ntis_vars_nm) %>%
  arrange(rank) %>%
  as.data.frame()
if (nrow(tmp4) > 0) print(tmp4) else cat("  (데이터 없음)\n")

cat("\n저장 파일:\n")
cat("  NTIS변수_이질성분석_결과.xlsx   — 8개 시트\n")
cat("  NTIS변수_CoefPlot.pdf           — OLS Interaction DID 계수\n")
cat("  NTIS변수_ResearchDur_BoxPlot.pdf — 연구기간 그룹별 성과\n")
cat("  NTIS변수_DoseResponse_요약.pdf   — Dose-Response 히트맵\n")
cat("  NTIS변수_RF중요도.pdf            — RF 변수 중요도\n")
cat(strrep("=",65), "\n")


