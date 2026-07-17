################################################################################
# 소부장 R&D 지원 정책 — 기술수명주기(TLC) 이질성 분석
# 처치군(정부지원 수혜 기업) 한정
#
# 수정 이력:
#   v1: 최초 작성
#   v2: [BUG-1] coeftest 컬럼명 → 위치 인덱스([,1]~[,4])로 변경
#       [BUG-2] AppleGothic → get_korean_font() 분기 + gpar(font=2)
#   v3: [BUG-2 완전 해결]
#       원인: fontface="bold", base_family=, gpar(font=) 등
#             ggplot2·grid 폰트 속성이 운영체제에 따라 에러 유발
#       해결: 모든 폰트 관련 속성 완전 제거
#             (face=, fontface=, base_family=, font= 속성 일체 삭제)
################################################################################


# ==============================================================================
# 0. 패키지
# ==============================================================================
pkgs <- c("readxl","dplyr","tidyr","ggplot2","writexl",
          "sandwich","lmtest","car","emmeans",
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
  "funded","seg","treat","tlc","age",
  "subclass","weights","distance",
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
# 2. 처치군 x TLC 유효 표본
# ==============================================================================
df_base <- raw %>%
  filter(treat == 1, tlc %in% c(1,2,3)) %>%
  mutate(
    tlc_f  = factor(tlc, levels = c(1,2,3),
                    labels = c("도입기","성장기","성숙기")),
    tlc_nm = as.character(tlc_f),
    seg_f  = factor(seg, levels = c(1,2,3),
                    labels = c("소재","부품","장비")),
    seg_nm = as.character(seg_f),
    total_vol = gfundvol2020 + gfundvol2021 + gfundvol2022,
    n_funded  = fund2020 + fund2021 + fund2022,
    ln_vol    = log(pmax(total_vol, 0) + 1)
  )

cat(sprintf("  처치군 TLC 유효 표본: %d개\n", nrow(df_base)))
cat("  부문 x TLC 분포:\n")
print(table(df_base$seg_nm, df_base$tlc_nm))

# ==============================================================================
# 3. 결과변수 차분 생성 함수
# ==============================================================================
make_diff_Y <- function(df, post_yr) {
  pre <- "2019"
  yr  <- as.character(post_yr)
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
    p < 0.001 ~ "***",
    p < 0.01  ~ "**",
    p < 0.05  ~ "*",
    p < 0.10  ~ ".",
    TRUE      ~ "ns"
  )
}

# ==============================================================================
# STEP 1. 기술통계
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 1. TLC별 사전(2019) 공변량 기술통계\n")
cat(strrep("=",60), "\n")

cov_stat_cols <- c("log_자산","log_매출","log_부채","log_업력",
                   "opm","log_연구개발비","log_수출","log_인건비",
                   "age","total_vol","n_funded")

res_stat <- list()
for (seg_nm in c("소재","부품","장비","전체")) {
  sub <- if (seg_nm == "전체") df_base else
    df_base %>% filter(seg_nm == !!seg_nm)
  for (cov in cov_stat_cols) {
    if (!cov %in% names(sub)) next
    by_tlc <- sub %>%
      group_by(tlc_f) %>%
      summarise(mean = mean(.data[[cov]], na.rm=TRUE),
                sd   = sd(.data[[cov]],   na.rm=TRUE),
                n    = n(), .groups="drop")
    frm <- as.formula(paste(cov, "~ tlc_f"))
    f_p <- tryCatch(
      summary(aov(frm, data=sub))[[1]]$`Pr(>F)`[1],
      error = function(e) NA_real_
    )
    for (i in seq_len(nrow(by_tlc))) {
      res_stat[[length(res_stat)+1]] <- data.frame(
        seg       = seg_nm,
        covariate = cov,
        tlc_nm    = as.character(by_tlc$tlc_f[i]),
        n         = by_tlc$n[i],
        mean      = round(by_tlc$mean[i], 4),
        sd        = round(by_tlc$sd[i],   4),
        F_p       = round(f_p, 4),
        F_sig     = sig_star(f_p)
      )
    }
  }
}
df_stat <- bind_rows(res_stat)
cat("  -> 완료\n")

# ==============================================================================
# STEP 2. OLS 회귀
# [BUG-1 수정] coeftest 결과를 위치 인덱스([,1]~[,4])로 추출
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 2. OLS 회귀 (TLC별 성과 차이, 공변량 통제)\n")
cat(strrep("=",60), "\n")

res_ols <- list()

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_base, yr)
  
  for (vi in var_info) {
    df_sub <- df_yr %>% filter(!is.na(.data[[vi$col]]))
    
    for (seg_nm in c("전체","소재","부품","장비")) {
      sub <- if (seg_nm == "전체") df_sub else
        df_sub %>% filter(seg_nm == !!seg_nm)
      if (nrow(sub) < 30) next
      
      sub$tlc_f <- relevel(sub$tlc_f, ref = "도입기")
      
      fmla <- if (seg_nm == "전체") {
        as.formula(paste(vi$col, "~ tlc_f + factor(seg_f) +", cov_fmla))
      } else {
        as.formula(paste(vi$col, "~ tlc_f +", cov_fmla))
      }
      
      mod <- tryCatch(lm(fmla, data=sub), error=function(e) NULL)
      if (is.null(mod)) next
      
      # [BUG-1 수정] 위치 인덱스로 추출 — 버전 무관
      coef_robust <- tryCatch({
        ct     <- coeftest(mod, vcov = vcovHC(mod, type="HC1"))
        ct_sub <- ct[grep("^tlc_f", rownames(ct)), , drop=FALSE]
        if (nrow(ct_sub) == 0) return(NULL)
        data.frame(
          term  = rownames(ct_sub),
          coef  = as.numeric(ct_sub[, 1]),
          se    = as.numeric(ct_sub[, 2]),
          t_val = as.numeric(ct_sub[, 3]),
          p_val = as.numeric(ct_sub[, 4]),
          stringsAsFactors = FALSE
        )
      }, error = function(e) NULL)
      if (is.null(coef_robust)) next
      
      for (i in seq_len(nrow(coef_robust))) {
        res_ols[[length(res_ols)+1]] <- data.frame(
          seg   = seg_nm,
          var   = vi$kr,
          year  = yr,
          N     = nrow(sub),
          term  = coef_robust$term[i],
          coef  = round(coef_robust$coef[i],  4),
          se    = round(coef_robust$se[i],    4),
          t_val = round(coef_robust$t_val[i], 3),
          p_val = round(coef_robust$p_val[i], 4),
          sig   = sig_star(coef_robust$p_val[i]),
          ci_lo = round(coef_robust$coef[i] - 1.96*coef_robust$se[i], 4),
          ci_hi = round(coef_robust$coef[i] + 1.96*coef_robust$se[i], 4)
        )
      }
    }
  }
}

df_ols <- bind_rows(res_ols) %>%
  mutate(tlc_comp = gsub("tlc_f","",term))
cat("  -> 완료\n")

# ==============================================================================
# STEP 3. ANOVA + Tukey 사후 비교
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 3. ANOVA + Tukey 사후비교\n")
cat(strrep("=",60), "\n")

res_anova <- list()
res_tukey <- list()

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_base, yr)
  
  for (vi in var_info) {
    df_sub <- df_yr %>% filter(!is.na(.data[[vi$col]]))
    
    for (seg_nm in c("전체","소재","부품","장비")) {
      sub <- if (seg_nm == "전체") df_sub else
        df_sub %>% filter(seg_nm == !!seg_nm)
      if (nrow(sub) < 20) next
      
      fmla    <- as.formula(paste(vi$col, "~ tlc_f"))
      aov_mod <- tryCatch(aov(fmla, data=sub), error=function(e) NULL)
      if (is.null(aov_mod)) next
      
      aov_sum <- summary(aov_mod)[[1]]
      grp_means <- sub %>%
        group_by(tlc_f) %>%
        summarise(mean_Y = round(mean(.data[[vi$col]], na.rm=TRUE), 4),
                  n=n(), .groups="drop")
      
      get_mean <- function(grp) {
        v <- grp_means$mean_Y[grp_means$tlc_f == grp]
        if (length(v) == 0) NA_real_ else v
      }
      
      res_anova[[length(res_anova)+1]] <- data.frame(
        seg         = seg_nm,
        var         = vi$kr,
        year        = yr,
        N           = nrow(sub),
        F_stat      = round(aov_sum$`F value`[1], 3),
        F_p         = round(aov_sum$`Pr(>F)`[1],  4),
        F_sig       = sig_star(aov_sum$`Pr(>F)`[1]),
        도입기_mean = get_mean("도입기"),
        성장기_mean = get_mean("성장기"),
        성숙기_mean = get_mean("성숙기")
      )
      
      tukey <- tryCatch({
        tk <- TukeyHSD(aov_mod, "tlc_f")$tlc_f
        as.data.frame(tk) %>%
          mutate(comparison=rownames(tk),
                 seg=seg_nm, var=vi$kr, year=yr,
                 sig=sig_star(`p adj`))
      }, error=function(e) NULL)
      if (!is.null(tukey)) res_tukey[[length(res_tukey)+1]] <- tukey
    }
  }
}

df_anova <- bind_rows(res_anova)
df_tukey <- bind_rows(res_tukey) %>%
  rename(diff_mean=diff, ci_lo=lwr, ci_hi=upr, p_adj=`p adj`) %>%
  select(seg, var, year, comparison, diff_mean, ci_lo, ci_hi, p_adj, sig)
cat("  -> 완료\n")

# ==============================================================================
# STEP 4. Random Forest — 변수 중요도
# ==============================================================================
cat("\n", strrep("=",60), "\n")
cat("  STEP 4. Random Forest 변수 중요도\n")
cat(strrep("=",60), "\n")

rf_covs <- c("tlc","log_자산","log_매출","log_부채","log_업력",
             "opm","log_연구개발비","log_수출","log_인건비",
             "age","ln_vol","n_funded","seg")

res_vi <- list()

for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_base, yr)
  
  for (vi in var_info) {
    df_sub <- df_yr %>%
      filter(!is.na(.data[[vi$col]])) %>%
      select(all_of(c(vi$col, rf_covs, "seg_nm"))) %>%
      mutate(tlc = as.numeric(tlc))
    
    for (seg_nm in c("전체","소재","부품","장비")) {
      sub <- if (seg_nm == "전체") df_sub else
        df_sub %>% filter(seg_nm == !!seg_nm)
      if (nrow(sub) < 30) next
      
      sub_rf <- sub %>%
        select(-seg_nm, -any_of("seg_f")) %>%
        mutate(across(where(is.character), ~as.numeric(factor(.))))
      
      fmla_rf <- as.formula(paste(vi$col, "~",
                                  paste(rf_covs, collapse=" + ")))
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
# STEP 5. 시각화
# [BUG-2 완전 해결] 모든 폰트 속성 제거
#   제거 항목: base_family=, fontface=, face=, font=, gpar(font=)
#   이유: ggplot2/grid 폰트 속성은 OS·R버전·PDF디바이스 조합에 따라
#         "invalid font type" 에러 유발 → 속성 자체를 사용하지 않음
# ==============================================================================
cat("\n-- 시각화 생성 중...\n")

tlc_cols <- c("도입기"="#2C8C7C", "성장기"="#1A3A5C", "성숙기"="#D98E2B")

# ── 5-1. Box Plot ─────────────────────────────────────────────────
pdf("TLC_처치군_BoxPlot.pdf", width=14, height=9)
for (yr in c(2024, 2025)) {
  df_yr <- make_diff_Y(df_base, yr)
  for (seg_nm in c("부품","소재","장비")) {
    sub <- df_yr %>% filter(seg_nm == !!seg_nm)
    if (nrow(sub) < 10) next
    
    plot_list <- lapply(var_info, function(vi) {
      sub_v   <- sub %>% filter(!is.na(.data[[vi$col]]))
      f_p_val <- df_anova %>%
        filter(seg==seg_nm, var==vi$kr, year==yr) %>% pull(F_p)
      subtitle_str <- if (length(f_p_val) > 0)
        sprintf("ANOVA p=%.4f %s", f_p_val[1], sig_star(f_p_val[1])) else ""
      
      ggplot(sub_v, aes(x=tlc_f, y=.data[[vi$col]], fill=tlc_f)) +
        geom_boxplot(alpha=0.75, outlier.size=1,
                     outlier.alpha=0.4, width=0.55) +
        geom_hline(yintercept=0, linetype="dashed",
                   color="red", linewidth=0.6) +
        stat_summary(fun=mean, geom="point", shape=23,
                     size=2.5, fill="white", color="black") +
        scale_fill_manual(values=tlc_cols) +
        labs(title=vi$kr, subtitle=subtitle_str,
             x=NULL, y="diff_Y (post - pre)") +
        theme_minimal() +
        theme(legend.position = "none",
              plot.title      = element_text(size=11),
              plot.subtitle   = element_text(size=9, color="gray40"),
              axis.text.x     = element_text(size=9))
    })
    
    title_str <- sprintf(
      "[%s] TLC별 성과 차이 (처치군 한정, Post=%d) N=%d",
      seg_nm, yr, nrow(sub))
    gridExtra::grid.arrange(
      grobs = plot_list,
      ncol  = 4,
      top   = grid::textGrob(title_str,
                             gp = grid::gpar(fontsize=12))
    )
  }
}
dev.off()
cat("  -> TLC_처치군_BoxPlot.pdf 저장\n")

# ── 5-2. OLS 계수 Plot ────────────────────────────────────────────
pdf("TLC_처치군_CoefPlot.pdf", width=13, height=8)
for (yr in c(2024, 2025)) {
  for (seg_nm in c("부품","소재","장비","전체")) {
    sub_ols <- df_ols %>% filter(seg==seg_nm, year==yr)
    if (nrow(sub_ols)==0) next
    
    p <- sub_ols %>%
      mutate(
        var = factor(var, levels=c("ln(자산)","ln(매출)","ln(부채)",
                                   "ln(개발비+1)","ln(인건비+1)",
                                   "OPM","ln(수출+1)")),
        tlc_comp = factor(tlc_comp, levels=c("성장기","성숙기"))
      ) %>%
      ggplot(aes(x=var, y=coef, color=tlc_comp, group=tlc_comp)) +
      geom_hline(yintercept=0, linetype="dashed", color="gray60") +
      geom_pointrange(
        aes(ymin=ci_lo, ymax=ci_hi),
        position = position_dodge(width=0.45),
        size=0.7, linewidth=0.9
      ) +
      geom_text(
        aes(label=sig, y=ci_hi+0.02),
        position = position_dodge(width=0.45),
        size=3.5
      ) +
      scale_color_manual(
        values = c("성장기"="#1A3A5C","성숙기"="#D98E2B"),
        name   = "TLC (기준: 도입기)"
      ) +
      labs(
        title    = sprintf("[%s] TLC별 성과 차이 (도입기 대비, Post=%d)",
                           seg_nm, yr),
        subtitle = paste(
          "공변량 통제 후 OLS 계수, HC1 강건 SE, 95% CI",
          "도입기=기준(0), 양수=도입기보다 높음, 음수=낮음",
          sep="\n"),
        x=NULL, y="계수 (도입기 대비 차이)",
        caption = "*** p<0.001  ** p<0.01  * p<0.05  . p<0.10"
      ) +
      theme_minimal() +
      theme(
        legend.position  = "top",
        axis.text.x      = element_text(size=10),
        plot.title       = element_text(size=13),
        plot.subtitle    = element_text(size=9, color="gray40"),
        panel.grid.minor = element_blank()
      )
    print(p)
  }
}
dev.off()
cat("  -> TLC_처치군_CoefPlot.pdf 저장\n")

# ── 5-3. 변수 중요도 Bar Plot ─────────────────────────────────────
vi_label <- c(
  "tlc"            = "TLC(기술수명주기)",
  "log_자산"       = "log(자산)",
  "log_매출"       = "log(매출)",
  "log_부채"       = "log(부채)",
  "log_업력"       = "log(업력)",
  "opm"            = "OPM",
  "log_연구개발비" = "log(R&D)",
  "log_수출"       = "log(수출)",
  "log_인건비"     = "log(인건비)",
  "age"            = "업력",
  "ln_vol"         = "ln(투자금액)",
  "n_funded"       = "투자횟수",
  "seg"            = "부문"
)

pdf("TLC_처치군_VarImportance.pdf", width=9, height=5)
for (yr in c(2024, 2025)) {
  for (seg_nm in c("전체","부품","소재","장비")) {
    sub_vi <- df_vi %>%
      filter(seg==seg_nm, year==yr) %>%
      group_by(variable) %>%
      summarise(importance=mean(importance), .groups="drop") %>%
      mutate(
        var_lbl = ifelse(variable %in% names(vi_label),
                         vi_label[variable], variable),
        is_tlc  = (variable=="tlc")
      ) %>%
      arrange(desc(importance))
    
    if (nrow(sub_vi)==0) next
    tlc_rank <- which(sub_vi$variable=="tlc")
    if (length(tlc_rank)==0) tlc_rank <- NA
    
    p <- sub_vi %>%
      ggplot(aes(x=reorder(var_lbl, importance),
                 y=importance, fill=is_tlc)) +
      geom_col(width=0.7, alpha=0.85) +
      scale_fill_manual(values=c("TRUE"="#D98E2B","FALSE"="#5B7494")) +
      coord_flip() +
      labs(
        title    = sprintf("[%s] 변수 중요도 (RF) — Post %d",
                           seg_nm, yr),
        subtitle = sprintf("TLC 순위: %s위 / %d개 변수",
                           ifelse(is.na(tlc_rank),"?",tlc_rank),
                           nrow(sub_vi)),
        x=NULL, y="Impurity 기반 중요도"
      ) +
      theme_minimal() +
      theme(legend.position = "none",
            axis.text.y     = element_text(size=9),
            plot.title      = element_text(size=12),
            plot.subtitle   = element_text(size=9, color="gray40"))
    print(p)
  }
}
dev.off()
cat("  -> TLC_처치군_VarImportance.pdf 저장\n")

# ==============================================================================
# 6. Excel 저장
# ==============================================================================
cat("\n-- Excel 저장 중...\n")

df_ols_wide <- df_ols %>%
  mutate(coef_sig = paste0(coef, " ", sig)) %>%
  select(seg, var, year, tlc_comp, coef_sig) %>%
  pivot_wider(names_from=tlc_comp, values_from=coef_sig) %>%
  rename(부문=seg, 변수=var, 연도=year)

sheets <- list(
  TLC_기술통계 = df_stat %>%
    rename(부문=seg, 공변량=covariate, TLC명=tlc_nm,
           N=n, 평균=mean, 표준편차=sd,
           ANOVA_p=F_p, ANOVA_sig=F_sig),
  
  TLC_OLS_상세 = df_ols %>%
    rename(부문=seg, 변수=var, 연도=year, 표본수=N,
           TLC비교=tlc_comp, 계수=coef, 표준오차=se,
           t값=t_val, p값=p_val, 유의성=sig,
           CI하한=ci_lo, CI상한=ci_hi),
  
  TLC_OLS_Wide = df_ols_wide,
  
  TLC_ANOVA = df_anova %>%
    rename(부문=seg, 변수=var, 연도=year, 표본수=N,
           F통계량=F_stat, p값=F_p, 유의성=F_sig),
  
  TLC_Tukey = df_tukey %>%
    rename(부문=seg, 변수=var, 연도=year,
           비교쌍=comparison, 평균차이=diff_mean,
           CI하한=ci_lo, CI상한=ci_hi, 조정p값=p_adj, 유의성=sig),
  
  TLC_RF_중요도 = df_vi %>%
    rename(부문=seg, 변수=var, 연도=year,
           공변량=variable, 중요도=importance, OOB_R2=r2_oob),
  
  TLC_메타 = data.frame(
    항목 = c("분석 대상","표본 수","TLC 범주","기준범주(OLS)",
           "사전연도","사후연도","결과변수 구조",
           "OLS 표준오차","ANOVA 사후비교","RF 나무 수",
           "BUG-1 수정","BUG-2 수정"),
    내용 = c("처치군(treat=1) 중 tlc 유효 기업",
           "913개 (쇠퇴기 1건 제외)",
           "1=도입기, 2=성장기, 3=성숙기",
           "도입기 (성장기·성숙기가 도입기 대비 차이)",
           "2019년",
           "2024년, 2025년",
           "diff_Y = log(Y_post)-log(Y_pre) / OPM: 수준 차분",
           "HC1 이분산 강건 표준오차",
           "Tukey HSD",
           "1,000",
           "coeftest 컬럼명 -> 위치인덱스 ct_sub[,1]~[,4]",
           "폰트 속성(fontface/face/base_family) 완전 제거")
  )
)

write_xlsx(sheets, "TLC_처치군_결과.xlsx")
cat("  -> TLC_처치군_결과.xlsx 저장\n")

# ==============================================================================
# 7. 콘솔 최종 요약
# ==============================================================================
cat("\n", strrep("=",65), "\n")
cat("  TLC 처치군 분석 완료\n")
cat(strrep("=",65), "\n")

cat("\n[1] ANOVA — 유의한 TLC 성과 차이\n")
tmp1 <- df_anova %>%
  filter(F_sig %in% c("***","**","*",".")) %>%
  select(seg, var, year, F_stat, F_p, F_sig,
         도입기_mean, 성장기_mean, 성숙기_mean) %>%
  as.data.frame()
if (nrow(tmp1) > 0) print(tmp1) else cat("  (유의한 조합 없음)\n")

cat("\n[2] OLS — 부품 부문 유의한 TLC 계수\n")
tmp2 <- df_ols %>%
  filter(seg=="부품", sig %in% c("***","**","*",".")) %>%
  select(var, year, tlc_comp, coef, se, p_val, sig) %>%
  as.data.frame()
if (nrow(tmp2) > 0) print(tmp2) else cat("  (유의한 계수 없음)\n")

cat("\n[3] RF 변수 중요도 — tlc 순위 (부품)\n")
tmp3 <- df_vi %>%
  filter(seg=="부품") %>%
  group_by(year, variable) %>%
  summarise(imp=mean(importance), .groups="drop") %>%
  group_by(year) %>%
  mutate(rank=rank(-imp)) %>%
  filter(variable=="tlc") %>%
  as.data.frame()
if (nrow(tmp3) > 0) print(tmp3) else cat("  (tlc 데이터 없음)\n")

cat("\n저장 파일:\n")
cat("  TLC_처치군_결과.xlsx\n")
cat("  TLC_처치군_BoxPlot.pdf\n")
cat("  TLC_처치군_CoefPlot.pdf\n")
cat("  TLC_처치군_VarImportance.pdf\n")
cat(strrep("=",65), "\n")

cat("  실행 완료\n")
