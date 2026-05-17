################################################################################
# 소부장 R&D 투자 총금액별 DID 분석 (부문별 세분화)
# 
# 분석 내용:
#   1. 투자 총금액 3분위(하위/중위/상위) x 부문(소재/부품/장비) DID
#   2. 더미모형 (0회 기준, ANOVA 이질성 검정 + pairwise 비교)
#   3. 평행추세 검정 (pre-treatment trend)
#   4. 연속변수 dose-response 모형
#   5. Wide format 요약표 + Excel 저장
#
# 데이터: matched_dataset_segPSM.xlsx (PSM 매칭 완료 데이터)
# pre = 2019, post = 2024, 2025, treatment period = 2020~2022
#
# ★ 주의: Excel 파일의 수치 컬럼이 text로 저장되어 있으므로 로드 후 변환 필수
# ★ 주의: data.frame 내부 컬럼명은 영문 사용 (한글 인코딩 이슈 방지)
#         Excel 저장 시 rename으로 한글 복원

# 분석 설계:
#현재는 1단계 PSM(funded=1 vs funded=0, 부문별)으로 매칭한 후, 
#2단계에서 처치군을 사후적으로 세분화(횟수 1·2·3회 또는 금액 하위·중위·상위)하여 각각의 DID를 통제군과 비교하는 구조입니다. 
#이때 모든 세분화 그룹이 동일한 통제군을 공유합니다.

# 본 연구는 PSM으로 처치·통제군의 사전 특성 균형을 확보한 후, 
#처치군 내 투자 강도(횟수·금액)의 이질적 효과를 사후 세분화(subgroup analysis)로 분석하는 설계를 채택했다. 
#투자 횟수와 금액은 처치 배정 이후 실현되는 사후 변수(post-treatment variable)이므로 
#이를 PSM 매칭 조건에 포함하면 처치 후 편향이 발생할 수 있다(Rosenbaum, 2002). 
#동일 통제군 대비 하위그룹별 DID를 추정하고, ANOVA 이질성 검정으로 
#강도별 효과 차이의 통계적 유의성을 확인하는 접근은 선행연구
#(예: Czarnitzki & Lopes-Bento, 2013; Hottenrott & Lopes-Bento, 2014)에서도 채택된 표준적 방법이다.
# 20260503
# 결과 변수에서 자본금 삭제. OPM 로그변환 하지 않음(opm). 노무비 삭제하고 인건비로 통합(lbcost) 
# Doubly Robust DID 추가 (자기자신 사전값 제외, 나머지 7개 통제)
# 2025년도 데이터 추가
# 20250517 : DOSE_RESPONSE - ln(금액)² 항 추가하여 역U형 관계 탐색 가능하도록 개선. 최적 투자금액 계산 추가.
################################################################################

packages <- c("readxl", "dplyr", "tidyr", "ggplot2",
              "sandwich", "lmtest", "writexl",
              "gridExtra", "grid", "scales", "car")

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}

select <- dplyr::select
lag    <- dplyr::lag

setwd("/Users/ultra/PSM-DID")

# ==============================================================================
# -- 데이터 로드 ---------------------------------------------------------------
# ==============================================================================
raw <- read_excel("matched_dataset_segPSM.xlsx", col_types = "text")

# ==============================================================================
# -- 1. 수치 컬럼 일괄 변환 (Excel text -> numeric) ---------------------------
# ==============================================================================

numeric_cols <- c(
  "fund2020", "fund2021", "fund2022",
  "gfundvol2020", "gfundvol2021", "gfundvol2022",
  "fundedpattern", "funded", "seg", "treat",
  grep("S15000|S18000|S18100|S21100|S25000|S05000|S21190|S21195|692080|692081|692082|692083|692084|692085|692086|692087|692088|124100|152000",
       names(raw), value = TRUE),
  paste0("exportamt", 2018:2025),
  paste0("rdcost", 2018:2025),
  paste0("lbcost", 2018:2025),
  paste0("mflbcost", 2018:2025),
  paste0("opm", 2018:2025),
  paste0("p", 2019:2025),
  paste0("export", 2018:2025),
  "age", "subclass", "weights", "distance",
  grep("^log_|^labor_prod", names(raw), value = TRUE)
)
numeric_cols <- intersect(numeric_cols, names(raw))
cat(sprintf("  -> %d columns converting to numeric...\n", length(numeric_cols)))
for (col in numeric_cols) {
  raw[[col]] <- suppressWarnings(as.numeric(raw[[col]]))
}
cat("  -> Done\n")

raw$total_gfundvol <- raw$gfundvol2020 + raw$gfundvol2021 + raw$gfundvol2022
raw$n_funded       <- raw$fund2020 + raw$fund2021 + raw$fund2022

# ==============================================================================
# -- 2. 변수 매핑 (list of lists) ---------------------------------------------
# [20260503]
# ==============================================================================
make_var_info <- function(post_yr)  {
  list(
    list(vn="ln_asset",  kr="ln자산",       cat="성장성", pre="2019/Annual S15000.자산총계",        post=paste0(post_yr,"/Annual S15000.자산총계"),        pt_pre="2018/Annual S15000.자산총계",        pt_post="2019/Annual S15000.자산총계"),
    list(vn="ln_rev",    kr="ln매출",       cat="성장성", pre="2019/Annual S21100.총수익",          post=paste0(post_yr,"/Annual S21100.총수익"),          pt_pre="2018/Annual S21100.총수익",          pt_post="2019/Annual S21100.총수익"),
    list(vn="ln_debt",   kr="ln부채",       cat="안정성", pre="2019/Annual S18000.부채총계",        post=paste0(post_yr,"/Annual S18000.부채총계"),        pt_pre="2018/Annual S18000.부채총계",        pt_post="2019/Annual S18000.부채총계"),
    #list(vn="ln_cap",    kr="ln자본금",     cat="성장성", pre="2019/Annual S18100.자본금",          post="2024/Annual S18100.자본금",          pt_pre="2018/Annual S18100.자본금",          pt_post="2019/Annual S18100.자본금"),
    list(vn="opm",       kr="OPM",          cat="수익성", pre="opm2019",         post=paste0("opm", post_yr),         pt_pre="opm2018",         pt_post="opm2019"),
    list(vn="ln_pat1",   kr="ln(특허+1)",   cat="혁신성", pre="p2019",           post=paste0("p", post_yr),           pt_pre="p2019",           pt_post="p2020"),
    list(vn="ln_exp1",   kr="ln(수출+1)",   cat="활동성", pre="exportamt2019",   post=paste0("exportamt", post_yr),   pt_pre="exportamt2018",   pt_post="exportamt2019"),
    list(vn="ln_rd1",    kr="ln(개발비+1)", cat="혁신성", pre="rdcost2019",      post=paste0("rdcost", post_yr),       pt_pre="rdcost2018",      pt_post="rdcost2019"),
    list(vn="ln_lb1",    kr="ln(인건비+1)", cat="활동성", pre="lbcost2019",      post=paste0("lbcost", post_yr),      pt_pre="lbcost2018",      pt_post="lbcost2019")
    #list(vn="ln_mflb1",  kr="ln(노무비+1)", cat="활동성", pre="mflbcost2019",   post="mflbcost2024",    pt_pre="mflbcost2018",    pt_post="mflbcost2019")
  )
}

# ==============================================================================
# -- 3. 유틸리티 함수 ---------------------------------------------------------
# ==============================================================================

# [20260503] OPM은 raw 값 그대로 사용. "ln_"로 시작하는 변수만 로그변환.
ln_transform <- function(x, vn) {
  x <- suppressWarnings(as.numeric(x))
  if (vn == "opm") return(x)                              # OPM raw
  if (grepl("1$", vn)) return(log(pmax(x, 0) + 1))        # ln_*1 → log1p
  else return(log(pmax(x, 1)))                             # ln_* → log
}

sig_label <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "***",
                ifelse(p < 0.01, "**",
                       ifelse(p < 0.05, "*",
                              ifelse(p < 0.1, ".", "ns")))))
}

pt_judge <- function(p) {
  ifelse(is.na(p), "N/A",
         ifelse(p < 0.01, "X_violation",
                ifelse(p < 0.1, "!_borderline", "O_pass")))
}

pw_test <- function(a, b) {
  if (length(a) >= 3 & length(b) >= 3) return(t.test(a, b, var.equal=FALSE)$p.value)
  return(NA)
}

# list-of-lists -> data.frame
to_df <- function(lst) bind_rows(lapply(lst, as.data.frame, stringsAsFactors=FALSE))

# [20260503] 부문 단위 사전값 행렬 생성 (DR-DID에서 공변량으로 사용)
#   - 각 행 = 기업, 각 열 = 변수의 2019년 값 (변수마다 ln_transform 적용 또는 raw)
#   - 컬럼명: <vn>_pre  (예: ln_asset_pre, opm_pre, ...)
make_pre_matrix <- function(seg_data) {
  vn_list <- sapply(var_info, function(vi) vi$vn)
  pre_mat <- as.data.frame(
    sapply(var_info, function(vi) ln_transform(seg_data[[vi$pre]], vi$vn))
  )
  colnames(pre_mat) <- paste0(vn_list, "_pre")
  pre_mat
}

# ==============================================================================
# -- 4. 금액 3분위 할당 -------------------------------------------------------
# ==============================================================================

assign_amt_group <- function(df_seg) {
  funded_pos <- df_seg[df_seg$funded == 1 & df_seg$total_gfundvol > 0, ]
  q33 <- quantile(funded_pos$total_gfundvol, 0.333, na.rm=TRUE)
  q67 <- quantile(funded_pos$total_gfundvol, 0.667, na.rm=TRUE)
  
  df_seg$amt_group <- ifelse(df_seg$funded == 0, "ctrl",
                             ifelse(df_seg$total_gfundvol <= q33, "low",
                                    ifelse(df_seg$total_gfundvol <= q67, "mid", "high")))
  
  cat(sprintf("  [%s] cutoffs: low<=%.1fM<=mid<=%.1fM<=high | ctrl:%d low:%d mid:%d high:%d\n",
              unique(df_seg$seg_name), q33/1e6, q67/1e6,
              sum(df_seg$amt_group=="ctrl"), sum(df_seg$amt_group=="low"),
              sum(df_seg$amt_group=="mid"),  sum(df_seg$amt_group=="high")))
  return(df_seg)
}


# ==============================================================================
# -- 5. 메인 분석 -------------------------------------------------------------
# ==============================================================================

segments   <- c("소재", "부품", "장비")
grp_labels <- c("low", "mid", "high")
grp_kr     <- c(low="하위", mid="중위", high="상위")

POST_YEARS  <- c(2024, 2025)
all_results <- list()

for (POST_YEAR in POST_YEARS) {
  
  cat("\n", strrep("=", 70), "\n")
  cat(sprintf("  ★ POST YEAR = %d  (Pre=2019)\n", POST_YEAR))
  cat(strrep("=", 70), "\n")
  
  # ★ var_info 를 POST_YEAR 에 맞게 동적 생성
  var_info <- make_var_info(POST_YEAR)
  
  # ★ 결과 컨테이너 초기화 (루프마다 새로 시작)
  res_did <- list(); res_dummy <- list()
  res_pt  <- list(); res_cont  <- list()
  res_dr  <- list()

  cat("\n====== Analysis Start ======\n")

  for (seg in segments) {
    cat(sprintf("\n-- Segment: %s --\n", seg))
    seg_data <- assign_amt_group(raw[raw$seg_name == seg, ])
    ctrl <- seg_data[seg_data$amt_group == "ctrl", ]
    
    # [20260503] 부문 단위 사전값 행렬 (DR-DID 공변량용)
    pre_mat_seg <- make_pre_matrix(seg_data)
    
    for (vi in var_info) {
      vn <- vi$vn; kr <- vi$kr; cat_ <- vi$cat
      
      # --- 5-1. Simple DID (per tercile) ----------------------------------------------
      cp <- ln_transform(ctrl[[ vi$pre  ]], vn)
      cq <- ln_transform(ctrl[[ vi$post ]], vn)
      cm <- !is.na(cp) & !is.na(cq); cd <- cq[cm] - cp[cm]
      
      g_diffs <- list(); g_dids <- list(); g_ns <- list(); g_ps <- list()
      
      for (g in grp_labels) {
        gd <- seg_data[seg_data$amt_group == g, ]
        tp <- ln_transform(gd[[ vi$pre  ]], vn)
        tq <- ln_transform(gd[[ vi$post ]], vn)
        tm <- !is.na(tp) & !is.na(tq); td <- tq[tm] - tp[tm]
        g_diffs[[g]] <- td; g_ns[[g]] <- sum(tm)
        
        if (sum(tm) >= 5 & sum(cm) >= 5) {
          dv <- mean(td) - mean(cd); tt <- t.test(td, cd, var.equal=FALSE)
          g_dids[[g]] <- dv; g_ps[[g]] <- tt$p.value
          res_did[[ length(res_did)+1 ]] <- list(
            seg=seg, grp=g, grp_kr=grp_kr[g], cat=cat_, var=kr,
            n_trt=sum(tm), n_ctrl=sum(cm),
            trt_pre=mean(tp[tm]), trt_post=mean(tq[tm]),
            ctrl_pre=mean(cp[cm]), ctrl_post=mean(cq[cm]),
            did=dv, t_val=tt$statistic[[1]], p_val=tt$p.value,
            sig=sig_label(tt$p.value))
        } else {
          g_dids[[g]] <- NA; g_ps[[g]] <- NA
          res_did[[ length(res_did)+1 ]] <- list(
            seg=seg, grp=g, grp_kr=grp_kr[g], cat=cat_, var=kr,
            n_trt=sum(tm), n_ctrl=sum(cm),
            trt_pre=NA, trt_post=NA, ctrl_pre=NA, ctrl_post=NA,
            did=NA, t_val=NA, p_val=NA, sig="")
        }
      }
      # ==============================================================================
      # --- 5-2. Dummy model: ANOVA + pairwise ----------------------------------
      # ==============================================================================
      
      al <- c(list(cd), g_diffs[sapply(g_diffs, function(x) length(x)>=3)])
      if (length(al)>=2) {
        adf <- data.frame(diff=unlist(al), group=factor(rep(seq_along(al), sapply(al, length))))
        f_p <- summary(aov(diff~group, data=adf))[[1]]$`Pr(>F)`[1]
      } else { f_p <- NA }
      
      res_dummy[[ length(res_dummy)+1 ]] <- list(
        seg=seg, cat=cat_, var=kr,
        n_ctrl=sum(cm), n_low=g_ns[["low"]], n_mid=g_ns[["mid"]], n_high=g_ns[["high"]],
        ctrl_mean_diff=mean(cd),
        did_low=g_dids[["low"]], did_mid=g_dids[["mid"]], did_high=g_dids[["high"]],
        p_low=g_ps[["low"]], p_mid=g_ps[["mid"]], p_high=g_ps[["high"]],
        sig_low=sig_label(g_ps[["low"]]), sig_mid=sig_label(g_ps[["mid"]]),
        sig_high=sig_label(g_ps[["high"]]),
        p_F=f_p, sig_F=sig_label(f_p),
        p_hv_l=pw_test(g_diffs[["high"]], g_diffs[["low"]]),
        p_hv_m=pw_test(g_diffs[["high"]], g_diffs[["mid"]]),
        p_mv_l=pw_test(g_diffs[["mid"]],  g_diffs[["low"]]))
  
      # ==============================================================================
      # --- 5-3. Parallel trends ------------------------------------------------
      # ==============================================================================
      
      cp2 <- ln_transform(ctrl[[ vi$pt_pre  ]], vn)
      cq2 <- ln_transform(ctrl[[ vi$pt_post ]], vn)
      cm2 <- !is.na(cp2) & !is.na(cq2); cd2 <- cq2[cm2] - cp2[cm2]
      
      for (g in grp_labels) {
        gd <- seg_data[seg_data$amt_group == g, ]
        tp2 <- ln_transform(gd[[ vi$pt_pre  ]], vn)
        tq2 <- ln_transform(gd[[ vi$pt_post ]], vn)
        tm2 <- !is.na(tp2) & !is.na(tq2); td2 <- tq2[tm2] - tp2[tm2]
        if (sum(tm2)>=3 & sum(cm2)>=3) {
          ptt <- t.test(td2, cd2, var.equal=FALSE)
          res_pt[[ length(res_pt)+1 ]] <- list(
            seg=seg, grp=g, grp_kr=grp_kr[g], var=kr,
            t_val=round(ptt$statistic[[1]],4), p_val=round(ptt$p.value,4),
            judge=pt_judge(ptt$p.value))
        } else {
          res_pt[[ length(res_pt)+1 ]] <- list(
            seg=seg, grp=g, grp_kr=grp_kr[g], var=kr,
            t_val=NA, p_val=NA, judge="N/A")
        }
      }
    
    # ==============================================================================
    # --- 5-4. [20260503] Doubly Robust DID -----------------------------------
    #  for each amount group g: lm(diff_y ~ treat_dummy + 7개 사전공변량)
    #  공변량 = pre_mat_seg에서 자기자신의 _pre 컬럼 제외한 나머지 7개
    # ==============================================================================
    pre_v_full  <- ln_transform(seg_data[[ vi$pre  ]], vn)
    post_v_full <- ln_transform(seg_data[[ vi$post ]], vn)
    
    cov_pre_cols <- setdiff(colnames(pre_mat_seg), paste0(vn, "_pre"))
    
    for (g in grp_labels) {
      sel <- seg_data$amt_group %in% c(g, "ctrl")
      
      df_dr <- data.frame(
        diff_y      = post_v_full[sel] - pre_v_full[sel],
        treat_dummy = as.integer(seg_data$amt_group[sel] == g),
        pre_mat_seg[sel, cov_pre_cols, drop=FALSE]
      )
      df_dr <- df_dr[complete.cases(df_dr), ]
      
      n_t <- sum(df_dr$treat_dummy == 1)
      n_c <- sum(df_dr$treat_dummy == 0)
      
      if (n_t < 5 || n_c < 5 || nrow(df_dr) < length(cov_pre_cols) + 5) {
        res_dr[[ length(res_dr)+1 ]] <- list(
          seg=seg, grp=g, grp_kr=grp_kr[g], cat=cat_, var=kr,
          n_trt=n_t, n_ctrl=n_c,
          DID_simple=NA, p_simple=NA, sig_simple="",
          DID_DR=NA, SE_DR=NA, t_DR=NA, p_DR=NA, sig_DR="")
        next
      }
      
      # (A) Simple DID (lm version, 비교용)
      reg_s <- tryCatch(lm(diff_y ~ treat_dummy, data=df_dr),
                        error=function(e) NULL)
      did_s <- if (!is.null(reg_s)) coef(reg_s)["treat_dummy"] else NA
      p_s   <- if (!is.null(reg_s))
        summary(reg_s)$coefficients["treat_dummy","Pr(>|t|)"] else NA
      
      # (B) Doubly Robust: 자기자신 _pre 제외한 나머지 7개 _pre 통제
      fmla <- as.formula(paste("diff_y ~ treat_dummy +",
                               paste(cov_pre_cols, collapse=" + ")))
      reg_dr <- tryCatch(lm(fmla, data=df_dr), error=function(e) NULL)
      
      if (is.null(reg_dr) || !("treat_dummy" %in% rownames(summary(reg_dr)$coefficients))) {
        res_dr[[ length(res_dr)+1 ]] <- list(
          seg=seg, grp=g, grp_kr=grp_kr[g], cat=cat_, var=kr,
          n_trt=n_t, n_ctrl=n_c,
          DID_simple=did_s, p_simple=p_s, sig_simple=sig_label(p_s),
          DID_DR=NA, SE_DR=NA, t_DR=NA, p_DR=NA, sig_DR="")
        next
      }
      
      coefs <- summary(reg_dr)$coefficients
      did_dr <- coefs["treat_dummy", "Estimate"]
      se_dr  <- coefs["treat_dummy", "Std. Error"]
      t_dr   <- coefs["treat_dummy", "t value"]
      p_dr   <- coefs["treat_dummy", "Pr(>|t|)"]
      
      res_dr[[ length(res_dr)+1 ]] <- list(
        seg=seg, grp=g, grp_kr=grp_kr[g], cat=cat_, var=kr,
        n_trt=n_t, n_ctrl=n_c,
        DID_simple=did_s,  p_simple=p_s,  sig_simple=sig_label(p_s),
        DID_DR=did_dr,     SE_DR=se_dr,   t_DR=t_dr,
        p_DR=p_dr,         sig_DR=sig_label(p_dr))
    }
  }
  
  
  # ============================================================================
    # --- 5-4. Continuous dose-response ----------------------------------------
    # diff_y ~ TR + TR × ln(금액) + TR × ln(금액)² 수정
    # β₂(선형항)와 β₃(2차항)를 동시에 추정하면 역U형 또는 U형 관계 포착
  # ============================================================================
    
    for (vi in var_info) {
      vn <- vi$vn; kr <- vi$kr
      pre_v  <- ln_transform(seg_data[[ vi$pre  ]], vn)
      post_v <- ln_transform(seg_data[[ vi$post ]], vn)
      valid  <- !is.na(pre_v) & !is.na(post_v)
      sub    <- seg_data[valid, ]
      sub$diff_y <- post_v[valid] - pre_v[valid]
      sub$tr     <- sub$funded
      sub$ln_fv  <- log(pmax(sub$total_gfundvol, 0) + 1)
      sub$ln_fv2 <- sub$ln_fv^2
      sub$tr_x_fv  <- sub$tr * sub$ln_fv
      sub$tr_x_fv2 <- sub$tr * sub$ln_fv2
      
      if (nrow(sub) < 10) next
      
      # ★ 공변량: 자기자신 _pre 제외한 나머지 7개
      cov_cols <- setdiff(colnames(pre_mat_seg), paste0(vn, "_pre"))
      cov_df   <- pre_mat_seg[valid, cov_cols, drop=FALSE]
      
      sub_full <- cbind(sub, cov_df)
      sub_full <- sub_full[complete.cases(sub_full[, c("diff_y","tr",
                                                       "tr_x_fv","tr_x_fv2", cov_cols)]), ]
      
      if (nrow(sub_full) < length(cov_cols) + 5) next
      
      tryCatch({
        # ── 선형 모형 (공변량 포함) ───────────────────────────
        fmla_lin  <- as.formula(paste(
          "diff_y ~ tr + tr_x_fv +",
          paste(cov_cols, collapse=" + ")))
        
        # ── 2차항 모형 (공변량 포함) ──────────────────────────
        fmla_quad <- as.formula(paste(
          "diff_y ~ tr + tr_x_fv + tr_x_fv2 +",
          paste(cov_cols, collapse=" + ")))
        
        mod_lin  <- lm(fmla_lin,  data=sub_full)
        mod_quad <- lm(fmla_quad, data=sub_full)
        rob_lin  <- coeftest(mod_lin,  vcov=vcovHC(mod_lin,  type="HC1"))
        rob_quad <- coeftest(mod_quad, vcov=vcovHC(mod_quad, type="HC1"))
        
        # ── F검정: 2차항 필요성 판단 ──────────────────────────
        anova_test <- anova(mod_lin, mod_quad)
        f_quad_p   <- anova_test$`Pr(>F)`[2]
        
        # ── 최적 투자금액 (역U형: β₃ < 0 일 때만) ─────────────
        b2 <- coef(mod_quad)["tr_x_fv"]
        b3 <- coef(mod_quad)["tr_x_fv2"]
        optimal_amt <- NA_real_
        if (!is.na(b3) && b3 < 0) {
          opt_ln_fv   <- -b2 / (2 * b3)
          optimal_amt <- exp(opt_ln_fv) - 1
        }
        
        # ── 결과 저장 ─────────────────────────────────────────
        res_cont[[ length(res_cont)+1 ]] <- list(
          seg         = seg, var = kr, N = nrow(sub_full),
          # 선형 모형
          b_treat     = round(rob_lin["tr",         "Estimate"], 4),
          p_treat     = round(rob_lin["tr",         "Pr(>|t|)"], 4),
          b_dose      = round(rob_lin["tr_x_fv",    "Estimate"], 6),
          t_dose      = round(rob_lin["tr_x_fv",    "t value"],  3),
          p_dose      = round(rob_lin["tr_x_fv",    "Pr(>|t|)"], 4),
          sig_dose    = sig_label(rob_lin["tr_x_fv",  "Pr(>|t|)"]),
          R2_lin      = round(summary(mod_lin)$r.squared, 4),
          # 2차항 모형
          b_dose2     = round(rob_quad["tr_x_fv2",  "Estimate"], 6),
          t_dose2     = round(rob_quad["tr_x_fv2",  "t value"],  3),
          p_dose2     = round(rob_quad["tr_x_fv2",  "Pr(>|t|)"], 4),
          sig_dose2   = sig_label(rob_quad["tr_x_fv2", "Pr(>|t|)"]),
          R2_quad     = round(summary(mod_quad)$r.squared, 4),
          # 모형 비교
          F_quad_p    = round(f_quad_p, 4),
          F_quad_sig  = sig_label(f_quad_p),
          optimal_amt = round(optimal_amt, 0)
        )
      }, error=function(e) NULL)
    }
  }

  # ★ 여기부터 추가 ────────────────────────────────────────────
  # 연도별 결과 누적
  all_results[[as.character(POST_YEAR)]] <- list(
    res_did   = res_did,
    res_dummy = res_dummy,
    res_pt    = res_pt,
    res_cont  = res_cont,
    res_dr    = res_dr
    # nf_did    = nf_did,
    # nf_dummy  = nf_dummy,
    # nf_pt     = nf_pt,
    # nf_dr     = nf_dr
  )
  cat(sprintf("\n★ POST_YEAR=%d 완료. 결과 누적.\n", POST_YEAR))

}  # ← for (POST_YEAR in POST_YEARS) 루프 닫기
cat("\n====== Analysis Complete ======\n")



#===============================================================================
# # 통합 저장
#===============================================================================

excel_sheets <- list()

make_wide <- function(df_sub) {
  df_sub %>%
    select(var, grp_kr, did, p_val, sig) %>%
    pivot_wider(names_from=grp_kr,
                values_from=c(did, p_val, sig),
                names_glue="{grp_kr}_{.value}")
}

make_wide_dr <- function(df_sub) {
  df_sub %>%
    select(var, grp_kr, DID_DR, p_DR, sig_DR) %>%
    pivot_wider(names_from=grp_kr,
                values_from=c(DID_DR, p_DR, sig_DR),
                names_glue="{grp_kr}_{.value}")
}

for (yr_key in names(all_results)) {
  res <- all_results[[yr_key]]
  sfx <- paste0("_", yr_key)          # "_2024" 또는 "_2025"
  
  # 데이터프레임 조립
  df_did   <- to_df(res$res_did);    df_dummy <- to_df(res$res_dummy)
  df_pt    <- to_df(res$res_pt);     df_cont  <- to_df(res$res_cont)
  df_dr    <- to_df(res$res_dr)

  # Wide format 생성
  wide_mat <- make_wide(df_did[df_did$seg == "소재", ])
  wide_prt <- make_wide(df_did[df_did$seg == "부품", ])
  wide_eqp <- make_wide(df_did[df_did$seg == "장비", ])
  wide_dr_mat <- if (nrow(df_dr)>0) make_wide_dr(df_dr[df_dr$seg=="소재",]) else NULL
  wide_dr_prt <- if (nrow(df_dr)>0) make_wide_dr(df_dr[df_dr$seg=="부품",]) else NULL
  wide_dr_eqp <- if (nrow(df_dr)>0) make_wide_dr(df_dr[df_dr$seg=="장비",]) else NULL
  
  # 시트 등록 (접미사 _2024 / _2025)
  excel_sheets[[paste0("amt_simple_DID",sfx)]] <- df_did
  excel_sheets[[paste0("amt_Anova_Pair",sfx)]] <- df_dummy
  excel_sheets[[paste0("amt_PT",        sfx)]] <- df_pt
  excel_sheets[[paste0("amt_DoseRes",   sfx)]] <- df_cont
  excel_sheets[[paste0("amt_DR_DID",    sfx)]] <- df_dr
  excel_sheets[[paste0("amt_Wide_mat",  sfx)]] <- wide_mat
  excel_sheets[[paste0("amt_Wide_prt",  sfx)]] <- wide_prt
  excel_sheets[[paste0("amt_Wide_eqp",  sfx)]] <- wide_eqp
  excel_sheets[[paste0("amt_Wide_DR_mat",  sfx)]] <- wide_mat
  excel_sheets[[paste0("amt_Wide_DR_part",  sfx)]] <- wide_prt
  excel_sheets[[paste0("amt_Wide_DR_eqp",  sfx)]] <- wide_eqp

}

write_xlsx(excel_sheets, path = "DID_Amount_Analysis.xlsx")
cat(sprintf("\nSaved: DID_Amount Analysis.xlsx (%d 시트)\n", length(excel_sheets)))
cat("====== All Done ======\n")


# # 통합 저장
# write_xlsx(
#   list(
#     "amt_DID"      = df_did,
#     "amt_dummy"    = df_dummy,
#     "amt_PT"       = df_pt,
#     "amt_cont"     = df_cont,
#     "amt_DR_DID"     = df_dr,        # [20260503]
#     "amt_Wide_mat" = wide_mat,
#     "amt_Wide_prt" = wide_prt,
#     "amt_Wide_eqp" = wide_eqp,
#     "nf_DID"       = df_nf_did,
#     "nf_dummy"     = df_nf_dummy,
#     "nf_PT"        = df_nf_pt,
#     "nf_DR_DID"      = df_nf_dr      # [20260503]
#   ),
#   path = "DID_Amount_and_Nfunded_Combined.xlsx"
# )
# 
# cat("\nSaved: DID_Amount_and_Nfunded_Combined.xlsx\n")
# cat("====== All Done ======\n")
