################################################################################
# 소부장 R&D 투자 총금액별 DID 분석 (부문별 세분화)
# 
# 분석 내용:
#   1. 투자 총금액 3분위(하위/중위/상위) x 부문(소재/부품/장비) DID
#   2. 더미모형 (0회 기준, ANOVA 이질성 검정 + pairwise 비교)
#   3. 평행추세 검정 (pre-treatment trend)
#   4. 연속변수 dose-response 모형
#   5. Wide format 요약표 + Excel 저장
#   6. (보너스) 투자 횟수(n_funded) 기준 DID 재현
#
# 데이터: matched_dataset_segPSM.xlsx (PSM 매칭 완료 데이터)
# pre = 2019, post = 2024, treatment period = 2020~2022
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
################################################################################

# -- 패키지 로드 ---------------------------------------------------------------
library(readxl)
library(dplyr)
library(tidyr)
library(writexl)
library(sandwich)
library(lmtest)

# -- 데이터 로드 ---------------------------------------------------------------
raw <- read_excel("matched_dataset_segPSM.xlsx")

# -- 1. 수치 컬럼 일괄 변환 (Excel text -> numeric) ---------------------------

numeric_cols <- c(
  "fund2020", "fund2021", "fund2022",
  "gfundvol2020", "gfundvol2021", "gfundvol2022",
  "fundedpattern", "funded", "seg", "treat",
  grep("S15000|S18000|S18100|S21100|S25000|S05000|S21190|S21195|692080|692081|692082|692083|692084|692085|692086|692087|692088|124100|152000",
       names(raw), value = TRUE),
  paste0("exportamt", 2018:2024),
  paste0("rdcost", 2018:2024),
  paste0("lbcost", 2018:2024),
  paste0("mflbcost", 2018:2024),
  paste0("p", 2019:2024),
  paste0("export", 2018:2024),
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


# -- 2. 변수 매핑 (list of lists) ---------------------------------------------

var_info <- list(
  list(vn="ln_asset",  kr="ln자산",       cat="성장성", pre="2019/Annual S15000.자산총계",        post="2024/Annual S15000.자산총계",        pt_pre="2018/Annual S15000.자산총계",        pt_post="2019/Annual S15000.자산총계"),
  list(vn="ln_rev",    kr="ln매출",       cat="성장성", pre="2019/Annual S21100.총수익",          post="2024/Annual S21100.총수익",          pt_pre="2018/Annual S21100.총수익",          pt_post="2019/Annual S21100.총수익"),
  list(vn="ln_debt",   kr="ln부채",       cat="안정성", pre="2019/Annual S18000.부채총계",        post="2024/Annual S18000.부채총계",        pt_pre="2018/Annual S18000.부채총계",        pt_post="2019/Annual S18000.부채총계"),
  list(vn="ln_cap",    kr="ln자본금",     cat="성장성", pre="2019/Annual S18100.자본금",          post="2024/Annual S18100.자본금",          pt_pre="2018/Annual S18100.자본금",          pt_post="2019/Annual S18100.자본금"),
  list(vn="ln_opinc",  kr="ln영업이익",   cat="수익성", pre="2019/Annual S25000.영업이익(손실)",  post="2024/Annual S25000.영업이익(손실)",  pt_pre="2018/Annual S25000.영업이익(손실)",  pt_post="2019/Annual S25000.영업이익(손실)"),
  list(vn="ln_pat1",   kr="ln(특허+1)",   cat="혁신성", pre="p2019",           post="p2024",           pt_pre="p2019",           pt_post="p2020"),
  list(vn="ln_exp1",   kr="ln(수출+1)",   cat="활동성", pre="exportamt2019",   post="exportamt2024",   pt_pre="exportamt2018",   pt_post="exportamt2019"),
  list(vn="ln_rd1",    kr="ln(개발비+1)", cat="혁신성", pre="rdcost2019",      post="rdcost2024",      pt_pre="rdcost2018",      pt_post="rdcost2019"),
  list(vn="ln_lb1",    kr="ln(인건비+1)", cat="활동성", pre="lbcost2019",      post="lbcost2024",      pt_pre="lbcost2018",      pt_post="lbcost2019"),
  list(vn="ln_mflb1",  kr="ln(노무비+1)", cat="활동성", pre="mflbcost2019",    post="mflbcost2024",    pt_pre="mflbcost2018",    pt_post="mflbcost2019")
)


# -- 3. 유틸리티 함수 ---------------------------------------------------------

ln_transform <- function(x, vn) {
  x <- suppressWarnings(as.numeric(x))
  if (grepl("1$", vn)) return(log(pmax(x, 0) + 1))
  else return(log(pmax(x, 1)))
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


# -- 4. 금액 3분위 할당 -------------------------------------------------------

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


################################################################################
# -- 5. 메인 분석 -------------------------------------------------------------
################################################################################

segments   <- c("소재", "부품", "장비")
grp_labels <- c("low", "mid", "high")
grp_kr     <- c(low="하위", mid="중위", high="상위")

res_did <- list(); res_dummy <- list(); res_pt <- list(); res_cont <- list()

cat("\n====== Analysis Start ======\n")

for (seg in segments) {
  cat(sprintf("\n-- Segment: %s --\n", seg))
  seg_data <- assign_amt_group(raw[raw$seg_name == seg, ])
  ctrl <- seg_data[seg_data$amt_group == "ctrl", ]
  
  for (vi in var_info) {
    vn <- vi$vn; kr <- vi$kr; cat_ <- vi$cat
    
    # --- 5-1. DID (per tercile) ----------------------------------------------
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
    
    # --- 5-2. Dummy model: ANOVA + pairwise ----------------------------------
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
    
    # --- 5-3. Parallel trends ------------------------------------------------
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
  }
  
  # --- 5-4. Continuous dose-response -----------------------------------------
  for (vi in var_info) {
    vn <- vi$vn; kr <- vi$kr
    pre_v  <- ln_transform(seg_data[[ vi$pre  ]], vn)
    post_v <- ln_transform(seg_data[[ vi$post ]], vn)
    valid  <- !is.na(pre_v) & !is.na(post_v)
    sub <- seg_data[valid, ]
    sub$diff_y  <- post_v[valid] - pre_v[valid]
    sub$tr      <- sub$funded
    sub$ln_fv   <- log(pmax(sub$total_gfundvol, 0) + 1)
    sub$tr_x_fv <- sub$tr * sub$ln_fv
    if (nrow(sub) < 10) next
    tryCatch({
      mod <- lm(diff_y ~ tr + tr_x_fv, data=sub)
      rob <- coeftest(mod, vcov=vcovHC(mod, type="HC1"))
      res_cont[[ length(res_cont)+1 ]] <- list(
        seg=seg, var=kr, N=nrow(sub),
        b_treat=round(rob["tr","Estimate"],4),
        p_treat=round(rob["tr","Pr(>|t|)"],4),
        b_dose=round(rob["tr_x_fv","Estimate"],6),
        t_dose=round(rob["tr_x_fv","t value"],3),
        p_dose=round(rob["tr_x_fv","Pr(>|t|)"],4),
        sig_dose=sig_label(rob["tr_x_fv","Pr(>|t|)"]),
        R2=round(summary(mod)$r.squared,4))
    }, error=function(e) NULL)
  }
}

cat("\n====== Analysis Complete ======\n")


################################################################################
# -- 6. 결과 조립 및 출력 -----------------------------------------------------
################################################################################

df_did   <- to_df(res_did)
df_dummy <- to_df(res_dummy)
df_pt    <- to_df(res_pt)
df_cont  <- to_df(res_cont)

# Wide format
make_wide <- function(df_sub) {
  df_sub %>%
    select(var, grp_kr, did, p_val, sig) %>%
    pivot_wider(names_from=grp_kr,
                values_from=c(did, p_val, sig),
                names_glue="{grp_kr}_{.value}")
}
wide_mat <- make_wide(df_did[df_did$seg == "소재", ])
wide_prt <- make_wide(df_did[df_did$seg == "부품", ])
wide_eqp <- make_wide(df_did[df_did$seg == "장비", ])

# Console output
cat("\n", strrep("=",80), "\n")
cat("Significant DID (p < 0.1)\n", strrep("=",80), "\n")
sig_did <- df_did[df_did$sig != "ns" & df_did$sig != "", ]
print(sig_did[, c("seg","grp_kr","var","did","t_val","p_val","sig")])

cat("\n", strrep("=",80), "\n")
cat("F-heterogeneity significant (p < 0.1)\n", strrep("=",80), "\n")
sig_f <- df_dummy[df_dummy$sig_F %in% c("***","**","*","."), ]
print(sig_f[, c("seg","var","did_low","did_mid","did_high","p_F","sig_F","p_hv_l","p_hv_m","p_mv_l")])

cat("\n", strrep("=",80), "\n")
cat("Parallel Trends\n", strrep("=",80), "\n")
print(table(df_pt$judge))
pt_prob <- df_pt[df_pt$judge != "O_pass", ]
if (nrow(pt_prob)>0) { cat("Issues:\n"); print(pt_prob) }

cat("\n", strrep("=",80), "\n")
cat("Continuous dose-response significant (p < 0.1)\n", strrep("=",80), "\n")
sig_c <- df_cont[df_cont$sig_dose != "ns" & df_cont$sig_dose != "", ]
if (nrow(sig_c)>0) print(sig_c) else cat("  None\n")


################################################################################
# -- 7. Excel 저장 (한글 컬럼명 복원) ------------------------------------------
################################################################################

write_xlsx(
  list(
    "DID_result"    = df_did %>% rename(부문=seg, 금액그룹=grp_kr, 카테고리=cat, 변수=var,
                                        N_처치=n_trt, N_통제=n_ctrl, 처치_pre=trt_pre, 처치_post=trt_post,
                                        통제_pre=ctrl_pre, 통제_post=ctrl_post, DID=did,
                                        t_value=t_val, p_value=p_val, 유의성=sig),
    "dummy_model"   = df_dummy %>% rename(부문=seg, 카테고리=cat, 변수=var,
                                          N_통제=n_ctrl, N_하위=n_low, N_중위=n_mid, N_상위=n_high,
                                          통제평균diff=ctrl_mean_diff,
                                          DID_하위=did_low, DID_중위=did_mid, DID_상위=did_high,
                                          p_하위=p_low, p_중위=p_mid, p_상위=p_high,
                                          sig_하위=sig_low, sig_중위=sig_mid, sig_상위=sig_high,
                                          p_F이질성=p_F, sig_F이질성=sig_F,
                                          p_상vs하=p_hv_l, p_상vs중=p_hv_m, p_중vs하=p_mv_l),
    "parallel_trend" = df_pt %>% rename(부문=seg, 금액그룹=grp_kr, 변수=var,
                                        t값=t_val, p값=p_val, 판정=judge),
    "continuous"    = df_cont %>% rename(부문=seg, 변수=var),
    "Wide_소재"     = wide_mat,
    "Wide_부품"     = wide_prt,
    "Wide_장비"     = wide_eqp
  ),
  path = "DID_FundAmount_Seg_R.xlsx"
)
cat("\nSaved: DID_FundAmount_Seg_R.xlsx\n")


################################################################################
# -- 8. (보너스) 투자 횟수(n_funded) 기준 DID ----------------------------------
################################################################################

cat("\n", strrep("=",80), "\n")
cat("=== N_funded analysis ===\n", strrep("=",80), "\n")

nf_did <- list(); nf_dummy <- list(); nf_pt <- list()

for (seg in segments) {
  seg_data <- raw[raw$seg_name == seg, ]
  ctrl <- seg_data[seg_data$funded == 0, ]
  
  for (vi in var_info) {
    vn <- vi$vn; kr <- vi$kr; cat_ <- vi$cat
    cp <- ln_transform(ctrl[[ vi$pre  ]], vn)
    cq <- ln_transform(ctrl[[ vi$post ]], vn)
    cm <- !is.na(cp) & !is.na(cq); cd <- cq[cm] - cp[cm]
    
    g_d <- list(); g_did <- list(); g_n <- list(); g_p <- list()
    for (nf in 1:3) {
      g <- paste0(nf, "x")
      gd <- seg_data[seg_data$n_funded == nf, ]
      tp <- ln_transform(gd[[ vi$pre ]], vn)
      tq <- ln_transform(gd[[ vi$post ]], vn)
      tm <- !is.na(tp) & !is.na(tq); td <- tq[tm] - tp[tm]
      g_d[[g]] <- td; g_n[[g]] <- sum(tm)
      if (sum(tm)>=5 & sum(cm)>=5) {
        dv <- mean(td)-mean(cd); tt <- t.test(td,cd,var.equal=FALSE)
        g_did[[g]] <- dv; g_p[[g]] <- tt$p.value
        nf_did[[ length(nf_did)+1 ]] <- list(
          seg=seg, nf=nf, nf_label=paste0(nf,"회"), cat=cat_, var=kr,
          n_trt=sum(tm), n_ctrl=sum(cm), did=dv,
          t_val=tt$statistic[[1]], p_val=tt$p.value, sig=sig_label(tt$p.value))
      } else { g_did[[g]] <- NA; g_p[[g]] <- NA }
    }
    
    al <- c(list(cd), g_d[sapply(g_d, function(x) length(x)>=3)])
    fp <- if(length(al)>=2) {
      adf <- data.frame(diff=unlist(al), group=factor(rep(seq_along(al), sapply(al,length))))
      summary(aov(diff~group, data=adf))[[1]]$`Pr(>F)`[1]
    } else NA
    
    nf_dummy[[ length(nf_dummy)+1 ]] <- list(
      seg=seg, cat=cat_, var=kr,
      n_0=sum(cm), n_1=g_n[["1x"]], n_2=g_n[["2x"]], n_3=g_n[["3x"]],
      ctrl_diff=mean(cd),
      did_1=g_did[["1x"]], did_2=g_did[["2x"]], did_3=g_did[["3x"]],
      p_1=g_p[["1x"]], p_2=g_p[["2x"]], p_3=g_p[["3x"]],
      sig_1=sig_label(g_p[["1x"]]), sig_2=sig_label(g_p[["2x"]]), sig_3=sig_label(g_p[["3x"]]),
      p_F=fp, sig_F=sig_label(fp),
      p_3v1=pw_test(g_d[["3x"]],g_d[["1x"]]),
      p_3v2=pw_test(g_d[["3x"]],g_d[["2x"]]),
      p_2v1=pw_test(g_d[["2x"]],g_d[["1x"]]))
    
    cp2 <- ln_transform(ctrl[[ vi$pt_pre  ]], vn)
    cq2 <- ln_transform(ctrl[[ vi$pt_post ]], vn)
    cm2 <- !is.na(cp2) & !is.na(cq2); cd2 <- cq2[cm2] - cp2[cm2]
    for (nf in 1:3) {
      g <- paste0(nf, "x")
      gd <- seg_data[seg_data$n_funded == nf, ]
      tp2 <- ln_transform(gd[[ vi$pt_pre  ]], vn)
      tq2 <- ln_transform(gd[[ vi$pt_post ]], vn)
      tm2 <- !is.na(tp2) & !is.na(tq2); td2 <- tq2[tm2] - tp2[tm2]
      if (sum(tm2)>=3 & sum(cm2)>=3) {
        ptt <- t.test(td2, cd2, var.equal=FALSE)
        nf_pt[[ length(nf_pt)+1 ]] <- list(seg=seg, nf=nf, nf_label=paste0(nf,"회"), var=kr,
                                           t_val=round(ptt$statistic[[1]],4), p_val=round(ptt$p.value,4), judge=pt_judge(ptt$p.value))
      } else {
        nf_pt[[ length(nf_pt)+1 ]] <- list(seg=seg, nf=nf, nf_label=paste0(nf,"회"), var=kr,
                                           t_val=NA, p_val=NA, judge="N/A")
      }
    }
  }
}

df_nf_did   <- to_df(nf_did)
df_nf_dummy <- to_df(nf_dummy)
df_nf_pt    <- to_df(nf_pt)

cat("\n[N_funded] Significant DID:\n")
if (nrow(df_nf_did)>0) {
  s <- df_nf_did[df_nf_did$sig != "ns" & df_nf_did$sig != "", ]
  print(s[, c("seg","nf_label","var","did","t_val","p_val","sig")])
}
cat("\n[N_funded] F-heterogeneity:\n")
if (nrow(df_nf_dummy)>0) {
  s2 <- df_nf_dummy[df_nf_dummy$sig_F %in% c("***","**","*","."), ]
  print(s2[, c("seg","var","did_1","did_2","did_3","p_F","sig_F")])
}

# 통합 저장
write_xlsx(
  list(
    "amt_DID"      = df_did,
    "amt_dummy"    = df_dummy,
    "amt_PT"       = df_pt,
    "amt_cont"     = df_cont,
    "amt_Wide_mat" = wide_mat,
    "amt_Wide_prt" = wide_prt,
    "amt_Wide_eqp" = wide_eqp,
    "nf_DID"       = df_nf_did,
    "nf_dummy"     = df_nf_dummy,
    "nf_PT"        = df_nf_pt
  ),
  path = "DID_Amount_and_Nfunded_Combined.xlsx"
)

cat("\nSaved: DID_Amount_and_Nfunded_Combined.xlsx\n")
cat("====== All Done ======\n")