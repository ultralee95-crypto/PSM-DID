# ==============================================================================
# PSM 분석 - 부문별(소재/부품/장비) 분리 매칭 버전
# 부문별 독립 PSM 비로 Main version of PSM
# ==============================================================================

packages <- c("readxl", "dplyr", "tidyr", "ggplot2", "gridExtra",
              "writexl", "MatchIt", "cobalt")

for (pkg in packages) {
  if (!require(pkg, character.only = TRUE, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org/")
    library(pkg, character.only = TRUE)
  }
}

select <- dplyr::select

#setwd("~/Downloads/Venture")
getwd()

# 한글 폰트
kfont <- if (Sys.info()["sysname"] == "Darwin") "AppleGothic" else
  if (Sys.info()["sysname"] == "Windows") "Malgun Gothic" else "NanumGothic"
theme_set(theme_minimal(base_family = kfont))

# ==============================================================================
# 1. 데이터 로드 및 전처리
# ==============================================================================

cat("=== 데이터 로드 ===\n")

analysis_data <- read_excel("sobujang_integrated_with_valuesearch+Patent.xlsx",
                   sheet = 1, col_types = "text")

head(analysis_data)
names(analysis_data)

to_numeric <- function(x) as.numeric(gsub(",", "", as.character(x)))

# 변환 전 결측치 확인하여 데이터 확인 함.
result = analysis_data %>%
  summarise(
    n_total = n(),
    na_자산 = sum(is.na(`2019/Annual S15000.자산총계`)),
    na_매출 = sum(is.na(`2019/Annual S21100.총수익`)),
    na_부채 = sum(is.na(`2019/Annual S18000.부채총계`)),
    na_자본금 = sum(is.na(`2019/Annual S18100.자본금`)),
    #na_자본총 = sum(is.na(`2019/Annual S18900.자본총계`)), 음수가 발생
    na_종업원 = sum(is.na(`2019/Annual S05000.종업원수`)),
    na_업력 = sum(is.na(`age`)),
    na_KSIC_표준 = sum(is.na(`KSIC_mid_code`)),
    na_지역 = sum(is.na(`region`)),
    #na_노동생산성= sum(is.na(`labor_prod2019`)), #삭제 해야 함. 
    na_영업이익= sum(is.na(`2019/Annual S25000.영업이익(손실)`)),
    na_연구개발비 = sum(is.na(`rdcost2019`)),
    na_수출 = sum(is.na(`exportamt2019`)),
    na_인건비 = sum(is.na('lbcost2019')),
    na_노무비 = sum(is.na(`mflbcost2019`))
   )

print(result, width = Inf)

analysis_data <- analysis_data %>%
  filter(group %in% c("Treatgroup", "Controlgroup")) %>%
  mutate(
    자산       = to_numeric(`2019/Annual S15000.자산총계`),
    매출       = to_numeric(`2019/Annual S21100.총수익`),
    부채       = to_numeric(`2019/Annual S18000.부채총계`),
    자본금     = to_numeric(`2019/Annual S18100.자본금`),
    #종업원수   = to_numeric(`2019/Annual S05000.종업원수`),
    업력       = to_numeric(`age`),
    산업       = to_numeric(`KSIC_mid_code`),
    지역       = to_numeric(`region`),
    영업이익   = to_numeric(`2019/Annual S25000.영업이익(손실)`),
    #노동생산성 = to_numeric(`labor_prod2019`),
    연구개발비 = to_numeric(rdcost2019),
    수출금액   = to_numeric(exportamt2019), 
    인건비     = to_numeric(lbcost2019),
    노무비     = to_numeric(mflbcost2019),
    
    treat      = ifelse(group == "Treatgroup", 1, 0),
    log_자산      = log(자산 + 1),
    log_매출      = log(매출 + 1),
    log_부채      = log(부채 + 1),
    log_자본금    = log(자본금 + 1),
    #log_종업원수  = log(종업원수 + 1),
    log_업력      = log(업력 + 1),
    log_산업      = log(산업 + 1),
    log_지역      = log(지역 + 1),
    #log_노동생산성 = log(노동생산성 + 1),
    log_연구개발비 = log(연구개발비 + 1), 
    log_수출       = log(수출금액 + 1),       
    log_인건비   = log(인건비 + 1),
    log_노무비   = log(노무비 + 1)
  ) %>%
  filter(!is.na(자산) & !is.na(매출) & !is.na(부채) & !is.na(자본금) &
           !is.na(업력) & !is.na(영업이익))

cat("전체 샘플:", nrow(analysis_data), "\n")
cat("seg 분포:\n"); print(table(analysis_data$seg, analysis_data$treat))

# PSM 공식
psm_formula <- treat ~ log_자산 + log_매출 + log_부채 + log_자본금 +
    log_업력 + factor(산업) + factor(지역) + log_연구개발비 + log_수출 + 
  log_인건비 + log_노무비

# ==============================================================================
# 2. 방법 A: 기존 방식 (통합 PSM + exact = ~seg)
# ==============================================================================

cat("\n\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
cat("  방법 A: 통합 PSM (exact = ~seg)\n")
cat(paste(rep("=", 70), collapse = ""), "\n")

matched_A <- matchit(
  psm_formula,
  data     = analysis_data,
  method   = "nearest",
  distance = "logit",
  ratio    = 1,
  caliper  = 0.2,
  replace  = FALSE,
  exact    = ~seg
)

matched_data_A <- match.data(matched_A)

cat("\n[방법 A] 매칭 결과:\n")
cat("  전체 처치:", sum(matched_data_A$treat == 1),
    "/ 전체 통제:", sum(matched_data_A$treat == 0), "\n")
print(table(matched_data_A$seg, matched_data_A$treat))

# ==============================================================================
# 3. 방법 B: 부문별 독립 PSM
# ==============================================================================

cat("\n\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
cat("  방법 B: 부문별 독립 PSM (소재/부품/장비 각각 매칭)\n")
cat(paste(rep("=", 70), collapse = ""), "\n")

seg_labels <- c("1" = "소재", "2" = "부품", "3" = "장비")
matched_list_B  <- list()
matchit_list_B  <- list()

for (s in names(seg_labels)) {
  seg_name <- seg_labels[s]
  sub_data <- analysis_data %>% filter(seg == s)
  
  n_t <- sum(sub_data$treat == 1)
  n_c <- sum(sub_data$treat == 0)
  cat(sprintf("\n  ─ %s (처치: %d, 통제: %d) ─\n", seg_name, n_t, n_c))
  
  # 부문 내 factor 수준이 1개인 경우 formula에서 제외
  safe_formula <- tryCatch({
    ind_lvl <- length(unique(sub_data$산업))
    reg_lvl <- length(unique(sub_data$지역))
    
    base_vars <- "log_자산 + log_매출 + log_부채 + log_자본금 + log_업력 + log_연구개발비 + log_수출 + 
    log_인건비 + log_노무비"

    ind_part  <- if (ind_lvl > 1) " + factor(산업)" else ""
    reg_part  <- if (reg_lvl > 1) " + factor(지역)" else ""
    as.formula(paste0("treat ~ ", base_vars, ind_part, reg_part))
  }, error = function(e) psm_formula)
  
  tryCatch({
    m <- matchit(
      safe_formula,
      data     = sub_data,
      method   = "nearest",
      distance = "logit",
      ratio    = 1,
      caliper  = 0.2,
      replace  = FALSE
    )
    md <- match.data(m)
    md$seg_name <- seg_name
    matchit_list_B[[seg_name]] <- m
    matched_list_B[[seg_name]] <- md
    cat(sprintf("    매칭 완료: 처치 %d, 통제 %d\n",
                sum(md$treat == 1), sum(md$treat == 0)))
  }, error = function(e) {
    cat(sprintf("    ⚠ %s 매칭 실패: %s\n", seg_name, e$message))
  })
}

# 부문별 매칭 데이터 통합
matched_data_B <- bind_rows(matched_list_B)

cat("\n[방법 B] 전체 매칭 결과:\n")
cat("  전체 처치:", sum(matched_data_B$treat == 1),
    "/ 전체 통제:", sum(matched_data_B$treat == 0), "\n")

# ==============================================================================
# 4. 균형 비교 (SMD: 방법 A vs 방법 B)
# ==============================================================================

calc_smd <- function(df) {
  vars <- c("log_자산", "log_매출", "log_부채", "log_자본금",
            # "log_종업원수" 삭제
            "log_업력", "영업이익", #"노동생산성",
            "log_연구개발비", "log_수출", "log_인건비", "log_노무비")    # ← 추가
  sapply(vars, function(v) {
    x1 <- df[[v]][df$treat == 1]; x0 <- df[[v]][df$treat == 0]
    (mean(x1, na.rm = TRUE) - mean(x0, na.rm = TRUE)) /
      sqrt((var(x1, na.rm = TRUE) + var(x0, na.rm = TRUE)) / 2)
  })
}

smd_before  <- calc_smd(analysis_data)
smd_A       <- calc_smd(matched_data_A)
smd_B       <- calc_smd(matched_data_B)

smd_compare <- data.frame(
  변수    = names(smd_before),
  매칭전  = round(smd_before, 4),
  방법A   = round(smd_A, 4),
  방법B   = round(smd_B, 4),
  A개선도 = round(abs(smd_before) - abs(smd_A), 4),
  B개선도 = round(abs(smd_before) - abs(smd_B), 4)
)

cat("\n\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
cat("  SMD 비교: 매칭전 / 방법A / 방법B\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
print(smd_compare)
cat("\n※ SMD < 0.1 이면 균형 달성\n")
cat("  방법A 기준 |SMD| 평균:", round(mean(abs(smd_A)), 4), "\n")
cat("  방법B 기준 |SMD| 평균:", round(mean(abs(smd_B)), 4), "\n")

# ==============================================================================
# 5. 부문별 SMD 비교
# ==============================================================================

cat("\n\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
cat("  부문별 SMD 비교 (방법 A vs 방법 B)\n")
cat(paste(rep("=", 70), collapse = ""), "\n")

seg_smd_list <- list()

for (s in names(seg_labels)) {
  seg_name <- seg_labels[s]
  
  sub_A <- matched_data_A %>% filter(seg == s)
  sub_B <- matched_data_B %>% filter(seg_name == !!seg_name)
  
  if (nrow(sub_A) > 0 & nrow(sub_B) > 0) {
    smd_a <- calc_smd(sub_A)
    smd_b <- calc_smd(sub_B)
    
    cat(sprintf("\n  ── %s ──\n", seg_name))
    cat(sprintf("    %-14s  %8s  %8s  %8s\n", "변수", "방법A", "방법B", "차이(A-B)"))
    for (v in names(smd_a)) {
      cat(sprintf("    %-14s  %8.4f  %8.4f  %8.4f\n",
                  v, smd_a[v], smd_b[v], smd_a[v] - smd_b[v]))
    }
    cat(sprintf("    %-14s  %8.4f  %8.4f\n",
                "│SMD│ 평균", mean(abs(smd_a)), mean(abs(smd_b))))
    
    seg_smd_list[[seg_name]] <- data.frame(
      부문    = seg_name,
      변수    = names(smd_a),
      방법A   = round(smd_a, 4),
      방법B   = round(smd_b, 4),
      개선_B  = round(abs(smd_a) - abs(smd_b), 4)
    )
  }
}

seg_smd_df <- bind_rows(seg_smd_list)

# ==============================================================================
# 6. 시각화 - Love Plot 비교
# ==============================================================================

cat("\n=== 시각화 생성 ===\n")

# 전체 Love Plot 비교
love_df <- data.frame(
  변수   = rep(names(smd_before), 3),
  방법   = rep(c("매칭전", "방법A(통합)", "방법B(부문별)"), each = length(smd_before)),
  SMD    = c(smd_before, smd_A, smd_B)
)

p_love <- ggplot(love_df, aes(x = abs(SMD), y = 변수, color = 방법, shape = 방법)) +
  geom_point(size = 4, position = position_dodge(width = 0.4)) +
  geom_vline(xintercept = 0.1, linetype = "dashed", color = "red", alpha = 0.6) +
  geom_vline(xintercept = 0,   color = "black") +
  scale_color_manual(values = c("매칭전" = "#888888",
                                "방법A(통합)"  = "#FF6B35",
                                "방법B(부문별)" = "#2196F3")) +
  scale_shape_manual(values = c("매칭전" = 16,
                                "방법A(통합)"  = 17,
                                "방법B(부문별)" = 15)) +
  labs(title  = "Love Plot: 방법A(통합) vs 방법B(부문별) 균형 비교",
       subtitle = "빨간 점선: SMD = 0.1 기준선",
       x = "│SMD│", y = "", color = "방법", shape = "방법") +
  theme_minimal(base_family = kfont) +
  theme(plot.title = element_text(face = "bold", size = 14),
        legend.position = "top")

ggsave("love_plot_AB_comparison.png", p_love, width = 12, height = 7, dpi = 300)
cat("✓ love_plot_AB_comparison.png\n")

# 부문별 SMD 비교 Facet
p_seg <- ggplot(seg_smd_df %>%
                  pivot_longer(cols = c(방법A, 방법B), names_to = "방법", values_to = "SMD"),
                aes(x = abs(SMD), y = 변수, color = 방법, shape = 방법)) +
  geom_point(size = 3, position = position_dodge(width = 0.5)) +
  geom_vline(xintercept = 0.1, linetype = "dashed", color = "red", alpha = 0.5) +
  facet_wrap(~부문, ncol = 3) +
  scale_color_manual(values = c("방법A" = "#FF6B35", "방법B" = "#2196F3")) +
  scale_shape_manual(values = c("방법A" = 17, "방법B" = 15)) +
  labs(title = "부문별 SMD 비교: 방법A vs 방법B",
       x = "│SMD│", y = "", color = "방법", shape = "방법") +
  theme_minimal(base_family = kfont) +
  theme(plot.title = element_text(face = "bold", size = 13), legend.position = "top")

ggsave("love_plot_by_segment.png", p_seg, width = 14, height = 6, dpi = 300)
cat("✓ love_plot_by_segment.png\n")

# 성향점수 분포 비교 (부문별 방법B)
ps_plots <- list()
for (seg_name in names(matchit_list_B)) {
  m     <- matchit_list_B[[seg_name]]
  md    <- matched_list_B[[seg_name]]
  sub_d <- analysis_data %>%
    filter(seg == names(seg_labels)[seg_labels == seg_name])
  
  ps_b <- sub_d %>%
    mutate(그룹 = ifelse(treat == 1, "처치", "통제"), 시점 = "매칭전")
  m_ps <- glm(psm_formula, data = ps_b, family = binomial())
  ps_b$PS <- predict(m_ps, type = "response")
  
  ps_a <- md %>%
    mutate(그룹 = ifelse(treat == 1, "처치", "통제"), 시점 = "매칭후", PS = distance)
  
  ps_plots[[seg_name]] <- ggplot(bind_rows(ps_b, ps_a),
                                 aes(x = PS, fill = 그룹)) +
    geom_histogram(alpha = 0.6, bins = 25, position = "identity", color = "white") +
    facet_wrap(~시점, ncol = 1) +
    scale_fill_manual(values = c("처치" = "#FF6B6B", "통제" = "#4ECDC4")) +
    labs(title = paste0("[방법B] ", seg_name), x = "성향점수", y = "", fill = "") +
    theme_minimal(base_family = kfont) +
    theme(plot.title = element_text(face = "bold", size = 11),
          legend.position = "top")
}

if (length(ps_plots) >= 3) {
  g_ps <- gridExtra::grid.arrange(grobs = ps_plots, ncol = 3,
                                  top = grid::textGrob("부문별 성향점수 분포 (방법B)",
                                                       gp = grid::gpar(fontsize = 14, fontface = "bold",
                                                                       fontfamily = kfont)))
  ggsave("ps_distribution_B.png", g_ps, width = 15, height = 8, dpi = 300)
  cat("✓ ps_distribution_B.png\n")
}

# ==============================================================================
# 7. 방법 선택 권고 출력
# ==============================================================================

avg_A <- mean(abs(smd_A))
avg_B <- mean(abs(smd_B))
winner <- if (avg_B < avg_A) "방법B (부문별 PSM)" else "방법A (통합 PSM)"

cat("\n\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
cat("  ★ 방법 비교 최종 요약\n")
cat(paste(rep("=", 70), collapse = ""), "\n")
cat(sprintf("  방법A 평균 │SMD│: %.4f\n", avg_A))
cat(sprintf("  방법B 평균 │SMD│: %.4f\n", avg_B))
cat(sprintf("  → 권장 방법: %s (SMD 기준)\n", winner))
cat("\n  부문별 매칭 샘플 수:\n")
for (seg_name in names(matched_list_B)) {
  md <- matched_list_B[[seg_name]]
  cat(sprintf("    %s: 처치 %d, 통제 %d\n",
              seg_name, sum(md$treat==1), sum(md$treat==0)))
}

# ==============================================================================
# 8. 방법 B 매칭 데이터 저장 (DID 분석용)
# ==============================================================================

cat("\n=== 결과 저장 ===\n")

write_xlsx(list(
  SMD_전체비교  = smd_compare,
  SMD_부문별    = seg_smd_df,
  방법A_요약    = data.frame(
    항목 = c("방법", "caliper", "전체처치N", "전체통제N", "평균SMD"),
    값   = c("통합 PSM + exact=~seg", "0.2",
            sum(matched_data_A$treat==1),
            sum(matched_data_A$treat==0),
            round(avg_A, 4))
  ),
  방법B_요약    = data.frame(
    항목 = c("방법", "caliper", "전체처치N", "전체통제N", "평균SMD"),
    값   = c("부문별 독립 PSM", "0.2",
            sum(matched_data_B$treat==1),
            sum(matched_data_B$treat==0),
            round(avg_B, 4))
  )
), "PSM_AB_comparison.xlsx")
cat("✓ PSM_AB_comparison.xlsx\n")

# 방법B 매칭 데이터 저장 (DID 분석에 사용 가능)
write_xlsx(matched_data_B, "matched_dataset_segPSM.xlsx")
cat("✓ matched_dataset_segPSM.xlsx\n")

cat("\n=== 분석 완료 ===\n")
cat("방법B 결과를 DID 분석에 쓰려면:\n")
cat("  #4 DID_Seg_log.R 상단에서 파일명을\n")
cat("  'matched_dataset_segPSM.xlsx' 로 변경하면 됩니다.\n")
