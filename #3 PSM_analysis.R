# ============================================================
# PSM (Propensity Score Matching) 분석
# Treatgroup vs Controlgroup 매칭 전후 비교
# 소재, 부품 ,장비를 구분해서 PSM - 20260306
# 결측치고려 : 종업원수를 대치 하기 위한 인건비, 노무비(회계 수치) - 필드워커, 사무실....labor cost, 
# R&D 투자 항목, 빈칸은 0으로 처리 할 수 있다. 
# ============================================================

# 필요한 패키지 설치
install.packages("MatchIt")      # PSM 분석
install.packages("cobalt")       # 균형 검정 및 시각화
install.packages("lmtest")       # 회귀분석
install.packages("sandwich")     # 로버스트 표준오차
install.packages("readxl")      # 엑셀 읽기
install.packages("dplyr")       # 데이터 조작
install.packages("tidyr")       # 데이터 정리
install.packages("writexl")     # 엑셀 쓰기
install.packages("stringr")     # 문자열 처리


# 필요한 패키지 로드
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(gridExtra)
library(writexl)
library(MatchIt)
library(cobalt)

#setwd("/Users/ultra/PSM-DID")
getwd()
# ============================================================
# 한글 폰트 설정
# ============================================================

cat("=== 한글 폰트 설정 중... ===\n")

if (Sys.info()["sysname"] == "Windows") {
  korean_font <- "Malgun Gothic"
} else if (Sys.info()["sysname"] == "Darwin") {
  korean_font <- "AppleGothic"
} else {
  korean_font <- "NanumGothic"
}

theme_set(theme_minimal(base_family = korean_font))
par(family = korean_font)

cat("한글 폰트 설정 완료!\n\n")

# ============================================================
# 데이터 읽기 및 전처리
# ============================================================

cat("=== 데이터 로드 중... ===\n")

# Text로 값을 읽어 와야 number 등으로 변환이 용이 하다. 
data <- read_excel("sobujang_integrated_with_valuesearch+Patent.xlsx", sheet=1,col_type = "text")
head(data)
names(data)

# Treatgroup과 Controlgroup만 필터링
analysis_data <- data %>%
  filter(group %in% c("Treatgroup", "Controlgroup"))

# 숫자 변환 함수
to_numeric <- function(x) {
  as.numeric(gsub(",", "", as.character(x)))
}


################################################################################
# 데이터에 포맷 및 구조를 확인 한다. 
# 이후 변수에 이름이 영어/짧게 해야 하는가 ? 
# 현재까지는 Value search, NTIS, Sobujang 원본을 그대로 사용하고.
# 추가는 영문 작명 한다.
################################################################################

# 각 열의 실제 이름과 데이터 구조 확인
names(analysis_data)

#===============================================================================
# 종업원수 결측치 처리 - 결측.....

# 년도별 결측치 개수 확인
analysis_data %>%
  group_by(group) %>%
  summarise(
    na2019_emp = sum(is.na(as.numeric(`2019/Annual S05000.종업원수`))),
    na2020_emp = sum(is.na(as.numeric(`2020/Annual S05000.종업원수`))),
    na2021_emp = sum(is.na(as.numeric(`2021/Annual S05000.종업원수`))),
    na2022_emp = sum(is.na(as.numeric(`2022/Annual S05000.종업원수`))),
    na2023_emp = sum(is.na(as.numeric(`2023/Annual S05000.종업원수`))),
    na2024_emp = sum(is.na(as.numeric(`2024/Annual S05000.종업원수`)))
  )


# 기업별 결측 패턴 확인
analysis_data <- analysis_data %>%
  mutate(
    emp2019 = as.numeric(`2019/Annual S05000.종업원수`),
    emp2020 = as.numeric(`2020/Annual S05000.종업원수`),
    emp2021 = as.numeric(`2021/Annual S05000.종업원수`),
    emp2022 = as.numeric(`2022/Annual S05000.종업원수`),
    emp2023 = as.numeric(`2023/Annual S05000.종업원수`),
    emp2024 = as.numeric(`2024/Annual S05000.종업원수`)
  ) %>%
  rowwise() %>%
  mutate(
    na_count = sum(is.na(c(emp2019, emp2020, emp2021, emp2022, emp2023, emp2024))),
    has_any_data = na_count < 6
  ) %>%
  ungroup()

# 결측 개수별 기업 수
analysis_data %>%
  count(na_count) %>%
  mutate(pct = n / sum(n) * 100)

#===============================================================================
# 변환 전 결측치 확인하여 데이터 확인 함.
analysis_data %>%
  summarise(
    n_total = n(),
    na_자산 = sum(is.na(`2019/Annual S15000.자산총계`)),
    na_매출 = sum(is.na(`2019/Annual S21100.총매출액`)),
    na_부채 = sum(is.na(`2019/Annual S18000.부채총계`)),
    na_자본금 = sum(is.na(`2019/Annual S18100.자본금`)),
    #na_자본총 = sum(is.na(`2019/Annual S18900.자본총계`)), 음수가 발생
    na_종업원 = sum(is.na(`2019/Annual S05000.종업원수`)),
    na_업력 = sum(is.na(`age`)),
    na_KSIC_표준 = sum(is.na(`KSIC_mid_code`)),
    na_지역 = sum(is.na(`region`)),
    na_노동생산성= sum(is.na(`labor_prod2019`)),
    na_영업이익= sum(is.na(`2019/Annual S25000.영업이익(손실)`)),
    na_연구개발비 = sum(is.na(`2019/Annual 692084.연구개발비용-연구개발비용계`)),
    na_수출 = sum(is.na(`2019/Annual S21195.[수출]`))
  )

# 변수 변환 및 처치 변수 생성
analysis_data <- analysis_data %>%
  mutate(
    # 숫자형 변환
    자산 = to_numeric(`2019/Annual S15000.자산총계`),
    매출 = to_numeric(`2019/Annual S21100.총매출액`),
    부채 = to_numeric(`2019/Annual S18000.부채총계`),
    자본금 = to_numeric(`2019/Annual S18100.자본금`),
    #자본총 = to_numeric(`2019/Annual S18900.자본총계`), 음수 발생
    종업원수 = to_numeric(`2019/Annual S05000.종업원수`),
    업력 = to_numeric(`age`),
    산업 = to_numeric(`KSIC_mid_code`),
    지역 = to_numeric(`region`),
    영업이익 = to_numeric(`2019/Annual S25000.영업이익(손실)`),
    노동생산성 = to_numeric(`labor_prod2019`),
    
    연구개발비 = ifelse(is.na(to_numeric(`2019/Annual 692084.연구개발비용-연구개발비용계`)), 0,
                   to_numeric(`2019/Annual 692084.연구개발비용-연구개발비용계`)),
    수출금액   = ifelse(is.na(to_numeric(`2019/Annual S21195.[수출]`)), 0,
                    to_numeric(`2019/Annual S21195.[수출]`)),
    # 처치 변수 (1: Treatgroup, 0: Controlgroup)
    treat = ifelse(group == "Treatgroup", 1, 0),
    
    # 로그 변환 (매칭용)
    log_자산 = log(자산 + 1),
    log_매출 = log(매출 + 1),
    log_부채 = log(부채 +1),
    log_자본금 = log(자본금 + 1),
    #log_자본총 = log(자본총 + 1), 음수 발생
    log_종업원수 = log(종업원수 + 1),
    log_업력 = log(업력 + 1),
    log_산업 = log(산업 + 1),
    log_지역 = log(지역 + 1),
    log_노동생산성 = log(노동생산성 + 1),
    log_연구개발비 = log(연구개발비 + 1),
    log_수출       = log(수출금액 + 1)
    # 영업 이익은 음수가 있어 log 치환 안됨.
  )%>%
  filter(!is.na(자산) & !is.na(매출) & !is.na(부채) & !is.na(자본금) & 
           !is.na(종업원수) & !is.na(업력) &!is.na(영업이익))


cat("데이터 준비 완료!\n")
cat("전체 샘플:", nrow(analysis_data), "\n")
cat("Treatgroup:", sum(analysis_data$treat == 1), "\n")
cat("Controlgroup:", sum(analysis_data$treat == 0), "\n\n")

# ============================================================
# 1. 매칭 전 기초통계량 및 균형 검정
# ============================================================

cat("=== 매칭 전 기초통계량 ===\n")

before_stats <- analysis_data %>%
  group_by(treat) %>%
  summarise(
    N = n(),
    자산_평균 = mean(자산, na.rm = TRUE),
    매출_평균 = mean(매출, na.rm = TRUE),
    부채_평균 = mean(부채, na.rm = TRUE),
    자본금_평균 = mean(자본금, na.rm = TRUE),
    #자본총_평균 = mean(자본총, na.rm = TRUE),
    종업원수_평균 = mean(종업원수, na.rm = TRUE),
    업력_평균 = mean(업력, na.rm = TRUE),
    영업이익_평균 = mean(영업이익, na.rm = TRUE),
    노동생산성_평균 = mean(노동생산성, na.rm = TRUE),
    지역_평균 = mean(지역, na.rm = TRUE),
    산업_평균 = mean(산업, na.rm = TRUE),
    연구개발비_평균 = mean(연구개발비, na.rm = TRUE),
    수출금액_평균   = mean(수출금액, na.rm = TRUE)
  )

print(before_stats)

# 표준화 차이 계산 (매칭 전) 
before_balance <- analysis_data %>%
  summarise(
    자산_SMD = (mean(log_자산[treat==1], na.rm = TRUE) - mean(log_자산[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_자산[treat==1], na.rm = TRUE)^2 + sd(log_자산[treat==0], na.rm = TRUE)^2) / 2),
    매출_SMD = (mean(log_매출[treat==1], na.rm = TRUE) - mean(log_매출[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_매출[treat==1], na.rm = TRUE)^2 + sd(log_매출[treat==0], na.rm = TRUE)^2) / 2),
    부채_SMD = (mean(log_부채[treat==1], na.rm = TRUE) - mean(log_부채[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_부채[treat==1], na.rm = TRUE)^2 + sd(log_부채[treat==0], na.rm = TRUE)^2) / 2),
    자본금_SMD = (mean(log_자본금[treat==1], na.rm = TRUE) - mean(log_자본금[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_자본금[treat==1], na.rm = TRUE)^2 + sd(log_자본금[treat==0], na.rm = TRUE)^2) / 2),
    #자본총_SMD = (mean(log_자본총[treat==1], na.rm = TRUE) - mean(log_자본총[treat==0], na.rm = TRUE)) / 
    #  sqrt((sd(log_자본총[treat==1], na.rm = TRUE)^2 + sd(log_자본총[treat==0], na.rm = TRUE)^2) / 2),
    종업원_SMD = (mean(log_종업원수[treat==1], na.rm = TRUE) - mean(log_종업원수[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_종업원수[treat==1], na.rm = TRUE)^2 + sd(log_종업원수[treat==0], na.rm = TRUE)^2) / 2),
    업력_SMD = (mean(log_업력[treat==1], na.rm = TRUE) - mean(log_업력[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_업력[treat==1], na.rm = TRUE)^2 + sd(log_업력[treat==0], na.rm = TRUE)^2) / 2),
    영업이익_SMD = (mean(영업이익[treat==1], na.rm = TRUE) - mean(영업이익[treat==0], na.rm = TRUE)) / 
      sqrt((sd(영업이익[treat==1], na.rm = TRUE)^2 + sd(영업이익[treat==0], na.rm = TRUE)^2) / 2),
    노동생산성_SMD = (mean(노동생산성[treat==1], na.rm = TRUE) - mean(노동생산성[treat==0], na.rm = TRUE)) / 
      sqrt((sd(노동생산성[treat==1], na.rm = TRUE)^2 + sd(노동생산성[treat==0], na.rm = TRUE)^2) / 2),
    지역_SMD = (mean(지역[treat==1], na.rm = TRUE) - mean(지역[treat==0], na.rm = TRUE)) / 
      sqrt((sd(지역[treat==1], na.rm = TRUE)^2 + sd(지역[treat==0], na.rm = TRUE)^2) / 2),
    산업_SMD = (mean(산업[treat==1], na.rm = TRUE) - mean(산업[treat==0], na.rm = TRUE)) / 
      sqrt((sd(산업[treat==1], na.rm = TRUE)^2 + sd(산업[treat==0], na.rm = TRUE)^2) / 2),
    연구개발비_SMD = (mean(log_연구개발비[treat==1], na.rm=TRUE) - mean(log_연구개발비[treat==0], na.rm=TRUE)) /
      sqrt((sd(log_연구개발비[treat==1], na.rm=TRUE)^2 + sd(log_연구개발비[treat==0], na.rm=TRUE)^2) / 2),
    수출_SMD = (mean(log_수출[treat==1], na.rm=TRUE) - mean(log_수출[treat==0], na.rm=TRUE)) /
      sqrt((sd(log_수출[treat==1], na.rm=TRUE)^2 + sd(log_수출[treat==0], na.rm=TRUE)^2) / 2)
  )

cat("\n=== 매칭 전 표준화 차이 (SMD) ===\n")
cat("SMD > 0.1이면 불균형\n")
print(before_balance)

# ============================================================
# 2. PSM 수행
# ============================================================

#names(analysis_data)

cat("\n=== PSM 수행 중... ===\n")

# 성향점수 모델 (로지스틱 회귀), 표준 산업 분류 추가 
psm_formula <- treat ~ log_자산 + log_매출 + log_부채 + log_자본금 +  
  log_종업원수 + log_업력 + factor(산업) + factor(지역) + log_연구개발비 + log_수출

# 로지스틱 회귀 모델 적합
psm_model <- glm(psm_formula, 
                 data = analysis_data, 
                 family = binomial(link = "logit"))

# 성향점수 추출
propensity_scores <- predict(psm_model, type = "response")
print(summary(propensity_scores))

# MatchIt을 사용한 매칭
# 소재, 부품, 장비에 정확 매칭을 위해서.
matched <- matchit(
  psm_formula,
  data = analysis_data,
  method = "nearest",
  distance = "logit",
  ratio = 1,
  caliper = 0.2,
  replace = FALSE,
  exact = ~seg
)

# ==============================================================================
# 매칭 할때 소재 = 0, 부품 = 1, 장비 = 2 를따로 매칭 할까 ? 
# 매칭 할때, Segment (0,1,2)를 넣어 줄까 ? 
# Implication은 ??
# 표준 산업 분류를 넣어 줘야 함. 
# ==============================================================================

cat("매칭 완료!\n")
print(summary(matched))

# 매칭된 데이터 추출
matched_data <- match.data(matched)

cat("\n매칭 후 샘플 크기:\n")
cat("Treatgroup:", sum(matched_data$treat == 1), "\n")
cat("Controlgroup:", sum(matched_data$treat == 0), "\n\n")

# ============================================================
# 3. 매칭 후 균형 검정
# ============================================================

cat("=== 매칭 후 기초통계량 ===\n")

after_stats <- matched_data %>%
  group_by(treat) %>%
  summarise(
    N = n(),
    자산_평균 = mean(자산, na.rm = TRUE),
    매출_평균 = mean(매출, na.rm = TRUE),
    부채_평균 = mean(부채, na.rm = TRUE),
    자본금_평균 = mean(자본금, na.rm = TRUE),
    #자본총_평균 = mean(자본총, na.rm = TRUE),
    종업원수_평균 = mean(종업원수, na.rm = TRUE),
    업력_평균 = mean(업력, na.rm = TRUE),
    영업이익_평균 = mean(영업이익, na.rm = TRUE),
    노동생산성_평균 = mean(노동생산성, na.rm = TRUE),
    지역_평균 = mean(지역, na.rm = TRUE),
    산업_평균 = mean(산업, na.rm = TRUE),
    연구개발비_평균 = mean(연구개발비, na.rm = TRUE),
    수출금액_평균   = mean(수출금액, na.rm = TRUE)
  )


print(after_stats)

# 표준화 차이 계산 (매칭 후)
after_balance <- matched_data %>%
  summarise(
    자산_SMD = (mean(log_자산[treat==1], na.rm = TRUE) - mean(log_자산[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_자산[treat==1], na.rm = TRUE)^2 + sd(log_자산[treat==0], na.rm = TRUE)^2) / 2),
    매출_SMD = (mean(log_매출[treat==1], na.rm = TRUE) - mean(log_매출[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_매출[treat==1], na.rm = TRUE)^2 + sd(log_매출[treat==0], na.rm = TRUE)^2) / 2),
    부채_SMD = (mean(log_부채[treat==1], na.rm = TRUE) - mean(log_부채[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_부채[treat==1], na.rm = TRUE)^2 + sd(log_부채[treat==0], na.rm = TRUE)^2) / 2),
    자본금_SMD = (mean(log_자본금[treat==1], na.rm = TRUE) - mean(log_자본금[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_자본금[treat==1], na.rm = TRUE)^2 + sd(log_자본금[treat==0], na.rm = TRUE)^2) / 2),
    #자본총_SMD = (mean(log_자본총[treat==1], na.rm = TRUE) - mean(log_자본총[treat==0], na.rm = TRUE)) / 
    #  sqrt((sd(log_자본총[treat==1], na.rm = TRUE)^2 + sd(log_자본총[treat==0], na.rm = TRUE)^2) / 2),
    종업원_SMD = (mean(log_종업원수[treat==1], na.rm = TRUE) - mean(log_종업원수[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_종업원수[treat==1], na.rm = TRUE)^2 + sd(log_종업원수[treat==0], na.rm = TRUE)^2) / 2),
    업력_SMD = (mean(log_업력[treat==1], na.rm = TRUE) - mean(log_업력[treat==0], na.rm = TRUE)) / 
      sqrt((sd(log_업력[treat==1], na.rm = TRUE)^2 + sd(log_업력[treat==0], na.rm = TRUE)^2) / 2),
    영업이익_SMD = (mean(영업이익[treat==1], na.rm = TRUE) - mean(영업이익[treat==0], na.rm = TRUE)) / 
      sqrt((sd(영업이익[treat==1], na.rm = TRUE)^2 + sd(영업이익[treat==0], na.rm = TRUE)^2) / 2),
    노동생산성_SMD = (mean(노동생산성[treat==1], na.rm = TRUE) - mean(노동생산성[treat==0], na.rm = TRUE)) / 
      sqrt((sd(노동생산성[treat==1], na.rm = TRUE)^2 + sd(노동생산성[treat==0], na.rm = TRUE)^2) / 2),
    지역_SMD = (mean(지역[treat==1], na.rm = TRUE) - mean(지역[treat==0], na.rm = TRUE)) / 
      sqrt((sd(지역[treat==1], na.rm = TRUE)^2 + sd(지역[treat==0], na.rm = TRUE)^2) / 2),
    산업_SMD = (mean(산업[treat==1], na.rm = TRUE) - mean(산업[treat==0], na.rm = TRUE)) / 
      sqrt((sd(산업[treat==1], na.rm = TRUE)^2 + sd(산업[treat==0], na.rm = TRUE)^2) / 2),
    연구개발비_SMD = (mean(log_연구개발비[treat==1], na.rm=TRUE) - mean(log_연구개발비[treat==0], na.rm=TRUE)) /
      sqrt((sd(log_연구개발비[treat==1], na.rm=TRUE)^2 + sd(log_연구개발비[treat==0], na.rm=TRUE)^2) / 2),
    수출_SMD = (mean(log_수출[treat==1], na.rm=TRUE) - mean(log_수출[treat==0], na.rm=TRUE)) /
      sqrt((sd(log_수출[treat==1], na.rm=TRUE)^2 + sd(log_수출[treat==0], na.rm=TRUE)^2) / 2)
  )

cat("\n=== 매칭 후 표준화 차이 (SMD) ===\n")
print(after_balance)

# ============================================================
# 4. 균형 검정 시각화 (Love Plot)
# ============================================================

balance_df <- data.frame(
  변수 = c("log(자산)", "log(매출)", "log(부채)", "log(자본금)",
         "log(종업원수)", "log(업력)", "영업이익", "노동생산성", "지역", "산업",
         "log(연구개발비)", "log(수출)"), 
  매칭전 = c(before_balance$자산_SMD, before_balance$매출_SMD, before_balance$부채_SMD,
          before_balance$자본금_SMD, before_balance$종업원_SMD, 
          before_balance$업력_SMD, before_balance$영업이익_SMD, before_balance$노동생산성_SMD,
          before_balance$지역_SMD, before_balance$산업_SMD,
          before_balance$연구개발비_SMD, before_balance$수출_SMD),   # ← 추가
  매칭후 = c(after_balance$자산_SMD, after_balance$매출_SMD, after_balance$부채_SMD,
          after_balance$자본금_SMD, after_balance$종업원_SMD, 
          after_balance$업력_SMD, after_balance$영업이익_SMD, after_balance$노동생산성_SMD,
          after_balance$지역_SMD, after_balance$산업_SMD,
          after_balance$연구개발비_SMD, after_balance$수출_SMD)    # ← 추가
) %>%
  pivot_longer(cols = c(매칭전, 매칭후), names_to = "시점", values_to = "SMD")

love_plot <- ggplot(balance_df, aes(x = abs(SMD), y = 변수, color = 시점, shape = 시점)) +
  geom_point(size = 4) +
  geom_vline(xintercept = 0.1, linetype = "dashed", color = "red", alpha = 0.5) +
  geom_vline(xintercept = 0, color = "black") +
  scale_color_manual(values = c("매칭전" = "#FF6B6B", "매칭후" = "#4ECDC4")) +
  scale_shape_manual(values = c("매칭전" = 16, "매칭후" = 17)) +
  labs(title = "매칭 전후 공변량 균형 비교 (Love Plot)",
       subtitle = "SMD < 0.1 이면 균형 달성",
       x = "표준화 평균 차이 (SMD)",
       y = "") +
  theme_minimal()

ggsave("love_plot.png", love_plot, width = 10, height = 8, dpi = 300)
# ============================================================
# 5. 성향점수 분포 비교
# ============================================================

cat("\n=== 성향점수 분포 그래프 생성 중... ===\n")

# 매칭 전 성향점수 분포
ps_before <- analysis_data %>%
  mutate(
    그룹 = ifelse(treat == 1, "Treatgroup", "Controlgroup"),
    시점 = "매칭 전"
  )

# 매칭 전 성향점수 계산
ps_model <- glm(psm_formula, data = ps_before, family = binomial())
ps_before$propensity_score <- predict(ps_model, type = "response")

# 매칭 후 성향점수
ps_after <- matched_data %>%
  mutate(
    그룹 = ifelse(treat == 1, "Treatgroup", "Controlgroup"),
    시점 = "매칭 후",
    propensity_score = distance
  )

# 통합
ps_combined <- bind_rows(ps_before, ps_after)

# 성향점수 분포 그래프
ps_plot <- ggplot(ps_combined, aes(x = propensity_score, fill = 그룹)) +
  geom_histogram(alpha = 0.6, position = "identity", bins = 30, color = "white") +
  facet_wrap(~시점, ncol = 1) +
  scale_fill_manual(values = c("Treatgroup" = "#FF6B6B", "Controlgroup" = "#4ECDC4")) +
  labs(title = "성향점수 분포 비교",
       subtitle = "매칭 전후",
       x = "성향점수 (Propensity Score)",
       y = "빈도",
       fill = "그룹") +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5, face = "bold", size = 14),
        plot.subtitle = element_text(hjust = 0.5, size = 10),
        legend.position = "top",
        strip.text = element_text(size = 12, face = "bold"))

ggsave("propensity_score_distribution.png", ps_plot, width = 10, height = 8, dpi = 300)
cat("✓ 저장 완료: propensity_score_distribution.png\n")

# ============================================================
# 6. 처치효과 추정 (ATT: Average Treatment effect on the Treated)
# ============================================================

test_vars <- c("자산", "매출", "부채", "자본금", "종업원수", "업력",
               "영업이익", "노동생산성", "지역", "산업",
               "연구개발비", "수출금액")   # ← 추가

# 매칭 전후 평균 차이 계산
before_att <- analysis_data %>%
  group_by(treat) %>%
  summarise(across(all_of(test_vars), ~mean(.x, na.rm = TRUE)))

before_diff <- data.frame(
  변수 = test_vars,
  매칭전_차이 = sapply(test_vars, function(v) {
    before_att[[v]][before_att$treat == 1] - before_att[[v]][before_att$treat == 0]
  })
)

after_att <- matched_data %>%
  group_by(treat) %>%
  summarise(across(all_of(test_vars), ~mean(.x, na.rm = TRUE)))

after_diff <- data.frame(
  변수 = test_vars,
  매칭후_차이 = sapply(test_vars, function(v) {
    after_att[[v]][after_att$treat == 1] - after_att[[v]][after_att$treat == 0]
  })
)

att_comparison <- before_diff %>%
  left_join(after_diff, by = "변수")

print(att_comparison)

# ============================================================
# 7. t-검정 (매칭 전후)
# ============================================================

cat("\n=== t-검정 결과 비교 ===\n")

# 매칭 전 t-검정
before_ttest <- data.frame(
  변수 = character(),
  t통계량_전 = numeric(),
  p값_전 = numeric(),
  유의성_전 = character(),
  stringsAsFactors = FALSE
)

for (var in test_vars) {
  treat_vals <- analysis_data %>% filter(treat == 1) %>% pull(!!sym(var))
  control_vals <- analysis_data %>% filter(treat == 0) %>% pull(!!sym(var))
  
  treat_vals <- treat_vals[!is.na(treat_vals)]
  control_vals <- control_vals[!is.na(control_vals)]
  
  if (length(treat_vals) > 0 && length(control_vals) > 0) {
    test <- t.test(treat_vals, control_vals)
    sig <- ifelse(test$p.value < 0.001, "***",
                  ifelse(test$p.value < 0.01, "**",
                         ifelse(test$p.value < 0.05, "*", "ns")))
    
    before_ttest <- rbind(before_ttest, data.frame(
      변수 = var,
      t통계량_전 = test$statistic,
      p값_전 = test$p.value,
      유의성_전 = sig
    ))
  }
}

# 매칭 후 t-검정
after_ttest <- data.frame(
  변수 = character(),
  t통계량_후 = numeric(),
  p값_후 = numeric(),
  유의성_후 = character(),
  stringsAsFactors = FALSE
)

for (var in test_vars) {
  treat_vals <- matched_data %>% filter(treat == 1) %>% pull(!!sym(var))
  control_vals <- matched_data %>% filter(treat == 0) %>% pull(!!sym(var))
  
  treat_vals <- treat_vals[!is.na(treat_vals)]
  control_vals <- control_vals[!is.na(control_vals)]
  
  if (length(treat_vals) > 0 && length(control_vals) > 0) {
    test <- t.test(treat_vals, control_vals)
    sig <- ifelse(test$p.value < 0.001, "***",
                  ifelse(test$p.value < 0.01, "**",
                         ifelse(test$p.value < 0.05, "*", "ns")))
    
    after_ttest <- rbind(after_ttest, data.frame(
      변수 = var,
      t통계량_후 = test$statistic,
      p값_후 = test$p.value,
      유의성_후 = sig
    ))
  }
}

ttest_comparison <- before_ttest %>%
  left_join(after_ttest, by = "변수")

print(ttest_comparison)

# ============================================================
# 8. 결과 저장
# ============================================================

cat("\n=== 결과 저장 중... ===\n")

# 종합 결과
psm_results <- list(
  "매칭전_기초통계" = before_stats,
  "매칭후_기초통계" = after_stats,
  "균형검정_매칭전" = before_balance,
  "균형검정_매칭후" = after_balance,
  "처치효과_비교" = att_comparison,
  "t검정_비교" = ttest_comparison,
  "매칭요약" = data.frame(
    항목 = c("매칭 방법", "매칭 비율", "Caliper", "매칭 전 N", "매칭 후 Treat N", "매칭 후 Control N"),
    값 = c("Nearest Neighbor", "1:1", "0.2", 
          nrow(analysis_data),
          sum(matched_data$treat == 1),
          sum(matched_data$treat == 0))
  )
)

write_xlsx(psm_results, "PSM_analysis_results_patent.xlsx")
cat("✓ 저장 완료: PSM_analysis_results_patent.xlsx\n")

# 매칭된 데이터 저장
write_xlsx(matched_data, "matched_dataset_patent.xlsx")
cat("✓ 저장 완료: matched_dataset_patent.xlsx\n")


# ============================================================
# 주요 결과 요약 출력
# ============================================================

cat("\n=== 주요 결과 요약 ===\n\n")

cat("2. 균형 개선 (SMD: 매칭전 → 매칭후):\n")
cat("   - 자산:", round(before_balance$자산_SMD, 3), "→", round(after_balance$자산_SMD, 3), "\n")
cat("   - 매출:", round(before_balance$매출_SMD, 3), "→", round(after_balance$매출_SMD, 3), "\n")
cat("   - 부채:", round(before_balance$부채_SMD, 3), "→", round(after_balance$부채_SMD, 3), "\n")
cat("   - 자본금:", round(before_balance$자본금_SMD, 3), "→", round(after_balance$자본금_SMD, 3), "\n")
#cat("   - 자본총:", round(before_balance$자본총_SMD, 3), "→", round(after_balance$자본총_SMD, 3), "\n")
cat("   - 종업원수:", round(before_balance$종업원_SMD, 3), "→", round(after_balance$종업원_SMD, 3), "\n")
cat("   - 업력:", round(before_balance$업력_SMD, 3), "→", round(after_balance$업력_SMD, 3), "\n")
cat("   - 영업이익:", round(before_balance$영업이익_SMD, 3), "→", round(after_balance$영업이익_SMD, 3), "\n")
cat("   - 노동생산성:", round(before_balance$노동생산성_SMD, 3), "→", round(after_balance$노동생산성_SMD, 3), "\n")
cat("   - 지역:", round(before_balance$지역_SMD, 3), "→", round(after_balance$지역_SMD, 3), "\n")
cat("   - 산업:", round(before_balance$산업_SMD, 3), "→", round(after_balance$산업_SMD, 3), "\n")
cat("   - 연구개발비:", round(before_balance$연구개발비_SMD, 3), "→", round(after_balance$연구개발비_SMD, 3), "\n")
cat("   - 수출:",       round(before_balance$수출_SMD, 3),       "→", round(after_balance$수출_SMD, 3), "\n")

