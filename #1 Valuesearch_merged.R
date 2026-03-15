install.packages("readxl")      # 엑셀 읽기
install.packages("dplyr")       # 데이터 조작
install.packages("tidyr")       # 데이터 정리
install.packages("writexl")     # 엑셀 쓰기
install.packages("stringr")     # 문자열 처리
install.packages("tidyverse")
install.packages(c("usethis", "gitcreds"))
# ============================================================
# 소부장 기업 정부 투자 여부 확인 (간결 버전)
# ============================================================

library(tidyverse)
library(readxl)
library(dplyr)
library(writexl)
library(stringr)

setwd("/Users/ultra/PSM-DID")
getwd()

#===============================================================================
# Valuesearch 기업 데이터 
# 평행 추세 분석을 위해서 모든 데이터에 2018-2019년도 데이터 포함.
# 2018 ~ 2024년도 데이터 
#===============================================================================
# 1 업체 이름 주소 사업자 자산 총계
valuesearch1 <- read_excel("1_VALUESearch_주소_자산총계20260315.xlsx", skip = 1, col_types = "text")
# 2 총수익  총매출에서 이름이 변경됨. 
valuesearch2 <- read_excel("2_VALUESearch_총수익20260315.xlsx", skip = 1, col_types = "text")
# 3 연구개발비용계 
valuesearch3 <- read_excel("3_VALUESearch_연구개발비용계20260315.xlsx", skip = 1, col_types = "text")
# 4 연구 개발 비용 세부 항목 
valuesearch4 <- read_excel("4_ValueSearch_연구개발비_비율20260315.xlsx", skip = 1, col_types = "text")
# 5 인건비 노무비 
valuesearch5 <- read_excel("5_ValueSearch_인건비_노무비20260315.xlsx", skip = 1, col_types = "text")

head(valuesearch1)
names(valuesearch1)
head(valuesearch2)
names(valuesearch2)
head(valuesearch3)
names(valuesearch3)
head(valuesearch4)
names(valuesearch4)
head(valuesearch5)
names(valuesearch5)

# 공통 컬럼 확인
common_cols_v2 <- intersect(names(valuesearch1), names(valuesearch2))
common_cols_v3 <- intersect(names(valuesearch1), names(valuesearch3))

cat("=== valuesearch1 & valuesearch2 공통 컬럼 ===\n")
print(common_cols_v2)

cat("\n=== valuesearch1 & valuesearch3 공통 컬럼 ===\n")
print(common_cols_v3)


# 1단계: valuesearch2에서 중복 컬럼 제거 후 병합
valuesearch_merged <- valuesearch1 %>%
  left_join(
    valuesearch2 %>%
      select(-all_of(setdiff(common_cols_v2, "691020.사업자번호"))) %>%  # 키 제외 공통컬럼 제거
      distinct(`691020.사업자번호`, .keep_all = TRUE),
    by = "691020.사업자번호"
  )

names(valuesearch_merged)

# 2단계: valuesearch3에서 중복 컬럼 제거 후 병합
common_cols_merged_v3 <- intersect(names(valuesearch_merged), names(valuesearch3))

print(common_cols_merged_v3)

valuesearch_final <- valuesearch_merged %>%
  left_join(
    valuesearch3 %>%
      select(-all_of(setdiff(common_cols_merged_v3, "691020.사업자번호"))) %>%  # 키 제외 공통컬럼 제거
      distinct(`691020.사업자번호`, .keep_all = TRUE),
    by = "691020.사업자번호"
  )

# 3단계: valuesearch4 병합
common_cols_merged_v4 <- intersect(names(valuesearch_final), names(valuesearch4))

valuesearch_final <- valuesearch_final %>%
  left_join(
    valuesearch4 %>%
      select(-all_of(setdiff(common_cols_merged_v4, "691020.사업자번호"))) %>%
      distinct(`691020.사업자번호`, .keep_all = TRUE),
    by = "691020.사업자번호"
  )

# 4단계: valuesearch5 병합
common_cols_merged_v5 <- intersect(names(valuesearch_final), names(valuesearch5))

valuesearch_final <- valuesearch_final %>%
  left_join(
    valuesearch5 %>%
      select(-all_of(setdiff(common_cols_merged_v5, "691020.사업자번호"))) %>%
      distinct(`691020.사업자번호`, .keep_all = TRUE),
    by = "691020.사업자번호"
  )
names(valuesearch_final)

# 결과 확인
cat("\n=== 병합 결과 ===\n")
cat("valuesearch1 행:", nrow(valuesearch1), "/ 컬럼:", ncol(valuesearch1), "\n")
cat("valuesearch2 행:", nrow(valuesearch2), "/ 컬럼:", ncol(valuesearch2), "\n")
cat("valuesearch3 행:", nrow(valuesearch3), "/ 컬럼:", ncol(valuesearch3), "\n")
cat("valuesearch4 행:", nrow(valuesearch4), "/ 컬럼:", ncol(valuesearch4), "\n")  # ← 추가
cat("valuesearch5 행:", nrow(valuesearch5), "/ 컬럼:", ncol(valuesearch5), "\n")  # ← 추가
cat("valuesearch_final 행:", nrow(valuesearch_final), "/ 컬럼:", ncol(valuesearch_final), "\n")


# 저장
write_xlsx(valuesearch_final, "valuesearch_merged.xlsx")


