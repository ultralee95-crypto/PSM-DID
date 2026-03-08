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

#setwd("~/Downloads/Venture")
getwd()

#===============================================================================
# Valuesearch 기업 데이터 
#===============================================================================
# 종업원 등 
valuesearch1 <- read_excel("VALUESearch_20251213_Region.xlsx", skip = 1, col_types = "text")
# 초기 데이터 
valuesearch2 <- read_excel("VALUESearch20251207.xlsx", skip = 1, col_types = "text")
# KSIC 중분류 및 기업 주소 등
valuesearch3 <- read_excel("VALUESearch 20251221_KSIC_Mid.xlsx", skip = 1, col_types = "text")
# R&D 비용 80-85
valuesearch4 <- read_excel("ValueSearch_R&D_692080-85.xlsx", skip = 1, col_types = "text")
# R&D 비용 86-88: 연구개발비/매출액 비율  
valuesearch5 <- read_excel("ValueSearch_R&D_692086-88.xlsx", skip = 1, col_types = "text")

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


