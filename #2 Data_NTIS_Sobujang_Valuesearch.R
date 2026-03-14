install.packages("readxl")      # 엑셀 읽기
install.packages("dplyr")       # 데이터 조작
install.packages("tidyr")       # 데이터 정리
install.packages("writexl")     # 엑셀 쓰기
install.packages("stringr")     # 문자열 처리
install.packages("tidyverse")
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

# Valuesearch 기업 데이터 
#===============================================================================
valuesearch <- read_excel("valuesearch_merged.xlsx" , col_types = "text")
head(valuesearch)
names(valuesearch)

# 소부장데이터 읽기
#===============================================================================
sobujang <- read_excel("sobujang_all_companies.xlsx")

#NTIS - 소재부품장비 정부 투자
#===============================================================================
# 소재부품장비 투자 
ntis <- read_excel("NTIS_corp_list.xls", skip = 1)
#NTIS - 소재부품기술 투자  
ntis1 <- read_excel("NTIS_soje_bupum_1.xls", skip = 1)
ntis2 <- read_excel("NTIS_soje_bupum_2.xls", skip = 1)
ntis3 <- read_excel("NTIS_soje_bupum_3.xls", skip = 1)
ntis4 <- read_excel("NTIS_soje_bupum_4.xls", skip = 1)
ntis5 <- read_excel("NTIS_soje_bupum_5.xls", skip = 1)
ntis6 <- read_excel("NTIS_soje_bupum_6.xls", skip = 1)
ntis7 <- read_excel("NTIS_soje_bupum_7.xls", skip = 1)

#단순히 행으로 합치기 (같은 구조일 때)
ntis <- bind_rows(ntis, ntis1, ntis2, ntis3, ntis4, ntis5, ntis7)

# 결과 확인
print(paste("총 행 수:", nrow(ntis)))
print(paste("총 열 수:", ncol(ntis)))

# 결과 저장
write_xlsx(ntis, "ntis_integration.xlsx")

# 2020,2021,2022년도 이상만 필터링
ntis <- ntis %>%
  filter(ntis$기준년도 %in% c(2020,2021,2022))  # 칼럼명에 맞춰 조정 필요

# 결과 확인
print(paste("총 행 수:", nrow(ntis)))
print(paste("총 열 수:", ncol(ntis)))

# 결과 저장
write_xlsx(ntis, "ntis_integration_2020-2022.xlsx")

#사업자 번호 매칭 
#==============================================================================

# 사업자번호 정리 (공백, 하이픈 제거)
sobujang$사업자번호_clean <- gsub("[[:space:]-]", "", sobujang$사업자번호)
ntis$사업자등록번호_clean <- gsub("[[:space:]-]", "", ntis$사업자등록번호)
valuesearch$사업자번호_clean <- gsub("[[:space:]-]", "", valuesearch$`691020.사업자번호`)

# 정부 투자 받은 기업 목록
#government_funded <- unique(ntis$사업자등록번호_clean[!is.na(ntis$사업자등록번호_clean)])
#연도별 정부 투자 기업 목록 생성
government_funded_2020 <- unique(ntis$사업자등록번호_clean[ntis$기준년도 == 2020 & !is.na(ntis$사업자등록번호_clean)])
government_funded_2021 <- unique(ntis$사업자등록번호_clean[ntis$기준년도 == 2021 & !is.na(ntis$사업자등록번호_clean)])
government_funded_2022 <- unique(ntis$사업자등록번호_clean[ntis$기준년도 == 2022 & !is.na(ntis$사업자등록번호_clean)])


# 정부투자여부 컬럼 추가 (1: 투자받음, 0: 투자안받음)
sobujang$fund2020 <- ifelse(sobujang$사업자번호_clean %in% government_funded_2020, 1, 0)
sobujang$fund2021 <- ifelse(sobujang$사업자번호_clean %in% government_funded_2021, 1, 0)
sobujang$fund2022 <- ifelse(sobujang$사업자번호_clean %in% government_funded_2022, 1, 0)

# 결과 확인 (선택사항)
table(sobujang$fund2020)
table(sobujang$fund2021)
table(sobujang$fund2022)

# 정부투자금 
sobujang$gfundvol2020 <- ifelse(sobujang$사업자번호_clean %in% government_funded_2020, ntis$민간연구비_소계, 0)
sobujang$gfundvol2021 <- ifelse(sobujang$사업자번호_clean %in% government_funded_2021, ntis$민간연구비_소계, 0)
sobujang$gfundvol2022 <- ifelse(sobujang$사업자번호_clean %in% government_funded_2022, ntis$민간연구비_소계, 0)

# 결과 확인 (선택사항)
table(sobujang$gfundvol2020)
table(sobujang$gfundvol2021)
table(sobujang$gfundvol2022)

#===============================================================================
# 펀딩 조합별 기업 수 확인
sobujang <- sobujang %>%
  mutate(
    fundedpattern = paste0(fund2020, fund2021, fund2022)
  )

cat("\n=== 펀딩 패턴 분포 ===\n")
print(table(sobujang$fundedpattern))

#===============================================================================
# 소재= 1 부품 = 2 장비 = 3

# 소재, 부품, 장비 분류 변수 생성
sobujang <- sobujang %>%
  mutate(
    seg = case_when(
      str_detect(분류, "소재") ~ 1,
      str_detect(분류, "부품") ~ 2,
      str_detect(분류, "장비") ~ 3,
      TRUE ~ NA_integer_
    )
  )

# 결과 확인
table(sobujang$seg, useNA = "ifany")

#===============================================================================
# 결과 확인
cat("총 기업 수:", nrow(sobujang), "\n")
cat("2020 정부 투자 받은 기업:", sum(sobujang$fund2020 == 1), "\n")
cat("2020 정부 투자 받지 않은 기업:", sum(sobujang$fund2020 == 0), "\n")
cat("2021 정부 투자 받은 기업:", sum(sobujang$fund2021 == 1), "\n")
cat("2021 정부 투자 받지 않은 기업:", sum(sobujang$fund2021 == 0), "\n")
cat("2022 정부 투자 받은 기업:", sum(sobujang$fund2022 == 1), "\n")
cat("2022 정부 투자 받지 않은 기업:", sum(sobujang$fund2022 == 0), "\n")


# 데이터 통합 (left join)
merged_data <- left_join(sobujang, valuesearch, by = "사업자번호_clean")

# 결과 저장
write_xlsx(merged_data, "sobujang_integrated_with_valuesearch2.xlsx")

# 매칭 결과 출력
matched <- sum(!is.na(merged_data$업체코드))
cat("총 기업 수:", nrow(merged_data), "\n")
cat("VALUESearch 매칭:", matched, "개 (", 
    round(matched/nrow(merged_data)*100, 2), "%)\n")

#===============================================================================
# 그룹 구분 컬럼 추가
merged_data <- merged_data %>%
  mutate(
    funded = ifelse(fund2020 == 1 | fund2021 == 1 | fund2022 == 1, 1, 0),
    
    group = case_when(
      funded == 1 & !is.na(업체코드) ~ "Treatgroup",
      funded == 0 & !is.na(업체코드) ~ "Controlgroup",
      TRUE ~ "기타"  # 업체코드가 없는 경우
    ))

#===============================================================================
#기업 업력

print(merged_data$`691040.설립일`)

merged_data <- merged_data %>%
  mutate(
    설립일_clean = na_if(`691040.설립일`, ""),   # 빈문자열 → NA
    설립일_date  = parse_date_time(설립일_clean,
                                orders = c("Y-m-d", "Ymd", "Y/m/d")),
    age          = year(Sys.Date()) - year(설립일_date)
  )

# 검증
cat("원본 NA    :", sum(is.na(merged_data$`691040.설립일`)), "\n")
cat("빈문자열   :", sum(merged_data$`691040.설립일` == "", na.rm = TRUE), "\n")
cat("변환 후 NA :", sum(is.na(merged_data$age)), "\n")

# 변환 성공 샘플
merged_data %>%
  filter(!is.na(age)) %>%
  select(`691040.설립일`, 설립일_date, age) %>%
  head(100)


print(merged_data$age)
print(merged_data$funded)
print(merged_data$group)

#===============================================================================
# 결과 저장
#write_xlsx(merged_data, "sobujang_integrated_with_valuesearch2.xlsx")

# 그룹별 통계 확인
group_summary <- merged_data %>%
  group_by(group) %>%
  summarise(
    기업수 = n(),
    비율 = round(n() / nrow(merged_data) * 100, 2)
  ) %>%
  arrange(desc(기업수))

print(group_summary)

# Treatment group과 Control group만 추출
treatment_group <- merged_data %>%
  filter(group == "Treatgroup")

control_group <- merged_data %>%
  filter(group == "Controlgroup")

# 결과 출력
cat("\n=== 그룹 구분 결과 ===\n")
cat("Treatment Group (정부투자 O, VALUESearch O):", nrow(treatment_group), "개\n")
cat("Control Group (정부투자 X, VALUESearch O):", nrow(control_group), "개\n")
cat("기타 (VALUESearch 매칭 안됨):", sum(merged_data$group == "기타"), "개\n")

# 기업 지역 구분 
#===============================================================================
names(merged_data)
#head(merged_data$`691090.본사주소`)
#head(merged_data$`691200.전화번호`)

# 지역 6개. 

merged_data <- merged_data %>%
  mutate(
    zipcode_num = as.numeric(str_extract(`691090.본사주소`, "\\d{5}")),
    region = case_when(
      zipcode_num >= 1000 & zipcode_num <= 9999 ~ 1,     # 서울
      zipcode_num >= 10000 & zipcode_num <= 20999 ~ 2,   # 경기
      zipcode_num >= 21000 & zipcode_num <= 23999 ~ 3,   # 인천
      zipcode_num >= 24000 & zipcode_num <= 26999 ~ 4,   # 강원
      zipcode_num >= 27000 & zipcode_num <= 29999 ~ 5,   # 충북
      zipcode_num >= 30000 & zipcode_num <= 33999 ~ 6,   # 세종/충남
      zipcode_num >= 34000 & zipcode_num <= 35999 ~ 7,   # 대전
      zipcode_num >= 36000 & zipcode_num <= 40499 ~ 8,   # 경북
      zipcode_num >= 40500 & zipcode_num <= 43999 ~ 9,   # 대구
      zipcode_num >= 44000 & zipcode_num <= 45999 ~ 10,  # 울산
      zipcode_num >= 46000 & zipcode_num <= 49999 ~ 11,  # 부산
      zipcode_num >= 50000 & zipcode_num <= 53999 ~ 12,  # 경남
      zipcode_num >= 54000 & zipcode_num <= 56999 ~ 13,  # 전북
      zipcode_num >= 57000 & zipcode_num <= 60999 ~ 14,  # 전남
      zipcode_num >= 61000 & zipcode_num <= 62999 ~ 15,  # 광주
      zipcode_num >= 63000 & zipcode_num <= 63999 ~ 16,  # 제주
      is.na(zipcode_num) ~ 0,
      TRUE ~ 99
    )
  )
# 0은 매칭 안됨 데이터.

# 우편번호 추출 (괄호 안 5자리 숫자)
#merged_data <- merged_data %>%
#  mutate(zipcode = str_extract(`691090.본사주소`, "\\d{5}"))

# 우편번호별 종류 및 개수
merged_data %>%
  count(region) %>%
  arrange(desc(n)) %>%
  print(n = 50)

# 기업 표준 산업 분류 - "691420.KSIC-중분류(11차)"   
#===============================================================================
# C00000.제조업(10~34) 소부장 기업 대부분 95% 제조업 이다. 
# E00000.수도, 하수 및 폐기물 처리, 원료 재생업(36~39) : 2
# 

names(merged_data)
#head(merged_data$`691300.NICS 산업분류`)
#head(merged_data$`691410.KSIC-대분류(11차)`)

# 전체 출력
merged_data %>%
  count(`691420.KSIC-중분류(11차)`) %>%
  arrange(desc(n)) %>%
  print(n = Inf)


#1 NA                                                       8305
#2 C29000.기타 기계 및 장비 제조업                           622
#3 C30000.자동차 및 트레일러 제조업                          381
#4 C26000.전자 부품, 컴퓨터, 영상, 음향 및 통신장비 제조업   303
#5 C25000.금속 가공제품 제조업                               209
#6 C28000.전기장비 제조업                                    186
#7 C24000.1차 금속 제조업                                    184
#8 C20000.화학 물질 및 화학제품 제조업                       183
#9 C27000.의료, 정밀, 광학 기기 및 시계 제조업               121
#10 C22000.고무 및 플라스틱제품 제조업                        105
#11 C31000.기타 운송장비 제조업                                56
#12 C13000.섬유제품 제조업                                     33
#13 C23000.비금속 광물제품 제조업                              33
#14 C21000.의료용 물질 및 의약품 제조업                        18
#15 G46000.도매 및 상품 중개업                                 13
#16 J58000.출판업                                              10
#17 C18000.인쇄 및 기록매체 복제업                              6
#18 C33000.기타 제품 제조업                                     6
#19 F42000.전문직별 공사업                                      6
#20 C10000.식료품 제조업                                        3
#21 C32000.가구 제조업                                          3
#22 E38000.폐기물 수집, 운반, 처리 및 원료 재생업               3
#23 M70000.연구개발업                                           3
#24 C17000.펄프, 종이 및 종이제품 제조업                        2
#25 C19000.코크스, 연탄 및 석유정제품 제조업                    2
#26 C14000.의복, 의복 액세서리 및 모피제품 제조업               1
#27 C15000.가죽, 가방 및 신발 제조업                            1
#28 C16000.목재 및 나무제품 제조업                              1
#29 J62000.컴퓨터 프로그래밍, 시스템 통합 및 관리업             1
#30 M71000.전문 서비스업                                        1
#31 M72000.건축 기술, 엔지니어링 및 기타 과학기술 서비스업      1

# ============================================================
# KSIC 중분류 코드 변환
# ============================================================

merged_data <- merged_data %>%
  mutate(KSIC_mid_code = case_when(
    is.na(`691420.KSIC-중분류(11차)`) ~ 0,                                      # 매칭 안됨
    grepl("^C29000", `691420.KSIC-중분류(11차)`) ~ 1,   # 기타 기계 및 장비 제조업
    grepl("^C30000", `691420.KSIC-중분류(11차)`) ~ 2,   # 자동차 및 트레일러 제조업
    grepl("^C26000", `691420.KSIC-중분류(11차)`) ~ 3,   # 전자 부품, 컴퓨터, 영상, 음향 및 통신장비 제조업
    grepl("^C25000", `691420.KSIC-중분류(11차)`) ~ 4,   # 금속 가공제품 제조업
    grepl("^C28000", `691420.KSIC-중분류(11차)`) ~ 5,   # 전기장비 제조업
    grepl("^C24000", `691420.KSIC-중분류(11차)`) ~ 6,   # 1차 금속 제조업
    grepl("^C20000", `691420.KSIC-중분류(11차)`) ~ 7,   # 화학 물질 및 화학제품 제조업
    grepl("^C27000", `691420.KSIC-중분류(11차)`) ~ 8,   # 의료, 정밀, 광학 기기 및 시계 제조업
    grepl("^C22000", `691420.KSIC-중분류(11차)`) ~ 9,   # 고무 및 플라스틱제품 제조업
    grepl("^C31000", `691420.KSIC-중분류(11차)`) ~ 10,  # 기타 운송장비 제조업
    grepl("^C13000", `691420.KSIC-중분류(11차)`) ~ 11,  # 섬유제품 제조업
    grepl("^C23000", `691420.KSIC-중분류(11차)`) ~ 12,  # 비금속 광물제품 제조업
    grepl("^C21000", `691420.KSIC-중분류(11차)`) ~ 13,  # 의료용 물질 및 의약품 제조업
    grepl("^G46000", `691420.KSIC-중분류(11차)`) ~ 14,  # 도매 및 상품 중개업
    grepl("^J58000", `691420.KSIC-중분류(11차)`) ~ 15,  # 출판업
    grepl("^C18000", `691420.KSIC-중분류(11차)`) ~ 16,  # 인쇄 및 기록매체 복제업
    grepl("^C33000", `691420.KSIC-중분류(11차)`) ~ 17,  # 기타 제품 제조업
    grepl("^F42000", `691420.KSIC-중분류(11차)`) ~ 18,  # 전문직별 공사업
    grepl("^C10000", `691420.KSIC-중분류(11차)`) ~ 19,  # 식료품 제조업
    grepl("^C32000", `691420.KSIC-중분류(11차)`) ~ 20,  # 가구 제조업
    grepl("^E38000", `691420.KSIC-중분류(11차)`) ~ 21,  # 폐기물 수집, 운반, 처리 및 원료 재생업
    grepl("^M70000", `691420.KSIC-중분류(11차)`) ~ 22,  # 연구개발업
    grepl("^C17000", `691420.KSIC-중분류(11차)`) ~ 23,  # 펄프, 종이 및 종이제품 제조업
    grepl("^C19000", `691420.KSIC-중분류(11차)`) ~ 24,  # 코크스, 연탄 및 석유정제품 제조업
    grepl("^C14000", `691420.KSIC-중분류(11차)`) ~ 25,  # 의복, 의복 액세서리 및 모피제품 제조업
    grepl("^C15000", `691420.KSIC-중분류(11차)`) ~ 26,  # 가죽, 가방 및 신발 제조업
    grepl("^C16000", `691420.KSIC-중분류(11차)`) ~ 27,  # 목재 및 나무제품 제조업
    grepl("^J62000", `691420.KSIC-중분류(11차)`) ~ 28,  # 컴퓨터 프로그래밍, 시스템 통합 및 관리업
    grepl("^M71000", `691420.KSIC-중분류(11차)`) ~ 29,  # 전문 서비스업
    grepl("^M72000", `691420.KSIC-중분류(11차)`) ~ 30,  # 건축 기술, 엔지니어링 및 기타 과학기술 서비스업
    TRUE ~ 99                                           # 기타
  ))

#merged_data <- merged_data %>%
#  mutate(KSIC_code = case_when(
#    is.na(`691410.KSIC-대분류(11차)`) ~ 0,            # 사업자번호 매칭 안됨
#    grepl("^C00000", `691410.KSIC-대분류(11차)`) ~ 1, # 제조 
#    grepl("^G00000", `691410.KSIC-대분류(11차)`) ~ 2, # 도매및소매업 
#    grepl("^J00000", `691410.KSIC-대분류(11차)`) ~ 3, # 정보통신업
#    grepl("^F00000", `691410.KSIC-대분류(11차)`) ~ 4, # 건설업 
#    grepl("^M00000", `691410.KSIC-대분류(11차)`) ~ 5, # 전문,과학및기술서비스업
#    grepl("^E00000", `691410.KSIC-대분류(11차)`) ~ 6, # 수도,하수및폐기물처리,원료재생업
#    TRUE ~ 99
#  ))

# 확인
merged_data %>%
  count(KSIC_mid_code) %>%
  arrange(KSIC_mid_code)

#===============================================================================
# 수출 여부를 0, 1로 표기 
# 수출, 개발비용 N/A 이면 0으로 처리.


# 전체 출력
merged_data %>%
  count(`2019/Annual S21195.[수출]`) %>%
  arrange(desc(n)) %>%
  print(n = Inf)

# 년도별 수출 유무 (수출 있으면 1, 없으면 0)
merged_data <- merged_data %>%
  mutate(
    export2019 = ifelse(!is.na(`2019/Annual S21195.[수출]`) & `2019/Annual S21195.[수출]` > 0, 1, 0),
    export2020 = ifelse(!is.na(`2020/Annual S21195.[수출]`) & `2020/Annual S21195.[수출]` > 0, 1, 0),
    export2021 = ifelse(!is.na(`2021/Annual S21195.[수출]`) & `2021/Annual S21195.[수출]` > 0, 1, 0),
    export2022 = ifelse(!is.na(`2022/Annual S21195.[수출]`) & `2022/Annual S21195.[수출]` > 0, 1, 0),
    export2023 = ifelse(!is.na(`2023/Annual S21195.[수출]`) & `2023/Annual S21195.[수출]` > 0, 1, 0),
    export2024 = ifelse(!is.na(`2024/Annual S21195.[수출]`) & `2024/Annual S21195.[수출]` > 0, 1, 0),
    
    # ── 수출금액 (NA → 0 처리) ──
    exportamt2019 = ifelse(is.na(as.numeric(`2019/Annual S21195.[수출]`)), 0, as.numeric(`2019/Annual S21195.[수출]`)),
    exportamt2020 = ifelse(is.na(as.numeric(`2020/Annual S21195.[수출]`)), 0, as.numeric(`2020/Annual S21195.[수출]`)),
    exportamt2021 = ifelse(is.na(as.numeric(`2021/Annual S21195.[수출]`)), 0, as.numeric(`2021/Annual S21195.[수출]`)),
    exportamt2022 = ifelse(is.na(as.numeric(`2022/Annual S21195.[수출]`)), 0, as.numeric(`2022/Annual S21195.[수출]`)),
    exportamt2023 = ifelse(is.na(as.numeric(`2023/Annual S21195.[수출]`)), 0, as.numeric(`2023/Annual S21195.[수출]`)),
    exportamt2024 = ifelse(is.na(as.numeric(`2024/Annual S21195.[수출]`)), 0, as.numeric(`2024/Annual S21195.[수출]`)),
    
    # ── 연구개발비용계 (NA → 0 처리) ──
    rdcost2019 = ifelse(is.na(as.numeric(`2019/Annual 692084.연구개발비용-연구개발비용계`)), 0, as.numeric(`2019/Annual 692084.연구개발비용-연구개발비용계`)),
    rdcost2020 = ifelse(is.na(as.numeric(`2020/Annual 692084.연구개발비용-연구개발비용계`)), 0, as.numeric(`2020/Annual 692084.연구개발비용-연구개발비용계`)),
    rdcost2021 = ifelse(is.na(as.numeric(`2021/Annual 692084.연구개발비용-연구개발비용계`)), 0, as.numeric(`2021/Annual 692084.연구개발비용-연구개발비용계`)),
    rdcost2022 = ifelse(is.na(as.numeric(`2022/Annual 692084.연구개발비용-연구개발비용계`)), 0, as.numeric(`2022/Annual 692084.연구개발비용-연구개발비용계`)),
    rdcost2023 = ifelse(is.na(as.numeric(`2023/Annual 692084.연구개발비용-연구개발비용계`)), 0, as.numeric(`2023/Annual 692084.연구개발비용-연구개발비용계`)),
    rdcost2024 = ifelse(is.na(as.numeric(`2024/Annual 692084.연구개발비용-연구개발비용계`)), 0, as.numeric(`2024/Annual 692084.연구개발비용-연구개발비용계`))
  )

# 확인
merged_data %>%
  summarise(
    exportamt2019 = sum(exportamt2019, na.rm = TRUE),
    exportamt2020 = sum(exportamt2020, na.rm = TRUE),
    exportamt2021 = sum(exportamt2021, na.rm = TRUE),
    exportamt2022 = sum(exportamt2022, na.rm = TRUE),
    exportamt2023 = sum(exportamt2023, na.rm = TRUE),
    exportamt2024 = sum(exportamt2024, na.rm = TRUE)
  )

#===============================================================================
# 노동생상성 지수 = 매출 / 종업원 수 

# 년도별 노동생산성 (매출/종업원수) 계산
merged_data <- merged_data %>%
  mutate(
    labor_prod2019 = as.numeric(`2019/Annual S21100.총매출액`) / as.numeric(`2019/Annual S05000.종업원수`),
    labor_prod2020 = as.numeric(`2020/Annual S21100.총매출액`) / as.numeric(`2020/Annual S05000.종업원수`),
    labor_prod2021 = as.numeric(`2021/Annual S21100.총매출액`) / as.numeric(`2021/Annual S05000.종업원수`),
    labor_prod2022 = as.numeric(`2022/Annual S21100.총매출액`) / as.numeric(`2022/Annual S05000.종업원수`),
    labor_prod2023 = as.numeric(`2023/Annual S21100.총매출액`) / as.numeric(`2023/Annual S05000.종업원수`),
    labor_prod2024 = as.numeric(`2024/Annual S21100.총매출액`) / as.numeric(`2024/Annual S05000.종업원수`)
  )

# 확인
merged_data %>%
  summarise(
    avg_labor_prod2019 = mean(labor_prod2019, na.rm = TRUE),
    avg_labor_prod2020 = mean(labor_prod2020, na.rm = TRUE),
    avg_labor_prod2021 = mean(labor_prod2021, na.rm = TRUE),
    avg_labor_prod2022 = mean(labor_prod2022, na.rm = TRUE),
    avg_labor_prod2023 = mean(labor_prod2023, na.rm = TRUE),
    avg_labor_prod2024 = mean(labor_prod2024, na.rm = TRUE)
  )



#===============================================================================
# 결과 저장
write_xlsx(merged_data, "sobujang_integrated_with_valuesearch2.xlsx")
#===============================================================================


#===============================================================================
# 특허 데이터를 매칭 해야 한다. 신규 매칭을 하면 데이터를 찾아서 사업자 
# 번호로 다시 매칭 해준다.
#1st 첫번째 특허 테이터 
#2nd 소재, 부품, 장비 숫자를 맞추고한 특허 데이터 
#===============================================================================
# 1st 특허 데이터 읽기
patent <- read_excel("PSM matched_dataset+Patent.xlsx", col_types = "text")

# 2st 특허 데이터 읽기
patent2 <- read_excel("matched_dataset_patent2_new_matching.xlsx", col_types = "text")

names(patent)
names(patent2)


cat("=== 원본 데이터 확인 ===\n")
cat("patent  행 수:", nrow(patent),  "| 고유 사업자:", length(unique(patent$사업자번호_clean)), "\n")
cat("patent2 행 수:", nrow(patent2), "| 고유 사업자:", length(unique(patent2$사업자번호_clean)), "\n\n")

#===============================================================================
# patent2에서 patent에 없는 사업자번호만 추출
#===============================================================================

# patent에 이미 있는 사업자번호 목록
existing_ids <- unique(patent$사업자번호_clean)

# patent2에서 중복되지 않는 항목만 필터링
patent2_new <- patent2 %>%
  filter(!(사업자번호_clean %in% existing_ids))

cat("=== 중복 확인 ===\n")
cat("patent2 전체:", nrow(patent2), "\n")
cat("patent과 중복:", nrow(patent2) - nrow(patent2_new), "\n")
cat("patent2에서 새로 추가할 기업:", nrow(patent2_new), "\n\n")


#===============================================================================
# patent2_new를 patent에 추가 (행 결합)
#===============================================================================

# patent2_new에서 patent과 동일한 컬럼만 선택하여 결합
# (두 데이터셋의 컬럼 구조가 다를 수 있으므로 공통 컬럼 확인)
common_cols <- intersect(names(patent), names(patent2_new))
cat("공통 컬럼 수:", length(common_cols), "\n")
cat("공통 컬럼:", paste(common_cols, collapse = ", "), "\n\n")

# 방법 A: 공통 컬럼 기준으로 결합
patent_combined <- bind_rows(
  patent,
  patent2_new %>% select(any_of(names(patent)))
)

cat("=== 통합 결과 ===\n")
cat("통합 후 patent 행 수:", nrow(patent_combined), "\n")
cat("통합 후 고유 사업자:", length(unique(patent_combined$사업자번호_clean)), "\n\n")

#===============================================================================
# 특허 컬럼 추출 + 숫자 변환 + 중복 사업자 합산
#===============================================================================

patent_subset <- patent_combined %>%
  select(사업자번호_clean, p2019, p2020, p2021, p2022, p2023, p2024) %>%
  mutate(across(p2019:p2024, as.numeric)) %>%
  # 중복 사업자번호가 있을 경우 합산
  group_by(사업자번호_clean) %>%
  summarise(across(p2019:p2024, ~sum(.x, na.rm = TRUE))) %>%
  ungroup()

cat("중복 합산 후 기업 수:", nrow(patent_subset), "\n\n")

#===============================================================================
# merged_data에 결합
#===============================================================================

# 기존 특허 컬럼이 있으면 제거 (재결합 시 충돌 방지)
merged_data <- merged_data %>%
  select(-any_of(c("p2019", "p2020", "p2021", "p2022", "p2023", "p2024")))

# left_join으로 특허 데이터 결합
merged_data <- merged_data %>%
  left_join(patent_subset, by = "사업자번호_clean")

#===============================================================================
# 6. 결과 확인
#===============================================================================

cat("=== 최종 매칭 결과 ===\n")
cat("전체 기업 수:", nrow(merged_data), "\n")
cat("특허 매칭 기업:", sum(!is.na(merged_data$p2019)), "\n")
cat("특허 미매칭 기업 (NA):", sum(is.na(merged_data$p2019)), "\n\n")

# 그룹별 연도별 평균 특허
cat("=== 그룹별 평균 특허 건수 ===\n")
merged_data %>%
  group_by(group) %>%
  summarise(
    N = n(),
    p2019 = round(mean(p2019, na.rm = TRUE), 2),
    p2020 = round(mean(p2020, na.rm = TRUE), 2),
    p2021 = round(mean(p2021, na.rm = TRUE), 2),
    p2022 = round(mean(p2022, na.rm = TRUE), 2),
    p2023 = round(mean(p2023, na.rm = TRUE), 2),
    p2024 = round(mean(p2024, na.rm = TRUE), 2)
  ) %>%
  print()

# 부문별(소재/부품/장비) 확인
cat("\n=== 부문별 특허 매칭 현황 ===\n")
merged_data %>%
  group_by(분류) %>%
  summarise(
    전체 = n(),
    매칭 = sum(!is.na(p2019)),
    미매칭 = sum(is.na(p2019)),
    매칭률 = paste0(round(매칭 / 전체 * 100, 1), "%")
  ) %>%
  print()




#===============================================================================
# 결측치 확인
# 8305 결측치는 소부장 데이터에 Value Search 데이터 맵핑이 안되어 발생 !!!
#===============================================================================

# ── 연도 고정 항목 (한 번만 확인) ──────────────────────────────
fixed_na <- merged_data %>%
  summarise(
    n_total       = n(),
    na_업력       = sum(is.na(age)),
    na_KSIC_표준  = sum(is.na(KSIC_mid_code)),
    na_지역       = sum(is.na(region))
  )

cat("=== 연도 고정 항목 ===\n")
print(fixed_na, width = Inf)

# ── 연도별 반복 항목 ───────────────────────────────────────────
years <- 2019:2024

result <- map_dfr(years, function(yr) {
  merged_data %>%
    summarise(
      연도                = yr,
      n_total             = n(),
      
      # 재무 변수 (원본 컬럼)
      na_자산             = sum(is.na(.data[[paste0(yr, "/Annual S15000.자산총계")]])),
      na_매출             = sum(is.na(.data[[paste0(yr, "/Annual S21100.총매출액")]])),
      na_부채             = sum(is.na(.data[[paste0(yr, "/Annual S18000.부채총계")]])),
      na_자본금           = sum(is.na(.data[[paste0(yr, "/Annual S18100.자본금")]])),
      na_종업원           = sum(is.na(.data[[paste0(yr, "/Annual S05000.종업원수")]])),
      na_영업이익         = sum(is.na(.data[[paste0(yr, "/Annual S25000.영업이익(손실)")]])),
      
      # 연구개발비 (원본 vs 변환)
      na_연구개발비_원본  = sum(is.na(.data[[paste0(yr, "/Annual 692084.연구개발비용-연구개발비용계")]])),
      na_연구개발비_변환  = sum(is.na(.data[[paste0("rdcost", yr)]])),   # 0이어야 정상
      
      # 수출 (원본 vs 변환)
      na_수출_원본        = sum(is.na(.data[[paste0(yr, "/Annual S21195.[수출]")]])),
      na_수출_변환        = sum(is.na(.data[[paste0("exportamt", yr)]])), # 0이어야 정상
      
      # 노동생산성 (변환 컬럼)
      na_노동생산성       = sum(is.na(.data[[paste0("labor_prod", yr)]])),
      
      # 특허 (변환 컬럼)
      na_특허             = sum(is.na(.data[[paste0("p", yr)]]))
    )
})

cat("\n=== 연도별 항목 ===\n")
print(result, width = Inf)

#===============================================================================
#  저장
#===============================================================================

write_xlsx(merged_data, "sobujang_integrated_with_valuesearch+Patent.xlsx")
cat("\n✓ 저장 완료\n")



