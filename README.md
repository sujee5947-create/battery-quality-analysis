# Battery Quality Data Analysis

**Python | Battery Quality | Data Analysis | Report Automation**

## Overview

가상(synthetic) 리튬이온 배터리 셀 측정 데이터를 규격과 비교해 OK/NG를 판정하고, 공정 성능 지수와 분석 결과를 HTML 대시보드와 PowerPoint 보고서로 자동 생성하는 Python 프로젝트입니다.

배터리 셀 개발·품질 업무 경험을 바탕으로, 반복적인 품질 데이터 정리·판정·보고 업무를 자동화하는 흐름을 구현했습니다.

## Key Features

- 측정 항목별 규격(LSL/USL) 판정 → 항목 하나라도 NG면 전체 NG
- 결측값은 "미측정", 규격이 없는 항목은 "미판정"으로 따로 집계 (합격률 계산에서 제외)
- 공정 성능 지수 **Pp / Ppk** (전체 표준편차 기반), 편측 규격은 Ppu / Ppl
- 기술 통계, OK/NG 비교, 항목 간 상관 분석, 분포 시각화
- 한글·영문 열 이름 자동 인식, CSV(UTF-8/CP949) · Excel 입력 지원
- HTML 대시보드 + PowerPoint 보고서 자동 생성

## Data Parameters

| 항목 | 단위 |
|---|---|
| Capacity (용량) | Ah |
| Voltage (전압) | V |
| Internal Resistance (내부저항) | mΩ |
| Thickness (두께) | mm |
| Weight (중량) | g |

## Example Results (포함된 가상 데이터 기준)

| 지표 | 결과 |
|---|---:|
| 전체 시료 | 400 |
| OK | 390 |
| NG | 4 |
| 미판정/미측정 | 6 |
| 합격률 (OK / (OK+NG)) | 98.98% |

## How to Run

```bash
python -m pip install -r requirements.txt
python battery_quality_app_updated.py
```

- `Raw_Data/` 폴더에 측정 데이터 파일(CSV 또는 Excel) **1개**와 `specifications.xlsx`를 둡니다.
- 결과는 `Results/` 폴더에 생성됩니다: `battery_quality_dashboard.html`, `battery_quality_report.pptx`, 판정 결과 CSV, 실행 로그

## Note on Pp/Ppk

이 데이터에는 시간 순서나 부분군(subgroup) 정보가 없어 군내 표준편차를 추정할 수 없으므로, **전체 표준편차 기반의 Pp/Ppk**를 계산합니다. 군내 변동 기반 Cp/Cpk와 관리도 분석은 [electrode-manufacturing-spc](https://github.com/sujee5947-create/electrode-manufacturing-spc)에서 다룹니다.

## Data Disclaimer

All datasets used in this project are synthetic data created for learning and demonstration purposes. They do not contain actual manufacturing data or confidential company information.

## Related Projects

- [electrode-manufacturing-spc](https://github.com/sujee5947-create/electrode-manufacturing-spc) — 전극 제조 SPC·공정능력 분석
- [measurement-validation-automation](https://github.com/sujee5947-create/measurement-validation-automation) — 측정 장비 간 일치도(MSA) 검증 자동화
