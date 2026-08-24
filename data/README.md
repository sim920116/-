# data/

`scripts/extract_weekly_activity.py` 실행 결과가 `weekly_activity_<이름>.csv` 형태로 이 폴더에 쌓입니다.

## CSV 컬럼

| 컬럼 | 설명 |
| --- | --- |
| weekOf | 실행 시 전달한 주차 시작일 |
| date | 엑셀에서 인식한 실제 발생일 (없으면 weekOf) |
| group | 배치/그룹명 |
| activity | 활동 종류 (폐사/출하/입식 등) |
| count | 두수 |
| weight_kg | 중량(kg), 있는 경우만 |
| note | 비고 |
| source_file | 원본 엑셀 파일명 |
| extracted_at | 추출 실행 시각 |

동일한 `weekOf`로 재실행하면 해당 주차 데이터가 새 결과로 교체됩니다(중복 누적 방지).
