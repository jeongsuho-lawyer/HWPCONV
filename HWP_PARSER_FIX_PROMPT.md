# HWP 파서 표/이미지 파싱 문제 해결 프롬프트

## 프로젝트 개요
- **경로**: `d:\Projects\HWPCONV`
- **테스트 파일**: `190820 (석간) .hwp 서식 유료 소프트웨어 없이 작성 가능(정보공개정책과)게시.hwp`
- **핵심 파서 파일**: `src/hwpconv/parsers/hwp.py`

---

## 해결된 문제

### 1. 본문 내용 대량 누락 (부분 해결)
- **증상**: 원본 HWP에 있는 "붙임1", "시간 계획", "복사/붙이기" 등 텍스트가 변환 결과에서 누락
- **원인**: `_parse_table_from_records_v2` 함수에서 표 파싱 시 `consumed` 값이 너무 커서 다음 표들과 본문을 건너뜀
- **해결**: `CTRL_HEADER`를 만나면 `break`하는 조건 추가 (Line 823-825)
- **결과**: Lines 68 → 182, "붙임1", "시간 계획" 포함됨

---

## 해결해야 할 문제

### 1. 표 셀 내용 빈칸
- **증상**: 
  ```markdown
  | 현행(as-is) |   |   |   |   |   |   |   |   |
  | --- | --- | --- | --- | --- | --- | --- | --- | --- |
  |   |   |   |   |   |   |   |   |   |
  ```
  표 셀이 빈칸이고, 셀 내용("서식파일 열기", "(개선to-be)" 등)이 표 아래 별도 문단으로 출력됨
  
- **원인**: `CTRL_HEADER`를 만나면 `break`하기 때문에 표 내 중첩 개체(화살표, 그림 등) 이후의 셀 내용이 파싱되지 않음

- **핵심 딜레마**:
  - `CTRL_HEADER`에서 `break` → 본문은 살지만 표 셀 빈칸
  - `CTRL_HEADER`에서 `break` 안 함 → 표 셀은 채워지지만 `consumed`가 너무 커서 본문 누락

### 2. 이미지 인라인 배치
- **증상**: 이미지 5개가 문서 맨 끝에 모임
- **원인**: `HwpParser.parse`의 fallback 로직 (Line 157-162)이 모든 이미지를 문서 끝에 추가
- **해결 방향**: `$pic` CTRL_HEADER 감지 → `SHAPE_COMPONENT_PICTURE`에서 `bin_data_id` 추출 → 인라인 삽입

---

## 참고해야 할 핵심 문서

### 1. HWP_Table_Image_Parsing_Guide.md
- **경로**: `d:\Projects\HWPCONV\HWP_Table_Image_Parsing_Guide.md`
- **핵심 함수**: 
  - `parse_table_from_records` (Line 289-368) - 표 파싱 전체 흐름
  - `parse_cell_property` (Line 236-283) - 셀 속성 파싱
  - `parse_picture_from_records` (Line 800-848) - 그림 파싱
  
- **중요 원칙**: 
  - **스킵하지 않음** - 모든 레코드를 순회하며 처리
  - **level 기반 종료**: `level <= base_level`일 때만 표 종료
  - **셀 개수**: `row_count * col_count` 개의 `LIST_HEADER`가 표의 셀

### 2. hwp-converter-prd.md
- **경로**: `d:\Projects\HWPCONV\hwp-converter-prd.md`
- **내용**: 프로젝트 요구사항, 데이터 모델, 전체 아키텍처

---

## 핵심 코드 위치

### `_parse_table_from_records_v2` (Line 795-910)
- 현재 문제의 핵심 함수
- `CTRL_HEADER` break 조건 (Line 823-825)이 표 셀 빈칸의 원인

### `_parse_section` (Line 391-465)
- 섹션 전체 파싱
- 표, 그림, 문단 처리 분기

### `_parse_picture_from_records` (Line 1018-1061)
- `$pic` CTRL_HEADER 이후 `bin_data_id` 추출
- 현재 작동하지 않음 (이미지가 인라인으로 배치되지 않음)

---

## 검증 방법
```bash
cd d:\Projects\HWPCONV
python quick_test.py
```
- **성공 기준**:
  - Lines: 180 이상
  - ✓ 붙임1 포함
  - ✓ 시간 계획 포함
  - 표 셀에 "현행(as-is)", "(개선to-be)", "서식파일 열기" 등 내용 포함

---

## 핵심 원칙
1. **스킵하지 않는다** - 레코드를 건너뛰는 것은 답이 아님
2. **표를 정확히 인식한다** - `TABLE` 레코드의 `row_count * col_count`로 셀 개수 파악
3. **가이드 문서를 따른다** - `HWP_Table_Image_Parsing_Guide.md`의 로직을 정확히 구현
