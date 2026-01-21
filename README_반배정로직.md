# 학급 개편 반 배정 로직

## 개요
학급 개편 추진 계획에 따른 A, B, C, D반의 남녀 인원 배정 로직입니다.

## 원칙
1. **학년 학생 수를 4학급으로 나누고 반별 학생 수 결정**
   - 예) 128명 ÷ 4 = 32명씩 배정

2. **학년의 남녀 학생 수를 구분하여 균등하게 배정**
   - 몫은 동일하게 나누고, 나머지는 A→B→C→D 순으로 1명씩 추가
   - 남녀 각각 배분 후, 남자가 총원보다 많은 반이 있으면 초과 인원을 여유 있는 반으로 이동
   - 여학생 수는 `(반별 총원 - 반별 남학생)`으로 계산해 합계를 자동 일치

## 사용 방법

### 기본 사용법

```python
from class_assignment_logic import calculate_class_distribution

# 학생 수 입력
total_students = 122  # 총 학생 수
male_count = 64      # 남학생 수
female_count = 58    # 여학생 수

# 배정 계산
distribution = calculate_class_distribution(total_students, male_count, female_count)

# 결과 확인
print(distribution)
# {
#     'total': {'A': 31, 'B': 31, 'C': 30, 'D': 30},
#     'male': {'A': 16, 'B': 16, 'C': 16, 'D': 16},
#     'female': {'A': 15, 'B': 15, 'C': 14, 'D': 14}
# }
```

### Excel 파일에서 직접 읽어서 사용

```python
import pandas as pd
from class_assignment_logic import calculate_class_distribution, print_distribution_summary

# Excel 파일 읽기
df = pd.read_excel('학생자료.xlsx')

# 남녀 학생 수 계산
total_students = len(df)
male_count = len(df[df['남녀'] == '남'])
female_count = len(df[df['남녀'] == '여'])

# 배정 계산
distribution = calculate_class_distribution(total_students, male_count, female_count)

# 결과 출력
print_distribution_summary(distribution, total_students, male_count, female_count)
```

## 함수 설명

### `calculate_class_distribution(total_students, male_count, female_count)`
총 학생 수와 남녀 학생 수를 바탕으로 A, B, C, D반의 목표 인원을 계산합니다.

**매개변수:**
- `total_students` (int): 총 학생 수
- `male_count` (int): 남학생 수
- `female_count` (int): 여학생 수

**반환값:**
```python
{
    'total': {'A': int, 'B': int, 'C': int, 'D': int},
    'male': {'A': int, 'B': int, 'C': int, 'D': int},
    'female': {'A': int, 'B': int, 'C': int, 'D': int}
}
```

### `distribute_equally(total_count, class_names)`
총 인원을 여러 반에 균등하게 분배합니다. 몫은 동일하게, 나머지는 A→B→C→D 순으로 1명씩 추가합니다.

**매개변수:**
- `total_count` (int): 총 인원
- `class_names` (list): 반 이름 리스트 (예: ['A', 'B', 'C', 'D'])

**반환값:**
- `dict`: 각 반의 인원 수

**예시:**
- 66명 → {'A': 17, 'B': 17, 'C': 16, 'D': 16}
- 62명 → {'A': 16, 'B': 16, 'C': 15, 'D': 15}

### `print_distribution_summary(distribution, total_students, male_count, female_count)`
배정 결과를 보기 좋게 출력합니다.

**매개변수:**
- `distribution`: `calculate_class_distribution()`의 반환값
- `total_students` (int): 총 학생 수
- `male_count` (int): 남학생 수
- `female_count` (int): 여학생 수

## 특징
- **범용성**: 매년 다른 학생 수와 구성을 가진 데이터에도 적용 가능
- **정확성**: 총 학생 수, 남학생 수, 여학생 수가 모두 정확히 일치하도록 보장
- **균등성**: 각 반의 인원 차이는 최대 1명 이내로 유지
- **자동 검증**: 배정 결과가 올바른지 자동으로 검증

## 예시 결과

### 현재 데이터 (122명, 남 64명, 여 58명)
```
반      총 인원     남학생     여학생    
----------------------------------------------------------------------
A      31         16         15        
B      31         16         15        
C      30         16         14        
D      30         16         14        
----------------------------------------------------------------------
합계     122        64         58        
```

### 예시 1 (128명, 남 66명, 여 62명)
```
반      총 인원     남학생     여학생    
----------------------------------------------------------------------
A      32         16         16        
B      32         16         16        
C      32         17         15        
D      32         17         15        
----------------------------------------------------------------------
합계     128        66         62        
```

### 예시 2 (120명, 남 60명, 여 60명)
```
반      총 인원     남학생     여학생    
----------------------------------------------------------------------
A      30         15         15        
B      30         15         15        
C      30         15         15        
D      30         15         15        
----------------------------------------------------------------------
합계     120        60         60        
```
