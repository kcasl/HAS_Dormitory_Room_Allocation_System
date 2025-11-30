"""
테스트용 엑셀 파일 생성 스크립트
100명의 학생 데이터를 생성하여 엑셀 파일로 저장

사용 방법:
    pip install pandas openpyxl numpy
    python create_test_data.py
"""
import pandas as pd
import numpy as np
import random

# 시드 설정 (재현 가능한 결과를 위해)
random.seed(42)
np.random.seed(42)

# 100명의 학생 데이터 생성
students_data = []

for student_id in range(1, 101):
    student = {
        "StudentID": student_id,
        "Prevloc": random.choice([1, 2, 3, 4]),  # 이전 좌석 위치
        "Prev1": random.choice([0, random.randint(1, 100)]) if random.random() > 0.7 else 0,  # 이전 룸메이트1
        "Prev2": random.choice([0, random.randint(1, 100)]) if random.random() > 0.8 else 0,  # 이전 룸메이트2
        "Prev3": random.choice([0, random.randint(1, 100)]) if random.random() > 0.9 else 0,  # 이전 룸메이트3
        "Avoid1": random.choice([0, random.randint(1, 100)]) if random.random() > 0.85 else 0,  # 회피 학생1
        "Avoid2": random.choice([0, random.randint(1, 100)]) if random.random() > 0.9 else 0,  # 회피 학생2
        
        # 유사도 기반 배정을 위한 특성 (1~10 척도)
        "ColdSensitivity": random.randint(1, 10),  # 추위 민감도 (1=추위 안탐, 10=추위 많이 탐)
        "NoiseSensitivity": random.randint(1, 10),  # 소음 민감도 (1=소음 무관, 10=소음 매우 민감)
        "LightSensitivity": random.randint(1, 10),  # 빛 민감도 (1=빛 무관, 10=빛 매우 민감)
        "Cleanliness": random.randint(1, 10),  # 청결도 선호도 (1=무관, 10=매우 깔끔함)
        "SocialLevel": random.randint(1, 10),  # 사교성 (1=내성적, 10=매우 사교적)
        "StudyTime": random.randint(1, 10),  # 학습 시간 선호도 (1=짧음, 10=매우 길음)
        "SleepTime": random.randint(1, 10),  # 수면 시간 선호도 (1=늦게, 10=일찍)
    }
    
    # Prev1, Prev2, Prev3가 자기 자신이거나 0이 아닌 경우만 유지
    if student["Prev1"] == student_id:
        student["Prev1"] = 0
    if student["Prev2"] == student_id:
        student["Prev2"] = 0
    if student["Prev3"] == student_id:
        student["Prev3"] = 0
    if student["Avoid1"] == student_id:
        student["Avoid1"] = 0
    if student["Avoid2"] == student_id:
        student["Avoid2"] = 0
    
    students_data.append(student)

# DataFrame 생성
df = pd.DataFrame(students_data)

# 엑셀 파일로 저장
output_file = "test_data.xlsx"
df.to_excel(output_file, index=False, engine='openpyxl')

print(f"테스트 데이터 파일이 생성되었습니다: {output_file}")
print(f"총 {len(df)}명의 학생 데이터가 포함되어 있습니다.")
print("\n컬럼 정보:")
print(df.columns.tolist())
print("\n첫 5개 행 미리보기:")
print(df.head())

