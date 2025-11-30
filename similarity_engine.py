import pandas as pd
import numpy as np


def calculate_similarity_score(student1_features, student2_features):
    """
    두 학생의 특성 벡터 간 유사도 점수 계산
    유클리드 거리를 사용하여 거리가 가까울수록 높은 점수 반환
    
    Args:
        student1_features: 학생1의 특성 벡터 (list or array)
        student2_features: 학생2의 특성 벡터 (list or array)
        
    Returns:
        float: 유사도 점수 (0~1, 1에 가까울수록 유사함)
    """
    if len(student1_features) != len(student2_features):
        return 0.0
    
    # NaN 값 처리
    features1 = np.array([x if not pd.isna(x) else 5.0 for x in student1_features])
    features2 = np.array([x if not pd.isna(x) else 5.0 for x in student2_features])
    
    # 유클리드 거리 계산
    distance = np.sqrt(np.sum((features1 - features2) ** 2))
    
    # 최대 거리 (모든 특성이 1~10 범위이므로)
    max_distance = np.sqrt(len(features1) * (10 - 1) ** 2)
    
    # 거리를 유사도 점수로 변환 (0~1 범위)
    similarity = 1.0 - (distance / max_distance) if max_distance > 0 else 1.0
    
    return max(0.0, min(1.0, similarity))


def calculate_room_similarity_score(room_members_features, candidate_features):
    """
    방의 현재 구성원들과 후보 학생의 평균 유사도 점수 계산
    
    Args:
        room_members_features: 방 구성원들의 특성 벡터 리스트
        candidate_features: 후보 학생의 특성 벡터
        
    Returns:
        float: 평균 유사도 점수
    """
    if not room_members_features:
        return 0.5  # 방이 비어있으면 중간 점수 반환
    
    similarities = []
    for member_features in room_members_features:
        similarity = calculate_similarity_score(member_features, candidate_features)
        similarities.append(similarity)
    
    return np.mean(similarities)


def get_student_features(df, student_id, feature_columns):
    """
    학생의 특성 벡터 추출
    
    Args:
        df: 데이터프레임
        student_id: 학생 ID
        feature_columns: 특성 컬럼 이름 리스트
        
    Returns:
        list: 특성 벡터
    """
    row = df.loc[df["StudentID"] == student_id]
    if row.empty:
        return [5.0] * len(feature_columns)  # 기본값 반환
    
    features = []
    for col in feature_columns:
        value = row[col].values[0]
        if pd.isna(value) or value == "":
            features.append(5.0)  # 기본값
        else:
            features.append(float(value))
    
    return features

