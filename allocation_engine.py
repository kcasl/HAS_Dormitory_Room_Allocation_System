import random as rd
import pandas as pd
import numpy as np
from similarity_engine import calculate_room_similarity_score, get_student_features


def allocate_rooms(excel_file_path, blacklist_pairs=None, selected_factors=None):
    """
    기숙사 방 배정 알고리즘
    
    Args:
        excel_file_path: xlsx 파일 경로
        blacklist_pairs: 블랙리스트 조합 리스트 [(학생1, 학생2), ...] - 이 학생들은 같은 방에 배정되지 않음
        selected_factors: 선택된 factor 컬럼 리스트 (예: ['factor1', 'factor2']) - None이면 유사도 미사용
        
    Returns:
        tuple: (room_id, failed_students) - 방 배정 결과와 실패한 좌석 목록
    """
    if blacklist_pairs is None:
        blacklist_pairs = []
    
    # 블랙리스트를 집합으로 변환 (양방향 체크를 위해)
    blacklist_set = set()
    for pair in blacklist_pairs:
        # 정렬하여 양방향 체크 가능하도록
        sorted_pair = tuple(sorted(pair))
        blacklist_set.add(sorted_pair)
    
    # 엑셀 파일 읽기
    df = pd.read_excel(excel_file_path)
    
    # "학번" 컬럼에서 학생 ID 리스트 읽기
    std_id = df["학번"].dropna().tolist()
    std_id = [int(x) for x in std_id if pd.notna(x)]
    rd.shuffle(std_id)

    # 방 개수 계산: len(std_id)/4 (나머지가 있으면 +1)
    num_rooms = len(std_id) // 4
    if len(std_id) % 4 != 0:
        num_rooms += 1

    room_id = []
    for _ in range(num_rooms):
        room_id.append({"seat1": "", "seat2": "", "seat3": "", "seat4": ""})

    # 유사도 계산에 사용할 factor 컬럼 확인
    similarity_features = []
    if selected_factors:
        # 선택된 factor 컬럼들만 사용
        for factor in selected_factors:
            if factor in df.columns:
                similarity_features.append(factor)
    
    hallway_seats = ["seat1", "seat4"]  # 1,4번
    window_seats = ["seat2", "seat3"]  # 2,3번

    for i in range(min(num_rooms, len(std_id))):
        student = std_id[i]

        prev_loc_row = df.loc[df["학번"] == student, "현재 좌석 번호"]
        if prev_loc_row.empty:
            prev_loc = 0
        else:
            prev_loc = prev_loc_row.values[0]
            if pd.isna(prev_loc):
                prev_loc = 0
            else:
                prev_loc = int(prev_loc)

        # 이전에 복도(1,4) 라면 이번엔 창가(2,3)
        if prev_loc in [1, 4]:
            candidate_seats = window_seats

        # 이전에 창가(2,3) 라면 이번엔 복도(1,4)
        else:
            candidate_seats = hallway_seats

        assigned = False
        for seat in candidate_seats:
            if room_id[i][seat] == "":
                room_id[i][seat] = student
                assigned = True
                break

        # 예외 상황: 만약 candidate 좌석 다 차 있으면 아무 빈자리 fallback
        if not assigned:
            for seat in ["seat1", "seat2", "seat3", "seat4"]:
                if room_id[i][seat] == "":
                    room_id[i][seat] = student
                    break

    failed_students = []  # 배정 실패 좌석 기록

    # 이미 배정된 학생 목록
    assigned = set()
    for room in room_id:
        for seat in room.values():
            if seat != "":
                assigned.add(seat)

    # 아직 배정 안된넘들
    remaining = [s for s in std_id if s not in assigned]

    # 좌석 순서
    seat_order = ["seat1", "seat2", "seat3", "seat4"]

    for room_idx in range(num_rooms):
        room = room_id[room_idx]

        # 현재 방에 있는 사람
        current_members = [p for p in room.values() if p != ""]

        # seat2 → seat3 → seat4 순서
        for seat in seat_order:
            if room[seat] != "":
                continue  # 이미 할당된 자리면

            assigned_flag = False

            # 유사도 기반 배정을 위한 후보 학생 리스트
            candidate_students = []
            
            # 배정 가능 여부 췤
            for student in remaining:
                row = df.loc[df["학번"] == student]
                if row.empty:
                    continue

                # 현재 룸메이트 1~3, 배려 학생 1~2 데이터 읽기
                prev_list = row[["현재 룸메이트 1", "현재 룸메이트 2", "현재 룸메이트 3"]].values[0]
                avoid_list = row[["배려 학생 1", "배려 학생 2"]].values[0]

                # NaN/빈칸 제거
                prev_list = [int(x) for x in prev_list if not (pd.isna(x) or x == "" or x == 0)]
                avoid_list = [int(x) for x in avoid_list if not (pd.isna(x) or x == "" or x == 0)]

                # 조건

                # 같이 하기 싫은 애
                avoid_conflict = any(a in current_members for a in avoid_list)
                if avoid_conflict:
                    continue

                # 이전 룸메 회피
                prev_conflict = any(p in current_members for p in prev_list)
                if prev_conflict:
                    continue
                
                # 블랙리스트 체크 - 현재 학생과 방의 다른 학생이 블랙리스트에 있는지 확인
                blacklist_conflict = False
                for member in current_members:
                    pair = tuple(sorted([student, member]))
                    if pair in blacklist_set:
                        blacklist_conflict = True
                        break
                if blacklist_conflict:
                    continue

                # 조건을 만족하는 학생을 후보 리스트에 추가
                candidate_students.append(student)
            
            # 후보 학생이 있으면 유사도 기반으로 선택
            if candidate_students:
                if similarity_features and current_members:
                    # 방 구성원들의 특성 벡터 추출
                    room_members_features = []
                    for member in current_members:
                        member_features = get_student_features(df, member, similarity_features)
                        room_members_features.append(member_features)
                    
                    # 각 후보 학생의 유사도 점수 계산
                    candidate_scores = []
                    for candidate in candidate_students:
                        candidate_features = get_student_features(df, candidate, similarity_features)
                        similarity_score = calculate_room_similarity_score(room_members_features, candidate_features)
                        candidate_scores.append((candidate, similarity_score))
                    
                    # 유사도 점수가 높은 순으로 정렬
                    candidate_scores.sort(key=lambda x: x[1], reverse=True)
                    selected_student = candidate_scores[0][0]
                else:
                    # 유사도 기반 배정을 사용하지 않으면 첫 번째 후보 선택
                    selected_student = candidate_students[0]
                
                # 배정
                room[seat] = selected_student
                assigned.add(selected_student)
                remaining.remove(selected_student)
                current_members.append(selected_student)
                assigned_flag = True

            # fail
            if not assigned_flag:
                failed_students.append(f"{room_idx+1}번방 {seat}")

    return room_id, failed_students

