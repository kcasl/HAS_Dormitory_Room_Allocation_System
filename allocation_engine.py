import random as rd
import pandas as pd
import numpy as np


def allocate_rooms(excel_file_path):
    """
    기숙사 방 배정 알고리즘
    
    Args:
        excel_file_path: xlsx 파일 경로
        
    Returns:
        tuple: (room_id, failed_students) - 방 배정 결과와 실패한 좌석 목록
    """
    std_id = [i for i in range(100)]
    rd.shuffle(std_id)

    room_id = []
    for _ in range(25):
        room_id.append({"seat1": "", "seat2": "", "seat3": "", "seat4": ""})

    df = pd.read_excel(excel_file_path)

    hallway_seats = ["seat1", "seat4"]  # 1,4번
    window_seats = ["seat2", "seat3"]  # 2,3번

    for i in range(25):
        student = std_id[i] + 1

        prev_loc = df.loc[df["StudentID"] == student, "Prevloc"].values[0]

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
    remaining = [s+1 for s in std_id if (s+1) not in assigned]

    # 좌석 순서
    seat_order = ["seat1", "seat2", "seat3", "seat4"]

    for room_idx in range(25):
        room = room_id[room_idx]

        # 현재 방에 있는 사람
        current_members = [p for p in room.values() if p != ""]

        # seat2 → seat3 → seat4 순서
        for seat in seat_order:
            if room[seat] != "":
                continue  # 이미 할당된 자리면

            assigned_flag = False

            # 배정 가능 여부 췤
            for student in remaining:
                row = df.loc[df["StudentID"] == student]
                if row.empty:
                    continue

                # Prev1~Prev3, Avoid1~Avoid2 데이터 읽기
                prev_list = row[["Prev1", "Prev2", "Prev3"]].values[0]
                avoid_list = row[["Avoid1", "Avoid2"]].values[0]

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

                # 어싸인
                room[seat] = student
                assigned.add(student)
                remaining.remove(student)
                current_members.append(student)
                assigned_flag = True
                break

            # fail
            if not assigned_flag:
                failed_students.append(f"{room_idx+1}번방 {seat}")

    return room_id, failed_students

