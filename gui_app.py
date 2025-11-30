import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from allocation_engine import allocate_rooms


class DormitoryAllocationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("기숙사 방 배정 시스템")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 선택된 파일 경로
        self.selected_file = None
        
        self.setup_ui()
        
    def setup_ui(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 그리드 가중치 설정
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="기숙사 방 배정 시스템", 
            font=("맑은 고딕", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 파일 선택 섹션
        file_frame = ttk.LabelFrame(main_frame, text="Excel 파일 선택", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="파일 경로:").grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        
        self.file_path_var = tk.StringVar(value="파일을 선택해주세요")
        file_path_label = ttk.Label(
            file_frame, 
            textvariable=self.file_path_var, 
            foreground="gray",
            wraplength=400
        )
        file_path_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        browse_button = ttk.Button(
            file_frame, 
            text="파일 선택", 
            command=self.browse_file
        )
        browse_button.grid(row=0, column=2)
        
        # 실행 버튼
        run_button = ttk.Button(
            main_frame,
            text="배정 실행",
            command=self.run_allocation,
            state="disabled"
        )
        run_button.grid(row=1, column=2, padx=(10, 0), sticky=(tk.W, tk.E))
        self.run_button = run_button
        
        # 결과 표시 섹션
        result_frame = ttk.LabelFrame(main_frame, text="배정 결과", padding="10")
        result_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        # 노트북 (탭) 생성
        notebook = ttk.Notebook(result_frame)
        notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 탭 1: 방 배정 결과
        room_frame = ttk.Frame(notebook, padding="10")
        notebook.add(room_frame, text="방 배정 결과")
        
        # 스크롤 가능한 텍스트 영역
        self.room_text = scrolledtext.ScrolledText(
            room_frame, 
            wrap=tk.WORD, 
            width=80, 
            height=25,
            font=("맑은 고딕", 10)
        )
        self.room_text.pack(fill=tk.BOTH, expand=True)
        
        # 탭 2: 실패 목록
        failed_frame = ttk.Frame(notebook, padding="10")
        notebook.add(failed_frame, text="배정 실패 목록")
        
        self.failed_text = scrolledtext.ScrolledText(
            failed_frame, 
            wrap=tk.WORD, 
            width=80, 
            height=25,
            font=("맑은 고딕", 10),
            foreground="red"
        )
        self.failed_text.pack(fill=tk.BOTH, expand=True)
        
        # 상태바
        self.status_var = tk.StringVar(value="파일을 선택해주세요")
        status_bar = ttk.Label(
            main_frame, 
            textvariable=self.status_var, 
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding="5"
        )
        status_bar.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def browse_file(self):
        """파일 선택 다이얼로그"""
        file_path = filedialog.askopenfilename(
            title="Excel 파일 선택",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.selected_file = file_path
            filename = os.path.basename(file_path)
            self.file_path_var.set(filename)
            self.run_button.config(state="normal")
            self.status_var.set(f"파일 선택됨: {filename}")
            
    def run_allocation(self):
        """배정 알고리즘 실행"""
        if not self.selected_file:
            messagebox.showerror("오류", "파일을 선택해주세요.")
            return
            
        if not os.path.exists(self.selected_file):
            messagebox.showerror("오류", "선택한 파일이 존재하지 않습니다.")
            return
        
        try:
            self.status_var.set("배정 중...")
            self.root.update()
            
            # 배정 알고리즘 실행
            room_id, failed_students = allocate_rooms(self.selected_file)
            
            # 결과 표시
            self.display_results(room_id, failed_students)
            
            self.status_var.set(f"배정 완료! (실패: {len(failed_students)}개)")
            
        except FileNotFoundError:
            messagebox.showerror("오류", "파일을 찾을 수 없습니다.")
            self.status_var.set("오류 발생")
        except Exception as e:
            messagebox.showerror("오류", f"배정 중 오류가 발생했습니다:\n{str(e)}")
            self.status_var.set("오류 발생")
            
    def display_results(self, room_id, failed_students):
        """배정 결과를 텍스트 영역에 표시"""
        # 방 배정 결과 탭 초기화
        self.room_text.delete(1.0, tk.END)
        
        # 방 배정 결과 출력
        self.room_text.insert(tk.END, "=" * 70 + "\n")
        self.room_text.insert(tk.END, "최종 방 배정 결과\n")
        self.room_text.insert(tk.END, "=" * 70 + "\n\n")
        
        for i, room in enumerate(room_id, start=1):
            room_info = f"방 {i:2d}번: "
            seats = []
            for seat_name in ["seat1", "seat2", "seat3", "seat4"]:
                student_id = room[seat_name]
                if student_id:
                    seats.append(f"{seat_name}={student_id}")
                else:
                    seats.append(f"{seat_name}=빈자리")
            room_info += " | ".join(seats)
            self.room_text.insert(tk.END, room_info + "\n")
            
            # 5개 방마다 구분선
            if i % 5 == 0:
                self.room_text.insert(tk.END, "-" * 70 + "\n")
        
        # 배정 실패 목록 탭 초기화
        self.failed_text.delete(1.0, tk.END)
        
        if failed_students:
            self.failed_text.insert(tk.END, "=" * 70 + "\n")
            self.failed_text.insert(tk.END, f"배정 실패 좌석 목록 (총 {len(failed_students)}개)\n")
            self.failed_text.insert(tk.END, "=" * 70 + "\n\n")
            
            for failed in failed_students:
                self.failed_text.insert(tk.END, f"  - {failed}\n")
        else:
            self.failed_text.insert(tk.END, "=" * 70 + "\n")
            self.failed_text.insert(tk.END, "배정 실패한 좌석이 없습니다!\n")
            self.failed_text.insert(tk.END, "=" * 70 + "\n")


def main():
    root = tk.Tk()
    app = DormitoryAllocationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

