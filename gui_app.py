import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
import pandas as pd
from datetime import datetime
from allocation_engine import allocate_rooms

# 플랫폼별 폰트 설정
if sys.platform == "win32":
    DEFAULT_FONT = ("Malgun Gothic",)
    DEFAULT_FONT_SMALL = ("Malgun Gothic",)
else:
    DEFAULT_FONT = ("맑은 고딕",)
    DEFAULT_FONT_SMALL = ("맑은 고딕",)


class DormitoryAllocationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("기숙사 방 배정 시스템")
        self.root.geometry("1200x900")
        self.root.resizable(True, True)

        # 배경색 설정 (윈도우 호환)
        try:
            self.root.configure(bg="#f5f5f5")
        except:
            pass  # 일부 시스템에서 색상 설정이 실패할 수 있음

        # 선택된 파일 경로
        self.selected_file = None

        # 블랙리스트 조합 저장 (튜플의 리스트)
        self.blacklist_pairs = []

        # 배정 결과 저장 (나중에 엑셀로 저장하기 위해)
        self.current_room_id = None
        self.current_failed_students = None
        self.student_name_map = {}  # 학번-이름 매핑 딕셔너리

        # Factor 체크박스 변수들
        self.factor_vars = {}
        self.available_factors = []

        # 스크롤 가능한 메인 UI 구성
        self.setup_scrollable_ui()

        # UI 요소 구성
        self.setup_ui_components()

    def setup_scrollable_ui(self):
        """
        전체 GUI를 스크롤 가능하게 만들기 위한 캔버스+프레임 구조 생성
        """
        # 외부 컨테이너 (root에 붙음)
        outer_frame = ttk.Frame(self.root)
        outer_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # 캔버스와 수직 스크롤바
        self.canvas = tk.Canvas(outer_frame, borderwidth=0, highlightthickness=0)
        self.v_scroll = ttk.Scrollbar(outer_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)

        self.v_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.canvas.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        outer_frame.columnconfigure(0, weight=1)
        outer_frame.rowconfigure(0, weight=1)

        # 내부 프레임 (실제 위젯들이 여기에 추가됨)
        self.main_frame = ttk.Frame(self.canvas, padding="30")
        # 윈도우로 캔버스에 붙이기
        self.canvas_window = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # 바인딩: 내부 프레임의 크기가 바뀔 때 캔버스 scrollregion 업데이트
        self.main_frame.bind("<Configure>", self._on_frame_configure)
        # 윈도우(캔버스) 크기가 바뀔 때 내부 프레임 폭을 조정 (responsive)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 마우스 휠 바인딩 (플랫폼별)
        self._bind_mousewheel(self.canvas)

    def _bind_mousewheel(self, widget):
        """마우스 휠을 스크롤에 연결 (Windows/macOS/Linux)"""
        if sys.platform == "darwin":
            # macOS
            widget.bind_all("<MouseWheel>", self._on_mousewheel_mac)
        elif sys.platform == "win32":
            widget.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        else:
            # X11
            widget.bind_all("<Button-4>", self._on_mousewheel_linux)
            widget.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _on_mousewheel_windows(self, event):
        # Windows: event.delta (120 단위)
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_mac(self, event):
        # macOS: event.delta (작음)
        self.canvas.yview_scroll(int(-1 * event.delta), "units")

    def _on_mousewheel_linux(self, event):
        # Linux: Button-4 up, Button-5 down
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")

    def _on_frame_configure(self, event):
        """내부 프레임이 바뀔 때 캔버스의 scrollregion을 업데이트"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """캔버스 크기 변경 시 내부 프레임 폭을 캔버스에 맞춤 (responsive)"""
        canvas_width = event.width
        # main_frame의 내부 위젯들이 잘 보이도록 main_frame의 폭을 캔버스 너비에 맞춤
        try:
            self.canvas.itemconfig(self.canvas_window, width=canvas_width)
        except Exception:
            pass

    def setup_ui_components(self):
        """
        기존 setup_ui 내용을 main_frame 기반으로 재구성
        """
        try:
            style = ttk.Style()
            style.configure("Gray.TLabel", foreground="gray")
            style.configure("Desc.TLabel", foreground="gray")
        except:
            pass

        # 그리드 가중치 설정 (main_frame는 이미 캔버스 내부)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(5, weight=1)

        # 제목 영역
        title_frame = ttk.Frame(self.main_frame)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 40))

        title_label = ttk.Label(
            title_frame,
            text="🏠 하나고등학교 기숙사 방 배정 시스템",
            font=(DEFAULT_FONT[0], 20, "bold")
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            title_frame,
            text="Excel 파일(xlsx)을 업로드하여 자동으로 방을 배정합니다",
            font=(DEFAULT_FONT_SMALL[0], 10)
        )
        subtitle_label.pack(pady=(5, 0))
        try:
            subtitle_label.configure(style="Gray.TLabel")
        except:
            pass

        # 파일 선택 및 실행 섹션
        control_frame = ttk.LabelFrame(
            self.main_frame,
            text=" 파일 선택 및 실행 ",
            padding="25"
        )
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 25))
        control_frame.columnconfigure(1, weight=1)

        # 파일 선택 영역
        file_select_frame = ttk.Frame(control_frame)
        file_select_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        file_select_frame.columnconfigure(1, weight=1)

        ttk.Label(
            file_select_frame,
            text="Excel 파일:",
            font=(DEFAULT_FONT[0], 11)
        ).grid(row=0, column=0, padx=(0, 20), sticky=tk.W)

        self.file_path_var = tk.StringVar(value="파일을 선택해주세요")
        file_path_entry = ttk.Entry(
            file_select_frame,
            textvariable=self.file_path_var,
            state="readonly",
            font=(DEFAULT_FONT_SMALL[0], 10),
            width=55
        )
        file_path_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 15))

        browse_button = ttk.Button(
            file_select_frame,
            text="📁 파일 선택",
            command=self.browse_file,
            width=18
        )
        browse_button.grid(row=0, column=2)

        # 실행 버튼 영역
        button_frame = ttk.Frame(control_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=(15, 0))

        self.run_button = ttk.Button(
            button_frame,
            text="▶ 배정 실행",
            command=self.run_allocation,
            state="disabled",
            width=25
        )
        self.run_button.pack()

        # 블랙리스트 관리 섹션
        blacklist_frame = ttk.LabelFrame(
            self.main_frame,
            text=" 🚫 블랙리스트 조합 관리 ",
            padding="20"
        )
        blacklist_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 25))
        blacklist_frame.columnconfigure(1, weight=1)

        # 설명
        desc_label = ttk.Label(
            blacklist_frame,
            text="같은 방에 배정되지 않아야 하는 학생 조합을 추가하세요 (예: 학생1과 학생2)",
            font=(DEFAULT_FONT_SMALL[0], 9)
        )
        desc_label.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 15))
        try:
            desc_label.configure(style="Desc.TLabel")
        except:
            pass

        # 입력 영역
        input_frame = ttk.Frame(blacklist_frame)
        input_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(1, weight=1)
        input_frame.columnconfigure(3, weight=1)

        ttk.Label(input_frame, text="학생 ID 1:", font=(DEFAULT_FONT_SMALL[0], 10)).grid(row=0, column=0, padx=(0, 10))
        self.blacklist_student1_var = tk.StringVar()
        student1_entry = ttk.Entry(input_frame, textvariable=self.blacklist_student1_var, width=12, font=(DEFAULT_FONT_SMALL[0], 10))
        student1_entry.grid(row=0, column=1, padx=(0, 20))

        ttk.Label(input_frame, text="학생 ID 2:", font=(DEFAULT_FONT_SMALL[0], 10)).grid(row=0, column=2, padx=(0, 10))
        self.blacklist_student2_var = tk.StringVar()
        student2_entry = ttk.Entry(input_frame, textvariable=self.blacklist_student2_var, width=12, font=(DEFAULT_FONT_SMALL[0], 10))
        student2_entry.grid(row=0, column=3, padx=(0, 15))

        add_blacklist_button = ttk.Button(
            input_frame,
            text="추가",
            command=self.add_blacklist_pair,
            width=12
        )
        add_blacklist_button.grid(row=0, column=4)

        # 블랙리스트 목록 표시
        list_frame = ttk.Frame(blacklist_frame)
        list_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # 리스트박스와 스크롤바
        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        listbox_frame.columnconfigure(0, weight=1)
        listbox_frame.rowconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        self.blacklist_listbox = tk.Listbox(
            listbox_frame,
            font=(DEFAULT_FONT_SMALL[0], 10),
            height=5,
            yscrollcommand=scrollbar.set
        )
        self.blacklist_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.config(command=self.blacklist_listbox.yview)

        # 삭제 버튼
        delete_button = ttk.Button(
            list_frame,
            text="선택 항목 삭제",
            command=self.remove_blacklist_pair,
            width=18
        )
        delete_button.grid(row=1, column=0, pady=(15, 0))

        # Factor 선택 섹션
        factor_frame = ttk.LabelFrame(
            self.main_frame,
            text=" 📊 Factor 선택 (유사도 기반 배정) ",
            padding="20"
        )
        factor_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 25))
        factor_frame.columnconfigure(0, weight=1)

        factor_desc_label = ttk.Label(
            factor_frame,
            text="파일을 선택하면 사용 가능한 factor들이 표시됩니다. 체크한 factor들로 유사한 학생들끼리 배정됩니다.",
            font=(DEFAULT_FONT_SMALL[0], 9)
        )
        factor_desc_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 15))
        try:
            factor_desc_label.configure(style="Desc.TLabel")
        except:
            pass

        # Factor 체크박스 영역
        self.factor_checkbox_frame = ttk.Frame(factor_frame)
        self.factor_checkbox_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # 결과 표시 섹션
        result_frame = ttk.LabelFrame(
            self.main_frame,
            text=" 배정 결과 ",
            padding="20"
        )
        result_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(1, weight=1)

        # 저장 버튼 영역
        save_button_frame = ttk.Frame(result_frame)
        save_button_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        save_button_frame.columnconfigure(0, weight=1)

        button_container = ttk.Frame(save_button_frame)
        button_container.grid(row=0, column=0)

        self.save_button = ttk.Button(
            button_container,
            text="💾 배정 결과를 엑셀로 저장",
            command=self.save_to_excel,
            state="disabled",
            width=25
        )
        self.save_button.pack()

        # 노트북 (탭) 생성
        notebook = ttk.Notebook(result_frame)
        notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 탭 1: 방 배정 결과
        room_frame = ttk.Frame(notebook, padding="20")
        notebook.add(room_frame, text="📋 방 배정 결과")
        room_frame.columnconfigure(0, weight=1)
        room_frame.rowconfigure(0, weight=1)

        # 스크롤 가능한 텍스트 영역 (내부 탭은 자체 스크롤 보유)
        self.room_text = scrolledtext.ScrolledText(
            room_frame,
            wrap=tk.WORD,
            width=95,
            height=32,
            font=(DEFAULT_FONT_SMALL[0], 10),
            relief=tk.FLAT,
            borderwidth=1
        )
        try:
            self.room_text.configure(bg="white")
        except:
            pass
        self.room_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 탭 2: 실패 목록
        failed_frame = ttk.Frame(notebook, padding="20")
        notebook.add(failed_frame, text="⚠ 배정 실패 목록")
        failed_frame.columnconfigure(0, weight=1)
        failed_frame.rowconfigure(0, weight=1)

        self.failed_text = scrolledtext.ScrolledText(
            failed_frame,
            wrap=tk.WORD,
            width=95,
            height=32,
            font=(DEFAULT_FONT_SMALL[0], 10),
            relief=tk.FLAT,
            borderwidth=1
        )
        try:
            self.failed_text.configure(bg="white", foreground="#d32f2f")
        except:
            try:
                self.failed_text.configure(bg="white", foreground="red")
            except:
                pass
        self.failed_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 상태바
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(20, 0))

        self.status_var = tk.StringVar(value="준비됨 - 파일을 선택해주세요")
        status_bar = ttk.Label(
            status_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding="10",
            font=(DEFAULT_FONT_SMALL[0], 9)
        )
        status_bar.pack(fill=tk.X)

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

            # Factor 컬럼 감지 및 체크박스 생성
            self.detect_and_create_factor_checkboxes(file_path)

            self.status_var.set(f"✓ 파일 선택됨: {filename} - Factor를 선택하고 배정 실행 버튼을 클릭하세요")

    def detect_and_create_factor_checkboxes(self, file_path):
        """엑셀 파일에서 factor 컬럼들을 감지하고 체크박스 생성
        '현재 룸메이트 3' 컬럼 이후의 모든 컬럼 중에서
        실수 또는 정수형 데이터를 가진 컬럼을 factor로 인식"""
        try:
            # 기존 체크박스 제거
            for widget in self.factor_checkbox_frame.winfo_children():
                widget.destroy()
            self.factor_vars.clear()
            self.available_factors.clear()
            
            # 엑셀 파일 읽기 (데이터 타입 확인을 위해 여러 행 읽기)
            df = pd.read_excel(file_path)
            
            # "현재 룸메이트 3" 컬럼의 인덱스 찾기
            target_column = "현재 룸메이트 3"
            if target_column not in df.columns:
                # "현재 룸메이트 3" 컬럼이 없으면 안내 메시지
                no_column_label = ttk.Label(
                    self.factor_checkbox_frame,
                    text=f"'{target_column}' 컬럼을 찾을 수 없습니다.",
                    font=(DEFAULT_FONT_SMALL[0], 9),
                    foreground="gray"
                )
                no_column_label.grid(row=0, column=0, sticky=tk.W)
                return
            
            # "현재 룸메이트 3" 컬럼의 인덱스 찾기
            target_idx = df.columns.get_loc(target_column)
            
            # "현재 룸메이트 3" 이후의 모든 컬럼 확인
            for col_idx in range(target_idx + 1, len(df.columns)):
                col_name = df.columns[col_idx]
                
                # 해당 컬럼의 데이터 타입 확인
                col_data = df[col_name].dropna()
                
                if len(col_data) == 0:
                    continue  # 데이터가 없으면 스킵
                
                # pandas의 숫자형 타입인지 확인
                is_numeric = pd.api.types.is_numeric_dtype(df[col_name])
                
                # 숫자형이 아니더라도 모든 값이 실수 또는 정수로 변환 가능한지 확인
                if not is_numeric:
                    is_numeric = True
                    for value in col_data:
                        try:
                            # 실수 또는 정수로 변환 가능한지 확인
                            float(value)
                        except (ValueError, TypeError):
                            is_numeric = False
                            break
                
                # 실수 또는 정수형 데이터인 경우 factor로 추가
                if is_numeric:
                    self.available_factors.append(col_name)
            
            if self.available_factors:
                # 체크박스 생성 (3열로 배치)
                cols_per_row = 3
                for idx, factor in enumerate(self.available_factors):
                    row = idx // cols_per_row
                    col = idx % cols_per_row
                    
                    var = tk.BooleanVar(value=True)  # 기본적으로 모두 체크
                    self.factor_vars[factor] = var
                    
                    checkbox = ttk.Checkbutton(
                        self.factor_checkbox_frame,
                        text=factor,
                        variable=var
                    )
                    checkbox.grid(row=row, column=col, sticky=tk.W, padx=15, pady=8)
            else:
                # Factor가 없으면 안내 메시지
                no_factor_label = ttk.Label(
                    self.factor_checkbox_frame,
                    text="'현재 룸메이트 3' 이후에 숫자형 데이터를 가진 컬럼이 없습니다.",
                    font=(DEFAULT_FONT_SMALL[0], 9),
                    foreground="gray"
                )
                no_factor_label.grid(row=0, column=0, sticky=tk.W)
                
        except Exception as e:
            # 오류 발생 시 메시지 표시
            error_label = ttk.Label(
                self.factor_checkbox_frame,
                text=f"Factor 감지 중 오류: {str(e)}",
                font=(DEFAULT_FONT_SMALL[0], 9),
                foreground="red"
            )
            error_label.grid(row=0, column=0, sticky=tk.W)

    def add_blacklist_pair(self):
        """블랙리스트 조합 추가"""
        try:
            student1 = int(self.blacklist_student1_var.get().strip())
            student2 = int(self.blacklist_student2_var.get().strip())

            if student1 == student2:
                messagebox.showwarning("경고", "같은 학생 ID를 입력할 수 없습니다.")
                return

            if student1 < 1 or student1 > 100 or student2 < 1 or student2 > 100:
                messagebox.showwarning("경고", "학생 ID는 1부터 100 사이의 숫자여야 합니다.")
                return

            # 정렬하여 중복 체크
            pair = tuple(sorted([student1, student2]))

            # 중복 체크
            if pair in self.blacklist_pairs:
                messagebox.showinfo("알림", "이미 추가된 조합입니다.")
                return

            # 추가
            self.blacklist_pairs.append(pair)
            self.update_blacklist_display()

            # 입력 필드 초기화
            self.blacklist_student1_var.set("")
            self.blacklist_student2_var.set("")

            self.status_var.set(f"블랙리스트 추가됨: 학생{student1} ↔ 학생{student2} (총 {len(self.blacklist_pairs)}개)")

        except ValueError:
            messagebox.showerror("오류", "올바른 숫자를 입력해주세요.")

    def remove_blacklist_pair(self):
        """블랙리스트 조합 삭제"""
        selection = self.blacklist_listbox.curselection()
        if not selection:
            messagebox.showwarning("경고", "삭제할 항목을 선택해주세요.")
            return

        index = selection[0]
        removed_pair = self.blacklist_pairs.pop(index)
        self.update_blacklist_display()
        self.status_var.set(f"블랙리스트 삭제됨: 학생{removed_pair[0]} ↔ 학생{removed_pair[1]} (총 {len(self.blacklist_pairs)}개)")

    def update_blacklist_display(self):
        """블랙리스트 목록 업데이트"""
        self.blacklist_listbox.delete(0, tk.END)
        for pair in self.blacklist_pairs:
            self.blacklist_listbox.insert(tk.END, f"학생{pair[0]} ↔ 학생{pair[1]}")

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

            # 선택된 factor들 추출
            selected_factors = []
            for factor, var in self.factor_vars.items():
                if var.get():
                    selected_factors.append(factor)

            # 배정 알고리즘 실행 (블랙리스트 및 선택된 factor 포함)
            room_id, failed_students = allocate_rooms(
                self.selected_file,
                self.blacklist_pairs,
                selected_factors if selected_factors else None
            )

            # 엑셀 파일에서 학번-이름 매핑 생성
            df = pd.read_excel(self.selected_file)
            if "이름" in df.columns:
                self.student_name_map = dict(zip(df["학번"], df["이름"]))
            else:
                # "이름" 컬럼이 없으면 학번만 사용
                self.student_name_map = {sid: str(sid) for sid in df["학번"].dropna()}

            # 배정 결과 저장 (엑셀 저장용)
            self.current_room_id = room_id
            self.current_failed_students = failed_students

            # 결과 표시
            self.display_results(room_id, failed_students)

            # 저장 버튼 활성화
            self.save_button.config(state="normal")

            self.status_var.set(f"배정 완료! (실패: {len(failed_students)}개) - 엑셀로 저장 가능")

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
        header = "=" * 85
        self.room_text.insert(tk.END, header + "\n")
        self.room_text.insert(tk.END, " " * 30 + "최종 방 배정 결과\n")
        self.room_text.insert(tk.END, header + "\n\n")

        # 방 배정 결과를 표 형식으로 출력
        for i, room in enumerate(room_id, start=1):
            room_info = f"방 {i:2d}번"
            seats_info = []
            for seat_name in ["seat1", "seat2", "seat3", "seat4"]:
                student_id = room[seat_name]
                seat_num = seat_name.replace("seat", "")
                if student_id:
                    # 학번-이름 형식으로 출력
                    student_name = self.student_name_map.get(student_id, str(student_id))
                    if student_name != str(student_id):
                        seats_info.append(f"좌석{seat_num}: {student_id}-{student_name}")
                    else:
                        seats_info.append(f"좌석{seat_num}: {student_id}")
                else:
                    seats_info.append(f"좌석{seat_num}: 빈자리  ")

            # 더 읽기 쉬운 형식으로 출력
            self.room_text.insert(tk.END, f"{room_info:8s} │ {' │ '.join(seats_info)}\n")

            # 5개 방마다 구분선
            if i % 5 == 0 and i < len(room_id):
                self.room_text.insert(tk.END, "-" * 85 + "\n")

        # 배정 실패 목록 탭 초기화
        self.failed_text.delete(1.0, tk.END)

        if failed_students:
            header = "=" * 85
            self.failed_text.insert(tk.END, header + "\n")
            self.failed_text.insert(tk.END, f" " * 25 + f"배정 실패 좌석 목록 (총 {len(failed_students)}개)\n")
            self.failed_text.insert(tk.END, header + "\n\n")

            for idx, failed in enumerate(failed_students, start=1):
                self.failed_text.insert(tk.END, f"  {idx:2d}. {failed}\n")
        else:
            header = "=" * 85
            self.failed_text.insert(tk.END, header + "\n")
            self.failed_text.insert(tk.END, " " * 30 + "✓ 배정 실패한 좌석이 없습니다!\n")
            self.failed_text.insert(tk.END, " " * 25 + "모든 학생이 성공적으로 배정되었습니다.\n")
            self.failed_text.insert(tk.END, header + "\n")

    def save_to_excel(self):
        """배정 결과를 엑셀 파일로 저장"""
        if self.current_room_id is None:
            messagebox.showwarning("경고", "저장할 배정 결과가 없습니다.")
            return

        # 파일 저장 다이얼로그
        default_filename = f"방배정결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="엑셀 파일로 저장",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ],
            initialfile=default_filename
        )

        if not file_path:
            return  # 사용자가 취소한 경우

        try:
            self.status_var.set("엑셀 파일 저장 중...")
            self.root.update()

            # 배정 결과를 DataFrame으로 변환 (학번과 이름을 별도로 저장)
            room_data = []
            for i, room in enumerate(self.current_room_id, start=1):
                # 각 좌석에 대해 학번과 이름을 별도로 저장
                def get_student_info(seat_value):
                    if seat_value:
                        student_id = seat_value
                        student_name = self.student_name_map.get(student_id, "")
                        return student_id, student_name if student_name else ""
                    return "", ""
                
                seat1_id, seat1_name = get_student_info(room["seat1"])
                seat2_id, seat2_name = get_student_info(room["seat2"])
                seat3_id, seat3_name = get_student_info(room["seat3"])
                seat4_id, seat4_name = get_student_info(room["seat4"])
                
                room_data.append({
                    "방 번호": i,
                    "좌석1_학번": seat1_id if seat1_id else "",
                    "좌석1_이름": seat1_name,
                    "좌석2_학번": seat2_id if seat2_id else "",
                    "좌석2_이름": seat2_name,
                    "좌석3_학번": seat3_id if seat3_id else "",
                    "좌석3_이름": seat3_name,
                    "좌석4_학번": seat4_id if seat4_id else "",
                    "좌석4_이름": seat4_name
                })

            df_rooms = pd.DataFrame(room_data)

            # 배정 실패 목록을 DataFrame으로 변환
            if self.current_failed_students:
                failed_data = []
                for idx, failed in enumerate(self.current_failed_students, start=1):
                    failed_data.append({
                        "번호": idx,
                        "실패 좌석": failed
                    })
                df_failed = pd.DataFrame(failed_data)
            else:
                df_failed = pd.DataFrame({"번호": [], "실패 좌석": []})

            # 엑셀 파일로 저장 (여러 시트 사용)
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # 시트 1: 방 배정 결과
                df_rooms.to_excel(writer, sheet_name='방 배정 결과', index=False)

                # 시트 2: 배정 실패 목록
                df_failed.to_excel(writer, sheet_name='배정 실패 목록', index=False)

                # 시트 3: 배정 정보 요약
                summary_data = {
                    "항목": [
                        "배정 일시",
                        "총 방 수",
                        "총 좌석 수",
                        "배정된 학생 수",
                        "배정 실패 좌석 수",
                        "사용된 Factor",
                        "블랙리스트 조합 수"
                    ],
                    "내용": [
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        len(self.current_room_id),
                        len(self.current_room_id) * 4,
                        sum(1 for room in self.current_room_id for seat in room.values() if seat),
                        len(self.current_failed_students),
                        ", ".join(self.available_factors) if self.available_factors else "없음",
                        len(self.blacklist_pairs)
                    ]
                }
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name='배정 정보', index=False)

                # 시트별 컬럼 너비 조정
                worksheet_rooms = writer.sheets['방 배정 결과']
                worksheet_failed = writer.sheets['배정 실패 목록']
                worksheet_summary = writer.sheets['배정 정보']

                # 방 배정 결과 시트 컬럼 너비 조정
                for column in worksheet_rooms.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet_rooms.column_dimensions[column_letter].width = adjusted_width

                # 배정 실패 목록 시트 컬럼 너비 조정
                for column in worksheet_failed.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet_failed.column_dimensions[column_letter].width = adjusted_width

                # 배정 정보 시트 컬럼 너비 조정
                for column in worksheet_summary.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet_summary.column_dimensions[column_letter].width = adjusted_width

            filename = os.path.basename(file_path)
            self.status_var.set(f"✓ 엑셀 파일 저장 완료: {filename}")
            messagebox.showinfo("저장 완료", f"배정 결과가 성공적으로 저장되었습니다.\n\n파일: {filename}")

        except Exception as e:
            messagebox.showerror("오류", f"엑셀 파일 저장 중 오류가 발생했습니다:\n{str(e)}")
            self.status_var.set("엑셀 파일 저장 실패")


def main():
    root = tk.Tk()
    app = DormitoryAllocationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
