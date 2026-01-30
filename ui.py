# ui.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import json
from preset import PRESET


class QuantitativeUI:
    def __init__(self, root):
        self.root = root
        self.root.title("양적기준분석")

        self.file_path = None
        self.output_dir_path = None
        self.rows = []

        self.build_ui()

    # -------------------------
    # UI 구성
    # -------------------------
    def build_ui(self):
        frame = ttk.Frame(self.root, padding=10)
        frame.pack(fill="both", expand=True)

        # =========================
        # 기본 정보 입력
        # =========================
        info = ttk.LabelFrame(frame, text="기본 정보", padding=10)
        info.grid(row=0, column=0, columnspan=4, sticky="w", pady=5)

        ttk.Label(info, text="법인명").grid(row=0, column=0, sticky="e")
        self.corp_name = ttk.Entry(info, width=20)
        self.corp_name.grid(row=0, column=1)

        ttk.Label(info, text="분석대상법인").grid(row=0, column=2, sticky="e")
        self.target_corp = ttk.Entry(info, width=20)
        self.target_corp.grid(row=0, column=3)

        ttk.Label(info, text="시작연도").grid(row=1, column=0, sticky="e")
        self.year_from = ttk.Entry(info, width=10)
        self.year_from.grid(row=1, column=1)

        ttk.Label(info, text="종료연도").grid(row=1, column=2, sticky="e")
        self.year_to = ttk.Entry(info, width=10)
        self.year_to.grid(row=1, column=3)
        

        ttk.Button(frame, text="Raw 파일 선택", command=self.select_file).grid(row=1, column=0, pady=5)
        self.file_label = ttk.Label(frame, text="선택된 파일 없음")
        self.file_label.grid(row=1, column=1, sticky="w")

        ttk.Button(frame, text="저장 경로 선택", command=self.select_output_dir).grid(row=2, column=0, pady=5)
        self.output_dir_label = ttk.Label(frame, text="선택된 폴더 없음")
        self.output_dir_label.grid(row=2, column=1, sticky="w")

        ttk.Button(frame, text="기준 추가", command=self.add_row).grid(row=3, column=0, pady=5)

        self.table = ttk.Frame(frame)
        self.table.grid(row=4, column=0, columnspan=4, sticky="w")

        headers = ["순번", "분석계정", "X값", "X비교", "포함", "삭제"]
        for i, h in enumerate(headers):
            ttk.Label(self.table, text=h).grid(row=0, column=i)

        ttk.Button(frame, text="변환", command=self.on_convert).grid(row=5, column=0, pady=10)

    # -------------------------
    # 파일 선택
    # -------------------------
    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_path:
            self.file_label.config(text=self.file_path)

    def select_output_dir(self):
        self.output_dir_path = filedialog.askdirectory()
        if self.output_dir_path:
            self.output_dir_label.config(text=self.output_dir_path)

    # -------------------------
    # 기준 row 추가
    # -------------------------
    def add_row(self):
        row_idx = len(self.rows) + 1

        seq = ttk.Label(self.table, text=str(row_idx))
        seq.grid(row=row_idx, column=0)

        account = ttk.Combobox(self.table, values=list(PRESET.keys()), width=20)
        account.grid(row=row_idx, column=1)
        account.bind("<<ComboboxSelected>>",
                     lambda e, r=row_idx-1: self.apply_preset(r))

        xvalue = ttk.Entry(self.table, width=10)
        xvalue.grid(row=row_idx, column=2)

        xcompare = ttk.Combobox(
            self.table,
            values=["미만", "초과", "같음", "텍스트 일치", "All equals", "존재함"],
            width=12
        )
        xcompare.grid(row=row_idx, column=3)


        include = ttk.Combobox(self.table, values=["포함", "제외"], width=20)
        include.grid(row=row_idx, column=4)
        include.bind("<<ComboboxSelected>>",
                     lambda e, r=row_idx-1: self.apply_preset(r))

        del_btn = ttk.Button(
            self.table, text="삭제",
            command=lambda r=row_idx-1: self.delete_row(r)
        )
        del_btn.grid(row=row_idx, column=5)

        # 버튼 레퍼런스 저장
        self.rows.append({
            "seq": seq,
            "account": account,
            "xValue": xvalue,
            "xCompare": xcompare,
            "include": include,
            "del_btn": del_btn
        })

    # -------------------------
    # preset 적용
    # -------------------------
    def apply_preset(self, idx):
        row = self.rows[idx]
        account = row["account"].get()
        preset = PRESET.get(account)

        if not preset:
            return

        row["xValue"].delete(0, tk.END)
        row["xValue"].insert(0, preset["xValue"])
        row["xCompare"].set(preset["xCompare"])
        row["include"].set(preset["include"])

    # -------------------------
    # row 삭제 + 순번 재정렬
    # -------------------------
    def delete_row(self, idx):
        # 해당 행의 위젯 제거
        row = self.rows[idx]
        row["seq"].destroy()
        row["account"].destroy()
        row["xValue"].destroy()
        row["xCompare"].destroy()
        row["include"].destroy()
        row["del_btn"].destroy()

        self.rows.pop(idx)
        self.resequence()

    def resequence(self):
        for i, row in enumerate(self.rows):
            # 순번 업데이트
            row["seq"].config(text=str(i+1))
            
            # Grid 위치 업데이트 (헤더가 row=3에 있고, 데이터는 row=4부터 시작한다고 가정하면 i+4 ?)
            # build_ui에서 headers가 row=3에 있고, add_row에서 row_idx = len + 1 (최초 1)
            # -> self.table 내에서의 row index는 i+1 이 맞음 (헤더 row=0)
            
            # self.table 내부 grid 재배치
            row["seq"].grid(row=i+1, column=0)
            row["account"].grid(row=i+1, column=1)
            row["xValue"].grid(row=i+1, column=2)
            row["xCompare"].grid(row=i+1, column=3)
            row["include"].grid(row=i+1, column=4)
            
            # 삭제 버튼 재배치 및 command 재바인딩
            del_btn = row["del_btn"] 
            del_btn.grid(row=i+1, column=5)
            del_btn.configure(command=lambda r=i: self.delete_row(r))
            
            # Combobox 바인딩 업데이트 
            row["account"].bind("<<ComboboxSelected>>", lambda e, r=i: self.apply_preset(r))
            row["include"].bind("<<ComboboxSelected>>", lambda e, r=i: self.apply_preset(r))

        # -------------------------
        # 변환 버튼
        # -------------------------
    def on_convert(self):
        # =========================
        # 1. 기본 입력값 수집
        # =========================
        corp_name = self.corp_name.get().strip()
        target_corp = self.target_corp.get().strip()
        year_from = self.year_from.get().strip()
        year_to = self.year_to.get().strip()

        if not corp_name or not target_corp or not year_from or not year_to:
            messagebox.showerror("입력 오류", "기본 정보를 모두 입력하세요.")
            return

        if not self.file_path:
            messagebox.showerror("입력 오류", "Raw 파일을 선택하세요.")
            return

        if not self.output_dir_path:
            messagebox.showerror("입력 오류", "저장 경로를 선택하세요.")
            return

        try:
            year_from = int(year_from)
            year_to = int(year_to)
        except ValueError:
            messagebox.showerror("입력 오류", "연도는 숫자로 입력해야 합니다.")
            return

        input_data = {
            "corpName": corp_name,
            "targetCorp": target_corp,
            "yearFrom": year_from,
            "yearTo": year_to,
            "rawFilePath": self.file_path,
            "outputDir": self.output_dir_path
        }

        # =========================
        # 2. 기준 테이블 수집
        # =========================
        criteria_list = []

        for idx, row in enumerate(self.rows, start=1):
            account = row["account"].get()
            x_value = row["xValue"].get()
            x_compare = row["xCompare"].get()
            include = row["include"].get() == "포함"

            if not account:
                continue  # 계정 선택 안 된 row는 무시

            criteria_list.append({
                "seq": idx,
                "account": account,
                "xValue": x_value,
                "xCompare": x_compare,
                "include": include
            })

        if not criteria_list:
            messagebox.showerror("입력 오류", "최소 1개의 기준을 입력하세요.")
            return

        # =========================
        # 3. 최종 payload
        # =========================
        payload = {
            "inputData": input_data,
            "criteriaList": criteria_list
        }
    
        print(payload)

        from processor import main_processor
        main_processor(payload)

