# ui.py
import logging
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from preset import PRESET

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)


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
        info.grid(row=0, column=0, columnspan=9, sticky="w", pady=5)

        ttk.Label(info, text="클라이언트명").grid(row=0, column=0, sticky="e")
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
        self.file_label.grid(row=1, column=1, columnspan=8, sticky="w")

        ttk.Button(frame, text="기준 추가", command=self.add_row).grid(row=3, column=0, pady=5)

        self.table = ttk.Frame(frame)
        self.table.grid(row=4, column=0, columnspan=9, sticky="w")

        headers = ["순번", "유형", "분석계정", "기준값", "비교연산자", "연도조건", "N", "포함", "삭제"]
        for i, h in enumerate(headers):
            ttk.Label(self.table, text=h, width=10 if i not in (0, 6, 8) else 4).grid(
                row=0, column=i, padx=2
            )

        ttk.Button(frame, text="변환", command=self.on_convert).grid(row=5, column=0, pady=10)

        # -------------------------
        # 설명 문구 (하단)
        # -------------------------
        desc_frame = ttk.LabelFrame(frame, text="사용 가이드", padding=10)
        desc_frame.grid(row=6, column=0, columnspan=9, sticky="ew", pady=10)

        guide_text = (
            "1. Raw 파일을 선택하세요. (기본 정보 입력 필수)\n"
            "2. 유형 → 분석계정 → 비교연산자 → 기준값 순서로 설정하세요.\n"
            "   비율계정의 기준값은 소수점으로 입력합니다. (예: 0.01 → 1%)\n"
            "3. '변환' 버튼을 누르면 분석 결과가 입력 파일과 동일한 폴더에 저장됩니다.\n"
            "   (파일명: [클라이언트명]_양적분석_[기간].xlsx)"
        )
        ttk.Label(desc_frame, text=guide_text).pack(anchor="w")

    # -------------------------
    # 파일 선택
    # -------------------------
    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.file_path:
            self.file_label.config(text=self.file_path)

    # -------------------------
    # 기준 row 추가 (9컬럼)
    # -------------------------
    def add_row(self):
        from criteria_config import TYPE_DISPLAY_NAMES, CRITERIA_TYPES, COUNT_REQUIREMENT_OPTIONS

        row_idx = len(self.rows) + 1

        seq = ttk.Label(self.table, text=str(row_idx), width=4)
        seq.grid(row=row_idx, column=0, padx=2)

        # col 1: 유형
        type_cb = ttk.Combobox(self.table, values=TYPE_DISPLAY_NAMES, width=14, state="readonly")
        type_cb.grid(row=row_idx, column=1, padx=2)

        # col 2: 분석계정 (유형에 따라 동적 갱신)
        account_cb = ttk.Combobox(self.table, values=[], width=22, state="readonly")
        account_cb.grid(row=row_idx, column=2, padx=2)

        # col 3: 기준값
        xvalue = ttk.Entry(self.table, width=12)
        xvalue.grid(row=row_idx, column=3, padx=2)

        # col 4: 비교연산자 (유형에 따라 동적 갱신)
        xcompare = ttk.Combobox(self.table, values=[], width=14, state="readonly")
        xcompare.grid(row=row_idx, column=4, padx=2)

        # col 5: 연도조건 (숫자-개별연도일 때만 표시)
        year_cond = ttk.Combobox(
            self.table, values=COUNT_REQUIREMENT_OPTIONS, width=14, state="readonly"
        )
        year_cond.grid(row=row_idx, column=5, padx=2)
        year_cond.grid_remove()

        # col 6: N개년 입력 (N개년이상 선택 시만 표시)
        n_years = ttk.Entry(self.table, width=4)
        n_years.grid(row=row_idx, column=6, padx=2)
        n_years.grid_remove()

        # col 7: 포함/제외
        include = ttk.Combobox(self.table, values=["포함", "제외"], width=7, state="readonly")
        include.grid(row=row_idx, column=7, padx=2)

        # col 8: 삭제
        del_btn = ttk.Button(
            self.table, text="삭제",
            command=lambda r=row_idx - 1: self.delete_row(r),
            width=4,
        )
        del_btn.grid(row=row_idx, column=8, padx=2)

        row_dict = {
            "seq": seq,
            "type": type_cb,
            "account": account_cb,
            "xValue": xvalue,
            "xCompare": xcompare,
            "yearCond": year_cond,
            "nYears": n_years,
            "include": include,
            "del_btn": del_btn,
        }
        self.rows.append(row_dict)

        idx = len(self.rows) - 1
        type_cb.bind("<<ComboboxSelected>>", lambda e, r=idx: self._on_type_changed(r))

    # -------------------------
    # 유형 변경 핸들러
    # -------------------------
    def _on_type_changed(self, idx):
        from criteria_config import CRITERIA_TYPES

        row = self.rows[idx]
        type_key = row["type"].get()
        if not type_key or type_key not in CRITERIA_TYPES:
            return

        cfg = CRITERIA_TYPES[type_key]

        # 계정 목록 갱신
        accounts = list(cfg["accounts"].keys())
        row["account"]["values"] = accounts
        row["account"].set("")

        # 연산자 목록 갱신
        row["xCompare"]["values"] = cfg["operators"]
        row["xCompare"].set("")

        # 기준값 활성/비활성
        if cfg["has_value"]:
            row["xValue"].config(state="normal")
        else:
            row["xValue"].delete(0, tk.END)
            row["xValue"].config(state="disabled")

        # 데이터가용성: 연산자 자동 선택
        if type_key == "데이터가용성":
            row["xCompare"].set("존재함")

        self._update_year_cond_visibility(idx)

    def _update_year_cond_visibility(self, idx):
        from criteria_config import CRITERIA_TYPES

        row = self.rows[idx]
        type_key = row["type"].get()
        cfg = CRITERIA_TYPES.get(type_key, {})

        if cfg.get("has_year_cond"):
            row["yearCond"].grid()
            row["yearCond"].bind(
                "<<ComboboxSelected>>", lambda e, r=idx: self._on_year_cond_changed(r)
            )
        else:
            row["yearCond"].grid_remove()
            row["nYears"].grid_remove()

    def _on_year_cond_changed(self, idx):
        row = self.rows[idx]
        if row["yearCond"].get() == "N개년이상":
            row["nYears"].grid()
        else:
            row["nYears"].grid_remove()

    # -------------------------
    # preset 적용
    # -------------------------
    def apply_preset(self, idx, preset_name):
        from criteria_config import PRESET_TO_TYPE_ACCOUNT, CRITERIA_TYPES

        preset = PRESET.get(preset_name)
        mapping = PRESET_TO_TYPE_ACCOUNT.get(preset_name)
        if not preset or not mapping or idx >= len(self.rows):
            return

        type_key, account_kor = mapping
        row = self.rows[idx]

        row["type"].set(type_key)
        self._on_type_changed(idx)
        row["account"].set(account_kor)

        row["xValue"].delete(0, tk.END)
        if preset.get("xValue"):
            row["xValue"].insert(0, str(preset["xValue"]))

        row["xCompare"].set(preset.get("xCompare", ""))

        if CRITERIA_TYPES[type_key]["has_year_cond"]:
            row["yearCond"].set(preset.get("yearCond", "모든연도"))
            self._on_year_cond_changed(idx)

        row["include"].set(preset.get("include", "포함"))

    # -------------------------
    # row 삭제 + 순번 재정렬
    # -------------------------
    def delete_row(self, idx):
        row = self.rows[idx]
        for widget in (
            row["seq"], row["type"], row["account"], row["xValue"],
            row["xCompare"], row["yearCond"], row["nYears"],
            row["include"], row["del_btn"],
        ):
            widget.destroy()

        self.rows.pop(idx)
        self.resequence()

    def resequence(self):
        from criteria_config import CRITERIA_TYPES

        for i, row in enumerate(self.rows):
            row["seq"].config(text=str(i + 1))

            row["seq"].grid(row=i + 1, column=0)
            row["type"].grid(row=i + 1, column=1)
            row["account"].grid(row=i + 1, column=2)
            row["xValue"].grid(row=i + 1, column=3)
            row["xCompare"].grid(row=i + 1, column=4)

            # yearCond / nYears: 현재 유형에 따라 조건부 배치
            type_key = row["type"].get()
            cfg = CRITERIA_TYPES.get(type_key, {})
            if cfg.get("has_year_cond"):
                row["yearCond"].grid(row=i + 1, column=5)
                if row["yearCond"].get() == "N개년이상":
                    row["nYears"].grid(row=i + 1, column=6)
                else:
                    row["nYears"].grid_remove()
            else:
                row["yearCond"].grid_remove()
                row["nYears"].grid_remove()

            row["include"].grid(row=i + 1, column=7)
            row["del_btn"].grid(row=i + 1, column=8)

            # 바인딩 재설정
            row["del_btn"].configure(command=lambda r=i: self.delete_row(r))
            row["type"].bind("<<ComboboxSelected>>", lambda e, r=i: self._on_type_changed(r))
            row["yearCond"].bind(
                "<<ComboboxSelected>>", lambda e, r=i: self._on_year_cond_changed(r)
            )

    # -------------------------
    # 변환 버튼
    # -------------------------
    def on_convert(self):
        from criteria_config import CRITERIA_TYPES, TEXT_OPERATORS

        # =========================
        # 1. 기본 입력값 수집
        # =========================
        corp_name   = self.corp_name.get().strip()
        target_corp = self.target_corp.get().strip()
        year_from   = self.year_from.get().strip()
        year_to     = self.year_to.get().strip()

        if not corp_name or not target_corp or not year_from or not year_to:
            messagebox.showerror("입력 오류", "기본 정보를 모두 입력하세요.")
            return

        if not self.file_path:
            messagebox.showerror("입력 오류", "Raw 파일을 선택하세요.")
            return

        self.output_dir_path = os.path.dirname(self.file_path)

        try:
            year_from = int(year_from)
            year_to   = int(year_to)
        except ValueError:
            messagebox.showerror("입력 오류", "연도는 숫자로 입력해야 합니다.")
            return

        if year_from > year_to:
            messagebox.showerror("입력 오류", "시작연도는 종료연도보다 클 수 없습니다.")
            return

        input_data = {
            "corpName":    corp_name,
            "targetCorp":  target_corp,
            "yearFrom":    year_from,
            "yearTo":      year_to,
            "rawFilePath": self.file_path,
            "outputDir":   self.output_dir_path,
        }

        # =========================
        # 2. 기준 테이블 수집
        # =========================
        # 연산자 중 기준값이 불필요한 것들
        NO_VALUE_OPERATORS = {"공란", "공란아님", "존재함"}

        criteria_list = []
        for idx, row in enumerate(self.rows, start=1):
            type_key    = row["type"].get()
            account     = row["account"].get()
            x_value     = row["xValue"].get().strip()
            x_compare   = row["xCompare"].get()
            year_cond   = row["yearCond"].get()
            n_years     = row["nYears"].get().strip()
            include_str = row["include"].get()

            if not type_key:
                continue  # 유형 미선택 row 무시

            if not account:
                messagebox.showerror("입력 오류", f"기준 {idx}: 분석계정을 선택해주세요.")
                return

            if not x_compare:
                messagebox.showerror("입력 오류", f"기준 {idx}: 비교연산자를 선택해주세요.")
                return

            cfg = CRITERIA_TYPES.get(type_key, {})
            if cfg.get("has_value") and x_compare not in NO_VALUE_OPERATORS and not x_value:
                messagebox.showerror("입력 오류", f"기준 {idx}: 기준값을 입력해주세요.")
                return

            if cfg.get("has_year_cond"):
                if not year_cond:
                    messagebox.showerror("입력 오류", f"기준 {idx}: 연도조건을 선택해주세요.")
                    return
                if year_cond == "N개년이상":
                    try:
                        n_int = int(n_years)
                        if n_int <= 0:
                            raise ValueError
                    except (ValueError, TypeError):
                        messagebox.showerror(
                            "입력 오류", f"기준 {idx}: N개년은 양의 정수로 입력해주세요."
                        )
                        return

            if not include_str:
                messagebox.showerror("입력 오류", f"기준 {idx}: 포함/제외를 선택해주세요.")
                return

            criteria_list.append({
                "seq":           idx,
                "type":          type_key,
                "account":       account,
                "xValue":        x_value,
                "xCompare":      x_compare,
                "yearCondition": year_cond,
                "nYears":        n_years,
                "include":       include_str == "포함",
            })

        if not criteria_list:
            messagebox.showerror("입력 오류", "최소 1개의 기준을 입력하세요.")
            return

        # =========================
        # 3. 변환 실행
        # =========================
        payload = {"inputData": input_data, "criteriaList": criteria_list}
        logger.info("변환 실행: %s", payload)

        try:
            from processor import main_processor
            main_processor(payload)
            messagebox.showinfo("완료", "분석 및 파일 저장이 완료되었습니다.")
        except PermissionError as pe:
            messagebox.showerror("저장 오류", str(pe))
        except Exception as e:
            logger.exception("변환 중 오류 발생")
            messagebox.showerror("오류", f"작업 중 오류가 발생했습니다:\n{e}")
