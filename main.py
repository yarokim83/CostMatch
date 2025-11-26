import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import threading

# Set theme
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class ExcelComparatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window setup
        self.title("CostMatch")
        self.geometry("900x700")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Variables
        self.file1_path = None
        self.file2_path = None

        # UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Header
        self.header_frame = ctk.CTkFrame(self, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 20))
        self.header_label = ctk.CTkLabel(self.header_frame, text="CostMatch", font=ctk.CTkFont(size=24, weight="bold"))
        self.header_label.pack(pady=15)

        # File Selection Area
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        self.input_frame.grid_columnconfigure(1, weight=1)

        # File 1
        self.file1_btn = ctk.CTkButton(self.input_frame, text="파일 1 (QC 입고)", command=self.select_file1, width=150)
        self.file1_btn.grid(row=0, column=0, padx=10, pady=10)
        self.file1_label = ctk.CTkLabel(self.input_frame, text="선택된 파일 없음", text_color="gray", anchor="w")
        self.file1_label.grid(row=0, column=1, sticky="ew", padx=10)

        # File 2
        self.file2_btn = ctk.CTkButton(self.input_frame, text="파일 2 (비용 정산)", command=self.select_file2, width=150, fg_color="#E53935", hover_color="#D32F2F")
        self.file2_btn.grid(row=1, column=0, padx=10, pady=10)
        self.file2_label = ctk.CTkLabel(self.input_frame, text="선택된 파일 없음", text_color="gray", anchor="w")
        self.file2_label.grid(row=1, column=1, sticky="ew", padx=10)

        # Result Area
        self.result_text = ctk.CTkTextbox(self, width=800, height=400, font=ctk.CTkFont(family="Consolas", size=12))
        self.result_text.grid(row=3, column=0, sticky="nsew", padx=20, pady=10)
        self.result_text.insert("0.0", "파일을 선택하고 비교 버튼을 눌러주세요.\n")
        self.result_text.configure(state="disabled")

        # Action Buttons
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=4, column=0, sticky="ew", padx=20, pady=20)
        
        self.compare_btn = ctk.CTkButton(self.action_frame, text="비교 분석 시작", command=self.start_comparison, height=50, font=ctk.CTkFont(size=18, weight="bold"))
        self.compare_btn.pack(fill="x")

    def select_file1(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if filename:
            self.file1_path = filename
            self.file1_label.configure(text=os.path.basename(filename), text_color="white")
            self.log(f"[설정] 파일 1 선택됨: {filename}")

    def select_file2(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if filename:
            self.file2_path = filename
            self.file2_label.configure(text=os.path.basename(filename), text_color="white")
            self.log(f"[설정] 파일 2 선택됨: {filename}")

    def log(self, message):
        self.result_text.configure(state="normal")
        self.result_text.insert("end", message + "\n")
        self.result_text.see("end")
        self.result_text.configure(state="disabled")
        
        # Save to file
        with open("analysis_log.txt", "a", encoding="utf-8") as f:
            f.write(message + "\n")

    def clear_log(self):
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.configure(state="disabled")
        
        # Clear log file
        with open("analysis_log.txt", "w", encoding="utf-8") as f:
            f.write("")

    def start_comparison(self):
        if not self.file1_path or not self.file2_path:
            messagebox.showwarning("경고", "두 파일을 모두 선택해주세요.")
            return

        self.compare_btn.configure(state="disabled", text="분석 중...")
        self.clear_log()
        thread = threading.Thread(target=self.run_analysis)
        thread.start()

    def load_excel_smart(self, filepath, required_columns):
        """
        Try to find the header row that contains the required columns.
        """
        # Read first few rows to inspect
        try:
            # First, try reading normally
            df = pd.read_excel(filepath)
            if all(col in df.columns for col in required_columns):
                return df
            
            # If not found, try to find the header in the first 20 rows
            df_raw = pd.read_excel(filepath, header=None, nrows=20)
            
            header_row_idx = -1
            for idx, row in df_raw.iterrows():
                # Check if this row contains ALL required columns (fuzzy match or exact)
                # We convert row values to string and check
                row_values = [str(v).strip() for v in row.values if pd.notna(v)]
                
                # Check if all required columns are present in this row
                if all(req in row_values for req in required_columns):
                    header_row_idx = idx
                    break
            
            if header_row_idx != -1:
                self.log(f"  -> {os.path.basename(filepath)}: 헤더를 {header_row_idx+1}행에서 찾았습니다.")
                return pd.read_excel(filepath, header=header_row_idx)
            
            # If still not found, return original to let the caller handle the error
            return df
            
        except Exception as e:
            raise e

    def run_analysis(self):
        try:
            self.log(">>> 데이터 로딩 및 분석 시작...")
            
            # Define required columns
            # File 1: QC (Needs 'Doc No.', 'Part Group', 'Total Price' and report columns)
            # We add a few key report columns to ensure we find the right header
            req_cols_1 = ['Doc No.', 'Part Group', 'Total Price', 'Part No.', 'Vendor']
            # File 2: Cost (Needs 'PR No.' or 'PR No..1', 'Account name', '발주금액')
            # We relax the requirement here because we handle column selection dynamically
            req_cols_2 = ['Account name', '발주금액']

            # Load Data with smart header detection
            self.log(f"파일 1 로드 중: {os.path.basename(self.file1_path)}")
            df1 = self.load_excel_smart(self.file1_path, req_cols_1)
            
            self.log(f"파일 2 로드 중: {os.path.basename(self.file2_path)}")
            df2 = self.load_excel_smart(self.file2_path, req_cols_2)

            # --- File 1 Processing (QC) ---
            # Filter: Part Group == 'WIRE ROPE' or 'INVENTORY'
            target_groups = ['WIRE ROPE', 'INVENTORY']
            df1_filtered = df1[df1['Part Group'].isin(target_groups)].copy()
            
            # --- Generate Report File (마감자료 with PRL.xlsx) ---
            try:
                report_cols = [
                    "Type", "Date", "Part No.", "Part Type", "Part Group", 
                    "Description", "Qty", "Unit Price", "Total Price", 
                    "Doc No.", "Mach No.", "Vendor"
                ]
                
                # Check if all columns exist
                missing_report_cols = [c for c in report_cols if c not in df1_filtered.columns]
                if missing_report_cols:
                    self.log(f"\n[주의] 리포트 생성 중 다음 컬럼이 없어 제외됩니다: {missing_report_cols}")
                    existing_report_cols = [c for c in report_cols if c in df1_filtered.columns]
                    df_report = df1_filtered[existing_report_cols].copy()
                else:
                    df_report = df1_filtered[report_cols].copy()
                
                report_filename = "마감자료 with PRL.xlsx"
                df_report.to_excel(report_filename, index=False)
                self.log(f"\n[알림] '{report_filename}' 파일이 생성되었습니다. (건수: {len(df_report)} 건)")
                
            except Exception as e:
                self.log(f"\n[오류] 리포트 파일 생성 실패: {str(e)}")

            # Group by 'Doc No.' and sum 'Total Price'
            # Convert Doc No. to string to ensure matching works
            df1_filtered['Doc No.'] = df1_filtered['Doc No.'].astype(str).str.strip()
            df1_grouped = df1_filtered.groupby('Doc No.')['Total Price'].sum().reset_index()

            self.log(f"\n[파일 1 (QC) 처리 결과]")
            self.log(f"- 필터 조건: Part Group in {target_groups}")
            self.log(f"- 원본 건수: {len(df1_filtered)} 건")
            self.log(f"- Doc No. 기준 그룹화 후: {len(df1_grouped)} 건 (Key)")
            self.log(f"- 총 합계: {df1_grouped['Total Price'].sum():,.0f}")

            # --- File 2 Processing (Cost) ---
            # Filter: Account name == '장비 자재비-QC'
            df2_filtered = df2[df2['Account name'] == '장비 자재비-QC'].copy()
            
            # Log all columns for debugging
            self.log(f"\n[디버깅] 파일 2 컬럼 목록: {list(df2.columns)}")

            # Check if 'PR No.' exists
            # Priority: 'PR No..1' > 'PR No.'
            pr_col = None
            if 'PR No..1' in df2.columns:
                pr_col = 'PR No..1'
                self.log(f"  -> 'PR No..1' 컬럼을 Key로 사용합니다.")
            elif 'PR No.' in df2.columns:
                pr_col = 'PR No.'
                self.log(f"  -> 'PR No.' 컬럼을 Key로 사용합니다.")
            else:
                # Try to find a similar column
                candidates = [c for c in df2.columns if 'PR' in str(c) and 'No' in str(c)]
                if candidates:
                    pr_col = candidates[0]
                    self.log(f"  -> 'PR No.' 컬럼을 찾지 못해 '{pr_col}' 컬럼을 사용합니다.")
                else:
                    self.log(f"  -> ⚠️ 'PR No.' 관련 컬럼을 찾을 수 없습니다.")

            # Group by 'PR No.' and sum '발주금액'
            if pr_col in df2_filtered.columns:
                df2_filtered[pr_col] = df2_filtered[pr_col].astype(str).str.strip()
                df2_grouped = df2_filtered.groupby(pr_col)['발주금액'].sum().reset_index()
                
                # Rename for consistency
                if pr_col != 'PR No.':
                    df2_grouped = df2_grouped.rename(columns={pr_col: 'PR No.'})

                self.log(f"\n[파일 2 (정산) 처리 결과]")
                self.log(f"- 필터 해제 (전체 데이터 사용)")
                self.log(f"- 원본 건수: {len(df2_filtered)} 건")
                self.log(f"- {pr_col} 기준 그룹화 후: {len(df2_grouped)} 건 (Key)")
                self.log(f"- 총 합계: {df2_grouped['발주금액'].sum():,.0f}")
            else:
                df2_grouped = pd.DataFrame(columns=['PR No.', '발주금액'])
                self.log(f"\n[파일 2 (정산) 처리 실패] PR No. 컬럼 없음")

            # --- Comparison Logic (Key Matching) ---
            self.log("\n>>> 상세 비교 분석 (Key: Doc No. vs PR No.)...")

            # DEBUG: Inspect specific ID
            target_id = "S202502180004"
            self.log(f"\n[디버깅] '{target_id}' 값 검사")
            
            # Check in File 1
            f1_match = df1_grouped[df1_grouped['Doc No.'].str.contains(target_id, na=False)]
            if not f1_match.empty:
                raw_val = f1_match.iloc[0]['Doc No.']
                self.log(f"  - File 1 (Doc No.): '{raw_val}' (길이: {len(raw_val)})")
                self.log(f"    -> repr: {repr(raw_val)}")
            else:
                self.log(f"  - File 1: 해당 ID 없음")

            # Check in File 2 (All columns)

            # Merge
            merged = pd.merge(
                df1_grouped, 
                df2_grouped, 
                left_on='Doc No.', 
                right_on='PR No.', 
                how='outer', 
                indicator=True
            )

            # 1. Matched but Amount Differs
            matched = merged[merged['_merge'] == 'both'].copy()
            matched['Diff'] = matched['Total Price'] - matched['발주금액']
            # Tolerance check (e.g., < 1.0 difference is ignored)
            diff_rows = matched[abs(matched['Diff']) > 1.0]

            # 2. Only in File 1 (Missing in File 2)
            only_file1 = merged[merged['_merge'] == 'left_only']

            # 3. Only in File 2 (Missing in File 1)
            only_file2 = merged[merged['_merge'] == 'right_only']

            # --- Report ---
            self.log(f"\n[분석 결과 요약]")
            self.log(f"✅ Key 매칭 성공: {len(matched)} 건")
            self.log(f"⚠️ 금액 불일치: {len(diff_rows)} 건")
            self.log(f"❌ File 1에만 존재 (정산 누락?): {len(only_file1)} 건")
            self.log(f"❓ File 2에만 존재 (매칭 불가): {len(only_file2)} 건")

            if not diff_rows.empty:
                self.log(f"\n[⚠️ 금액 불일치 상세]")
                for _, row in diff_rows.iterrows():
                    self.log(f"Key: {row['Doc No.']}")
                    self.log(f"  - QC(File1): {row['Total Price']:,.0f}")
                    self.log(f"  - 정산(File2): {row['발주금액']:,.0f}")
                    self.log(f"  - 차이: {row['Diff']:,.0f}")

            if not only_file1.empty:
                self.log(f"\n[❌ File 1에만 존재 (Doc No.)]")
                for _, row in only_file1.iterrows():
                    self.log(f"- {row['Doc No.']} (금액: {row['Total Price']:,.0f})")

            if not only_file2.empty:
                self.log(f"\n[❓ File 2에만 존재 (PR No.)]")
                for _, row in only_file2.iterrows():
                    self.log(f"- {row['PR No.']} (금액: {row['발주금액']:,.0f})")

            # Total Diff Calculation
            total_diff = df1_grouped['Total Price'].sum() - df2_grouped['발주금액'].sum()
            self.log(f"\n💰 전체 차액 (File1 - File2): {total_diff:,.0f}")

        except Exception as e:
            self.log(f"\n[오류 발생] {str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            self.compare_btn.configure(state="normal", text="비교 분석 시작")

if __name__ == "__main__":
    app = ExcelComparatorApp()
    app.mainloop()
