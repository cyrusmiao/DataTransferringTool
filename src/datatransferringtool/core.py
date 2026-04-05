import pandas as pd
from pathlib import Path
from thefuzz import fuzz
from typing import List, Dict, Any
from .config import TransferConfig

class DataTransfer:
    def __init__(self, config: TransferConfig):
        self.config = config
        self.report = []
        self.target_df = self._load_file(config.target_file)
        
    def _col_to_index(self, col) -> int:
        if isinstance(col, int):
            return col
        col = str(col).strip()
        if col.isdigit():
            return int(col)
        num = 0
        for c in col.upper():
            if c.isalpha():
                num = num * 26 + (ord(c) - ord('A')) + 1
        return num - 1
        
    def _load_file(self, file_path: str) -> pd.DataFrame:
        path = Path(file_path)
        if path.suffix.lower() == '.csv':
            return pd.read_csv(path)
        elif path.suffix.lower() in ['.xls', '.xlsx']:
            return pd.read_excel(path)
        else:
            raise ValueError(f"Unsupported file format: {path.suffix}")

    def _save_file(self, df: pd.DataFrame, file_path: str):
        path = Path(file_path)
        if path.suffix.lower() == '.csv':
            df.to_csv(path, index=False)
        elif path.suffix.lower() in ['.xls', '.xlsx']:
            # For .xls it's generally better to output as .xlsx if possible, 
            # but pandas to_excel supports xlsx via openpyxl and xls via xlwt (deprecated).
            engine = 'openpyxl' if path.suffix.lower() == '.xlsx' else 'xlwt'
            df.to_excel(path, index=False, engine=engine)
        else:
            raise ValueError(f"Unsupported file format: {path.suffix}")

    def _interactive_resolve(self, old_val, new_val, src_file, target_file, ref_val) -> str:
        print(f"\nConflict Detected!")
        print(f"Target file: {target_file}")
        print(f"Source file: {src_file}")
        print(f"Reference value: {ref_val}")
        print(f"Original Data: {old_val}")
        print(f"New Data: {new_val}")
        print("Options: [1] Keep Original [2] Overwrite")
        while True:
            choice = input("Enter choice (1/2): ").strip()
            if choice == '1':
                return 'keep_original'
            elif choice == '2':
                return 'overwrite'
            else:
                print("Invalid choice, try again.")

    def run(self):
        # We need to process each source
        # Make a copy of the target to modify
        out_df = self.target_df.copy()
        
        for source_idx, source in enumerate(self.config.sources):
            try:
                src_df = self._load_file(source.file_path)
            except Exception as e:
                print(f"Error loading source file {source.file_path}: {e}")
                continue
                
            if not source.reference_column:
                print(f"Warning: No reference column defined for {source.file_path}. Skipping.")
                continue

            ref_src_key, ref_tgt_key = list(source.reference_column.items())[0]
            ref_src_idx = self._col_to_index(ref_src_key)
            ref_tgt_idx = self._col_to_index(ref_tgt_key)
            
            if ref_src_idx >= len(src_df.columns):
                print(f"Warning: Reference column '{ref_src_key}' out of bounds in {source.file_path}. Skipping source.")
                continue
            if ref_tgt_idx >= len(out_df.columns):
                print(f"Warning: Reference column '{ref_tgt_key}' out of bounds in target file. Skipping source.")
                continue
                
            ref_col_src = src_df.columns[ref_src_idx]
            ref_col_tgt = out_df.columns[ref_tgt_idx]
            
            # Map target dataframe to dictionary of indices for quick lookup by reference column
            # Since target might have duplicate reference values, we store a list of indices
            tgt_dict = {}
            for target_idx, ref_val in out_df[ref_col_tgt].items():
                if pd.isna(ref_val):
                    continue
                if ref_val not in tgt_dict:
                    tgt_dict[ref_val] = []
                tgt_dict[ref_val].append(target_idx)
                
            # Parse mappings
            valid_mappings = []
            for src_key, tgt_key in source.mapping.items():
                src_idx = self._col_to_index(src_key)
                tgt_idx = self._col_to_index(tgt_key)
                if src_idx < len(src_df.columns) and tgt_idx < len(out_df.columns):
                    valid_mappings.append({
                        'src_col': src_df.columns[src_idx],
                        'tgt_col': out_df.columns[tgt_idx],
                        'tgt_letter': tgt_key
                    })

            for src_idx, src_row in src_df.iterrows():
                ref_val = src_row[ref_col_src]
                if pd.isna(ref_val):
                    continue
                
                if ref_val not in tgt_dict:
                    # Target row not found, report and skip
                    print(f"Report: Source row with reference '{ref_src_key}={ref_val}' from '{source.file_path}' not found in target file '{self.config.target_file}'. Skipping.")
                    self.report.append({
                        "conflict_resolution": "skipped_not_in_target",
                        "source_file": source.file_path,
                        "target_file": self.config.target_file,
                        "reference_value": ref_val,
                        "target_column": None,
                        "original_data": None,
                        "new_data": None,
                        "similarity_score": None
                    })
                    continue
                
                # For each matching target row, transfer data
                target_indices = tgt_dict[ref_val]
                for target_idx in target_indices:
                    for mapping in valid_mappings:
                        src_col = mapping['src_col']
                        tgt_col = mapping['tgt_col']
                        
                        new_val = src_row[src_col]
                        old_val = out_df.at[target_idx, tgt_col]
                        
                        if pd.isna(new_val):
                            continue # Nothing to transfer
                        
                        action_taken = "transferred"
                        similarity = 0.0
                        
                        if pd.isna(old_val):
                            out_df.at[target_idx, tgt_col] = new_val
                        else:
                            # Both have values, calculate similarity
                            similarity = fuzz.ratio(str(old_val), str(new_val))
                            
                            if str(old_val) == str(new_val):
                                action_taken = "identical_skipped"
                            else:
                                resolution = self.config.conflict_resolution
                                if resolution == 'manual':
                                    resolution = self._interactive_resolve(old_val, new_val, source.file_path, self.config.target_file, ref_val)
                                
                                if resolution == 'keep_original':
                                    action_taken = "conflict_kept_original"
                                elif resolution == 'overwrite':
                                    out_df.at[target_idx, tgt_col] = new_val
                                    action_taken = "conflict_overwritten"
                        
                        if action_taken != "identical_skipped":
                            self.report.append({
                                "conflict_resolution": action_taken,
                                "source_file": source.file_path,
                                "target_file": self.config.target_file,
                                "reference_value": ref_val,
                                "target_column": mapping['tgt_letter'],
                                "original_data": old_val if not pd.isna(old_val) else None,
                                "new_data": new_val,
                                "similarity_score": similarity
                            })

        self._save_file(out_df, self.config.output_file)
        self._generate_report()

    def _generate_report(self):
        report_df = pd.DataFrame(self.report)
        report_path = "transfer_report.xlsx"
        if not report_df.empty:
            report_df.to_excel(report_path, index=False)
            print(f"Report generated successfully: {report_path}")
        else:
            print("No transfers or conflicts to report.")
