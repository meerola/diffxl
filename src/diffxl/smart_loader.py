import pandas as pd
import difflib
from dataclasses import dataclass
from typing import List, Dict, Optional, Tuple

@dataclass
class ColumnStats:
    name: str
    uniqueness: float  # 0.0 to 1.0
    non_null_count: int
    total_rows: int
    dtype: str

@dataclass
class AnalysisReport:
    file_path: str
    sheet_name: Optional[str]
    missing_key: str
    candidates: List[Tuple[str, float]]  # (Column Name, Confidence Score)
    all_columns: List[ColumnStats]
    sheets_found: List[str]
    header_row_index: int = -1

class SmartLoadError(Exception):
    """Raised when loading fails, containing analysis data."""
    def __init__(self, message: str, report: AnalysisReport):
        super().__init__(message)
        self.report = report

class SmartLoader:
    def __init__(self):
        pass

    def _find_header_row(self, df: pd.DataFrame, key_column: str) -> int:
        """Scans the dataframe to find the row index containing the key_column."""
        # Limit scan to first 50 rows for performance
        limit = min(len(df), 50)
        for idx in range(limit):
            row = df.iloc[idx]
            # Check if the key_column string is in this row's values
            row_values = [str(x).strip() for x in row.values]
            if key_column in row_values:
                return idx
        return -1

    def _guess_header_row(self, df: pd.DataFrame) -> int:
        """Heuristic to find the header row based on column density."""
        scan_limit = min(len(df), 50)
        if scan_limit == 0:
             return 0
        
        best_idx = 0
        max_cols = 0
        
        for i in range(scan_limit):
            row_vals = df.iloc[i]
            # Count non-null and non-empty strings
            count = 0
            for val in row_vals:
                if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                    count += 1
            
            # Prioritize rows with more populated columns
            if count > max_cols:
                max_cols = count
                best_idx = i
        return best_idx

    def _deduplicate_columns(self, columns: pd.Index) -> Tuple[pd.Index, List[str]]:
        """
        Ensures column names are unique by appending suffixes.
        Returns: (New Index, List of duplicated column names)
        """
        new_cols = []
        counts: Dict[str, int] = {}
        duplicates = set()
        
        for col in columns:
            col_str = str(col).strip()
            if pd.isna(col) or col_str == "" or col_str.lower() == "nan":
                 col_str = "Unnamed"
            
            if col_str in counts:
                counts[col_str] += 1
                duplicates.add(col_str)
                new_cols.append(f"{col_str}_{counts[col_str]}")
            else:
                counts[col_str] = 0
                new_cols.append(col_str)
        
        return pd.Index(new_cols), list(duplicates)

    def analyze_dataframe(self, df: pd.DataFrame, missing_key: str) -> List[Tuple[str, float]]:
        """
        Analyzes a dataframe to find columns that look like good keys.
        Returns list of (col_name, score).
        """
        candidates = []
        total_rows = len(df)
        if total_rows == 0:
            return []

        norm_key = missing_key.lower().replace(" ", "").replace("_", "")

        for col in df.columns:
            col_str = str(col).strip()
            if not col_str:
                continue
                
            vals = df[col].astype(str)
            unique_count = vals.nunique()
            uniqueness = unique_count / total_rows
            
            norm_col = col_str.lower().replace(" ", "").replace("_", "")
            name_similarity = difflib.SequenceMatcher(None, norm_key, norm_col).ratio()
            
            score = (uniqueness * 0.6) + (name_similarity * 0.4)
            
            if missing_key.lower() in col_str.lower():
                score += 0.1
                
            candidates.append((col_str, score))

        candidates.sort(key=lambda x: x[1], reverse=True)
        return candidates

    def get_column_stats(self, df: pd.DataFrame) -> List[ColumnStats]:
        stats = []
        for col in df.columns:
            total = len(df)
            non_null = df[col].count()
            unique = df[col].nunique()
            stats.append(ColumnStats(
                name=str(col),
                uniqueness=unique / total if total > 0 else 0,
                non_null_count=int(non_null),
                total_rows=total,
                dtype=str(df[col].dtype)
            ))
        return stats

    def load_with_report(self, file_path: str, key_column: Optional[str] = None, sheet_name: Optional[str] = None) -> Tuple[pd.DataFrame, AnalysisReport]:
        """
        Smart load function that returns both the DataFrame and the analysis report.
        """
        import pandas as pd # Local import
        
        raw_dfs = [] # List of (sheet_name, df)
        available_sheets = []

        try:
            if file_path.lower().endswith('.csv'):
                # Robust CSV reading for ragged files (metadata headers)
                # Pre-scan to find max columns
                import csv
                max_cols = 1
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                        sample_lines = [f.readline() for _ in range(50)]
                        # Simple comma counting heuristic (fallback to 1 if empty)
                        for line in sample_lines:
                            if not line: continue
                            # Use csv module to count fields properly (handle quoted commas)
                            reader = csv.reader([line])
                            for row in reader:
                                if len(row) > max_cols:
                                    max_cols = len(row)
                except Exception:
                    pass # Fallback to standard read if file scan fails
                
                # Use range(max_cols) as names to force pandas to read all columns
                df = pd.read_csv(file_path, header=None, names=range(max_cols), engine='python')
                raw_dfs.append(("<CSV>", df))
            else:
                xl = pd.ExcelFile(file_path)
                available_sheets = xl.sheet_names
                
                if sheet_name:
                    if sheet_name not in available_sheets:
                         # For reporting purposes, we need a dummy key if none provided
                         dummy_key = key_column if key_column else "<Default Key>"
                         raise SmartLoadError(
                             f"Sheet '{sheet_name}' not found.",
                             AnalysisReport(file_path, sheet_name, dummy_key, [], [], available_sheets)
                         )
                    df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                    raw_dfs.append((sheet_name, df))
                else:
                    for s in available_sheets:
                        df = pd.read_excel(xl, sheet_name=s, header=None)
                        raw_dfs.append((s, df))

        except FileNotFoundError:
             raise FileNotFoundError(f"File not found: {file_path}")
        except Exception as e:
             if isinstance(e, SmartLoadError): raise e
             raise RuntimeError(f"Error reading file structure: {e}")

        # --- PATH 1: Default Mode (No Key) ---
        if key_column is None:
            # Use first available sheet
            target_sheet, target_df = raw_dfs[0]
            
            # Smart Detection of Header Row
            if len(target_df) == 0:
                 raise RuntimeError(f"File '{file_path}' (Sheet: '{target_sheet}') is empty.")
            
            header_idx = self._guess_header_row(target_df)
            
            headers = target_df.iloc[header_idx]
            clean_headers, dups = self._deduplicate_columns(headers)
            
            final_df = target_df.iloc[header_idx + 1:].copy()
            final_df.columns = clean_headers
            
            if dups:
                 raise SmartLoadError(f"Duplicate column names found in sheet '{target_sheet}': {dups}. Please ensure column names are unique.", None)

            if len(final_df.columns) == 0:
                 raise RuntimeError(f"No columns found in '{file_path}'.")

            # Determine key (First Column)
            detected_key = final_df.columns[0]
            
            # Clean empty rows
            # Vectorized filtering: Keep all rows where key is not empty/nan
            # This allows gaps in data while filtering out truly empty rows
            key_series = final_df[detected_key].astype(str).str.strip()
            # Check for validity: not 'nan', not '', not 'None' (pandas might convert None to 'nan' str)
            valid_mask = ~key_series.isin(['nan', '', 'None']) & final_df[detected_key].notna()
            
            valid_indices = final_df[valid_mask].index
            
            cleaned_df = final_df.loc[valid_indices].copy()

            report = AnalysisReport(
                file_path=file_path,
                sheet_name=target_sheet,
                missing_key=str(detected_key),
                candidates=[],
                all_columns=self.get_column_stats(cleaned_df),
                sheets_found=available_sheets,
                header_row_index=header_idx
            )
            return cleaned_df, report

        # --- PATH 2: Explicit Key Mode ---
        best_candidate_sheet = None
        best_candidate_idx = -1
        found_sheets_list = []
        
        for s_name, df in raw_dfs:
            idx = self._find_header_row(df, key_column)
            if idx != -1:
                found_sheets_list.append(s_name)
                if best_candidate_sheet is None:
                    best_candidate_sheet = s_name
                    best_candidate_idx = idx

        if len(found_sheets_list) > 1:
             report = AnalysisReport(
                file_path=file_path,
                sheet_name=None,
                missing_key=key_column,
                candidates=[],
                all_columns=[],
                sheets_found=found_sheets_list
             )
             raise SmartLoadError(f"Ambiguous key column. Found '{key_column}' in multiple sheets: {found_sheets_list}. Please specify the sheet to use with --sheet.", report)
        
        if best_candidate_sheet is not None:
            target_df = next(df for s, df in raw_dfs if s == best_candidate_sheet)
            
            headers = target_df.iloc[best_candidate_idx]
            clean_headers, dups = self._deduplicate_columns(headers)
            
            final_df = target_df.iloc[best_candidate_idx + 1:].copy()
            final_df.columns = clean_headers
            
            if dups:
                 raise SmartLoadError(f"Duplicate column names found in sheet '{best_candidate_sheet}': {dups}. Please ensure column names are unique.", None)
            
            actual_key = None
            for col in final_df.columns:
                if str(col).strip() == key_column:
                    actual_key = col
                    break
            
            if actual_key and actual_key != key_column:
                final_df.rename(columns={actual_key: key_column}, inplace=True)
            
            # Key Column Validation (should be there since we found it)
            if key_column not in final_df.columns:
                 # Logic fallback if deduplication somehow messed it up, but strict check above catches dups
                 report = AnalysisReport(
                        file_path=file_path,
                        sheet_name=best_candidate_sheet,
                        missing_key=key_column,
                        candidates=self.analyze_dataframe(final_df, key_column)[:5],
                        all_columns=self.get_column_stats(final_df),
                        sheets_found=available_sheets,
                        header_row_index=best_candidate_idx
                 )
                 raise SmartLoadError(f"Column '{key_column}' not found in sheet '{best_candidate_sheet}'.", report)

            valid_indices = []
            try:
                key_series = final_df[key_column]
            except Exception:
                 raise SmartLoadError(f"Error accessing key column '{key_column}'.", None)
            
            # Vectorized filtering for explicit key mode
            key_series_str = key_series.astype(str).str.strip()
            valid_mask = ~key_series_str.isin(['nan', '', 'None']) & key_series.notna()
            
            valid_indices = final_df[valid_mask].index
            
            cleaned_df = final_df.loc[valid_indices].copy()

            report = AnalysisReport(
                file_path=file_path,
                sheet_name=best_candidate_sheet,
                missing_key=key_column,
                candidates=[],
                all_columns=self.get_column_stats(cleaned_df),
                sheets_found=available_sheets,
                header_row_index=best_candidate_idx
            )
            return cleaned_df, report

        # Key Not Found -> Analysis
        target_sheet_name, target_df = raw_dfs[0]
        
        # Use heuristic for analysis report
        candidate_header_row = self._guess_header_row(target_df)
        
        headers = target_df.iloc[candidate_header_row]
        clean_headers, _ = self._deduplicate_columns(headers) # Ignore duplicates here for robustness
        
        analysis_df = target_df.iloc[candidate_header_row + 1:].copy()
        analysis_df.columns = clean_headers
        
        candidates = self.analyze_dataframe(analysis_df, key_column)
        col_stats = self.get_column_stats(analysis_df)
        
        report = AnalysisReport(
            file_path=file_path,
            sheet_name=target_sheet_name,
            missing_key=key_column,
            candidates=candidates[:5], 
            all_columns=col_stats,
            sheets_found=available_sheets,
            header_row_index=candidate_header_row
        )
        
        raise SmartLoadError(f"Column '{key_column}' not found.", report)

    def load(self, file_path: str, key_column: Optional[str] = None, sheet_name: Optional[str] = None) -> pd.DataFrame:
        """
        Smart load function.
        - Scans for header.
        - If key found: returns clean DataFrame.
        - If key NOT found: raises SmartLoadError with suggestions.
        """
        df, _ = self.load_with_report(file_path, key_column, sheet_name)
        return df