import pandas as pd
from typing import Tuple, Optional
from rich.console import Console

# Import Smart Loader
from .smart_loader import SmartLoader, SmartLoadError, AnalysisReport

console = Console()
error_console = Console(stderr=True)

class DiffXLError(Exception):
    """Custom exception for known DiffXL errors."""
    pass

class KeyUniquenessError(DiffXLError):
    """Raised when the selected key is not unique, carrying analysis data."""
    def __init__(self, message: str, report: AnalysisReport):
        super().__init__(message)
        self.report = report

def read_data_table(file_path: str, key_column: Optional[str] = None, sheet_name: Optional[str] = None) -> Tuple[pd.DataFrame, AnalysisReport]:
    """
    Reads an Excel or CSV file using SmartLoader.
    """
    loader = SmartLoader()
    try:
        return loader.load_with_report(file_path, key_column, sheet_name)
    except SmartLoadError as e:
        # Re-raise as is, so main.py can catch it and access .report
        raise e 
    except FileNotFoundError:
        raise DiffXLError(f"File not found: {file_path}")
    except Exception as e:
        raise DiffXLError(f"Error reading file {file_path}: {e}")

def compare_dataframes(df_old: pd.DataFrame, df_new: pd.DataFrame, key_col: str, raw_mode: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Compares two dataframes and returns (added, removed, changed).
    'changed' dataframe will have columns: [Key, Column, Old Value, New Value]
    
    raw_mode: If True, uses simple string conversion for comparison (NaN becomes 'nan').
              If False, uses smart normalization (NaN becomes 'MISSING_VALUE' to avoid clashing with literal 'nan').
    """
    
    df_old[key_col] = df_old[key_col].astype(str).str.strip()
    df_new[key_col] = df_new[key_col].astype(str).str.strip()
    
    # Check for duplicates
    dup_old = df_old[key_col].duplicated().sum()
    dup_new = df_new[key_col].duplicated().sum()
    
    if dup_old > 0 or dup_new > 0:
        msg_parts = []
        if dup_old > 0:
            msg_parts.append(f"Found {dup_old} duplicate keys in original file")
        if dup_new > 0:
            msg_parts.append(f"Found {dup_new} duplicate keys in new file")
            
        # Run analysis to find better candidates
        loader = SmartLoader()
        # Analyze df_old (assuming it's representative) to find potential UIDs
        candidates = loader.analyze_dataframe(df_old, key_col)
        
        # Filter out the current key_col from suggestions since we know it has duplicates
        candidates = [c for c in candidates if c[0] != key_col]
        
        col_stats = loader.get_column_stats(df_old)
        
        report = AnalysisReport(
            file_path="Original File", # Placeholder as we don't have path here
            sheet_name=None,
            missing_key=key_col, # It's not missing but it failed
            candidates=candidates[:5],
            all_columns=col_stats,
            sheets_found=[]
        )
            
        raise KeyUniquenessError(
            f"Key Uniqueness Violation: {'; '.join(msg_parts)}. Use --dedup to remove duplicates, or ensure Key column '{key_col}' is unique.",
            report
        )
    
    # Set index to key for easier comparison
    df_old_indexed = df_old.set_index(key_col)
    df_new_indexed = df_new.set_index(key_col)
    
    old_keys = set(df_old_indexed.index)
    new_keys = set(df_new_indexed.index)
    
    # Added and Removed
    added_keys = new_keys - old_keys
    removed_keys = old_keys - new_keys
    common_keys = old_keys.intersection(new_keys)
    
    df_added = df_new_indexed.loc[list(added_keys)].reset_index()
    df_removed = df_old_indexed.loc[list(removed_keys)].reset_index()
    
    # Changed Rows Logic
    changes_list = []
    
    # Align columns: valid columns are those present in both
    common_columns = set(df_old_indexed.columns).intersection(set(df_new_indexed.columns))
    
    # Pre-filter to only common keys and common columns
    df_old_common = df_old_indexed.loc[list(common_keys), list(common_columns)]
    df_new_common = df_new_indexed.loc[list(common_keys), list(common_columns)]
    
    # Sort by index to ensure consistent iteration order
    df_old_common.sort_index(inplace=True)
    df_new_common.sort_index(inplace=True)
    
    # Efficient comparison using pandas
    
    # Function to normalize for comparison
    def normalize_for_diff(df):
        if raw_mode:
            # Strictly convert to string. NaN becomes 'nan'.
            return df.astype(str)
        else:
            # Use where to replace NaNs without triggering downcasting warnings
            # df.where(cond, other) replaces where condition is False
            return df.where(pd.notna(df), "MISSING_VALUE").astype(str)

    diff_mask = normalize_for_diff(df_old_common) != normalize_for_diff(df_new_common)
    
    # diff_mask is True where values differ
    # We now extract those values
    
    # Stack the mask to get (Key, Column) index of all True values
    # stack() creates a Series with MultiIndex (Key, Column)
    changed_cells = diff_mask.stack()
    changed_cells = changed_cells[changed_cells] # Filter only True
    
    for (key, col), _ in changed_cells.items():
        old_val = df_old_common.at[key, col]
        new_val = df_new_common.at[key, col]
        
        changes_list.append({
            key_col: key,
            "Column": col,
            "Old Value": old_val,
            "New Value": new_val
        })
            
    if not changes_list:
        df_changed = pd.DataFrame(columns=[key_col, "Column", "Old Value", "New Value"])
    else:
        df_changed = pd.DataFrame(changes_list)
    
    return df_added, df_removed, df_changed
