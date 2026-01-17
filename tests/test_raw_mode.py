import pandas as pd
import numpy as np
from diffxl.diff_engine import compare_dataframes

def test_raw_mode_nan_vs_literal_nan():
    # Setup
    # Old: Real NaN
    # New: Literal string "nan"
    df_old = pd.DataFrame({'ID': ['1'], 'Val': [np.nan]})
    df_new = pd.DataFrame({'ID': ['1'], 'Val': ['nan']})
    
    # Default Mode: "MISSING_VALUE" vs "nan" -> Changed
    _, _, changed = compare_dataframes(df_old, df_new, 'ID', raw_mode=False)
    assert len(changed) == 1
    
    # Raw Mode: "nan" vs "nan" -> Unchanged (String equality)
    # This proves raw mode is doing simple string conversion
    _, _, changed_raw = compare_dataframes(df_old, df_new, 'ID', raw_mode=True)
    assert len(changed_raw) == 0

def test_raw_mode_none_vs_nan():
    # Old: None (Object type)
    # New: NaN (Float type)
    df_old = pd.DataFrame({'ID': ['1'], 'Val': [None]}, dtype=object)
    df_new = pd.DataFrame({'ID': ['1'], 'Val': [np.nan]})
    
    # Default: Both "MISSING_VALUE" -> Unchanged
    _, _, changed = compare_dataframes(df_old, df_new, 'ID', raw_mode=False)
    assert len(changed) == 0
    
    # Raw: "None" vs "nan" -> Changed
    _, _, changed_raw = compare_dataframes(df_old, df_new, 'ID', raw_mode=True)
    assert len(changed_raw) == 1
    assert changed_raw.iloc[0]['Old Value'] is None
    # assert changed_raw.iloc[0]['New Value'] is np.nan # NaN comparison is tricky
    assert pd.isna(changed_raw.iloc[0]['New Value'])
