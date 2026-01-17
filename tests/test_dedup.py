
import pandas as pd
import pytest
from diffxl.diff_engine import compare_dataframes, DiffXLError

def test_compare_raises_on_duplicates():
    df_old = pd.DataFrame({
        'ID': ['1', '1', '2'],
        'Value': ['A', 'A', 'B']
    })
    df_new = pd.DataFrame({
        'ID': ['1', '2'],
        'Value': ['A', 'B']
    })
    
    with pytest.raises(DiffXLError, match="Key Uniqueness Violation"):
        compare_dataframes(df_old, df_new, 'ID')

def test_dedup_flag_logic_simulation():
    # Simulate what main.py does: drop duplicates (keep=False) then compare
    key_col = 'ID'
    df_old = pd.DataFrame({
        'ID': ['1', '1', '2'],
        'Value': ['A', 'A_dup', 'B']
    })
    df_new = pd.DataFrame({
        'ID': ['1', '2', '2'],
        'Value': ['A', 'B', 'B_dup']
    })
    
    # Pre-processing (Dedup - strict)
    old_mask = df_old[key_col].duplicated(keep=False)
    new_mask = df_new[key_col].duplicated(keep=False)
    
    df_old_clean = df_old[~old_mask]
    df_new_clean = df_new[~new_mask]
    
    # Logic:
    # df_old had 1,1,2. 1s are duplicates. Remaining: 2
    # df_new had 1,2,2. 2s are duplicates. Remaining: 1
    
    assert len(df_old_clean) == 1
    assert df_old_clean.iloc[0]['ID'] == '2'
    
    assert len(df_new_clean) == 1
    assert df_new_clean.iloc[0]['ID'] == '1'
    
    # Compare
    added, removed, changed = compare_dataframes(df_old_clean, df_new_clean, key_col)
    
    # Remaining Old: 2 (Value B)
    # Remaining New: 1 (Value A)
    # Result: 
    # - 1 was Added (in new, not in old)
    # - 2 was Removed (in old, not in new)
    
    assert len(added) == 1
    assert added.iloc[0]['ID'] == '1'
    
    assert len(removed) == 1
    assert removed.iloc[0]['ID'] == '2'
    
    assert len(changed) == 0
