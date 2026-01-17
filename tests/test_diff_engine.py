import pandas as pd
from diffxl.diff_engine import compare_dataframes

def test_compare_dataframes_basic_addition():
    # Setup
    df_old = pd.DataFrame({
        'ID': ['1', '2'],
        'Value': ['A', 'B']
    })
    df_new = pd.DataFrame({
        'ID': ['1', '2', '3'],
        'Value': ['A', 'B', 'C']
    })
    
    # Execute
    added, removed, changed = compare_dataframes(df_old, df_new, 'ID')
    
    # Verify
    assert len(added) == 1
    assert added.iloc[0]['ID'] == '3'
    assert len(removed) == 0
    assert len(changed) == 0

def test_compare_dataframes_basic_removal():
    # Setup
    df_old = pd.DataFrame({
        'ID': ['1', '2', '3'],
        'Value': ['A', 'B', 'C']
    })
    df_new = pd.DataFrame({
        'ID': ['1', '2'],
        'Value': ['A', 'B']
    })
    
    # Execute
    added, removed, changed = compare_dataframes(df_old, df_new, 'ID')
    
    # Verify
    assert len(added) == 0
    assert len(removed) == 1
    assert removed.iloc[0]['ID'] == '3'
    assert len(changed) == 0

def test_compare_dataframes_basic_change():
    # Setup
    df_old = pd.DataFrame({
        'ID': ['1'],
        'Value': ['A']
    })
    df_new = pd.DataFrame({
        'ID': ['1'],
        'Value': ['A_modified']
    })
    
    # Execute
    added, removed, changed = compare_dataframes(df_old, df_new, 'ID')
    
    # Verify
    assert len(added) == 0
    assert len(removed) == 0
    assert len(changed) == 1
    assert changed.iloc[0]['Old Value'] == 'A'
    assert changed.iloc[0]['New Value'] == 'A_modified'
