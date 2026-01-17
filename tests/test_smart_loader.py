
import pandas as pd
from diffxl.smart_loader import SmartLoader

def test_load_stops_at_empty_key(tmp_path):
    # Create a dataframe with an empty key in the first row of data
    df = pd.DataFrame([
        {'ID': None, 'Value': 'SkipMe'},
        {'ID': '1', 'Value': 'KeepMe'},
        {'ID': '2', 'Value': 'KeepMeAlso'}
    ])
    
    file_path = tmp_path / "test_gap.xlsx"
    df.to_excel(file_path, index=False)
    
    loader = SmartLoader()
    # Logic currently stops at first empty key, so we expect this to return an empty DF or just fail to find the valid rows
    loaded_df = loader.load(str(file_path), key_column='ID')
    
    # If the bug is real, loaded_df will be empty because of the first None
    # We WANT it to contain ID '1' and '2', or at least '1'. 
    # But for reproduction, I just want to see what it does. 
    # I'll assert what correct behavior SHOULD be, and expect failure.
    
    assert len(loaded_df) == 2
    assert '1' in loaded_df['ID'].values
    assert '2' in loaded_df['ID'].values
