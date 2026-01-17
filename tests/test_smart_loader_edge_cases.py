import pandas as pd
import pytest
from diffxl.smart_loader import SmartLoader, SmartLoadError

def test_smart_loader_duplicate_keys(tmp_path):
    # Create an Excel file with duplicate headers
    file_path = tmp_path / "dup_headers.xlsx"
    
    # Create a DataFrame with duplicate columns manually
    df = pd.DataFrame([[1, 2, 3]], columns=['A', 'B', 'B'])  # Duplicate 'B'
    df.to_excel(file_path, index=False)
    
    loader = SmartLoader()
    
    # Should raise SmartLoadError because 'B' is duplicated (caught by generic duplicates check now)
    with pytest.raises(SmartLoadError, match="Duplicate column names found"):
        loader.load(str(file_path), key_column="B")

def test_smart_loader_missing_key_after_detection(tmp_path):
    # This simulates a case where _find_header_row might return positive, but the key is somehow not in the final columns
    # This is tricky to reproduce exactly without mocking, but we can verify that basic missing key raises SmartLoadError
    file_path = tmp_path / "missing_key.xlsx"
    df = pd.DataFrame({'id': [1, 2], 'name': ['alice', 'bob']})
    df.to_excel(file_path, index=False)
    
    loader = SmartLoader()
    
    with pytest.raises(SmartLoadError, match="Column 'foo' not found"):
        loader.load(str(file_path), key_column="foo")

def test_smart_loader_duplicate_non_key_cols(tmp_path):
    # Create an Excel file with duplicate non-key columns
    file_path = tmp_path / "dup_cols.xlsx"
    
    # Create a DataFrame with duplicate columns 'Data'
    df = pd.DataFrame([['1', 'x', 'y']], columns=['ID', 'Data', 'Data'])
    df.to_excel(file_path, index=False)
    
    loader = SmartLoader()
    
    # Should raise SmartLoadError because 'Data' is duplicated
    with pytest.raises(SmartLoadError, match="Duplicate column names found"):
        loader.load(str(file_path), key_column="ID")

