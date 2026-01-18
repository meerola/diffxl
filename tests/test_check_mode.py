import pytest
import pandas as pd
from diffxl.diff_engine import read_data_table, SmartLoadError

@pytest.fixture
def multi_sheet_excel(tmp_path):
    # Create a dummy Excel file with multiple sheets
    file_path = tmp_path / "multi_sheet.xlsx"
    
    # Sheet 1: Irrelevant info
    df1 = pd.DataFrame({"Info": ["Cover Page", "Date: 2024"]})
    
    # Sheet 2: The actual data with Key 'ID'
    # Start on row 2 (index 2)
    # We construct the data as a list of lists to ensure "ID" is written as a cell value
    rows = [
        ["", "", ""],   # Row 0
        ["", "", ""],   # Row 1
        ["ColA", "ID", "Value"], # Row 2: Header
        ["A", 1, 10],   # Row 3
        ["B", 2, 20],   # Row 4
        ["C", 3, 30]    # Row 5
    ]
    df2_final = pd.DataFrame(rows)
    
    # Sheet 3: Another random sheet
    df3 = pd.DataFrame({"Notes": ["Some notes"]})
    
    with pd.ExcelWriter(file_path) as writer:
        df1.to_excel(writer, sheet_name="Cover", index=False)
        # Write without header because the header is already in the data rows
        df2_final.to_excel(writer, sheet_name="DataSheet", index=False, header=False)
        df3.to_excel(writer, sheet_name="Notes", index=False)
        
    return str(file_path)

@pytest.fixture
def ambiguous_excel(tmp_path):
    file_path = tmp_path / "ambiguous.xlsx"
    df = pd.DataFrame({"ID": [1, 2], "Val": [1, 2]})
    with pd.ExcelWriter(file_path) as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False)
        df.to_excel(writer, sheet_name="Sheet2", index=False)
    return str(file_path)

def test_smart_sheet_detection(multi_sheet_excel):
    # Should automatically find "ID" in "DataSheet"
    df, _ = read_data_table(multi_sheet_excel, "ID")
    assert "ID" in df.columns
    assert len(df) == 3
    assert df.iloc[0]["ID"] == 1

def test_smart_sheet_detection_ambiguous(ambiguous_excel):
    # Should exit or raise error. 
    # read_data_table raises SmartLoadError for ambiguity now
    with pytest.raises(SmartLoadError) as exc:
        read_data_table(ambiguous_excel, "ID")
    assert "Ambiguous key column" in str(exc.value)

def test_smart_sheet_explicit(multi_sheet_excel):
    # Should work if we explicitly say DataSheet
    df, _ = read_data_table(multi_sheet_excel, "ID", sheet_name="DataSheet")
    assert "ID" in df.columns
    
    # Should fail if we say Cover
    # This might raise SmartLoadError (key not found) instead of DiffXLError now
    with pytest.raises(SmartLoadError):
        read_data_table(multi_sheet_excel, "ID", sheet_name="Cover")