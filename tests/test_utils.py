import pandas as pd
from diffxl.utils import save_diff_report

def test_save_diff_report_creates_file(tmp_path):
    # Setup
    output_file = tmp_path / "test_diff_report.xlsx"
    
    df_new = pd.DataFrame({
        'ID': ['1', '2', '3'],
        'Value': ['A', 'B', 'C']
    })
    
    df_added = pd.DataFrame({
        'ID': ['3'],
        'Value': ['C']
    })
    
    df_removed = pd.DataFrame({
        'ID': ['0'],
        'Value': ['Z']
    })
    
    df_changed = pd.DataFrame({
        'ID': ['2'],
        'Column': ['Value'],
        'Old Value': ['B_old'],
        'New Value': ['B']
    })
    
    # Execute
    save_diff_report(
        df_added=df_added,
        df_removed=df_removed,
        df_changed=df_changed,
        df_new=df_new,
        key_col='ID',
        output_path=str(output_file)
    )
    
    # Verify
    assert output_file.exists()
    
    # Optional: Read back to check sheets exist
    xls = pd.ExcelFile(output_file)
    assert "Added Rows" in xls.sheet_names
    assert "Removed Rows" in xls.sheet_names
    assert "Changed Details" in xls.sheet_names
    assert "Full Diff" in xls.sheet_names
    
    # Check Full Diff content (simple check)
    df_full = pd.read_excel(output_file, sheet_name="Full Diff")
    assert len(df_full) == 3
