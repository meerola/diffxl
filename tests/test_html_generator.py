import pandas as pd
from diffxl.html_generator import generate_html_report

def test_generate_html_report(tmp_path):
    # Setup Data
    df_old = pd.DataFrame({'ID': ['1', '2'], 'Value': ['A', 'B']})
    df_new = pd.DataFrame({'ID': ['1', '2', '3'], 'Value': ['A', 'B_mod', 'C']})
    
    df_added = pd.DataFrame({'ID': ['3'], 'Value': ['C']})
    df_removed = pd.DataFrame({'ID': [], 'Value': []})
    df_changed = pd.DataFrame({
        'ID': ['2'], 
        'Column': ['Value'], 
        'Old Value': ['B'], 
        'New Value': ['B_mod']
    })
    
    output_file = tmp_path / "report.html"
    
    # Execute
    generate_html_report(
        df_old, df_new, df_added, df_removed, df_changed,
        key_col='ID',
        output_path=str(output_file),
        old_file_name="old.csv",
        new_file_name="new.csv"
    )
    
    # Verify
    assert output_file.exists()
    content = output_file.read_text(encoding='utf-8')
    assert "<!DOCTYPE html>" in content
    assert "DiffXL Report" in content
    assert "B_mod" in content
    # Check for new features
    assert "class=\"col-filter\"" in content
    assert "Clear" in content
