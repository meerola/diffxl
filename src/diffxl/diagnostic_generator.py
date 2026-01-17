from jinja2 import Template
from typing import Optional
from .smart_loader import AnalysisReport

# Updated Template for Dual-File Comparison
DIAGNOSTIC_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>DiffXL Diagnostic Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 2rem; background: #f4f6f8; color: #333; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 0; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); overflow: hidden; }
        
        .header { background: linear-gradient(135deg, #2c3e50, #4ca1af); color: white; padding: 2rem; }
        .header h1 { margin: 0; font-size: 1.8rem; font-weight: 300; }
        .header p { margin: 0.5rem 0 0; opacity: 0.9; }

        .content { padding: 2rem; }

        .section { margin-bottom: 2.5rem; }
        h2 { border-bottom: 2px solid #eee; padding-bottom: 0.5rem; color: #444; margin-top: 0; }
        
        /* Grid Layout for Side-by-Side */
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; }
        
        .card { background: #fff; border: 1px solid #e1e4e8; border-radius: 8px; padding: 1.5rem; }
        .card h3 { margin-top: 0; color: #0066cc; font-size: 1.1rem; }
        
        .stat-row { display: flex; justify-content: space-between; padding: 0.5rem 0; border-bottom: 1px solid #f0f0f0; }
        .stat-row:last-child { border-bottom: none; }
        .stat-label { color: #666; }
        .stat-val { font-weight: 600; }

        .error-tag { background: #ffebee; color: #c62828; padding: 0.2rem 0.5rem; border-radius: 4px; font-size: 0.85rem; }
        .success-tag { background: #e8f5e9; color: #2e7d32; padding: 0.2rem 0.5rem; border-radius: 4px; font-size: 0.85rem; }

        table { width: 100%; border-collapse: collapse; margin-top: 1rem; font-size: 0.9rem; }
        th, td { text-align: left; padding: 0.6rem; border-bottom: 1px solid #eee; }
        th { background: #f8f9fa; font-weight: 600; color: #555; }
        
        .col-match { color: #2e7d32; }
        .col-miss { color: #c62828; }
        
        .badge { display: inline-block; padding: 0.25em 0.6em; font-size: 75%; font-weight: 700; line-height: 1; text-align: center; white-space: nowrap; vertical-align: baseline; border-radius: 0.25rem; }
        .badge-primary { color: #fff; background-color: #007bff; }
        .badge-secondary { color: #fff; background-color: #6c757d; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>DiffXL Diagnostic Report</h1>
            <p>Generated to help resolve comparison issues.</p>
        </div>

        <div class="content">
            <!-- 1. Logic & Settings Used -->
            <div class="section">
                <h2>1. Loading & Detection Summary</h2>
                <div class="grid-2">
                    <!-- File A -->
                    <div class="card">
                        <h3>Original File (A)</h3>
                        <div class="stat-row">
                            <span class="stat-label">File Path:</span>
                            <span class="stat-val" title="{{ report_old.file_path }}">{{ report_old.file_path.split('/')[-1] }}</span>
                        </div>
                        <div class="stat-row">
                            <span class="stat-label">Sheet Used:</span>
                            <span class="stat-val">
                                {% if report_old.sheet_name %}
                                    <span class="success-tag">{{ report_old.sheet_name }}</span>
                                {% else %}
                                    {% if report_old.sheets_found %}
                                        <span class="error-tag">None (Found: {{ report_old.sheets_found|join(', ') }})</span>
                                    {% else %}
                                        <span class="error-tag">Detection Failed</span>
                                    {% endif %}
                                {% endif %}
                            </span>
                        </div>
                        <div class="stat-row">
                            <span class="stat-label">Header Row:</span>
                            <span class="stat-val">{{ report_old.header_row_index if report_old.header_row_index >= 0 else 'Unknown' }}</span>
                        </div>
                        <div class="stat-row">
                            <span class="stat-label">Columns Found:</span>
                            <span class="stat-val">{{ report_old.all_columns|length }}</span>
                        </div>
                    </div>

                    <!-- File B -->
                    <div class="card">
                        <h3>New File (B)</h3>
                        <div class="stat-row">
                            <span class="stat-label">File Path:</span>
                            <span class="stat-val" title="{{ report_new.file_path }}">{{ report_new.file_path.split('/')[-1] }}</span>
                        </div>
                        <div class="stat-row">
                            <span class="stat-label">Sheet Used:</span>
                            <span class="stat-val">
                                {% if report_new.sheet_name %}
                                    <span class="success-tag">{{ report_new.sheet_name }}</span>
                                {% else %}
                                    {% if report_new.sheets_found %}
                                        <span class="error-tag">None (Found: {{ report_new.sheets_found|join(', ') }})</span>
                                    {% else %}
                                        <span class="error-tag">Detection Failed</span>
                                    {% endif %}
                                {% endif %}
                            </span>
                        </div>
                        <div class="stat-row">
                            <span class="stat-label">Header Row:</span>
                            <span class="stat-val">{{ report_new.header_row_index if report_new.header_row_index >= 0 else 'Unknown' }}</span>
                        </div>
                        <div class="stat-row">
                            <span class="stat-label">Columns Found:</span>
                            <span class="stat-val">{{ report_new.all_columns|length }}</span>
                        </div>
                    </div>
                </div>
            </div>

            <!-- 2. Column Matching -->
            <div class="section">
                <h2>2. Column Matching Analysis</h2>
                <p style="color: #666; margin-bottom: 1rem;">
                    Comparison requires identical column names. Whitespace or case differences will cause mismatches.
                </p>
                
                <div class="grid-2">
                    <div class="card">
                        <h3>Matches ({{ common_cols|length }})</h3>
                        {% if common_cols %}
                            <div style="max-height: 200px; overflow-y: auto;">
                                <ul style="margin: 0; padding-left: 1.2rem;">
                                {% for c in common_cols %}
                                    <li class="col-match">{{ c }}</li>
                                {% endfor %}
                                </ul>
                            </div>
                        {% else %}
                            <p class="error-tag">No matching columns found!</p>
                        {% endif %}
                    </div>

                    <div class="card">
                        <h3>Mismatches</h3>
                        {% if unique_old_cols %}
                            <div style="margin-bottom: 1rem;">
                                <strong>Only in Old File ({{ unique_old_cols|length }}):</strong>
                                <ul style="margin: 0; padding-left: 1.2rem; color: #d9534f; max-height: 100px; overflow-y: auto;">
                                {% for c in unique_old_cols %}
                                    <li>{{ c }}</li>
                                {% endfor %}
                                </ul>
                            </div>
                        {% endif %}
                        
                        {% if unique_new_cols %}
                            <div>
                                <strong>Only in New File ({{ unique_new_cols|length }}):</strong>
                                <ul style="margin: 0; padding-left: 1.2rem; color: #d9534f; max-height: 100px; overflow-y: auto;">
                                {% for c in unique_new_cols %}
                                    <li>{{ c }}</li>
                                {% endfor %}
                                </ul>
                            </div>
                        {% endif %}
                        
                        {% if not unique_old_cols and not unique_new_cols %}
                            <p class="success-tag">All columns match perfectly.</p>
                        {% endif %}
                    </div>
                </div>
            </div>

            <!-- 3. Unique Value Analysis -->
            <div class="section">
                <h2>3. Top Unique Columns (Potential Keys)</h2>
                <div class="grid-2">
                    <div>
                        <h3>Old File Top 3</h3>
                        <table>
                            <thead>
                                <tr>
                                    <th>Column</th>
                                    <th>Uniqueness</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for col in report_old.all_columns|sort(attribute='uniqueness', reverse=True)|slice(3)|first %}
                                <tr>
                                    <td>{{ col.name }}</td>
                                    <td>{{ "%.1f"|format(col.uniqueness * 100) }}%</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                    <div>
                        <h3>New File Top 3</h3>
                        <table>
                            <thead>
                                <tr>
                                    <th>Column</th>
                                    <th>Uniqueness</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for col in report_new.all_columns|sort(attribute='uniqueness', reverse=True)|slice(3)|first %}
                                <tr>
                                    <td>{{ col.name }}</td>
                                    <td>{{ "%.1f"|format(col.uniqueness * 100) }}%</td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

        </div>
    </div>
</body>
</html>
"""

def generate_diagnostic_report(report_old: Optional[AnalysisReport], report_new: Optional[AnalysisReport], output_path: str):
    # Fallback for None reports to avoid crashes (create dummy empty/error report)
    # But caller should ideally provide both or we should handle it.
    
    # We will assume caller does their best to provide reports.
    # If one is totally missing (file load failed completely early on), we can pass a dummy.
    
    # Actually, let's create a minimal dummy if None
    def create_dummy(path):
        return AnalysisReport(path, "Failed to Load", "", [], [], [])

    if report_old is None: report_old = create_dummy("Old File")
    if report_new is None: report_new = create_dummy("New File")

    # Pre-calculate column intersection logic to avoid using Jinja 'do'
    cols_old = [c.name for c in report_old.all_columns]
    cols_new = [c.name for c in report_new.all_columns]
    
    common_cols = []
    unique_old_cols = []
    
    # Calculate common and unique_old
    for c in cols_old:
        if c in cols_new:
            common_cols.append(c)
        else:
            unique_old_cols.append(c)
            
    # Calculate unique_new
    unique_new_cols = [c for c in cols_new if c not in cols_old]

    template = Template(DIAGNOSTIC_TEMPLATE)
    html = template.render(
        report_old=report_old, 
        report_new=report_new,
        common_cols=common_cols,
        unique_old_cols=unique_old_cols,
        unique_new_cols=unique_new_cols
    )
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
