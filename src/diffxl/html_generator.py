import pandas as pd
import datetime
from jinja2 import Template

# --- Modern HTML Template (Enhanced) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title_prefix }}DiffXL Report</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
    <style>
        :root {
            /* Colors */
            --bg-body: #f3f4f6;
            --bg-card: #ffffff;
            --bg-header: #ffffff;
            
            --text-main: #111827;
            --text-muted: #6b7280;
            --text-light: #9ca3af;
            
            --primary: #3b82f6;
            --primary-hover: #2563eb;
            --border: #e5e7eb;
            
            /* Status Colors */
            --added-bg: #dcfce7; --added-text: #166534; --added-border: #86efac;
            --removed-bg: #fee2e2; --removed-text: #991b1b; --removed-border: #fca5a5;
            --changed-bg: #fef9c3; --changed-text: #854d0e; --changed-border: #fde047;
            
            --diff-old-text: #ef4444;
            --diff-new-text: #22c55e;
            
            --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
            --shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1);
            
            --font-main: 'Inter', sans-serif;
            --font-mono: 'JetBrains Mono', monospace;
        }

        /* Reset & Base */
        * { box-sizing: border-box; }
        body { font-family: var(--font-main); background: var(--bg-body); color: var(--text-main); margin: 0; line-height: 1.5; font-size: 0.925rem; }
        
        /* Layout */
        .app-container { max-width: 1800px; margin: 0 auto; padding: 1rem; height: 100vh; display: flex; flex-direction: column; }
        
        /* Compact Header */
        header { 
            display: flex; justify-content: space-between; align-items: center; 
            background: var(--bg-card); padding: 0.75rem 1rem; border-radius: 0.5rem; 
            box-shadow: var(--shadow-sm); border: 1px solid var(--border);
            margin-bottom: 0.75rem; flex-shrink: 0;
        }
        .header-left { display: flex; align-items: baseline; gap: 1rem; }
        .header-left h1 { margin: 0; font-size: 1.25rem; font-weight: 700; color: var(--text-main); letter-spacing: -0.025em; }
        .header-left .meta { font-size: 0.85rem; color: var(--text-muted); display: flex; align-items: center; gap: 0.5rem; }
        .header-left .meta strong { color: var(--text-main); font-weight: 600; }
        
        .key-badge { 
            background: #e0e7ff; color: #1e40af; padding: 0.1rem 0.5rem; 
            border-radius: 4px; font-size: 0.75rem; font-weight: 600; border: 1px solid #c7d2fe;
            font-family: var(--font-mono);
        }
        
        .header-right { display: flex; align-items: center; gap: 1rem; }
        .timestamp { font-size: 0.75rem; color: var(--text-light); }

        /* Unified Control Bar */
        .controls-header { 
            display: flex; flex-wrap: wrap; gap: 1rem; align-items: center; justify-content: space-between; 
            background: var(--bg-card); padding: 0.75rem 1rem; border-radius: 0.5rem 0.5rem 0 0; 
            border: 1px solid var(--border); border-bottom: none; flex-shrink: 0;
        }
        
        .stat-group { display: flex; gap: 0.5rem; flex-wrap: wrap; }
        
        /* Stat Buttons */
        .stat-btn { 
            background: #f3f4f6; border: 1px solid transparent; border-radius: 0.375rem; 
            padding: 0.35rem 0.75rem; cursor: pointer; display: flex; align-items: center; gap: 0.5rem;
            transition: all 0.15s ease; font-size: 0.85rem; color: var(--text-muted);
        }
        .stat-btn:hover { background: #e5e7eb; color: var(--text-main); }
        .stat-btn.active { background: white; border-color: var(--border); box-shadow: var(--shadow-sm); color: var(--text-main); font-weight: 600; }
        
        .stat-btn .value { font-weight: 700; font-family: var(--font-mono); }
        .stat-btn.added .value { color: var(--added-text); }
        .stat-btn.removed .value { color: var(--removed-text); }
        .stat-btn.changed .value { color: var(--changed-text); }
        
        .actions-group { display: flex; gap: 0.5rem; align-items: center; flex-wrap: wrap; }
        
        /* Standard Buttons */
        .btn { 
            padding: 0.35rem 0.75rem; border-radius: 0.375rem; border: 1px solid var(--border); 
            background: white; color: var(--text-main); font-size: 0.85rem; font-weight: 500; 
            cursor: pointer; transition: all 0.1s; display: inline-flex; align-items: center; gap: 0.4rem;
        }
        .btn:hover { background: #f9fafb; border-color: #9ca3af; }
        .btn.active { background: var(--text-main); color: white; border-color: var(--text-main); }
        
        /* Highlight Only Edits */
        #btnEdits { 
            background-color: var(--primary); 
            color: white; 
            border-color: var(--primary); 
            box-shadow: 0 4px 6px -1px rgba(59, 130, 246, 0.2);
        }
        #btnEdits:hover { background-color: var(--primary-hover); border-color: var(--primary-hover); }
        
        .search-input { padding: 0.35rem 0.75rem 0.35rem 2rem; font-size: 0.85rem; width: 200px; border-radius: 0.375rem; border: 1px solid var(--border); }

        /* Table Area */
        .table-window {
            background: var(--bg-card); border: 1px solid #d1d5db; border-radius: 0 0 0.5rem 0.5rem; 
            overflow: hidden; display: flex; flex-direction: column; flex: 1; 
            box-shadow: var(--shadow);
        }
        
        .table-scroll { overflow: auto; height: 100%; }
        
        table { width: 100%; border-collapse: separate; border-spacing: 0; font-family: var(--font-main); font-size: 0.85rem; }
        
        /* Sticky Headers */
        thead { z-index: 20; }
        thead tr:nth-child(1) th { position: sticky; top: 0; z-index: 40; height: 3rem; }
        thead tr:nth-child(2) th { position: sticky; top: 3rem; z-index: 40; height: 2.5rem; }
        
        thead th { 
            text-align: left; padding: 0.75rem 1rem; font-weight: 600; color: var(--text-muted); 
            border-bottom: 1px solid var(--border); border-right: 1px solid var(--border); white-space: nowrap; user-select: none;
            background: #f9fafb;
        }
        thead th:last-child { border-right: none; }
        thead th.sortable { cursor: pointer; }
        thead th.sortable:hover { background: #f3f4f6; color: var(--text-main); }
        thead th .sort-icon { font-size: 0.7rem; margin-left: 0.25rem; color: var(--text-light); }
        
        /* Sticky Columns */
        .sticky-left-1 { 
            position: sticky; left: 0; z-index: 30; 
            border-right: 1px solid var(--border); 
            width: 90px; min-width: 90px; max-width: 100px;
            /* background handled by td/th rules (white or row color) */
        }
        
        .sticky-left-2 { 
            position: sticky; left: 90px; z-index: 30; 
            border-right: 2px solid var(--border); 
            white-space: nowrap !important;
        }
        
        /* High Z-Index for Header intersections */
        thead th.sticky-left-1, thead th.sticky-left-2 { z-index: 50 !important; background: #f9fafb; }
        
        /* Header Highlighting */
        th.col-added { background-color: var(--added-bg) !important; color: var(--added-text); border-bottom: 2px solid var(--added-text) !important; }
        
        /* Filter Row */
        .filter-row th { background: #ffffff; padding: 0.5rem; border-bottom: 1px solid var(--border); }
        .col-filter { width: 100%; padding: 0.2rem 0.4rem; font-size: 0.75rem; border: 1px solid var(--border); border-radius: 0.25rem; }
        
        tbody td { 
            padding: 0.5rem 0.75rem; border-bottom: 1px solid var(--border); border-right: 1px solid var(--border); vertical-align: top;
            background: white; transition: background 0.1s;
        }
        tbody td:last-child { border-right: none; }
        tbody tr:hover td { background-color: #f9fafb !important; }

        /* Status Styles */
        .row-added td { background-color: var(--added-bg) !important; color: var(--added-text); }
        .row-added td:first-child { border-left: 4px solid var(--added-text); padding-left: calc(0.75rem - 4px); }
        
        .row-removed td { background-color: var(--removed-bg) !important; color: var(--removed-text); }
        .row-removed td:first-child { border-left: 4px solid var(--removed-text); padding-left: calc(0.75rem - 4px); }
        
        .row-changed td:first-child { border-left: 4px solid var(--changed-text); padding-left: calc(0.75rem - 4px); }
        
        .badge { display: inline-flex; align-items: center; padding: 0.1rem 0.3rem; border-radius: 99px; font-size: 0.65rem; font-weight: 700; text-transform: uppercase; }
        .badge-added { background: white; color: var(--added-text); border: 1px solid var(--added-text); }
        .badge-removed { background: white; color: var(--removed-text); border: 1px solid var(--removed-text); }
        .badge-changed { background: var(--changed-bg); color: var(--changed-text); border: 1px solid var(--changed-border); }
        .badge-unchanged { background: #f3f4f6; color: #6b7280; border: 1px solid #d1d5db; }
        
        /* Cell Diffs */
        td.cell-mod { background-color: #fef9c3 !important; border-left: 4px solid #ca8a04; padding-left: calc(0.75rem - 4px); }
        td.cell-add-col { background-color: #dcfce7 !important; }
        
        .diff-container { display: flex; flex-direction: column; gap: 0.25rem; font-family: inherit; font-size: inherit; }
        
        /* Old Values Hidden by Default */
        .val-old { color: #dc2626; text-decoration: line-through; opacity: 0.9; font-weight: 500; display: none; }
        body.show-old-values .val-old { display: block; }
        
        .val-new { color: #15803d; font-weight: 700; display: block; }
        
        .empty-state { padding: 2rem; }
        
        /* Helpers */
        .hidden { display: none !important; }
        
        /* Print Optimization */
        @media print {
            @page { size: A1 landscape; margin: 0.5cm; }
            body { 
                background: white; 
                -webkit-print-color-adjust: exact; 
                print-color-adjust: exact; 
                font-size: 9pt; 
                zoom: 80%; /* Fit more content */
            }
            .app-container { height: auto; padding: 0; max-width: none; display: block; }
            
            /* Hide UI controls */
            .actions-group, .filter-row, .btn-help, #btnToggleOld { display: none !important; }
            .col-filter { display: none !important; }
            
            /* Simplify Header */
            header { box-shadow: none; border: none; padding: 0; margin-bottom: 1rem; }
            .controls-header { border: none; background: none; padding: 0; margin-bottom: 1rem; }
            .stat-btn { border: 1px solid #ccc; background: white; }
            .stat-btn.active { box-shadow: none; border-color: #000; }
            
            /* Table Layout */
            .table-window { border: none; box-shadow: none; height: auto; overflow: visible; display: block; }
            .table-scroll { overflow: visible; height: auto; }
            
            table { 
                width: 100%; 
                border-collapse: collapse; 
                table-layout: auto; /* Allow shrink */
            }
            
            /* Remove Sticky/Fixed */
            th, td, .sticky-left-1, .sticky-left-2 { 
                position: static !important; 
                left: auto !important; 
                top: auto !important; 
                height: auto !important;
                z-index: auto !important;
                overflow: visible !important;
            }
            
            /* Print Borders */
            th, td { border: 1px solid #999 !important; padding: 4px 6px; }
            
            /* Ensure Colors */
            .row-added td { background-color: #dcfce7 !important; }
            .row-removed td { background-color: #fee2e2 !important; }
            td.cell-mod { background-color: #fef9c3 !important; }
            
            /* Wrap text to fit width */
            td, th { white-space: normal !important; word-wrap: break-word; }
            

        }
        /* Help Tooltip */
        .btn-help {
            width: 1.5rem; height: 1.5rem; border-radius: 50%; background: #e5e7eb; color: var(--text-muted);
            border: none; font-size: 0.85rem; font-weight: 700; cursor: help; display: inline-flex; align-items: center; justify-content: center;
            position: relative; margin-left: 0.5rem; vertical-align: middle;
        }
        .btn-help:hover { background: #d1d5db; color: var(--text-main); }
        
        .btn-help::after {
            content: attr(data-tooltip);
            position: absolute; top: 100%; left: 0; margin-top: 0.5rem;
            background: #1f2937; color: white; padding: 0.75rem; border-radius: 0.375rem;
            font-size: 0.75rem; font-weight: 400; width: 220px; text-align: left;
            visibility: hidden; opacity: 0; transition: opacity 0.2s; z-index: 100;
            pointer-events: none; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); line-height: 1.4;
        }
        .btn-help:hover::after { visibility: visible; opacity: 1; }
    </style>
</head>
<body class="show-old-values">
    <div class="app-container">
        <header>
            <div class="header-left">
                <h1>{{ title_prefix }}DiffXL Report</h1>
                <button class="btn-help" data-tooltip="DiffXL treats empty values (NaN, None, '') as equal by default. Use --raw for strict string comparison. When using the 'Show old values' button, an empty value is replaced with the !empty! placeholder">?</button>
                <div class="meta">
                    <span class="key-badge">UID: {{ key_col }}</span>
                    <span style="color: var(--border)">|</span>
                    <span>{{ old_file }}</span> 
                    <span style="color: var(--text-light)">→</span> 
                    <span>{{ new_file }}</span>
                </div>
            </div>
            <div class="header-right">
                <button class="btn" onclick="toggleOldValues(this)" id="btnToggleOld">Hide old values</button>
                <div class="timestamp">{{ timestamp }}</div>
            </div>
        </header>

        <div class="controls-header">
            <div class="stat-group">
                <button class="stat-btn active" onclick="setFilter('all', this)">
                    <span>Total</span>
                    <span class="value">{{ stats.added + stats.removed + stats.changed + stats.unchanged }}</span>
                </button>
                <button class="stat-btn added" onclick="setFilter('added', this)">
                    <span>Added</span>
                    <span class="value">+{{ stats.added }}</span>
                </button>
                <button class="stat-btn removed" onclick="setFilter('removed', this)">
                    <span>Removed</span>
                    <span class="value">-{{ stats.removed }}</span>
                </button>
                <button class="stat-btn changed" onclick="setFilter('changed', this)">
                    <span>Changed</span>
                    <span class="value">{{ stats.changed }}</span>
                </button>
                <button class="stat-btn" onclick="setFilter('unchanged', this)">
                    <span>Unchanged</span>
                    <span class="value">{{ stats.unchanged }}</span>
                </button>
                {% if stats.ignored > 0 %}
                <button class="stat-btn ignored" onclick="toggleIgnored(this)" style="color: #ef4444; border-color: #fecaca; background: #fef2f2;">
                    <span>Ignored Duplicates</span>
                    <span class="value">{{ stats.ignored }}</span>
                </button>
                {% endif %}
            </div>
            
            <div class="actions-group">
                <button class="btn" onclick="setFilter('edits', this)" id="btnEdits">Only Edits</button>
                <button class="btn" onclick="window.print()">Print / Save PDF</button>
                
                <div class="search-wrapper">
                    <input type="text" class="search-input" id="globalSearch" placeholder="Search..." onkeyup="applyFilters()">
                </div>
                
                <button class="btn" onclick="clearFilters()">Clear</button>
            </div>
        </div>

        {% if removed_cols %}
        <div style="background: #fff1f2; border: 1px solid #fecaca; color: #991b1b; padding: 0.5rem 1rem; border-radius: 0.5rem; margin-bottom: 0.75rem; display: flex; align-items: center; gap: 0.5rem; font-size: 0.85rem;">
            <span>⚠️</span>
            <div><strong>Deleted Columns:</strong> {{ removed_cols|join(', ') }}</div>
        </div>
        {% endif %}

        <div class="table-window">
            <div class="table-scroll">
                <table id="diffTable">
                    <thead>
                        <tr>
                            <th class="sticky-left-1" onclick="sortTable(0)" class="sortable">Status <span class="sort-icon">⇅</span></th>
                            {% for col in columns %}
                            <th class="{{ 'col-added' if col in added_cols else '' }} {{ 'sticky-left-2' if loop.first else '' }}" 
                                onclick="sortTable({{ loop.index }})" 
                                class="sortable"
                                style="width: {{ col_widths[col] }}; min-width: {{ col_widths[col] }};"
                                data-relevant="{{ 'true' if col in relevant_cols else 'false' }}">
                                {{ col }} <span class="sort-icon">⇅</span>
                            </th>
                            {% endfor %}
                        </tr>
                        <tr class="filter-row">
                            <th class="sticky-left-1"></th>
                            {% for col in columns %}
                            <th class="{{ 'sticky-left-2' if loop.first else '' }}" 
                                data-relevant="{{ 'true' if col in relevant_cols else 'false' }}"
                                style="width: {{ col_widths[col] }}; min-width: {{ col_widths[col] }};">
                                <input type="text" class="col-filter" placeholder="Filter..." onkeyup="applyFilters()">
                            </th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody id="tableBody">
                        {% for row_data in rows %}
                        <tr class="row-{{ row_data.status|lower }}" data-status="{{ row_data.status|lower }}">
                            <td class="sticky-left-1">
                                <span class="badge badge-{{ row_data.status|lower }}">{{ row_data.status }}</span>
                            </td>
                            {% for col in columns %}
                            {% set change_info = row_data.changes.get(col) %}
                            <td class="{{ 'cell-mod' if change_info else ('cell-add-col' if col in added_cols and row_data.status != 'removed' else '') }} {{ 'sticky-left-2' if loop.first else '' }}" 
                                data-relevant="{{ 'true' if col in relevant_cols else 'false' }}"
                                data-val="{{ row_data.data[col] }}"
                                style="width: {{ col_widths[col] }}; min-width: {{ col_widths[col] }};">
                                
                                {% if change_info %}
                                <div class="diff-container">
                                    <div class="val-old">{{ change_info.old }}</div>
                                    <div class="val-new">{{ row_data.data[col] }}</div>
                                </div>
                                {% else %}
                                {{ row_data.data[col] }}
                                {% endif %}
                            </td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
                <div id="noResults" class="empty-state hidden">
                    <div class="empty-icon">🔍</div>
                    <h3>No matching records found</h3>
                    <p>Try adjusting your filters or search query.</p>
                </div>
            </div>
        </div>
        
        <div id="ignoredWindow" class="table-window hidden" style="margin-top: 1rem; border-color: #fca5a5;">
            <div style="padding: 0.75rem 1rem; background: #fef2f2; border-bottom: 1px solid #fca5a5; font-weight: 700; color: #991b1b;">
                Ignored Duplicate Keys (Not Compared)
            </div>
            <div class="table-scroll">
                <table>
                    <thead>
                        <tr>
                            <th class="sticky-left-1">Source</th>
                            {% for col in columns %}
                            <th style="width: {{ col_widths[col] }}; min-width: {{ col_widths[col] }};">{{ col }}</th>
                            {% endfor %}
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in ignored_rows %}
                        <tr>
                            <td class="sticky-left-1" style="font-weight: 600; color: #7f1d1d;">{{ row._source }}</td>
                            {% for col in columns %}
                            <td>{{ row[col] }}</td>
                            {% endfor %}
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        // State
        let currentStatusFilter = 'all';
        let showOldValues = true; // Default shown

        function toggleOldValues(btn) {
            showOldValues = !showOldValues;
            document.body.classList.toggle('show-old-values', showOldValues);
            if (btn) btn.innerText = showOldValues ? "Hide old values" : "Show old values";
        }

        function setFilter(status, btn) {
            // UI: Clear active from stats and edit button
            document.querySelectorAll('.stat-btn, #btnEdits').forEach(b => b.classList.remove('active'));
            // If it's the ignored button, don't set active class here as it's a toggle logic handled elsewhere or we treat it as a mode
            if (btn && !btn.classList.contains('ignored')) btn.classList.add('active');
            
            // Hide ignored window if switching back to normal filters
            if (status !== 'ignored_view') {
                 document.getElementById('ignoredWindow').classList.add('hidden');
                 document.querySelector('.table-window:not(#ignoredWindow)').classList.remove('hidden');
                 // Reset ignored button style
                 const igBtn = document.querySelector('.stat-btn.ignored');
                 if (igBtn) igBtn.classList.remove('active');
            }
            
            currentStatusFilter = status;
            applyFilters();
        }

        function toggleIgnored(btn) {
            const igWindow = document.getElementById('ignoredWindow');
            const mainWindow = document.querySelector('.table-window:not(#ignoredWindow)');
            const isActive = btn.classList.contains('active');
            
            if (!isActive) {
                // Show Ignored
                document.querySelectorAll('.stat-btn, #btnEdits').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                
                mainWindow.classList.add('hidden');
                igWindow.classList.remove('hidden');
            } else {
                // Hide Ignored (Go back to All)
                btn.classList.remove('active');
                igWindow.classList.add('hidden');
                mainWindow.classList.remove('hidden');
                
                // Reset to Total
                document.querySelector('.stat-btn').classList.add('active'); // First one is Total
                currentStatusFilter = 'all';
                applyFilters();
            }
        }
        
        function clearFilters() {
            if (document.getElementById('globalSearch')) document.getElementById('globalSearch').value = '';
            document.querySelectorAll('.col-filter').forEach(input => input.value = '');
            applyFilters();
        }

        function applyFilters() {
            const searchInput = document.getElementById('globalSearch');
            const globalSearch = searchInput ? searchInput.value.toLowerCase() : '';
            
            const colInputs = document.querySelectorAll('.col-filter');
            const colFilters = {};
            colInputs.forEach((input, idx) => {
                if (input.value.trim()) colFilters[idx + 1] = input.value.toLowerCase();
            });

            // Column Visibility (Only Edits Mode)
            const table = document.getElementById('diffTable');
            const allCells = table.querySelectorAll('th, td');
            
            if (currentStatusFilter === 'edits') {
                // Hide irrelevant columns
                allCells.forEach(cell => {
                    if (cell.getAttribute('data-relevant') === 'false') cell.classList.add('hidden');
                    else cell.classList.remove('hidden');
                });
            } else {
                // Show all columns
                allCells.forEach(cell => cell.classList.remove('hidden'));
            }

            const rows = document.querySelectorAll('#tableBody tr');
            let visibleCount = 0;

            rows.forEach(row => {
                const status = row.getAttribute('data-status');
                
                // 1. Status Filter
                let matchStatus = false;
                if (currentStatusFilter === 'all') matchStatus = true;
                else if (currentStatusFilter === 'edits') matchStatus = (status !== 'unchanged');
                else matchStatus = (status === currentStatusFilter);

                if (!matchStatus) {
                    row.classList.add('hidden');
                    return; // Skip rest
                }

                // 2. Global Search
                if (globalSearch) {
                    const rowText = row.innerText.toLowerCase();
                    if (!rowText.includes(globalSearch)) {
                        row.classList.add('hidden');
                        return;
                    }
                }

                // 3. Column Filters
                let matchCols = true;
                for (const [colIdx, term] of Object.entries(colFilters)) {
                    const cell = row.children[colIdx];
                    if (!cell || !cell.innerText.toLowerCase().includes(term)) {
                        matchCols = false;
                        break;
                    }
                }

                if (matchCols) {
                    row.classList.remove('hidden');
                    visibleCount++;
                } else {
                    row.classList.add('hidden');
                }
            });

            const noRes = document.getElementById('noResults');
            if (noRes) noRes.classList.toggle('hidden', visibleCount > 0);
        }

        // Sorting
        function sortTable(n) {
            const table = document.getElementById("diffTable");
            const tbody = document.getElementById("tableBody");
            let rows = Array.from(tbody.querySelectorAll("tr"));
            let dir = table.getAttribute("data-sort-dir") === "asc" ? "desc" : "asc";
            table.setAttribute("data-sort-dir", dir);

            rows.sort((a, b) => {
                let x = a.children[n].innerText.toLowerCase();
                let y = b.children[n].innerText.toLowerCase();
                
                let numX = parseFloat(x);
                let numY = parseFloat(y);
                if (!isNaN(numX) && !isNaN(numY)) {
                    x = numX;
                    y = numY;
                }

                if (dir === "asc") return x > y ? 1 : -1;
                else return x < y ? 1 : -1;
            });

            rows.forEach(row => tbody.appendChild(row));
        }
    </script>
</body>
</html>
"""

def generate_html_report(
    df_old: pd.DataFrame,
    df_new: pd.DataFrame,
    df_added: pd.DataFrame,
    df_removed: pd.DataFrame,
    df_changed: pd.DataFrame,
    key_col: str,
    output_path: str,
    old_file_name: str,
    new_file_name: str,
    prefix: str = "",
    df_old_dups: pd.DataFrame = None,
    df_new_dups: pd.DataFrame = None,
) -> None:
    """
    Generates a standalone HTML report using the Jinja2 template.
    Combines all data into a unified view for the table.
    """
    
    # 1. Prepare Stats
    # Calculate unchanged count
    # Note: df_new contains Unchanged + Added + Changed (New Values)
    # So Unchanged = Total New - Added - Changed Rows
    total_new = len(df_new)
    count_added = len(df_added)
    
    # df_changed contains individual cell changes. We need the number of affected rows.
    count_changed_rows = df_changed[key_col].nunique() if not df_changed.empty else 0
    
    count_removed = len(df_removed)
    
    # Correct calculation for unchanged:
    count_unchanged = total_new - count_added - count_changed_rows
    
    # Duplicates stats
    count_ignored = 0
    if df_old_dups is not None: count_ignored += len(df_old_dups)
    if df_new_dups is not None: count_ignored += len(df_new_dups)

    stats = {
        "added": count_added,
        "removed": count_removed,
        "changed": count_changed_rows,
        "unchanged": max(0, count_unchanged), # Safety floor
        "ignored": count_ignored
    }
    
    # 2. Prepare Display Data
    # We want a unified list of dicts: {status: '...', data: {col: val...}}
    
    # Identify keys for quick lookup
    added_keys = set(df_added[key_col].astype(str))
    removed_keys = set(df_removed[key_col].astype(str))
    changed_keys = set(df_changed[key_col].astype(str))
    
    # Build the main list from df_new (contains Unchanged, Added, Changed)
    display_rows = []
    
    # Convert to string to avoid NaNs and ensure display consistency
    # This also silences pandas FutureWarning about downcasting
    df_new = df_new.astype(str).replace("nan", "")
    df_removed = df_removed.astype(str).replace("nan", "")
    
    # Process Changed Cells logic
    # df_changed has [Key, Column, Old Value, New Value]
    # We want a map: key_val -> { col_name: { old: ..., new: ... } }
    changes_map = {}
    if not df_changed.empty:
        # iterate rows of df_changed
        for _, row in df_changed.iterrows():
            k = str(row[key_col])
            c = row['Column']
            old_v = str(row['Old Value'])
            
            # Handle empty/nan values
            if old_v.lower() == 'nan' or old_v == 'None' or old_v.strip() == '':
                old_v = '!empty!'
                
            if k not in changes_map:
                changes_map[k] = {}
            changes_map[k][c] = {'old': old_v}

    # Iterate new file rows
    for _, row in df_new.iterrows():
        key_val = str(row[key_col])
        
        status = "unchanged"
        row_changes = {} # Dict of col -> change_info
        
        if key_val in added_keys:
            status = "added"
        elif key_val in changed_keys:
            status = "changed"
            # Get changes for this row
            row_changes = changes_map.get(key_val, {})
            
        display_rows.append({
            "status": status,
            "data": row.to_dict(),
            "changes": row_changes
        })
        
    # Append removed rows
    for _, row in df_removed.iterrows():
        display_rows.append({
            "status": "removed",
            "data": row.to_dict(),
            "changes": {}
        })
        
    # 3. Determine Columns to Show
    # Use columns from new file, putting key_col first.
    cols = [key_col] + [c for c in df_new.columns if c != key_col]
    
    # 4. Schema Diff Logic
    old_cols = set(df_old.columns)
    new_cols = set(df_new.columns)
    
    # Calculate added/removed columns (excluding key which is always present or matched)
    added_cols_list = sorted(list(new_cols - old_cols))
    removed_cols_list = sorted(list(old_cols - new_cols))

    # Calculate relevant columns (Key + Added Cols + Changed Cols)
    changed_cols_set = set(df_changed["Column"].unique()) if not df_changed.empty else set()
    relevant_cols = set([key_col]) | set(added_cols_list) | changed_cols_set
    
    # Prepare title prefix
    title_prefix = f"{prefix} - " if prefix else ""
    
    # Calculate Column Widths
    col_widths = {}
    for col in cols:
        # Header length
        max_len = len(str(col))
        
        # Data length (from stringified df_new)
        # Note: df_new contains the "New" state.
        # Check df_new for length.

        
        if col in df_new.columns:
            # Check length of strings
            series_len = df_new[col].str.len()
            if not series_len.empty:
                val_max = series_len.max()
                if pd.notna(val_max):
                    max_len = max(max_len, int(val_max))
        
        # Check removed rows too for robustness (if col exists)
        if col in df_removed.columns:
             series_len = df_removed[col].str.len()
             if not series_len.empty:
                val_max = series_len.max()
                if pd.notna(val_max):
                    max_len = max(max_len, int(val_max))

        # Heuristic: ~9px per char + padding
        px = (max_len * 9) + 24
        
        if col == key_col:
            # UID: No max width, force fit
            col_widths[col] = f"{max(100, px)}px"
        else:
            # Data: Clamp between 120 and 500
            col_widths[col] = f"{max(120, min(px, 500))}px"

    # Prepare Duplicates for template
    ignored_dups_list = []
    if count_ignored > 0:
        if df_old_dups is not None and not df_old_dups.empty:
             for _, row in df_old_dups.iterrows():
                 r = row.to_dict()
                 r["_source"] = "Original File"
                 ignored_dups_list.append(r)
        if df_new_dups is not None and not df_new_dups.empty:
             for _, row in df_new_dups.iterrows():
                 r = row.to_dict()
                 r["_source"] = "New File"
                 ignored_dups_list.append(r)

    # Rendering
    template = Template(HTML_TEMPLATE)
    html_content = template.render(
        stats=stats,
        rows=display_rows,
        ignored_rows=ignored_dups_list,
        columns=cols,
        old_file=old_file_name,
        new_file=new_file_name,
        added_cols=added_cols_list,
        removed_cols=removed_cols_list,
        timestamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        relevant_cols=relevant_cols,
        title_prefix=title_prefix,
        key_col=key_col,
        col_widths=col_widths
    )
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
