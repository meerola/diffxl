import typer
from typing import Annotated, Optional
import sys
import os
from rich.console import Console, Group
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from .diff_engine import read_data_table, compare_dataframes, DiffXLError, SmartLoadError, AnalysisReport, KeyUniquenessError
import webbrowser
from .html_generator import generate_html_report
from .diagnostic_generator import generate_diagnostic_report
from .smart_loader import AnalysisReport # Need to import AnalysisReport for typing if used

console = Console()

def _fmt_list(items: set[str] | list[str], max_items: int = 5) -> str:
    """Format a list of items, truncating if too many."""
    sorted_items = sorted(list(items))
    if len(sorted_items) <= max_items:
        return ", ".join(sorted_items)
    else:
        remaining = len(sorted_items) - max_items
        return ", ".join(sorted_items[:max_items]) + f" and {remaining} others"

def _resolve_sheet(sheet_arg: Optional[str], file_path: str) -> Optional[str]:
    """Resolve a sheet argument to a sheet name.

    If ``sheet_arg`` is a numeric string (e.g. "2"), it is treated as a
    1-based sheet index and converted to the corresponding sheet name.
    Otherwise it is returned as-is.  Returns ``None`` when no sheet is
    specified.
    """
    if sheet_arg is None:
        return None

    # Check if the argument is a numeric index (1-based)
    if sheet_arg.isdigit():
        index = int(sheet_arg) - 1  # Convert to 0-based
        if file_path.lower().endswith('.csv'):
            return None  # CSV files don't have sheets
        try:
            import openpyxl as _xl
            wb = _xl.load_workbook(file_path, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
        except Exception:
            return sheet_arg  # Fall back to treating as name
        if 0 <= index < len(names):
            return names[index]
        # Index out of range – return as-is so SmartLoader raises a clear error
        return sheet_arg

    return sheet_arg

def get_candidates_renderables(report: AnalysisReport, title: str = "[bold]Did you mean one of these?[/bold]") -> list:
    """Returns a list of renderables for the candidate columns table."""
    renderables = []
    if report.candidates:
        renderables.append(Text.from_markup(title))
        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("Candidate Column", style="cyan")
        table.add_column("Confidence", justify="right")
        table.add_column("Uniqueness", justify="right")
        
        for cand_name, score in report.candidates[:3]:
            stats = next((s for s in report.all_columns if s.name == cand_name), None)
            uniq_str = f"{stats.uniqueness:.1%}" if stats else "N/A"
            table.add_row(cand_name, f"{score:.2f}", uniq_str)
            
        renderables.append(table)
        renderables.append(Text.from_markup("[dim]Confidence is based on name similarity and data uniqueness (UIDs).[/dim]"))
    else:
        renderables.append(Text.from_markup("[yellow]No obvious candidates found.[/yellow]"))
    return renderables

def display_combined_analysis_report(console: Console, err_old: Optional[SmartLoadError], err_new: Optional[SmartLoadError], args):
    """Displays the combined failure analysis report using Rich."""
    renderables = []
    
    missing_key = args.key if args.key else "Leftmost Column"
    
    if err_old and err_new:
        renderables.append(Text.from_markup(f"[bold red]Validation Failed:[/bold red] Column [yellow]'{missing_key}'[/yellow] not found in either file."))
    elif err_old:
        renderables.append(Text.from_markup(f"[bold red]Validation Failed:[/bold red] Column [yellow]'{err_old.report.missing_key}'[/yellow] not found in '{os.path.basename(err_old.report.file_path)}'."))
    elif err_new:
        renderables.append(Text.from_markup(f"[bold red]Validation Failed:[/bold red] Column [yellow]'{err_new.report.missing_key}'[/yellow] not found in '{os.path.basename(err_new.report.file_path)}'."))
    
    renderables.append(Text(""))
    
    if err_old:
        renderables.extend(get_candidates_renderables(err_old.report, title=f"[bold]Candidates for '{os.path.basename(err_old.report.file_path)}':[/bold]"))
        if err_old.report.sheets_found and len(err_old.report.sheets_found) > 1:
            renderables.append(Text(""))
            renderables.append(Text.from_markup(f"[bold]Sheets scanned in original file:[/bold] {', '.join(err_old.report.sheets_found)}"))
            if not err_old.report.sheet_name:
                renderables.append(Text.from_markup("[dim]Try specifying a sheet with [cyan]--sheet[/cyan].[/dim]"))
        renderables.append(Text(""))
        
    if err_new:
        renderables.extend(get_candidates_renderables(err_new.report, title=f"[bold]Candidates for '{os.path.basename(err_new.report.file_path)}':[/bold]"))
        if err_new.report.sheets_found and len(err_new.report.sheets_found) > 1:
            renderables.append(Text(""))
            renderables.append(Text.from_markup(f"[bold]Sheets scanned in new file:[/bold] {', '.join(err_new.report.sheets_found)}"))
            if not err_new.report.sheet_name:
                renderables.append(Text.from_markup("[dim]Try specifying a sheet with [cyan]--sheet[/cyan].[/dim]"))
        renderables.append(Text(""))

    suggested_key = "NewKey"
    if err_old and err_old.report.candidates:
        suggested_key = err_old.report.candidates[0][0]
    elif err_new and err_new.report.candidates:
        suggested_key = err_new.report.candidates[0][0]
        
    renderables.append(Text.from_markup(f"Tip: Run with a different key using [cyan]--key \"{suggested_key}\"[/cyan]"))

    console.print()
    console.print(Panel(
        Group(*renderables),
        title="[bold red]Analysis Failed[/bold red]",
        border_style="red",
        expand=False
    ))

def diff_command(
    old_file: Annotated[str, typer.Argument(help="Path to the original file (Excel or CSV)")],
    new_file: Annotated[str, typer.Argument(help="Path to the new file (Excel or CSV)")],
    key: Annotated[Optional[str], typer.Option("--key", "-k", help="Column name to use as unique identifier (default: Leftmost column)")] = None,
    sheet: Annotated[Optional[str], typer.Option("--sheet", "-s", help="Sheet name or number for both files (or just the original if --sheet2 is given)")] = None,
    sheet2: Annotated[Optional[str], typer.Option("--sheet2", "-s2", help="Sheet name or number for the new file (overrides --sheet for the second file)")] = None,
    output: Annotated[str, typer.Option("--output", "-o", help="Output Excel file name")] = "diff_report.xlsx",
    prefix: Annotated[Optional[str], typer.Option("--prefix", "-p", help="Add a prefix to output filenames (e.g., 'ABC' -> 'ABC_diff_report.xlsx')")] = None,
    raw: Annotated[bool, typer.Option("--raw", help="Perform exact string comparison (disable smart normalization)")] = False,
    excel: Annotated[bool, typer.Option("--excel/--no-excel", help="Generate an Excel report")] = True,
    web: Annotated[bool, typer.Option("--web/--no-web", help="Generate an HTML report")] = True,
    open_web: Annotated[bool, typer.Option("--open/--no-open", help="Automatically open the HTML report in browser")] = False,
    diagnostic: Annotated[bool, typer.Option("--diagnostic", "--diagnostics", "-d", help="Generate a diagnostic report showing column status and exit without diffing")] = False,
    dedup: Annotated[bool, typer.Option("--dedup", help="Remove duplicate rows based on Key column (keeps first occurrence)")] = False,
) -> None: 
    class Args: pass
    args = Args()
    args.old_file = old_file
    args.new_file = new_file
    args.key = key
    args.sheet = sheet
    args.sheet2 = sheet2
    args.output = output
    args.prefix = prefix
    args.raw = raw
    args.excel = excel
    args.web = web
    args.open_web = open_web
    args.diagnostic = diagnostic
    args.dedup = dedup

    # console.print(f"[bold blue]DiffXL[/bold blue]: Comparing [green]'{args.old_file}'[/green] vs [green]'{args.new_file}'[/green]")
    
    if os.path.abspath(args.old_file) == os.path.abspath(args.new_file):
        console.print("[bold red]Error:[/bold red] You are comparing the same file against itself. Operation aborted.", style="red")
        sys.exit(1)
    
    df_old, report_old = None, None
    df_new, report_new = None, None
    err_old = None
    err_new = None

    # Resolve sheet arguments (supports numeric 1-based indices)
    sheet_old = _resolve_sheet(args.sheet, args.old_file)
    sheet_new = _resolve_sheet(args.sheet2 if args.sheet2 is not None else args.sheet, args.new_file)

    output_renderables = []

    try:
        with console.status("[bold green]Processing files...") as status:
            status.update(f"Loading '{args.old_file}'...")
            try:
                df_old, report_old = read_data_table(args.old_file, args.key, sheet_old)
            except SmartLoadError as e:
                err_old = e
            
            status.update(f"Loading '{args.new_file}'...")
            try:
                df_new, report_new = read_data_table(args.new_file, args.key, sheet_new)
            except SmartLoadError as e:
                err_new = e
            
        if args.diagnostic:
            with console.status("[bold green]Generating diagnostic report..."):
                report_o = err_old.report if err_old else report_old
                report_n = err_new.report if err_new else report_new
                
                # In case a different exception happened and we don't have a report, try one more time
                if not report_o:
                    try:
                        _, report_o = read_data_table(args.old_file, args.key, sheet_old)
                    except SmartLoadError as e2:
                        report_o = e2.report
                    except Exception:
                        pass
                if not report_n:
                    try:
                        _, report_n = read_data_table(args.new_file, args.key, sheet_new)
                    except SmartLoadError as e2:
                        report_n = e2.report
                    except Exception:
                        pass
                        
                diag_path = "diffxl_diagnostic.html"
                if report_o or report_n:
                    generate_diagnostic_report(report_o, report_n, diag_path)
                    
            if report_o or report_n:
                console.print(f"[bold green]Diagnostic report generated:[/bold green] {diag_path}")
                webbrowser.open(f"file://{os.path.abspath(diag_path)}")
            else:
                console.print("[bold red]Failed to generate diagnostic report (files could not be read).[/bold red]")
            sys.exit(0)
            
        if err_old or err_new:
            display_combined_analysis_report(console, err_old, err_new, args)
            sys.exit(1)
            
        output_renderables.append(Text(f"Original: '{args.old_file}' ({len(df_old)} rows)", style="green"))
        output_renderables.append(Text(f"New:      '{args.new_file}' ({len(df_new)} rows)", style="green"))
        
        # Determine actual key to use for comparison
        if args.key:
            key_col = args.key
        else:
            # Default Mode: Leftmost Column
            old_key = df_old.columns[0]
            new_key = df_new.columns[0]
            key_col = old_key
            
            if old_key != new_key:
                output_renderables.append(Text(f"Different first column in the files. Renaming second file's first column from '{new_key}' to '{old_key}'.", style="yellow"))
                df_new.rename(columns={new_key: old_key}, inplace=True)

        df_old_dups = None
        df_new_dups = None

        if args.dedup:
            output_renderables.append(Text(""))
            output_renderables.append(Text("Dedup Mode Enabled: checking for duplicates...", style="bold yellow"))
            
            # Normalize keys before dedup, just to be safe they match what compare logic sees
            df_old[key_col] = df_old[key_col].astype(str).str.strip()
            df_new[key_col] = df_new[key_col].astype(str).str.strip()
            
            # Identify ALL duplicates (keep=False)
            old_dup_mask = df_old[key_col].duplicated(keep=False)
            new_dup_mask = df_new[key_col].duplicated(keep=False)
            
            old_dups_count = old_dup_mask.sum()
            new_dups_count = new_dup_mask.sum()
            
            if old_dups_count > 0:
                df_old_dups = df_old[old_dup_mask].copy()
                df_old = df_old[~old_dup_mask].copy() # Keep only non-duplicates
                output_renderables.append(Text.from_markup(f"[yellow]Removed {old_dups_count} duplicate rows[/yellow] from original file (dropping ALL instances)."))
            
            if new_dups_count > 0:
                df_new_dups = df_new[new_dup_mask].copy()
                df_new = df_new[~new_dup_mask].copy()
                output_renderables.append(Text.from_markup(f"[yellow]Removed {new_dups_count} duplicate rows[/yellow] from new file (dropping ALL instances)."))
            
            if old_dups_count == 0 and new_dups_count == 0:
                output_renderables.append(Text("No duplicates found.", style="dim"))

        # Apply prefix to output filename
        final_output = args.output
        if args.prefix:
            final_output = f"{args.prefix}_{args.output}"

        # Perform Diff
        try:
            with console.status("[bold green]Comparing files...") as status:
                df_added, df_removed, df_changed = compare_dataframes(df_old, df_new, key_col, raw_mode=args.raw)
            
            common_keys = set(df_old[key_col]).intersection(set(df_new[key_col]))
            if not df_changed.empty:
                changed_keys = set(df_changed[key_col].unique())
            else:
                changed_keys = set()
            unchanged_count = len(common_keys) - len(changed_keys)

            old_cols = set(df_old.columns)
            new_cols = set(df_new.columns)
            common_cols = old_cols.intersection(new_cols)
            data_cols = {c for c in common_cols if c != key_col}
            
            if not data_cols:
                 console.print(Panel(f"[bold red]Error: No common columns found to compare (besides key '{key_col}').[/bold red]", border_style="red", expand=False))
                 sys.exit(1)

            added_cols = new_cols - old_cols
            removed_cols = old_cols - new_cols
            if added_cols or removed_cols:
                output_renderables.append(Text(""))
                output_renderables.append(Text("Columns added/removed:", style="bold"))
                if added_cols:
                    output_renderables.append(Text.from_markup(f"  [green]+ Added:[/green] {_fmt_list(added_cols)}"))
                if removed_cols:
                    output_renderables.append(Text.from_markup(f"  [red]- Removed:[/red] {_fmt_list(removed_cols)}"))

            # Summary
            output_renderables.append(Text(""))
            table = Table(show_header=False, header_style="bold magenta")
            table.add_column("Metric", style="cyan")
            table.add_column("Count", justify="right")
            table.add_row("Added Rows", str(len(df_added)), style="green" if len(df_added) > 0 else "dim")
            table.add_row("Removed Rows", str(len(df_removed)), style="red" if len(df_removed) > 0 else "dim")
            table.add_row("Changed Cells", str(len(df_changed)), style="yellow" if len(df_changed) > 0 else "dim")
            table.add_row("Unchanged Rows", str(unchanged_count), style="blue" if unchanged_count > 0 else "dim")
            output_renderables.append(table)
            
            # Reports
            output_renderables.append(Text(""))
            output_renderables.append(Text("Reports Generated:", style="bold"))
            if args.excel:
                from .utils import save_diff_report
                save_diff_report(df_added, df_removed, df_changed, df_new, key_col, final_output, df_old_dups, df_new_dups)
                output_renderables.append(Text.from_markup(f"  [blue]- {final_output}[/blue] (Excel)"))
            
            if args.web:
                html_output = final_output.rsplit('.', 1)[0] + ".html"
                generate_html_report(
                    df_old, df_new, df_added, df_removed, df_changed, 
                    key_col, html_output, 
                    os.path.basename(args.old_file), os.path.basename(args.new_file),
                    prefix=args.prefix if args.prefix else "",
                    df_old_dups=df_old_dups,
                    df_new_dups=df_new_dups
                )
                output_renderables.append(Text.from_markup(f"  [blue]- {html_output}[/blue] (HTML)"))
                if args.open_web:
                    webbrowser.open(f"file://{os.path.abspath(html_output)}")
                    output_renderables.append(Text.from_markup("    [green]Opening web report...[/green]"))
            
            console.print()
            console.print(Panel(
                Group(*output_renderables),
                title="[bold blue]DiffXL Report[/bold blue]",
                border_style="blue",
                expand=False
            ))

        except KeyUniquenessError as e:
            err_renderables = []
            err_renderables.append(Text.from_markup(f"[bold yellow]UID Column '{key_col}' is not unique.[/bold yellow]"))
            err_renderables.append(Text.from_markup("Use different column with [cyan]--key <column>[/cyan]"))
            err_renderables.append(Text.from_markup(" or use [cyan]--dedup[/cyan] to ignore duplicates."))
            err_renderables.append(Text(""))
            err_renderables.extend(get_candidates_renderables(e.report))
            
            console.print()
            console.print(Panel(
                Group(*err_renderables),
                title="[bold yellow]UID Column not unique[/bold yellow]",
                border_style="yellow",
                expand=False
            ))
            sys.exit(1)

        except Exception as e:
            console.print(f"[bold red]Error in Diff Process:[/bold red] {e}")
            sys.exit(1)

    except DiffXLError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected Error:[/bold red] {e}")
        sys.exit(1)

def main():
    typer.run(diff_command)

if __name__ == "__main__":
    main()