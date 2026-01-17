import argparse
import sys
import os
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
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

def display_candidates_table(console: Console, report: AnalysisReport):
    """Displays the table of candidate columns."""
    if report.candidates:
        console.print("[bold]Did you mean one of these?[/bold]")
        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("Candidate Column", style="cyan")
        table.add_column("Confidence", justify="right")
        table.add_column("Uniqueness", justify="right")
        
        for cand_name, score in report.candidates[:3]:
            stats = next((s for s in report.all_columns if s.name == cand_name), None)
            uniq_str = f"{stats.uniqueness:.1%}" if stats else "N/A"
            table.add_row(cand_name, f"{score:.2f}", uniq_str)
            
        console.print(table)
        console.print("[dim]Confidence is based on name similarity and data uniqueness (UIDs).[/dim]\n")
    else:
        console.print("[yellow]No obvious candidates found.[/yellow]\n")

def display_analysis_report(console: Console, report: AnalysisReport):
    """Displays the failure analysis report using Rich."""
    console.print()
    console.print(Panel(
        f"[bold red]Validation Failed:[/bold red] Column [yellow]'{report.missing_key}'[/yellow] not found in file.",
        title=f"Analysis: {os.path.basename(report.file_path)}",
        border_style="red"
    ))
    
    display_candidates_table(console, report)

    if report.sheets_found and len(report.sheets_found) > 1:
         console.print(f"[bold]Sheets scanned:[/bold] {', '.join(report.sheets_found)}")
         if not report.sheet_name:
             console.print("[dim]Try specifying a sheet with [cyan]--sheet[/cyan].[/dim]")
             
    console.print(Panel(
        "Tip: Run with a different key using [cyan]--key " + (f'\"{report.candidates[0][0]}\"' if report.candidates else '"NewKey"') + "[/cyan]",
        border_style="blue",
        expand=False
    ))

def main() -> None: 
    parser = argparse.ArgumentParser(description="DiffXL: Excel/CSV Comparison Tool")
    
    # Simplified CLI
    parser.add_argument("old_file", help="Path to the original file (Excel or CSV)")
    parser.add_argument("new_file", help="Path to the new file (Excel or CSV)")
    parser.add_argument("--key", "-k", help="Column name to use as unique identifier (default: Leftmost column)")
    parser.add_argument("--sheet", "-s", help="Specific sheet name to compare (Excel only)")
    parser.add_argument("--output", "-o", default="diff_report.xlsx", help="Output Excel file name (default: diff_report.xlsx)")
    parser.add_argument("--prefix", "-p", help="Add a prefix to output filenames (e.g., 'ABC' -> 'ABC_diff_report.xlsx')")
    parser.add_argument("--raw", action="store_true", help="Perform exact string comparison (disable smart normalization)")
    parser.add_argument("--no-web", action="store_true", help="Disable HTML report generation")
    parser.add_argument("--diagnostic", "-d", "--diagnostics", action="store_true", help="Generate a diagnostic report if validation fails")
    parser.add_argument("--dedup", action="store_true", help="Remove duplicate rows based on Key column (keeps first occurrence)")

    args = parser.parse_args()

    # console.print(f"[bold blue]DiffXL[/bold blue]: Comparing [green]'{args.old_file}'[/green] vs [green]'{args.new_file}'[/green]")
    
    if os.path.abspath(args.old_file) == os.path.abspath(args.new_file):
        console.print("[bold red]Error:[/bold red] You are comparing the same file against itself. Operation aborted.", style="red")
        sys.exit(1)
    
    report_old = None
    report_new = None

    try:
        with console.status("[bold green]Processing files...") as status:
            status.update(f"Loading '{args.old_file}'...")
            df_old, report_old = read_data_table(args.old_file, args.key, args.sheet)
            console.print(f"[green]✔[/green] Loaded '{args.old_file}' ({len(df_old)} rows)")
            
            status.update(f"Loading '{args.new_file}'...")
            df_new, report_new = read_data_table(args.new_file, args.key, args.sheet)
            console.print(f"[green]✔[/green] Loaded '{args.new_file}' ({len(df_new)} rows)")
            
        # Determine actual key to use for comparison
        if args.key:
            key_col = args.key
        else:
            # Default Mode: Leftmost Column
            old_key = df_old.columns[0]
            new_key = df_new.columns[0]
            key_col = old_key
            
            if old_key != new_key:
                console.print(f"Assuming they represent the same Key. Renaming '{new_key}' to '{old_key}' in new file.")
                df_new.rename(columns={new_key: old_key}, inplace=True)

        df_old_dups = None
        df_new_dups = None

        if args.dedup:
            console.print("[bold yellow]Dedup Mode Enabled:[/bold yellow] checking for duplicates...")
            
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
                console.print(f"[yellow]Removed {old_dups_count} duplicate rows[/yellow] from original file (dropping ALL instances).")
            
            if new_dups_count > 0:
                df_new_dups = df_new[new_dup_mask].copy()
                df_new = df_new[~new_dup_mask].copy()
                console.print(f"[yellow]Removed {new_dups_count} duplicate rows[/yellow] from new file (dropping ALL instances).")
            
            if old_dups_count == 0 and new_dups_count == 0:
                console.print("[dim]No duplicates found.[/dim]")

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
                 console.print(Panel(f"[bold red]Error: No common columns found to compare (besides key '{key_col}').[/bold red]", border_style="red"))
                 sys.exit(1)

            added_cols = new_cols - old_cols
            removed_cols = old_cols - new_cols
            if added_cols or removed_cols:
                warn_msg = "[bold yellow]Warning: Column Mismatch Detected[/bold yellow]\n"
                if added_cols:
                    warn_msg += f"- [green]Added Cols[/green]: {_fmt_list(added_cols)}\n"
                if removed_cols:
                    warn_msg += f"- [red]Removed Cols[/red]: {_fmt_list(removed_cols)}"
                console.print(Panel(warn_msg, title="Scope Warning", border_style="yellow", expand=False))

            # Summary
            table = Table(title="Comparison Summary", show_header=True, header_style="bold magenta")
            table.add_column("Metric", style="cyan")
            table.add_column("Count", justify="right")
            table.add_row("Added Rows", str(len(df_added)), style="green" if len(df_added) > 0 else "dim")
            table.add_row("Removed Rows", str(len(df_removed)), style="red" if len(df_removed) > 0 else "dim")
            table.add_row("Changed Cells", str(len(df_changed)), style="yellow" if len(df_changed) > 0 else "dim")
            table.add_row("Unchanged Rows", str(unchanged_count), style="blue" if unchanged_count > 0 else "dim")
            console.print(table)
            
            # Reports
            console.print(f"\nGenerating report [bold]'{final_output}'[/bold]...")
            from .utils import save_diff_report
            save_diff_report(df_added, df_removed, df_changed, df_new, key_col, final_output, df_old_dups, df_new_dups)
            
            if not args.no_web:
                html_output = final_output.rsplit('.', 1)[0] + ".html"
                console.print(f"Generating web report [bold]'{html_output}'[/bold]...")
                generate_html_report(
                    df_old, df_new, df_added, df_removed, df_changed, 
                    key_col, html_output, 
                    os.path.basename(args.old_file), os.path.basename(args.new_file),
                    prefix=args.prefix if args.prefix else "",
                    df_old_dups=df_old_dups,
                    df_new_dups=df_new_dups
                )
                console.print("[bold green]Opening web report...[/bold green]")
                webbrowser.open(f"file://{os.path.abspath(html_output)}")

            console.print("[bold green]Done![/bold green] :sparkles:")

        except KeyUniquenessError as e:
            console.print(Panel(
                f"[bold red]Key Violation:[/bold red] {str(e)}",
                title="Analysis: Key Uniqueness",
                border_style="red"
            ))
            display_candidates_table(console, e.report)
            console.print(Panel(
                "Tip: Run with a different key using [cyan]--key " + (f'\"{e.report.candidates[0][0]}\"' if e.report.candidates else '"NewKey"') + "[/cyan]\nOr use [yellow]--dedup[/yellow] to ignore duplicates.",
                border_style="blue",
                expand=False
            ))
            sys.exit(1)

        except Exception as e:
            console.print(f"[bold red]Error in Diff Process:[/bold red] {e}")
            sys.exit(1)

    except SmartLoadError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        
        if e.report:
            display_analysis_report(console, e.report)
            
            if args.diagnostic:
                # Capture the report from the error
                if os.path.abspath(e.report.file_path) == os.path.abspath(args.old_file):
                    report_old = e.report
                elif os.path.abspath(e.report.file_path) == os.path.abspath(args.new_file):
                    report_new = e.report
                
                # Attempt to get the missing report for a complete picture
                try:
                    if report_old is None:
                        _, report_old = read_data_table(args.old_file, args.key, args.sheet)
                except SmartLoadError as e2:
                    report_old = e2.report
                except Exception:
                    pass

                try:
                    if report_new is None:
                        _, report_new = read_data_table(args.new_file, args.key, args.sheet)
                except SmartLoadError as e2:
                    report_new = e2.report
                except Exception:
                    pass

                diag_path = "diffxl_diagnostic.html"
                generate_diagnostic_report(report_old, report_new, diag_path)
                console.print(f"[bold green]Diagnostic report generated:[/bold green] {diag_path}")
                webbrowser.open(f"file://{os.path.abspath(diag_path)}")
        sys.exit(1)
    except DiffXLError as e:
        console.print(f"[bold red]Error:[/bold red] {e}")
        sys.exit(1)
    except Exception as e:
        console.print(f"[bold red]Unexpected Error:[/bold red] {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()