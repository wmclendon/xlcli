"""main.py - Main entry point for the xcli tool"""

from pathlib import Path
import typer
from rich import print as rprint
from rich.console import Console
from rich.table import Table
from openpyxl import load_workbook


app = typer.Typer()

console = Console()


@app.command()
def read_xlsx(
    file: str,
    sheet: str
    | None = typer.Option(None, help="Sheet name to read. Defaults to the first sheet"),
):
    """Read an Excel file and display its contents"""

    file_path = Path(file)
    if not file_path.exists():
        console.print(f"[bold red]File '{file}' not found.")
        raise typer.Exit(code=1)
    wb = load_workbook(file)
    if sheet:
        ws = wb[sheet]
    else:
        ws = wb.active

    table = Table(title=f"{file_path.stem} - {ws.title}", show_lines=True)
    for col in ws.iter_cols(1, ws.max_column):
        table.add_column(str(col[0].value))
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        table.add_row(*[str(cell) for cell in row])

    rprint(table)


if __name__ == "__main__":
    app()
