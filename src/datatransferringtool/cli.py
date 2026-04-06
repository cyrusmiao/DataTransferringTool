import typer
import sys
import os
from pathlib import Path
from rich.console import Console
from .config import load_config
from .core import DataTransfer

app = typer.Typer(help="Data Transferring Tool CLI")
console = Console()

def print_third_party_notices():
    # Determine the path to ThirdPartyNotices.txt
    # In PyInstaller, files added via --add-data are extracted to sys._MEIPASS
    if hasattr(sys, '_MEIPASS'):
        base_path = Path(sys._MEIPASS)
    else:
        # In development, it's at the project root
        base_path = Path(__file__).parent.parent.parent
        
    notices_path = base_path / "ThirdPartyNotices.txt"
    
    if notices_path.exists():
        with open(notices_path, "r", encoding="utf-8") as f:
            console.print(f.read())
    else:
        console.print("[red]ThirdPartyNotices.txt not found.[/red]")

@app.callback(invoke_without_command=True)
def main_callback(
    ctx: typer.Context,
    third_party_notices: bool = typer.Option(
        False, 
        "--third-party-notices", 
        help="Show third-party notices and licenses."
    ),
):
    """
    Data Transferring Tool
    """
    if third_party_notices:
        print_third_party_notices()
        raise typer.Exit()
    
    if ctx.invoked_subcommand is None:
        console.print(ctx.get_help())

@app.command()
def run(
    config_path: Path = typer.Argument(..., help="Path to the YAML configuration file"),
):
    """
    Run the data transfer process using the provided YAML configuration.
    """
    if not config_path.exists():
        console.print(f"[red]Error: Configuration file '{config_path}' does not exist.[/red]")
        raise typer.Exit(code=1)

    try:
        console.print(f"[green]Loading configuration from {config_path}...[/green]")
        config = load_config(config_path)
        
        console.print("[green]Starting data transfer...[/green]")
        transfer = DataTransfer(config)
        transfer.run()
        
        console.print(f"[bold green]Data transfer completed successfully![/bold green]")
        console.print(f"Output file: {config.output_file}")
        if config.generate_transfer_report:
            console.print("Report file: transfer_report.xlsx")
        if config.generate_reference_report:
            console.print("Reference report file: reference_report.md")
        
    except Exception as e:
        console.print(f"[bold red]An error occurred during transfer:[/bold red] {e}")
        raise typer.Exit(code=1)

@app.command()
def gui():
    """
    Launch the graphical user interface.
    """
    from .gui import run_gui
    run_gui()

if __name__ == "__main__":
    app()
