"""Command-line interface for EML to PDF converter."""

from pathlib import Path
from typing import Optional

import click
from rich.console import Console

from .converter import EMLToPDFConverter

console = Console()


@click.command()
@click.argument('input_path', type=click.Path(exists=True, path_type=Path))
@click.option(
    '--output', '-o',
    type=click.Path(path_type=Path),
    help='Output file or directory. If not specified, creates PDF next to EML file.'
)
@click.option(
    '--batch/--single', 
    default=False,
    help='Process all EML files in directory (batch mode) or single file.'
)
@click.option(
    '--recursive', '-r',
    is_flag=True,
    default=False,
    help='Recursively process EML files in subdirectories (requires --batch).'
)
def main(input_path: Path, output: Optional[Path], batch: bool, recursive: bool):
    """Convert EML files to PDF format.
    
    INPUT_PATH can be either a single EML file or a directory containing EML files.
    
    Examples:
        eml-to-pdf email.eml                    # Convert single file
        eml-to-pdf email.eml -o output.pdf     # Convert with custom output name
        eml-to-pdf ./emails/ --batch           # Convert all EML files in directory
        eml-to-pdf ./emails/ --batch -r        # Convert all EML files recursively
        eml-to-pdf ./emails/ --batch -o ./pdfs/  # Convert all with custom output directory
        eml-to-pdf ./emails/ --batch -r -o ./pdfs/  # Convert recursively with output directory
    """
    converter = EMLToPDFConverter()
    
    try:
        # Validate recursive option
        if recursive and not batch:
            console.print("❌ --recursive flag requires --batch mode")
            raise click.Abort()
            
        if input_path.is_file() and input_path.suffix.lower() == '.eml':
            # Single file conversion
            if batch:
                console.print("❌ Cannot use --batch with single file input")
                raise click.Abort()
            if recursive:
                console.print("❌ Cannot use --recursive with single file input")
                raise click.Abort()
            
            result = converter.convert_eml_to_pdf(input_path, output)
            console.print(f"📄 PDF created: [bold green]{result}[/bold green]")
            
        elif input_path.is_dir():
            # Directory conversion
            if not batch:
                console.print("💡 Use --batch flag to convert all EML files in directory")
                raise click.Abort()
            
            if recursive:
                results = converter.recursive_batch_convert(input_path, output)
            else:
                results = converter.batch_convert(input_path, output)
                
            if results:
                console.print(f"📁 Converted {len(results)} files to: [bold green]{results[0].parent}[/bold green]")
            else:
                console.print("❌ No files were converted")
                
        else:
            console.print("❌ Input must be an EML file or directory containing EML files")
            raise click.Abort()
            
    except Exception as e:
        console.print(f"❌ Error: {e}")
        raise click.Abort()


if __name__ == '__main__':
    main()