#!/usr/bin/env python3
"""
Excel/CSV to Markdown Converter

This script converts Excel (.xlsx, .xls) or CSV files to Markdown tables.
It provides both a command-line interface and functions for programmatic use.
"""

import os
import sys
import pandas as pd
import click
from rich.console import Console
from rich.table import Table
from rich.progress import Progress
from rich import print as rprint
from pathlib import Path
import glob
import chardet


def detect_encoding(file_path):
    """
    Detect the encoding of a file
    
    Args:
        file_path (str): Path to the file
    
    Returns:
        str: Detected encoding
    """
    # Read the first 10000 bytes to detect encoding
    with open(file_path, 'rb') as f:
        raw_data = f.read(10000)
        
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    confidence = result['confidence']
    
    # Default to common Chinese encodings if detection fails or has low confidence
    if encoding is None or confidence < 0.7:
        # Try common Chinese encodings
        for enc in ['gb18030', 'gbk', 'gb2312', 'utf-8', 'utf-16', 'big5']:
            try:
                with open(file_path, 'r', encoding=enc) as f:
                    f.read(100)  # Try reading a small sample
                return enc
            except UnicodeDecodeError:
                continue
        
        # If all fail, default to utf-8
        return 'utf-8'
    
    return encoding


def read_file(file_path):
    """
    Read Excel or CSV file into a pandas DataFrame
    
    Args:
        file_path (str): Path to the Excel or CSV file
    
    Returns:
        pandas.DataFrame: The content of the file
    """
    file_ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if file_ext in ['.xlsx', '.xls']:
            return pd.read_excel(file_path)
        elif file_ext == '.csv':
            # 检测CSV文件编码
            encoding = detect_encoding(file_path)
            console = Console()
            console.print(f"Detected encoding for [cyan]{os.path.basename(file_path)}[/cyan]: [yellow]{encoding}[/yellow]")
            
            # 使用检测到的编码读取CSV文件
            return pd.read_csv(file_path, encoding=encoding)
        else:
            raise ValueError(f"Unsupported file extension: {file_ext}. Only .xlsx, .xls, and .csv are supported.")
    except Exception as e:
        raise Exception(f"Error reading file: {e}")


def dataframe_to_markdown(df, headers=True):
    """
    Convert a pandas DataFrame to a Markdown table
    
    Args:
        df (pandas.DataFrame): The DataFrame to convert
        headers (bool): Whether to include the column headers
    
    Returns:
        str: Markdown table representation
    """
    # Handle empty dataframe
    if df.empty:
        return "Empty table"
    
    # Convert DataFrame to markdown
    markdown_table = []
    
    # Add headers if requested
    if headers:
        header_row = "| " + " | ".join(str(col) for col in df.columns) + " |"
        separator_row = "| " + " | ".join(["---"] * len(df.columns)) + " |"
        markdown_table.append(header_row)
        markdown_table.append(separator_row)
    
    # Add data rows
    for _, row in df.iterrows():
        data_row = "| " + " | ".join(str(val) if pd.notna(val) else "" for val in row) + " |"
        markdown_table.append(data_row)
    
    return "\n".join(markdown_table)


def convert_excel_to_markdown(input_file, output_file=None, sheet_name=None, preview=False, progress=None):
    """
    Convert an Excel/CSV file to Markdown format
    
    Args:
        input_file (str): Path to the input Excel/CSV file
        output_file (str, optional): Path to the output Markdown file. If None, output to console
        sheet_name (str, optional): For Excel files, the name of the sheet to convert. If None, convert all sheets
        preview (bool): Whether to preview the output
        progress (Progress, optional): An existing Progress instance to use
    
    Returns:
        str or tuple: The Markdown table(s) if successful, or (None, error_message) if error in batch mode
    """
    console = Console()
    
    try:
        # For Excel files with multiple sheets
        file_ext = os.path.splitext(input_file)[1].lower()
        all_markdown = []

        if file_ext in ['.xlsx', '.xls'] and sheet_name is None:
            # Get all sheet names
            excel_file = pd.ExcelFile(input_file)
            sheet_names = excel_file.sheet_names
            
            # Use provided progress or create a new one if not in batch mode
            if progress:
                sheets_task = progress.add_task(f"[blue]Converting sheets in {os.path.basename(input_file)}...", total=len(sheet_names))
                for sheet in sheet_names:
                    df = pd.read_excel(input_file, sheet_name=sheet)
                    markdown = f"## Sheet: {sheet}\n\n" + dataframe_to_markdown(df)
                    all_markdown.append(markdown)
                    progress.update(sheets_task, advance=1)
            else:
                # Not in batch mode, create a new progress bar
                with Progress() as local_progress:
                    sheets_task = local_progress.add_task("[cyan]Converting sheets...", total=len(sheet_names))
                    for sheet in sheet_names:
                        df = pd.read_excel(input_file, sheet_name=sheet)
                        markdown = f"## Sheet: {sheet}\n\n" + dataframe_to_markdown(df)
                        all_markdown.append(markdown)
                        local_progress.update(sheets_task, advance=1)
                    
        else:
            # For CSV or specific Excel sheet
            if file_ext in ['.xlsx', '.xls'] and sheet_name is not None:
                df = pd.read_excel(input_file, sheet_name=sheet_name)
            else:
                df = read_file(input_file)
            
            markdown = dataframe_to_markdown(df)
            all_markdown.append(markdown)
        
        # Combine all markdown content
        full_markdown = "\n\n".join(all_markdown)
        
        # Preview if requested and not in batch mode
        if preview and not progress:
            for i, markdown_section in enumerate(all_markdown):
                if i > 0:
                    console.print("\n---\n")
                console.print(markdown_section)
            
        # Output to file if specified
        if output_file:
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(full_markdown)
            if not progress:  # Only print if not in batch mode
                console.print(f"[green]Successfully converted to Markdown and saved to [bold]{output_file}[/bold][/green]")
        
        return full_markdown
        
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        if progress:
            # When in batch mode, return a tuple with None and error message
            console.print(f"[red]{error_msg}[/red]")
            return (None, error_msg)
        else:
            # When in single file mode, print the error and exit
            console.print(f"[bold red]{error_msg}[/bold red]")
            sys.exit(1)


def convert_directory(input_dir, output_dir=None, recursive=False, preview=False):
    """
    Convert all Excel/CSV files in a directory to Markdown
    
    Args:
        input_dir (str): Path to the input directory
        output_dir (str, optional): Path to the output directory. If None, use the input directory
        recursive (bool): Whether to search for files recursively
        preview (bool): Whether to preview the output
    
    Returns:
        int: Number of files successfully converted
    """
    console = Console()
    console.print(f"Processing directory [cyan]{input_dir}[/cyan] (recursive={recursive})")
    
    # Make sure input directory exists
    if not os.path.isdir(input_dir):
        console.print(f"[bold red]Error: Input directory '{input_dir}' does not exist[/bold red]")
        sys.exit(1)
    
    # Set output directory
    if output_dir is None:
        output_dir = input_dir
    else:
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        console.print(f"Output directory: [green]{output_dir}[/green]")
    
    # Find all Excel/CSV files
    all_files = []
    
    # 查找文件的方式改进
    if recursive:
        # 使用Path.rglob进行递归查找
        for ext in ['xlsx', 'xls', 'csv']:
            found_files = [str(p) for p in Path(input_dir).rglob(f'*.{ext}')]
            all_files.extend(found_files)
    else:
        # 使用Path.glob进行非递归查找
        for ext in ['xlsx', 'xls', 'csv']:
            found_files = [str(p) for p in Path(input_dir).glob(f'*.{ext}')]
            all_files.extend(found_files)
    
    # 输出找到的文件详情
    if not all_files:
        console.print(f"[yellow]No Excel or CSV files found in '{input_dir}'[/yellow]")
        return 0
    
    # 显示找到的文件类型统计
    excel_count = sum(1 for f in all_files if f.lower().endswith(('.xlsx', '.xls')))
    csv_count = sum(1 for f in all_files if f.lower().endswith('.csv'))
    console.print(f"Found [cyan]{excel_count}[/cyan] Excel files and [cyan]{csv_count}[/cyan] CSV files.")
    
    # 输出所有找到的文件列表
    console.print("[blue]Files to convert:[/blue]")
    for i, file_path in enumerate(all_files):
        file_type = "Excel" if file_path.lower().endswith(('.xlsx', '.xls')) else "CSV"
        console.print(f"  {i+1}. [{file_type}] {file_path}")
    
    # Convert each file
    success_count = 0
    error_count = 0
    with Progress() as progress:
        task = progress.add_task("[cyan]Converting files...", total=len(all_files))
        
        for file_path in all_files:
            try:
                # Generate output filename
                rel_path = os.path.relpath(file_path, input_dir)
                output_path = os.path.join(output_dir, os.path.splitext(rel_path)[0] + '.md')
                
                # Create subdirectories if needed
                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                
                # Convert the file
                console.print(f"Converting: [cyan]{file_path}[/cyan] -> [green]{output_path}[/green]")
                
                # 使用同一个函数处理所有文件类型（CSV和Excel）
                result = convert_excel_to_markdown(file_path, output_path, preview=preview, progress=progress)
                
                # 检查返回值类型
                if isinstance(result, tuple) and result[0] is None:
                    # 转换失败
                    error_count += 1
                    console.print(f"[red]✘ Failed to convert {file_path}: {result[1]}[/red]")
                else:
                    # 转换成功
                    success_count += 1
                    console.print(f"[green]✓ Successfully converted {file_path}[/green]")
                    
            except Exception as e:
                error_count += 1
                console.print(f"[red]✘ Error converting {file_path}: {str(e)}[/red]")
            
            progress.update(task, advance=1)
    
    # 显示处理结果摘要
    console.print(f"[green]Successfully converted {success_count} out of {len(all_files)} files[/green]")
    if error_count > 0:
        console.print(f"[red]Failed to convert {error_count} files[/red]")
    
    return success_count


@click.command()
@click.argument('input_path', type=click.Path(exists=True))
@click.option('-o', '--output', type=click.Path(), help='Output Markdown file path or directory')
@click.option('-s', '--sheet', help='Sheet name (for Excel files)')
@click.option('-p', '--preview', is_flag=True, help='Preview the Markdown output')
@click.option('-d', '--directory', is_flag=True, help='Process all Excel/CSV files in the directory')
@click.option('-r', '--recursive', is_flag=True, help='Recursively process subdirectories (when --directory is used)')
@click.option('-e', '--encoding', help='Specify encoding for CSV files (e.g., utf-8, gbk, gb18030)')
def main(input_path, output, sheet, preview, directory, recursive, encoding):
    """Convert Excel/CSV file(s) to Markdown table format
    
    INPUT_PATH can be a single file or a directory (when --directory is used).
    """
    # Print a welcome message
    rprint("[bold blue]Excel/CSV to Markdown Converter[/bold blue]")
    
    # 首先检查是否需要安装chardet
    try:
        import chardet
    except ImportError:
        console = Console()
        console.print("[yellow]Package 'chardet' is required for automatic encoding detection.[/yellow]")
        console.print("[yellow]Installing 'chardet'...[/yellow]")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "chardet"])
            console.print("[green]Successfully installed 'chardet'.[/green]")
        except subprocess.CalledProcessError:
            console.print("[red]Failed to install 'chardet'. Please install it manually: pip install chardet[/red]")
            sys.exit(1)
    
    # Check if we're processing a directory
    if directory:
        rprint(f"Processing directory: [cyan]{input_path}[/cyan]")
        convert_directory(input_path, output, recursive, preview)
        return
    
    # Otherwise process a single file
    rprint(f"Converting: [cyan]{input_path}[/cyan]")
    
    # If no output file is specified and not preview mode, default to preview
    if not output and not preview:
        preview = True
    
    # If no output file is specified but we want to save, use the same name with .md extension
    if not output and not preview:
        output = str(Path(input_path).with_suffix('.md'))
    
    convert_excel_to_markdown(input_path, output, sheet, preview)


if __name__ == '__main__':
    main() 