#!/usr/bin/env python3
"""
PDF to Markdown Converter - Hybrid Approach
Downloads PDF, extracts text with PyMuPDF, then formats with Claude
"""

import os
import sys
import json
import requests
import argparse
import fitz  # PyMuPDF
from pathlib import Path
from typing import Optional, Dict
from dotenv import load_dotenv
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
from rich.console import Console
from rich.progress import (
    Progress,
    SpinnerColumn,
    TextColumn,
    BarColumn,
    TaskProgressColumn,
    TimeElapsedColumn,
    DownloadColumn,
    TransferSpeedColumn
)
from rich.panel import Panel
from rich.table import Table


class PDFToMarkdownConverter:
    def __init__(self, output_dir: str = "docs_mistral"):
        """Initialize the converter."""
        load_dotenv()

        self.api_key = os.getenv("OPENROUTER_API_KEY")
        if not self.api_key:
            raise ValueError(
                "OPENROUTER_API_KEY not found in environment variables.\n"
                "Please create a .env file with your API key"
            )

        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)

        self.api_url = "https://openrouter.ai/api/v1/chat/completions"
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/devinmlowe/Deltek-OPP-Docs",
            "X-Title": "Deltek OPP Documentation Converter"
        }

        self.console = Console()

    def download_pdf(self, url: str, save_path: Path) -> None:
        """Download PDF with progress bar."""
        self.console.print(f"\n[bold cyan]📥 Downloading PDF...[/bold cyan]")

        response = requests.get(url, stream=True, timeout=60)
        response.raise_for_status()

        total_size = int(response.headers.get('content-length', 0))

        with Progress(
            SpinnerColumn(),
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            DownloadColumn(),
            TransferSpeedColumn(),
            TimeElapsedColumn(),
            console=self.console
        ) as progress:
            task = progress.add_task("[cyan]Downloading...", total=total_size)

            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        progress.update(task, advance=len(chunk))

        self.console.print(f"[green]✓ Downloaded to {save_path}[/green]\n")

    def extract_text_from_pdf(self, pdf_path: Path, max_pages: Optional[int] = None) -> list:
        """Extract text from PDF using PyMuPDF."""
        self.console.print("[bold yellow]📄 Extracting text from PDF...[/bold yellow]")

        doc = fitz.open(pdf_path)
        total_pages = len(doc) if max_pages is None else min(max_pages, len(doc))

        pages_text = []

        with Progress(
            SpinnerColumn(),
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            TimeElapsedColumn(),
            console=self.console
        ) as progress:
            task = progress.add_task(
                f"[cyan]Extracting text from {total_pages} pages...",
                total=total_pages
            )

            for page_num in range(total_pages):
                page = doc[page_num]
                text = page.get_text()
                pages_text.append({
                    'page_num': page_num + 1,
                    'text': text
                })
                progress.update(task, advance=1)

        doc.close()
        self.console.print(f"[green]✓ Extracted text from {total_pages} pages[/green]\n")

        return pages_text

    def _process_single_chunk_api(self, chunk_pages: list, chunk_label: str) -> str:
        """Make a single API call to process pages and return markdown content."""
        # Combine chunk text
        chunk_text = "\n\n---PAGE BREAK---\n\n".join([
            f"PAGE {p['page_num']}:\n{p['text']}"
            for p in chunk_pages
        ])

        prompt = f"""Convert this PDF text extract to clean, well-formatted GitHub Flavored Markdown.

The text is from pages {chunk_pages[0]['page_num']} to {chunk_pages[-1]['page_num']} of a technical document.

Requirements:
- Remove headers, footers, and page numbers
- Preserve all headings using proper markdown levels (# ## ### etc.)
- Maintain code examples in proper code blocks with language identifiers
- Convert tables to markdown table format
- Preserve the document structure
- Keep all technical content
- Use proper markdown formatting for lists, emphasis, and links
- Remove artifacts like repeated headers or footers

Here is the extracted text:

{chunk_text}

Please convert this to clean markdown format."""

        messages = [{
            "role": "user",
            "content": prompt
        }]

        payload = {
            "model": "anthropic/claude-sonnet-4",
            "messages": messages,
            "max_tokens": 50000
        }

        # Retry logic
        max_retries = 3
        retry_delay = 5

        for attempt in range(max_retries):
            try:
                response = requests.post(
                    self.api_url,
                    headers=self.headers,
                    json=payload,
                    timeout=300
                )
                response.raise_for_status()

                data = response.json()
                if 'choices' in data and len(data['choices']) > 0:
                    markdown = data['choices'][0]['message']['content']
                    return markdown
                else:
                    return f"<!-- No content generated for {chunk_label} -->"

            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                    retry_delay *= 2
                else:
                    raise Exception(f"{chunk_label} failed after {max_retries} attempts: {e}")

            except Exception as e:
                raise Exception(f"Error processing {chunk_label}: {e}")

    def _process_single_chunk(self, chunk_idx: int, chunk_pages: list, num_chunks: int,
                             min_subdivision_size: int = 5) -> tuple:
        """
        Process a single chunk with automatic subdivision on failure.

        If processing fails, the chunk is subdivided into smaller chunks and each
        is processed recursively. This ensures resilience against timeouts and
        transient failures.

        Args:
            chunk_idx: Index of the chunk in the overall document
            chunk_pages: List of page dictionaries to process
            num_chunks: Total number of chunks (for progress reporting)
            min_subdivision_size: Minimum pages per subdivision (default: 5)

        Returns:
            tuple: (chunk_idx, markdown_content)
        """
        chunk_label = f"Chunk {chunk_idx + 1}/{num_chunks}"

        try:
            # Try to process the entire chunk
            markdown = self._process_single_chunk_api(chunk_pages, chunk_label)
            return (chunk_idx, markdown)

        except Exception as e:
            # If chunk is already at minimum size, can't subdivide further
            if len(chunk_pages) <= min_subdivision_size:
                self.console.print(
                    f"[red]✗ {chunk_label} failed and cannot be subdivided further "
                    f"({len(chunk_pages)} pages)[/red]"
                )
                raise

            # Subdivide the chunk into smaller pieces
            num_pages = len(chunk_pages)
            sub_chunk_size = max(min_subdivision_size, num_pages // 2)

            self.console.print(
                f"[yellow]⚠ {chunk_label} failed with {len(chunk_pages)} pages. "
                f"Subdividing into ~{sub_chunk_size}-page chunks...[/yellow]"
            )

            # Process subdivisions recursively
            markdown_parts = []
            for i in range(0, num_pages, sub_chunk_size):
                sub_pages = chunk_pages[i:i + sub_chunk_size]
                sub_label = f"{chunk_label} (sub-chunk {i//sub_chunk_size + 1}, pages {sub_pages[0]['page_num']}-{sub_pages[-1]['page_num']})"

                try:
                    # Recursive call with subdivision
                    sub_markdown = self._process_single_chunk_api(sub_pages, sub_label)
                    markdown_parts.append(sub_markdown)
                    self.console.print(f"[green]✓ {sub_label} processed successfully[/green]")

                except Exception as sub_e:
                    # If subdivision still fails, try even smaller chunks
                    if len(sub_pages) > 1:
                        self.console.print(
                            f"[yellow]⚠ {sub_label} failed. "
                            f"Further subdividing into 1-page chunks...[/yellow]"
                        )

                        # Process one page at a time as last resort
                        for page in sub_pages:
                            page_label = f"{chunk_label} (page {page['page_num']})"
                            try:
                                page_markdown = self._process_single_chunk_api([page], page_label)
                                markdown_parts.append(page_markdown)
                                self.console.print(f"[green]✓ {page_label} processed[/green]")
                            except Exception as page_e:
                                self.console.print(f"[red]✗ {page_label} failed: {page_e}[/red]")
                                markdown_parts.append(f"<!-- Failed to process page {page['page_num']} -->")
                    else:
                        raise sub_e

            # Combine all subdivisions
            combined_markdown = "\n\n".join(markdown_parts)
            self.console.print(f"[green]✓ {chunk_label} completed via subdivision[/green]")
            return (chunk_idx, combined_markdown)

    def convert_to_markdown(self, pages_text: list, chunk_size: int = 25,
                          checkpoint_file: Optional[Path] = None, max_workers: int = 4) -> str:
        """Convert extracted text to markdown using Claude with parallel processing."""
        self.console.print("[bold yellow]🚀 Converting to Markdown with Claude Sonnet 4...[/bold yellow]")
        self.console.print(f"[dim]Using {max_workers} parallel workers for faster processing[/dim]\n")

        total_pages = len(pages_text)
        num_chunks = (total_pages + chunk_size - 1) // chunk_size

        # Load checkpoint if exists
        completed_chunks: Dict[int, str] = {}

        if checkpoint_file and checkpoint_file.exists():
            with open(checkpoint_file, 'r', encoding='utf-8') as f:
                checkpoint_data = json.load(f)
                completed_chunks = {int(k): v for k, v in checkpoint_data.get('completed_chunks', {}).items()}
                self.console.print(f"[cyan]📋 Resuming: {len(completed_chunks)}/{num_chunks} chunks already completed[/cyan]\n")

        # Prepare all chunks
        chunks_to_process = []
        for chunk_idx in range(num_chunks):
            if chunk_idx not in completed_chunks:
                start_idx = chunk_idx * chunk_size
                end_idx = min((chunk_idx + 1) * chunk_size, total_pages)
                chunk_pages = pages_text[start_idx:end_idx]
                chunks_to_process.append((chunk_idx, chunk_pages))

        # Process chunks in parallel
        if not chunks_to_process:
            self.console.print(f"[green]✓ All {num_chunks} chunks already completed![/green]\n")
        else:
            with Progress(
                SpinnerColumn(),
                TextColumn("[bold blue]{task.description}"),
                BarColumn(),
                TaskProgressColumn(),
                TimeElapsedColumn(),
                console=self.console
            ) as progress:
                task = progress.add_task(
                    f"[cyan]Processing {len(chunks_to_process)} chunks in parallel...",
                    total=len(chunks_to_process)
                )

                # Use ThreadPoolExecutor for parallel processing
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # Submit all chunks
                    future_to_chunk = {
                        executor.submit(self._process_single_chunk, chunk_idx, chunk_pages, num_chunks): chunk_idx
                        for chunk_idx, chunk_pages in chunks_to_process
                    }

                    # Process completed chunks as they finish
                    for future in as_completed(future_to_chunk):
                        chunk_idx = future_to_chunk[future]
                        try:
                            chunk_idx, markdown = future.result()
                            completed_chunks[chunk_idx] = markdown

                            # Save checkpoint
                            if checkpoint_file:
                                with open(checkpoint_file, 'w', encoding='utf-8') as f:
                                    json.dump({
                                        'completed_chunks': completed_chunks
                                    }, f)

                            progress.update(task, advance=1)

                        except Exception as e:
                            self.console.print(f"\n[red]✗ {str(e)}[/red]")
                            raise

            self.console.print(f"[green]✓ Converted {num_chunks} chunks to markdown[/green]\n")

        # Combine chunks in order
        markdown_parts = [completed_chunks[i] for i in range(num_chunks)]
        return "\n\n".join(markdown_parts)

    def convert(self, pdf_url: str, output_filename: str = "DeltekOpenPlanDeveloperGuide.md",
                max_pages: Optional[int] = None, chunk_size: int = 25, max_workers: int = 4) -> Path:
        """Main conversion workflow."""
        output_path = self.output_dir / output_filename

        # Display info
        self.console.print("\n" + "="*80)
        self.console.print(Panel.fit(
            "[bold cyan]PDF to Markdown Converter (Parallel Processing)[/bold cyan]\n"
            f"[white]PDF URL:[/white] {pdf_url}\n"
            f"[white]Output:[/white] {output_path}\n"
            f"[white]Parallel workers:[/white] {max_workers}",
            border_style="cyan"
        ))
        self.console.print("="*80)

        # Download PDF
        temp_pdf = self.output_dir.parent / "temp_downloaded.pdf"
        self.download_pdf(pdf_url, temp_pdf)

        # Extract text
        pages_text = self.extract_text_from_pdf(temp_pdf, max_pages)

        # Convert to markdown with checkpoint support and parallel processing
        checkpoint_file = self.output_dir / ".checkpoint.json"
        markdown_content = self.convert_to_markdown(pages_text, chunk_size, checkpoint_file, max_workers)

        # Save
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        # Cleanup
        temp_pdf.unlink()
        if checkpoint_file.exists():
            checkpoint_file.unlink()  # Remove checkpoint after successful completion

        # Display success
        file_size = output_path.stat().st_size
        self.console.print(f"[bold green]✓ Conversion completed successfully![/bold green]")

        summary = Table(show_header=False, box=None, padding=(0, 2))
        summary.add_row("[cyan]Output file:", f"[white]{output_path}[/white]")
        summary.add_row("[cyan]File size:", f"[white]{file_size:,} bytes[/white]")
        summary.add_row("[cyan]Pages processed:", f"[white]{len(pages_text)}[/white]")

        self.console.print(summary)
        self.console.print()

        return output_path


def main():
    """Main entry point."""
    parser = argparse.ArgumentParser(
        description='Convert PDF documents to Markdown using text extraction + Claude',
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    parser.add_argument(
        '--pdf-url',
        type=str,
        default='https://dsm.deltek.com/DeltekSoftwareManagerWebServices/downloadFile.ashx?documentid=C6E40CBC-E0A5-4722-8E62-1E827AD56D8A',
        help='URL to the PDF document'
    )
    parser.add_argument(
        '--output',
        type=str,
        default='DeltekOpenPlanDeveloperGuide.md',
        help='Output markdown filename'
    )
    parser.add_argument(
        '--output-dir',
        type=str,
        default='docs_mistral',
        help='Output directory'
    )
    parser.add_argument(
        '--max-pages',
        type=int,
        help='Maximum pages to process'
    )
    parser.add_argument(
        '--chunk-size',
        type=int,
        default=25,
        help='Pages per processing chunk (default: 25, smaller = more reliable)'
    )
    parser.add_argument(
        '--workers',
        type=int,
        default=4,
        help='Number of parallel workers (default: 4, max recommended: 6)'
    )

    args = parser.parse_args()

    try:
        converter = PDFToMarkdownConverter(output_dir=args.output_dir)
        converter.convert(
            pdf_url=args.pdf_url,
            output_filename=args.output,
            max_pages=args.max_pages,
            chunk_size=args.chunk_size,
            max_workers=args.workers
        )
        return 0

    except KeyboardInterrupt:
        Console().print("\n[yellow]Conversion cancelled[/yellow]")
        return 1
    except Exception as e:
        Console().print(f"\n[bold red]Error:[/bold red] {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
