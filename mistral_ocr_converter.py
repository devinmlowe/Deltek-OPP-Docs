#!/usr/bin/env python3
"""
PDF to Markdown Converter using Mistral OCR via OpenRouter
Downloads and processes PDF documents using Mistral's OCR capabilities
"""

import os
import sys
import json
import requests
import base64
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv
import time
from rich.console import Console
from rich.progress import (
    Progress,
    SpinnerColumn,
    TextColumn,
    BarColumn,
    TaskProgressColumn,
    TimeRemainingColumn,
    TimeElapsedColumn,
    DownloadColumn,
    TransferSpeedColumn
)
from rich.panel import Panel
from rich.table import Table
from rich.live import Live
from rich.layout import Layout

class MistralOCRConverter:
    def __init__(self, output_dir: str = "docs_mistral"):
        """
        Initialize the Mistral OCR converter.

        Args:
            output_dir: Directory for output files
        """
        # Load environment variables from .env file
        load_dotenv()

        self.api_key = os.getenv("OPENROUTER_API_KEY")
        if not self.api_key:
            raise ValueError(
                "OPENROUTER_API_KEY not found in environment variables.\n"
                "Please create a .env file with your API key:\n"
                "OPENROUTER_API_KEY=your_key_here"
            )

        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)

        self.api_url = "https://openrouter.ai/api/v1/chat/completions"
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/deltek-opp-docs",
            "X-Title": "Deltek OPP Documentation Converter"
        }

        # Initialize rich console for beautiful output
        self.console = Console()

        # Create progress bar for different phases
        self.progress = Progress(
            SpinnerColumn(),
            TextColumn("[bold blue]{task.description}"),
            BarColumn(),
            TaskProgressColumn(),
            TimeElapsedColumn(),
            TimeRemainingColumn(),
            console=self.console,
            transient=False
        )

    def download_pdf(self, url: str, save_path: Optional[str] = None) -> Path:
        """
        Download a PDF from a URL with rich progress display.

        Args:
            url: URL to download from
            save_path: Optional path to save the file

        Returns:
            Path to the downloaded file
        """
        if save_path is None:
            save_path = self.output_dir.parent / "downloaded_document.pdf"
        else:
            save_path = Path(save_path)

        self.console.print(f"\n[bold cyan]📥 Downloading PDF from URL...[/bold cyan]")
        self.console.print(f"   URL: [link={url}]{url}[/link]")

        try:
            # Download with streaming to handle large files
            response = requests.get(url, stream=True, timeout=60)
            response.raise_for_status()

            # Get total file size if available
            total_size = int(response.headers.get('content-length', 0))

            # Create a download progress bar
            download_progress = Progress(
                SpinnerColumn(),
                TextColumn("[bold blue]{task.description}"),
                BarColumn(),
                DownloadColumn(),
                TransferSpeedColumn(),
                TaskProgressColumn(),
                TimeRemainingColumn(),
                console=self.console
            )

            # Write the file with progress
            with download_progress:
                download_task = download_progress.add_task(
                    "Downloading",
                    total=total_size
                )

                with open(save_path, 'wb') as f:
                    if total_size == 0:
                        f.write(response.content)
                        download_progress.update(download_task, completed=len(response.content))
                    else:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                            download_progress.update(download_task, advance=len(chunk))

            file_size_mb = save_path.stat().st_size / (1024*1024)
            self.console.print(f"[green]✓[/green] Downloaded: {save_path}")
            self.console.print(f"   Size: [yellow]{file_size_mb:.2f} MB[/yellow]")
            return save_path

        except requests.exceptions.RequestException as e:
            self.console.print(f"[red]❌ Error downloading PDF: {e}[/red]")
            raise

    def convert_pdf_with_mistral_ocr(
        self,
        pdf_path: str,
        prompt: Optional[str] = None,
        chunk_pages: int = 50
    ) -> str:
        """
        Convert a PDF to markdown using Mistral OCR via OpenRouter.

        Args:
            pdf_path: Path to PDF file or URL
            prompt: Custom prompt for the conversion (optional)
            chunk_pages: Number of pages to process at once (for large PDFs)

        Returns:
            Markdown formatted content
        """
        # If pdf_path is a URL, use it directly; otherwise convert to file URL
        if pdf_path.startswith('http://') or pdf_path.startswith('https://'):
            file_url = pdf_path
        else:
            # For local files, we need to upload or convert to base64
            # For now, we'll use the URL approach
            file_url = pdf_path

        if prompt is None:
            prompt = """Please convert this PDF document to well-formatted GitHub Flavored Markdown.

Instructions:
1. Extract all text content preserving the document structure
2. Format headings appropriately (# for main titles, ## for sections, etc.)
3. Convert all tables to markdown table format
4. Preserve code blocks and examples with proper formatting
5. Extract and describe any figures, diagrams, or images with descriptive alt text
6. Remove repetitive headers and footers
7. Maintain proper paragraph breaks and spacing
8. Format lists (numbered and bulleted) correctly
9. Preserve any important formatting like bold, italic, or code inline

Focus on creating a clean, readable markdown document that preserves the technical content and structure of the original PDF."""

        messages = [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": prompt
                    },
                    {
                        "type": "file",
                        "file": {
                            "filename": "document.pdf",
                            "file_data": file_url
                        }
                    }
                ]
            }
        ]

        # Configure to use Mistral OCR
        plugins = [
            {
                "id": "file-parser",
                "pdf": {
                    "engine": "mistral-ocr"
                }
            }
        ]

        payload = {
            "model": "anthropic/claude-sonnet-4",  # Using Claude for markdown generation
            "messages": messages,
            "plugins": plugins,
            "max_tokens": 200000  # Large token limit for comprehensive conversion
        }

        print(f"\n🔄 Processing PDF with Mistral OCR...")
        print(f"   API: OpenRouter")
        print(f"   OCR Engine: Mistral OCR")
        print(f"   Model: Claude Sonnet 4")

        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=payload,
                timeout=600  # 10 minute timeout for large documents
            )
            response.raise_for_status()

            result = response.json()

            # Extract the markdown content from the response
            if 'choices' in result and len(result['choices']) > 0:
                markdown_content = result['choices'][0]['message']['content']

                # Show token usage if available
                if 'usage' in result:
                    usage = result['usage']
                    print(f"\n📊 Token Usage:")
                    print(f"   Prompt tokens: {usage.get('prompt_tokens', 'N/A')}")
                    print(f"   Completion tokens: {usage.get('completion_tokens', 'N/A')}")
                    print(f"   Total tokens: {usage.get('total_tokens', 'N/A')}")

                return markdown_content
            else:
                print(f"❌ Unexpected API response format:")
                print(json.dumps(result, indent=2))
                raise ValueError("Could not extract content from API response")

        except requests.exceptions.Timeout:
            print(f"❌ Request timed out. The document may be too large.")
            print(f"   Try processing a smaller PDF or increasing the timeout.")
            raise
        except requests.exceptions.RequestException as e:
            print(f"❌ API request failed: {e}")
            if hasattr(e.response, 'text'):
                print(f"   Response: {e.response.text}")
            raise

    def save_markdown(self, content: str, filename: str = "document.md") -> Path:
        """
        Save markdown content to a file with rich output.

        Args:
            content: Markdown content
            filename: Output filename

        Returns:
            Path to saved file
        """
        output_path = self.output_dir / filename

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)

        file_size_kb = len(content.encode('utf-8')) / 1024
        self.console.print(f"\n[green]✓[/green] Saved markdown: [link=file://{output_path}]{output_path}[/link]")
        self.console.print(f"   Size: [yellow]{file_size_kb:.1f} KB[/yellow]")
        return output_path

    def split_pdf(self, pdf_path: Path, pages_per_chunk: int = 50) -> list[Path]:
        """
        Split a large PDF into smaller chunks with progress display.

        Args:
            pdf_path: Path to the PDF file
            pages_per_chunk: Number of pages per chunk

        Returns:
            List of paths to the chunk files
        """
        import fitz  # PyMuPDF

        self.console.print(f"\n[bold cyan]📄 Splitting PDF into chunks...[/bold cyan]")

        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        self.console.print(f"   Total pages: [yellow]{total_pages}[/yellow]")
        self.console.print(f"   Pages per chunk: [yellow]{pages_per_chunk}[/yellow]")

        chunk_paths = []
        num_chunks = (total_pages + pages_per_chunk - 1) // pages_per_chunk

        with self.progress:
            split_task = self.progress.add_task(
                "[cyan]Splitting PDF",
                total=num_chunks
            )

            for chunk_num in range(num_chunks):
                start_page = chunk_num * pages_per_chunk
                end_page = min(start_page + pages_per_chunk, total_pages)

                # Create a new PDF with just these pages
                chunk_doc = fitz.open()
                chunk_doc.insert_pdf(doc, from_page=start_page, to_page=end_page - 1)

                # Save the chunk
                chunk_filename = f"chunk_{chunk_num + 1:03d}_pages_{start_page + 1}-{end_page}.pdf"
                chunk_path = self.output_dir.parent / "temp_chunks" / chunk_filename
                chunk_path.parent.mkdir(exist_ok=True)

                chunk_doc.save(chunk_path)
                chunk_doc.close()

                chunk_paths.append(chunk_path)
                self.progress.update(split_task, advance=1)

        self.console.print(f"[green]✓[/green] Created {num_chunks} chunks")
        doc.close()
        return chunk_paths

    def convert_pdf_chunks(
        self,
        pdf_path: Path,
        pages_per_chunk: int = 50,
        delay_between_chunks: float = 2.0,
        max_pages: Optional[int] = None
    ) -> str:
        """
        Convert a large PDF by splitting it into chunks and processing each with rich progress.

        Args:
            pdf_path: Path to the PDF file
            pages_per_chunk: Pages per chunk (max ~50 to stay under 100 image limit)
            delay_between_chunks: Seconds to wait between API calls
            max_pages: Maximum number of pages to process (for testing)

        Returns:
            Combined markdown content
        """
        import fitz

        # First, check the PDF size
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        doc.close()

        if max_pages:
            total_pages = min(total_pages, max_pages)
            self.console.print(f"\n[bold cyan]📚 Processing PDF in chunks[/bold cyan] [yellow](limited to {max_pages} pages)[/yellow]")
        else:
            self.console.print(f"\n[bold cyan]📚 Processing PDF in chunks...[/bold cyan]")

        self.console.print(f"   Total pages: [yellow]{total_pages}[/yellow]")
        num_chunks = (total_pages + pages_per_chunk - 1) // pages_per_chunk
        self.console.print(f"   Estimated chunks: [yellow]{num_chunks}[/yellow]")

        # Split the PDF
        chunk_paths = self.split_pdf(pdf_path, pages_per_chunk)

        # Limit chunks if max_pages is set
        if max_pages:
            num_chunks_needed = (max_pages + pages_per_chunk - 1) // pages_per_chunk
            chunk_paths = chunk_paths[:num_chunks_needed]

        # Process each chunk
        all_markdown = []

        for i, chunk_path in enumerate(chunk_paths):
            self.console.print(f"\n[bold blue]🔄 Processing chunk {i + 1}/{len(chunk_paths)}...[/bold blue]")

            # For now, we still need to handle the file upload issue
            # Let's try a different approach: extract pages as images
            chunk_markdown = self.process_chunk_with_images(
                chunk_path,
                i + 1,
                max_pages=max_pages - (i * pages_per_chunk) if max_pages else None
            )
            all_markdown.append(chunk_markdown)

            # Delay between chunks to avoid rate limiting
            if i < len(chunk_paths) - 1:
                self.console.print(f"   [yellow]⏳ Waiting {delay_between_chunks}s before next chunk...[/yellow]")
                time.sleep(delay_between_chunks)

        # Combine all markdown
        combined_markdown = "\n\n---\n\n".join(all_markdown)
        return combined_markdown

    def process_chunk_with_images(self, chunk_path: Path, chunk_num: int, max_pages: Optional[int] = None) -> str:
        """
        Process a PDF chunk by converting pages to images and using vision API with rich progress.

        Args:
            chunk_path: Path to the chunk PDF
            chunk_num: Chunk number for reference
            max_pages: Maximum pages to process in this chunk

        Returns:
            Markdown content for this chunk
        """
        import fitz
        import base64

        # Open the PDF chunk
        doc = fitz.open(chunk_path)

        # Limit pages if needed
        num_pages = len(doc) if not max_pages else min(len(doc), max_pages)

        # For large chunks, we might still hit limits
        # Let's process page by page for better control
        page_markdowns = []

        with self.progress:
            page_task = self.progress.add_task(
                f"[cyan]Processing chunk {chunk_num}",
                total=num_pages
            )

            for page_num in range(num_pages):
                page = doc[page_num]

                # Render page as image
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better quality
                img_bytes = pix.tobytes("png")
                img_base64 = base64.b64encode(img_bytes).decode('utf-8')

                # Send to API with vision
                messages = [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": f"""Extract and convert this page to markdown. Follow these rules:
- Use appropriate heading levels (# ## ###)
- Convert tables to markdown table format
- Remove headers/footers like "Developer's Guide", "Integration", page numbers
- Preserve technical content, code examples, and important details
- Format lists properly
- Describe any diagrams or figures briefly"""
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{img_base64}"
                                }
                            }
                        ]
                    }
                ]

                payload = {
                    "model": "anthropic/claude-sonnet-4",
                    "messages": messages,
                    "max_tokens": 4000
                }

                try:
                    response = requests.post(
                        self.api_url,
                        headers=self.headers,
                        json=payload,
                        timeout=120
                    )
                    response.raise_for_status()
                    result = response.json()

                    if 'choices' in result and len(result['choices']) > 0:
                        page_markdown = result['choices'][0]['message']['content']
                        page_markdowns.append(page_markdown)
                        self.progress.update(
                            page_task,
                            advance=1,
                            description=f"[green]✓[/green] [cyan]Processing chunk {chunk_num} - Page {page_num + 1}/{num_pages}"
                        )
                    else:
                        self.console.print(f"   [yellow]⚠️  Page {page_num + 1} returned unexpected response[/yellow]")
                        self.progress.update(page_task, advance=1)

                    # Small delay between pages
                    time.sleep(0.5)

                except Exception as e:
                    self.console.print(f"   [yellow]⚠️  Error processing page {page_num + 1}: {e}[/yellow]")
                    self.progress.update(page_task, advance=1)
                    continue

        doc.close()

        # Combine pages
        return "\n\n".join(page_markdowns)

    def convert_from_url(
        self,
        pdf_url: str,
        output_filename: str = "DeltekOpenPlan85DeveloperGuide.md",
        download_first: bool = True,
        pages_per_chunk: int = 50,
        max_pages: Optional[int] = None
    ) -> Path:
        """
        Complete conversion pipeline from URL to markdown with rich UI.

        Args:
            pdf_url: URL of the PDF to convert
            output_filename: Name for the output markdown file
            download_first: If True, download PDF first and process in chunks
            pages_per_chunk: Pages per chunk for large PDFs
            max_pages: Maximum pages to process (for testing)

        Returns:
            Path to the output markdown file
        """
        # Print beautiful header
        self.console.print()
        self.console.rule("[bold cyan]📚 Mistral OCR PDF to Markdown Converter[/bold cyan]")
        self.console.print()

        try:
            # Download the PDF first (required for chunking)
            pdf_path = self.download_pdf(pdf_url)

            # Convert using chunked approach
            markdown_content = self.convert_pdf_chunks(
                pdf_path,
                pages_per_chunk=pages_per_chunk,
                max_pages=max_pages
            )

            # Save the result
            output_path = self.save_markdown(markdown_content, output_filename)

            # Success panel
            self.console.print()
            self.console.rule("[bold green]✨ Conversion Complete![/bold green]")
            self.console.print(f"\n[green]✓[/green] Output: [link=file://{output_path}]{output_path}[/link]")

            # Clean up chunks
            import shutil
            chunk_dir = self.output_dir.parent / "temp_chunks"
            if chunk_dir.exists():
                shutil.rmtree(chunk_dir)
                self.console.print(f"[green]✓[/green] Cleaned up temporary chunks")

            return output_path

        except Exception as e:
            self.console.print(f"\n[red]❌ Conversion failed: {e}[/red]")
            raise


def main():
    """Main entry point for the script."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert PDF to Markdown using Mistral OCR via OpenRouter"
    )
    parser.add_argument(
        "--url",
        default="https://dsm.deltek.com/DeltekSoftwareManagerWebServices/downloadFile.ashx?documentid=C6E40CBC-E0A5-4722-8E62-1E827AD56D8A",
        help="URL of the PDF to convert"
    )
    parser.add_argument(
        "--output",
        default="DeltekOpenPlan85DeveloperGuide_Mistral.md",
        help="Output markdown filename"
    )
    parser.add_argument(
        "--output-dir",
        default="docs_mistral",
        help="Output directory"
    )
    parser.add_argument(
        "--pages-per-chunk",
        type=int,
        default=10,
        help="Number of pages to process per chunk (default: 10 for large PDFs)"
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        help="Maximum number of pages to process (for testing small samples)"
    )

    args = parser.parse_args()

    try:
        # Create converter
        converter = MistralOCRConverter(output_dir=args.output_dir)

        # Run conversion
        converter.convert_from_url(
            pdf_url=args.url,
            output_filename=args.output,
            pages_per_chunk=args.pages_per_chunk,
            max_pages=args.max_pages
        )

        return 0

    except ValueError as e:
        console = Console()
        console.print(f"\n[red]❌ Configuration error: {e}[/red]")
        return 1
    except Exception as e:
        console = Console()
        console.print(f"\n[red]❌ Error: {e}[/red]")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
