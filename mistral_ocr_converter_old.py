#!/usr/bin/env python3
"""
PDF to Markdown Converter using Mistral OCR via OpenRouter
Processes PDF documents using Mistral's OCR capabilities via OpenRouter's file parser plugin
"""

import os
import sys
import json
import requests
import argparse
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
    TimeElapsedColumn
)
from rich.panel import Panel
from rich.table import Table


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
            "HTTP-Referer": "https://github.com/devinmlowe/Deltek-OPP-Docs",
            "X-Title": "Deltek OPP Documentation Converter"
        }

        # Initialize rich console for beautiful output
        self.console = Console()

    def convert_pdf_to_markdown(
        self,
        pdf_url: str,
        output_filename: str = "DeltekOpenPlanDeveloperGuide.md",
        max_pages: Optional[int] = None
    ) -> Path:
        """
        Convert a PDF document to Markdown using Mistral OCR via OpenRouter.

        Args:
            pdf_url: URL to the PDF document
            output_filename: Name of the output markdown file
            max_pages: Optional limit on number of pages to process (for testing)

        Returns:
            Path to the generated markdown file
        """
        output_path = self.output_dir / output_filename

        # Display conversion info
        self.console.print("\n" + "="*80)
        self.console.print(Panel.fit(
            "[bold cyan]Mistral OCR PDF to Markdown Converter[/bold cyan]\n"
            f"[white]PDF URL:[/white] {pdf_url}\n"
            f"[white]Output:[/white] {output_path}",
            border_style="cyan"
        ))
        self.console.print("="*80 + "\n")

        # Prepare the API request
        prompt = """Convert this PDF document to clean, well-formatted GitHub Flavored Markdown.

Requirements:
- Remove all headers, footers, and page numbers
- Preserve all headings using proper markdown heading levels (# ## ### etc.)
- Maintain code examples in proper code blocks with language identifiers
- Convert tables to markdown table format
- Preserve the document structure and hierarchy
- Keep all technical content and details
- Use proper markdown formatting for lists, emphasis, and links
- Ensure the output is clean and readable

Please convert the entire document to markdown format."""

        messages = [{
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
                        "file_data": pdf_url
                    }
                }
            ]
        }]

        plugins = [{
            "id": "file-parser",
            "pdf": {"engine": "mistral-ocr"}
        }]

        payload = {
            "model": "anthropic/claude-sonnet-4",
            "messages": messages,
            "plugins": plugins,
            "max_tokens": 100000,
            "stream": True,
            "transforms": ["middle-out"]  # Compress long prompts automatically
        }

        self.console.print("[bold yellow]🚀 Sending PDF to OpenRouter for processing...[/bold yellow]")
        self.console.print("[dim]Using Mistral OCR engine with Claude Sonnet 4 for markdown generation[/dim]")
        self.console.print("[dim]Applying middle-out compression for large documents[/dim]\n")

        try:
            # Make the API request with streaming
            response = requests.post(
                self.api_url,
                headers=self.headers,
                json=payload,
                stream=True,
                timeout=600  # 10 minute timeout for large documents
            )
            response.raise_for_status()

            # Process the streaming response
            markdown_content = []
            chunk_count = 0
            error_message = None

            with Progress(
                SpinnerColumn(),
                TextColumn("[bold blue]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeElapsedColumn(),
                console=self.console,
                transient=False
            ) as progress:
                task = progress.add_task(
                    "[cyan]Receiving markdown content...",
                    total=None
                )

                for line in response.iter_lines():
                    if line:
                        line_text = line.decode('utf-8')
                        if line_text.startswith('data: '):
                            data_text = line_text[6:]  # Remove 'data: ' prefix

                            if data_text == '[DONE]':
                                break

                            try:
                                data = json.loads(data_text)

                                # Check for errors in the response
                                if 'error' in data:
                                    error_data = data['error']
                                    if isinstance(error_data, dict):
                                        error_message = error_data.get('message', str(error_data))
                                        error_code = error_data.get('code', 'Unknown')
                                        error_metadata = error_data.get('metadata', {})
                                        self.console.print(f"\n[bold red]API Error ({error_code}):[/bold red] {error_message}")
                                        if error_metadata:
                                            self.console.print(f"  [dim]Metadata: {error_metadata}[/dim]")
                                    else:
                                        error_message = str(error_data)
                                        self.console.print(f"\n[bold red]API Error:[/bold red] {error_message}")
                                    break

                                if 'choices' in data and len(data['choices']) > 0:
                                    delta = data['choices'][0].get('delta', {})
                                    if 'content' in delta:
                                        content = delta['content']
                                        markdown_content.append(content)
                                        chunk_count += 1

                                        # Update progress every 10 chunks
                                        if chunk_count % 10 == 0:
                                            progress.update(
                                                task,
                                                description=f"[cyan]Received {chunk_count} chunks..."
                                            )
                            except json.JSONDecodeError as e:
                                # Log the problematic line for debugging
                                if len(data_text) < 200:
                                    self.console.print(f"[dim]Skipping invalid JSON: {data_text[:100]}[/dim]")
                                continue

                progress.update(task, description=f"[green]✓ Received {chunk_count} chunks")

            # If there was an error, raise it
            if error_message:
                raise ValueError(f"API returned an error: {error_message}")

            # Combine all content
            full_markdown = ''.join(markdown_content)

            if not full_markdown.strip():
                raise ValueError("No markdown content was generated from the PDF")

            # Save to file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(full_markdown)

            # Display success summary
            file_size = output_path.stat().st_size
            self.console.print(f"\n[bold green]✓ Conversion completed successfully![/bold green]")

            summary = Table(show_header=False, box=None, padding=(0, 2))
            summary.add_row("[cyan]Output file:", f"[white]{output_path}[/white]")
            summary.add_row("[cyan]File size:", f"[white]{file_size:,} bytes[/white]")
            summary.add_row("[cyan]Chunks received:", f"[white]{chunk_count}[/white]")

            self.console.print(summary)
            self.console.print()

            return output_path

        except requests.exceptions.RequestException as e:
            self.console.print(f"\n[bold red]✗ API request failed:[/bold red] {str(e)}")
            raise
        except Exception as e:
            self.console.print(f"\n[bold red]✗ Conversion failed:[/bold red] {str(e)}")
            raise


def main():
    """Main entry point for the converter."""
    parser = argparse.ArgumentParser(
        description='Convert PDF documents to Markdown using Mistral OCR',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Test with first 10 pages
  python3 mistral_ocr_converter.py --max-pages 10

  # Convert full document
  python3 mistral_ocr_converter.py

  # Custom output file and directory
  python3 mistral_ocr_converter.py --output MyDoc.md --output-dir custom_output
        """
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
        help='Output directory for markdown files'
    )
    parser.add_argument(
        '--max-pages',
        type=int,
        help='Maximum number of pages to process (for testing)'
    )

    args = parser.parse_args()

    try:
        converter = MistralOCRConverter(output_dir=args.output_dir)

        output_path = converter.convert_pdf_to_markdown(
            pdf_url=args.pdf_url,
            output_filename=args.output,
            max_pages=args.max_pages
        )

        return 0

    except KeyboardInterrupt:
        console = Console()
        console.print("\n[yellow]Conversion cancelled by user[/yellow]")
        return 1
    except Exception as e:
        console = Console()
        console.print(f"\n[bold red]Error:[/bold red] {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
