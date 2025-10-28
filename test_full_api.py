#!/usr/bin/env python3
"""
Test script to process the full PDF with middle-out transform
"""

import os
import json
import requests
from dotenv import load_dotenv
from pathlib import Path

load_dotenv()

api_key = os.getenv("OPENROUTER_API_KEY")
if not api_key:
    print("ERROR: OPENROUTER_API_KEY not found")
    exit(1)

pdf_url = "https://dsm.deltek.com/DeltekSoftwareManagerWebServices/downloadFile.ashx?documentid=C6E40CBC-E0A5-4722-8E62-1E827AD56D8A"

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
    "transforms": ["middle-out"]
}

headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json",
    "HTTP-Referer": "https://github.com/devinmlowe/Deltek-OPP-Docs",
    "X-Title": "Deltek OPP Documentation Converter"
}

print("Sending API request for full PDF conversion...")
print("Using middle-out transform to handle large document\n")

response = requests.post(
    "https://openrouter.ai/api/v1/chat/completions",
    headers=headers,
    json=payload,
    stream=True,
    timeout=600
)

print(f"Response status: {response.status_code}\n")

if response.status_code != 200:
    print(f"Error: {response.text}")
    exit(1)

processing_count = 0
data_count = 0
content_chunks = []
error_occurred = False

print("Processing stream...")

for line in response.iter_lines():
    if line:
        line_text = line.decode('utf-8')

        # Count processing messages
        if line_text == ": OPENROUTER PROCESSING":
            processing_count += 1
            if processing_count % 50 == 0:
                print(f"  Processing... ({processing_count} messages)")
            continue

        if not line_text.startswith('data: '):
            continue

        data_text = line_text[6:]
        if data_text == '[DONE]':
            print("\nReceived [DONE]")
            break

        try:
            data = json.loads(data_text)

            # Check for errors
            if 'error' in data:
                error_data = data['error']
                print(f"\nERROR in response: {json.dumps(error_data, indent=2)}")
                error_occurred = True
                break

            # Check for content
            if 'choices' in data and len(data['choices']) > 0:
                delta = data['choices'][0].get('delta', {})
                if 'content' in delta:
                    content = delta['content']
                    content_chunks.append(content)
                    data_count += 1

                    if data_count % 100 == 0:
                        print(f"  Received {data_count} chunks ({len(''.join(content_chunks))} chars)...")
        except json.JSONDecodeError:
            continue

if error_occurred:
    print("Conversion failed due to error")
    exit(1)

print(f"\nProcessing complete!")
print(f"  Processing messages: {processing_count}")
print(f"  Data chunks: {data_count}")
print(f"  Content chunks: {len(content_chunks)}")

if content_chunks:
    full_content = ''.join(content_chunks)
    print(f"  Total content: {len(full_content)} characters")

    # Save to file
    output_path = Path("docs_mistral/test-full-api.md")
    output_path.parent.mkdir(exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(full_content)

    print(f"\n✓ Saved to: {output_path}")
    print(f"  File size: {output_path.stat().st_size:,} bytes")
else:
    print("\n✗ No content was generated")
