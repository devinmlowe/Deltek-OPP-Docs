#!/usr/bin/env python3
"""
Simple test script to debug the OpenRouter API response
"""

import os
import json
import requests
from dotenv import load_dotenv

load_dotenv()

api_key = os.getenv("OPENROUTER_API_KEY")
if not api_key:
    print("ERROR: OPENROUTER_API_KEY not found")
    exit(1)

pdf_url = "https://dsm.deltek.com/DeltekSoftwareManagerWebServices/downloadFile.ashx?documentid=C6E40CBC-E0A5-4722-8E62-1E827AD56D8A"

messages = [{
    "role": "user",
    "content": [
        {
            "type": "text",
            "text": "Extract the first page of this PDF and convert it to markdown. Just give me the markdown content for the first page only."
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
    "max_tokens": 10000,
    "stream": True,
    "transforms": ["middle-out"]  # Compress long prompts automatically
}

headers = {
    "Authorization": f"Bearer {api_key}",
    "Content-Type": "application/json",
    "HTTP-Referer": "https://github.com/devinmlowe/Deltek-OPP-Docs",
    "X-Title": "Deltek OPP Documentation Converter"
}

print("Sending API request...")
print(f"Payload: {json.dumps(payload, indent=2)}")
print("\n" + "="*80 + "\n")

response = requests.post(
    "https://openrouter.ai/api/v1/chat/completions",
    headers=headers,
    json=payload,
    stream=True,
    timeout=120
)

print(f"Response status: {response.status_code}")
print(f"Response headers: {dict(response.headers)}")
print("\n" + "="*80 + "\n")

line_count = 0
processing_count = 0
data_count = 0
content_chunks = []

for line in response.iter_lines():
    if line:
        line_count += 1
        line_text = line.decode('utf-8')

        # Count processing messages
        if line_text == ": OPENROUTER PROCESSING":
            processing_count += 1
            if processing_count % 10 == 0:
                print(f"Processing... ({processing_count} messages)")
            continue

        # Print non-data lines
        if not line_text.startswith('data: '):
            print(f"Line {line_count}: {line_text[:200]}")
            continue

        # Process data lines
        data_text = line_text[6:]
        if data_text == '[DONE]':
            print("\nReceived [DONE]")
            break

        try:
            data = json.loads(data_text)
            data_count += 1

            # Check for content
            if 'choices' in data and len(data['choices']) > 0:
                delta = data['choices'][0].get('delta', {})
                if 'content' in delta:
                    content = delta['content']
                    content_chunks.append(content)
                    print(f"Data {data_count}: Got {len(content)} chars of content")

            # Check for errors
            if 'error' in data:
                print(f"ERROR in response: {data['error']}")

        except json.JSONDecodeError as e:
            print(f"JSON decode error: {e}")
            print(f"  Data was: {data_text[:200]}")

print(f"\nTotal lines: {line_count}")
print(f"Processing messages: {processing_count}")
print(f"Data messages: {data_count}")
print(f"Content chunks: {len(content_chunks)}")

if content_chunks:
    full_content = ''.join(content_chunks)
    print(f"\nTotal content length: {len(full_content)} chars")
    print(f"\nFirst 500 chars of content:\n{full_content[:500]}")
