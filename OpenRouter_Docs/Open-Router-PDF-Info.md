---
title: PDF Inputs
subtitle: How to send PDFs to OpenRouter models
headline: OpenRouter PDF Inputs | Complete Documentation
canonical-url: 'https://openrouter.ai/docs/features/multimodal/pdfs'
'og:site_name': OpenRouter Documentation
'og:title': OpenRouter PDF Inputs - Complete Documentation
'og:description': Send PDF documents to any model on OpenRouter.
'og:image':
  type: url
  value: >-
    https://openrouter.ai/dynamic-og?title=OpenRouter%20PDF%20Inputs&description=Send%20PDF%20documents%20to%20any%20model%20on%20OpenRouter.
'og:image:width': 1200
'og:image:height': 630
'twitter:card': summary_large_image
'twitter:site': '@OpenRouterAI'
noindex: false
nofollow: false
---

import {
  API_KEY_REF,
  DEFAULT_PDF_ENGINE,
  MISTRAL_OCR_USER_COST_PER_1K_PAGE as MISTRAL_OCR_COST,
  PDFParserEngine,
} from '../../../../imports/constants';

OpenRouter supports PDF processing through the `/api/v1/chat/completions` API. PDFs can be sent as **direct URLs** or **base64-encoded data URLs** in the messages array, via the file content type. This feature works on **any** model on OpenRouter.

**URL support**: Send publicly accessible PDFs directly without downloading or encoding  
**Base64 support**: Required for local files or private documents that aren't publicly accessible

PDFs also work in the chat room for interactive testing.

<Info>
  When a model supports file input natively, the PDF is passed directly to the
  model. When the model does not support file input natively, OpenRouter will
  parse the file and pass the parsed results to the requested model.
</Info>

<Tip>You can send both PDFs and other file types in the same request.</Tip>

## Plugin Configuration

To configure PDF processing, use the `plugins` parameter in your request. OpenRouter provides several PDF processing engines with different capabilities and pricing:

```typescript
{
  plugins: [
    {
      id: 'file-parser',
      pdf: {
        engine: 'pdf-text', // or 'mistral-ocr' or 'native'
      },
    },
  ],
}
```

## Pricing

OpenRouter provides several PDF processing engines:

1. <code>"{PDFParserEngine.MistralOCR}"</code>: Best for scanned documents or
   PDFs with images (${MISTRAL_OCR_COST.toString()} per 1,000 pages).
2. <code>"{PDFParserEngine.PDFText}"</code>: Best for well-structured PDFs with
   clear text content (Free).
3. <code>"{PDFParserEngine.Native}"</code>: Only available for models that
   support file input natively (charged as input tokens).

If you don't explicitly specify an engine, OpenRouter will default first to the model's native file processing capabilities, and if that's not available, we will use the <code>"{DEFAULT_PDF_ENGINE}"</code> engine.

## Using PDF URLs

For publicly accessible PDFs, you can send the URL directly without needing to download and encode the file:

<Template data={{
  API_KEY_REF,
  MODEL: 'anthropic/claude-sonnet-4',
  ENGINE: PDFParserEngine.MistralOCR,
}}>
<CodeGroup>

```python
import requests
import json

url = "https://openrouter.ai/api/v1/chat/completions"
headers = {
    "Authorization": f"Bearer {API_KEY_REF}",
    "Content-Type": "application/json"
}

messages = [
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "What are the main points in this document?"
            },
            {
                "type": "file",
                "file": {
                    "filename": "document.pdf",
                    "file_data": "https://bitcoin.org/bitcoin.pdf"
                }
            },
        ]
    }
]

# Optional: Configure PDF processing engine
plugins = [
    {
        "id": "file-parser",
        "pdf": {
            "engine": "{{ENGINE}}"
        }
    }
]

payload = {
    "model": "{{MODEL}}",
    "messages": messages,
    "plugins": plugins
}

response = requests.post(url, headers=headers, json=payload)
print(response.json())
```

```typescript
const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
  method: 'POST',
  headers: {
    Authorization: `Bearer ${API_KEY_REF}`,
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    model: '{{MODEL}}',
    messages: [
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: 'What are the main points in this document?',
          },
          {
            type: 'file',
            file: {
              filename: 'document.pdf',
              file_data: 'https://bitcoin.org/bitcoin.pdf',
            },
          },
        ],
      },
    ],
    // Optional: Configure PDF processing engine
    plugins: [
      {
        id: 'file-parser',
        pdf: {
          engine: '{{ENGINE}}',
        },
      },
    ],
  }),
});

const data = await response.json();
console.log(data);
```

</CodeGroup>
</Template>

<Info>
  PDF URLs work with all processing engines. For Mistral OCR, the URL is passed directly to the service. For other engines, OpenRouter fetches the PDF and processes it internally.
</Info>

## Using Base64 Encoded PDFs

For local PDF files or when you need to send PDF content directly, you can base64 encode the file:

<Template data={{
  API_KEY_REF,
  MODEL: 'google/gemma-3-27b-it',
  ENGINE: PDFParserEngine.PDFText,
  DEFAULT_PDF_ENGINE,
}}>
<CodeGroup>

```python
import requests
import json
import base64
from pathlib import Path

def encode_pdf_to_base64(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        return base64.b64encode(pdf_file.read()).decode('utf-8')

url = "https://openrouter.ai/api/v1/chat/completions"
headers = {
    "Authorization": f"Bearer {API_KEY_REF}",
    "Content-Type": "application/json"
}

# Read and encode the PDF
pdf_path = "path/to/your/document.pdf"
base64_pdf = encode_pdf_to_base64(pdf_path)
data_url = f"data:application/pdf;base64,{base64_pdf}"

messages = [
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "What are the main points in this document?"
            },
            {
                "type": "file",
                "file": {
                    "filename": "document.pdf",
                    "file_data": data_url
                }
            },
        ]
    }
]

# Optional: Configure PDF processing engine
# PDF parsing will still work even if the plugin is not explicitly set
plugins = [
    {
        "id": "file-parser",
        "pdf": {
            "engine": "{{ENGINE}}"  # defaults to "{{DEFAULT_PDF_ENGINE}}". See Pricing above
        }
    }
]

payload = {
    "model": "{{MODEL}}",
    "messages": messages,
    "plugins": plugins
}

response = requests.post(url, headers=headers, json=payload)
print(response.json())
```

```typescript
async function encodePDFToBase64(pdfPath: string): Promise<string> {
  const pdfBuffer = await fs.promises.readFile(pdfPath);
  const base64PDF = pdfBuffer.toString('base64');
  return `data:application/pdf;base64,${base64PDF}`;
}

// Read and encode the PDF
const pdfPath = 'path/to/your/document.pdf';
const base64PDF = await encodePDFToBase64(pdfPath);

const response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
  method: 'POST',
  headers: {
    Authorization: `Bearer ${API_KEY_REF}`,
    'Content-Type': 'application/json',
  },
  body: JSON.stringify({
    model: '{{MODEL}}',
    messages: [
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: 'What are the main points in this document?',
          },
          {
            type: 'file',
            file: {
              filename: 'document.pdf',
              file_data: base64PDF,
            },
          },
        ],
      },
    ],
    // Optional: Configure PDF processing engine
    // PDF parsing will still work even if the plugin is not explicitly set
    plugins: [
      {
        id: 'file-parser',
        pdf: {
          engine: '{{ENGINE}}', // defaults to "{{DEFAULT_PDF_ENGINE}}". See Pricing above
        },
      },
    ],
  }),
});

const data = await response.json();
console.log(data);
```

</CodeGroup>
</Template>

## Skip Parsing Costs

When you send a PDF to the API, the response may include file annotations in the assistant's message. These annotations contain structured information about the PDF document that was parsed. By sending these annotations back in subsequent requests, you can avoid re-parsing the same PDF document multiple times, which saves both processing time and costs.

Here's how to reuse file annotations:

<Template data={{
  API_KEY_REF,
  MODEL: 'google/gemma-3-27b-it'
}}>
<CodeGroup>

```python
import requests
import json
import base64
from pathlib import Path

# First, encode and send the PDF
def encode_pdf_to_base64(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        return base64.b64encode(pdf_file.read()).decode('utf-8')

url = "https://openrouter.ai/api/v1/chat/completions"
headers = {
    "Authorization": f"Bearer {API_KEY_REF}",
    "Content-Type": "application/json"
}

# Read and encode the PDF
pdf_path = "path/to/your/document.pdf"
base64_pdf = encode_pdf_to_base64(pdf_path)
data_url = f"data:application/pdf;base64,{base64_pdf}"

# Initial request with the PDF
messages = [
    {
        "role": "user",
        "content": [
            {
                "type": "text",
                "text": "What are the main points in this document?"
            },
            {
                "type": "file",
                "file": {
                    "filename": "document.pdf",
                    "file_data": data_url
                }
            },
        ]
    }
]

payload = {
    "model": "{{MODEL}}",
    "messages": messages
}

response = requests.post(url, headers=headers, json=payload)
response_data = response.json()

# Store the annotations from the response
file_annotations = None
if response_data.get("choices") and len(response_data["choices"]) > 0:
    if "annotations" in response_data["choices"][0]["message"]:
        file_annotations = response_data["choices"][0]["message"]["annotations"]

# Follow-up request using the annotations (without sending the PDF again)
if file_annotations:
    follow_up_messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": "What are the main points in this document?"
                },
                {
                    "type": "file",
                    "file": {
                        "filename": "document.pdf",
                        "file_data": data_url
                    }
                }
            ]
        },
        {
            "role": "assistant",
            "content": "The document contains information about...",
            "annotations": file_annotations
        },
        {
            "role": "user",
            "content": "Can you elaborate on the second point?"
        }
    ]

    follow_up_payload = {
        "model": "{{MODEL}}",
        "messages": follow_up_messages
    }

    follow_up_response = requests.post(url, headers=headers, json=follow_up_payload)
    print(follow_up_response.json())
```

```typescript
import fs from 'fs/promises';

async function encodePDFToBase64(pdfPath: string): Promise<string> {
  const pdfBuffer = await fs.readFile(pdfPath);
  const base64PDF = pdfBuffer.toString('base64');
  return `data:application/pdf;base64,${base64PDF}`;
}

// Initial request with the PDF
async function processDocument() {
  // Read and encode the PDF
  const pdfPath = 'path/to/your/document.pdf';
  const base64PDF = await encodePDFToBase64(pdfPath);

  const initialResponse = await fetch(
    'https://openrouter.ai/api/v1/chat/completions',
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${API_KEY_REF}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: '{{MODEL}}',
        messages: [
          {
            role: 'user',
            content: [
              {
                type: 'text',
                text: 'What are the main points in this document?',
              },
              {
                type: 'file',
                file: {
                  filename: 'document.pdf',
                  file_data: base64PDF,
                },
              },
            ],
          },
        ],
      }),
    },
  );

  const initialData = await initialResponse.json();

  // Store the annotations from the response
  let fileAnnotations = null;
  if (initialData.choices && initialData.choices.length > 0) {
    if (initialData.choices[0].message.annotations) {
      fileAnnotations = initialData.choices[0].message.annotations;
    }
  }

  // Follow-up request using the annotations (without sending the PDF again)
  if (fileAnnotations) {
    const followUpResponse = await fetch(
      'https://openrouter.ai/api/v1/chat/completions',
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${API_KEY_REF}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          model: '{{MODEL}}',
          messages: [
            {
              role: 'user',
              content: [
                {
                  type: 'text',
                  text: 'What are the main points in this document?',
                },
                {
                  type: 'file',
                  file: {
                    filename: 'document.pdf',
                    file_data: base64PDF,
                  },
                },
              ],
            },
            {
              role: 'assistant',
              content: 'The document contains information about...',
              annotations: fileAnnotations,
            },
            {
              role: 'user',
              content: 'Can you elaborate on the second point?',
            },
          ],
        }),
      },
    );

    const followUpData = await followUpResponse.json();
    console.log(followUpData);
  }
}

processDocument();
```

</CodeGroup>
</Template>

<Info>
  When you include the file annotations from a previous response in your
  subsequent requests, OpenRouter will use this pre-parsed information instead
  of re-parsing the PDF, which saves processing time and costs. This is
  especially beneficial for large documents or when using the `mistral-ocr`
  engine which incurs additional costs.
</Info>

## Response Format

The API will return a response in the following format:

```json
{
  "id": "gen-1234567890",
  "provider": "DeepInfra",
  "model": "google/gemma-3-27b-it",
  "object": "chat.completion",
  "created": 1234567890,
  "choices": [
    {
      "message": {
        "role": "assistant",
        "content": "The document discusses..."
      }
    }
  ],
  "usage": {
    "prompt_tokens": 1000,
    "completion_tokens": 100,
    "total_tokens": 1100
  }
}
```
