# MacroFlow - API Documentation

This document provides comprehensive documentation for the MacroFlow REST API.

## Base URL

```
http://localhost:5000  # Development
```

## Authentication

Currently, MacroFlow uses OpenAI API keys configured via environment variables. No additional API authentication is required for local development.

## Core Endpoints

### 1. Generate VBA Code

**Endpoint:** `POST /generate`

Generates VBA code using the AI agent system with intelligent intent classification.

**Request Body:**
```json
{
  "prompt": "string",                    // User's natural language request
  "context": {                          // Optional context object
    "hasSelection": boolean,            // Whether user has selected code
    "selectedText": "string",           // Selected VBA code text
    "surroundingCode": "string",        // Code surrounding selection
    "currentSubroutine": "string",      // Name of current subroutine
    "activeModule": "string",           // Name of active VBA module
    "sheetContext": {                   // Excel sheet context
      "workbook_name": "string",
      "sheet_name": "string",
      "sheet_type": "string",
      "used_range": {
        "address": "string",
        "rows": number,
        "columns": number
      },
      "column_headers": ["string"],
      "data_structure": {},
      "named_ranges": [],
      "chart_objects": []
    }
  },
  "conversation_history": []            // Optional conversation history
}
```

**Response:**
```json
{
  "type": "string",                     // "vba_code" or "chat_response"
  "has_vba": boolean,                   // Whether response contains VBA code
  "vba_code": "string",                 // Generated VBA code (if applicable)
  "explanation": "string",              // AI explanation or chat response
  "classification": {                   // Intent classification metadata
    "intent": "string",                 // "vba_generation", "vba_modification", or "conversation"
    "confidence": "string",             // "high", "medium", "low"
    "classifier": "string"              // Classification method used
  }
}
```

**Example Request:**
```bash
curl -X POST http://localhost:5000/generate \
  -H "Content-Type: application/json" \
  -d '{
    "prompt": "Create a subroutine to format selected cells with bold and yellow background",
    "context": {
      "activeModule": "Module1",
      "sheetContext": {
        "workbook_name": "Budget.xlsx",
        "sheet_name": "Summary"
      }
    }
  }'
```

### 2. Inject VBA Code

**Endpoint:** `POST /inject`

Injects VBA code directly into the active Excel workbook.

**Request Body:**
```json
{
  "macro": "string"                     // VBA code to inject
}
```

**Response:**
```json
{
  "success": boolean,                   // Whether injection succeeded
  "message": "string",                  // Success/error message
  "error": "string"                     // Error details (if applicable)
}
```

### 3. List VBA Modules

**Endpoint:** `GET /list-modules`

Retrieves all VBA modules from the active Excel workbook.

**Response:**
```json
{
  "workbook": "string",                 // Workbook name
  "modules": [                          // Array of modules
    {
      "name": "string",                 // Module name
      "content": "string",              // Full VBA code content
      "line_count": number,             // Number of lines
      "subroutines": [                  // Parsed subroutines
        {
          "name": "string",             // Subroutine name
          "type": "string",             // "Sub" or "Function"
          "line": "string",             // Declaration line
          "line_number": number         // Line number in module
        }
      ],
      "isModified": boolean             // Modification status
    }
  ]
}
```

### 4. Create VBA Module

**Endpoint:** `POST /create-module`

Creates a new VBA module in the active workbook.

**Request Body:**
```json
{
  "name": "string",                     // Module name (required)
  "content": "string"                   // Initial VBA content (optional)
}
```

**Response:**
```json
{
  "success": boolean,
  "message": "string",
  "error": "string"                     // If creation failed
}
```

### 5. Update VBA Module

**Endpoint:** `POST /update-module`

Updates the content of an existing VBA module.

**Request Body:**
```json
{
  "name": "string",                     // Module name (required)
  "content": "string"                   // New VBA content (required)
}
```

**Response:**
```json
{
  "success": boolean,
  "message": "string",
  "error": "string"
}
```

### 6. Delete VBA Module

**Endpoint:** `POST /delete-module`

Deletes a VBA module from the active workbook.

**Request Body:**
```json
{
  "name": "string"                      // Module name to delete
}
```

**Response:**
```json
{
  "success": boolean,
  "message": "string",
  "error": "string"
}
```

### 7. Run VBA Subroutine

**Endpoint:** `POST /run-subroutine`

Executes a specific VBA subroutine in Excel.

**Request Body:**
```json
{
  "module": "string",                   // Module name
  "subroutine": "string"                // Subroutine name
}
```

**Response:**
```json
{
  "success": boolean,
  "message": "string",
  "error": "string"
}
```

### 8. Get Sheet Context

**Endpoint:** `GET /get-sheet-context`

Retrieves comprehensive context from the active Excel sheet.

**Response:**
```json
{
  "workbook_name": "string",
  "sheet_name": "string",
  "sheet_type": "string",
  "used_range": {
    "address": "string",
    "rows": number,
    "columns": number,
    "first_row": number,
    "first_col": number
  },
  "data_sample": [["string"]],          // Sample data rows
  "column_headers": ["string"],         // Column headers
  "data_structure": {                   // Data type analysis
    "ColumnName": {
      "type": "string",                 // "number", "text", "date_or_text"
      "sample_values": ["string"]       // Sample values
    }
  },
  "named_ranges": [                     // Named ranges in workbook
    {
      "name": "string",
      "refers_to": "string",
      "scope": "string"
    }
  ],
  "chart_objects": [                    // Chart objects on sheet
    {
      "name": "string",
      "chart_type": number,
      "position": {
        "left": number,
        "top": number,
        "width": number,
        "height": number
      }
    }
  ],
  "error": "string"                     // Error message if context retrieval failed
}
```

### 9. Get Active Module

**Endpoint:** `GET /get-active-module`

Gets the currently active VBA module for ribbon functionality.

**Response:**
```json
{
  "activeModule": {                     // Active module details
    "name": "string",
    "content": "string"
  },
  "modules": [                          // All available modules
    {
      "name": "string",
      "content": "string"
    }
  ],
  "error": "string"
}
```

## Error Handling

All endpoints return appropriate HTTP status codes:

- `200 OK` - Request successful
- `400 Bad Request` - Invalid request data
- `403 Forbidden` - VBA project access denied
- `404 Not Found` - Resource not found (e.g., module doesn't exist)
- `500 Internal Server Error` - Server error

**Error Response Format:**
```json
{
  "error": "string",                    // Error description
  "success": false,                     // Always false for errors
  "message": "string"                   // Detailed error message
}
```

## Agent System

### Intent Classification

The `/generate` endpoint uses an intelligent intent classification system:

**Intent Types:**
- `vba_generation` - User wants to generate new VBA code
- `vba_modification` - User wants to modify existing VBA code
- `conversation` - User wants conversational help or explanation

**Confidence Levels:**
- `high` - Very confident in classification
- `medium` - Moderately confident
- `low` - Low confidence, may need clarification

### Context Enhancement

The API supports rich context awareness:

**VBA Context:**
- Selected code analysis
- Surrounding code structure
- Active module information
- Subroutine detection

**Excel Context:**
- Workbook and sheet information
- Data structure analysis
- Named ranges and charts
- Sample data for AI processing

## Rate Limiting

Currently, no rate limiting is implemented for local development. Production deployments should implement appropriate rate limiting based on usage requirements.

## Development Notes

### Local Development
- Backend runs on `http://localhost:5000`
- Frontend serves from `https://localhost:3000`
- CORS is enabled for localhost origins

### Excel COM Integration
- Requires Excel to be running with an active workbook
- VBA project access must be enabled in Excel Trust Center
- COM objects are properly cleaned up to prevent memory leaks

### Caching
- Sheet context is cached for 30 seconds to improve performance
- Module information is cached until workbook changes

## Testing

### Manual Testing
Use tools like Postman or curl to test endpoints manually.

### Automated Testing
```python
# Example test with Python requests
import requests

response = requests.post('http://localhost:5000/generate', json={
    'prompt': 'Create a hello world macro',
    'context': {}
})

print(response.json())
```

### Office Add-in Testing
The frontend provides a complete testing interface through the Excel task pane.

## Security Considerations

### Development Environment
- API keys should never be committed to version control
- Use environment variables for sensitive configuration
- Excel COM integration requires appropriate trust settings

### Production Environment
- Implement HTTPS for all endpoints
- Add authentication/authorization as needed
- Validate all input data
- Implement rate limiting and monitoring
- Secure storage of API keys and sensitive data

---

For additional support or questions about the API, please refer to the main [README](../README.md) or open an issue on GitHub.