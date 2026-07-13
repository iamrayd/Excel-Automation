# Sheetshift

Upload a source ("FROM") and target ("TO") Excel file, describe a transformation in plain English,
and get back a working VBA macro — injected directly into the source workbook.

## Project status: Phase 2 (AI generation wired up)

This phase includes everything from Phase 1, plus:
- Parsing FROM/TO files: sheet names, header row, and up to 3 sample data rows per sheet (via `exceljs`)
- A real call to OpenRouter (`deepseek/deepseek-chat-v3-0324:free` by default) that generates a VBA macro based on your prompt + both files' structure
- The generated code is displayed in a read-only panel in the UI, with a copy button
- A mock mode (`OPENROUTER_MOCK=true` in `.env`) to test the whole pipeline without spending API credits

Not yet built (later phases):
- Phase 3: Python + pywin32 script that injects the generated VBA into the FROM file via Excel COM
- Phase 4: Making the code panel editable, so you can tweak the macro before it's injected
- Phase 5: Overwrite vs. save-as choice, polish, error handling

## Requirements

- Node.js 18+
- Python 3.9+ with `pywin32` installed (`pip install pywin32`) — only needed starting Phase 3
- Microsoft Excel installed (Windows) — only needed starting Phase 3
- An OpenRouter API key — get one at https://openrouter.ai/keys

## Running it

**Server:**
```bash
cd server
npm install
copy .env.example .env
```
Then open `.env` and paste in your `OPENROUTER_API_KEY`. Leave `OPENROUTER_MOCK=false` to use the real API, or set it to `true` to test without spending credits.
```bash
npm run dev
```

**Client:**
```bash
cd client
npm install
npm run dev
```

Client runs on http://localhost:5173, server on http://localhost:3001. Upload a FROM and TO `.xlsx` file, type a request, and click Generate — the VBA code will appear below the card.

Note: legacy `.xls` files can be uploaded but can't be parsed for structure yet (`exceljs` only reads the modern `.xlsx`/`.xlsm` format) — you'll get a clear error asking you to re-save as `.xlsx`.
