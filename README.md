# Sheetshift

Sheetshift is a fast, secure, and minimalist Excel column-mapping and data copy-paste utility. 

Upload a source (`FROM`) Excel sheet and a target (`TO`) Excel sheet, customize the automatically matched column mappings, and execute the data transfer. Sheetshift automatically handles row appending, in-flight deduplication, and performs an automated post-transfer AI quality audit.

## Key Features

- **Header Auto-Matching:** Automatically matches columns case-insensitively upon file upload.
- **Interactive Mapping Editor:** Edit column associations, add new columns, or skip unwanted fields in a clean, responsive layout.
- **Duplicate Prevention:** Identifies the primary ID column (e.g., Customer ID) and filters out duplicate rows dynamically.
- **AI Quality Auditor:** Generates a post-transfer audit report via OpenRouter to verify data copy alignment (with warning fallback if API key is not configured).
- **Vercel Cloud Ready:** Out-of-the-box serverless monorepo configuration with dynamic `os.tmpdir()` filesystem allocations.

## Requirements

- Node.js 18+
- An OpenRouter API key (optional; get one at [openrouter.ai](https://openrouter.ai/keys))

## Local Development

### 1. Backend Server Setup

```bash
cd server
npm install
copy .env.example .env
```
Open `.env` and configure your `OPENROUTER_API_KEY` and target model.
Start the dev server (listening on port `3001`):
```bash
npm run dev
```

### 2. Frontend Client Setup

```bash
cd client
npm install
npm run dev
```
Open the dev server at the local URL (usually `http://localhost:5174` or `http://localhost:5173`).

---

## Cloud Deployment (Vercel)

This project is configured as a monorepo ready for instant deploy via Vercel:

1. Import the repository directly on your **[Vercel Dashboard](https://vercel.com/dashboard)**.
2. Vercel automatically reads the root `vercel.json` to handle React static builds and compile the Node backend server as a Serverless Function.
3. Add your environment variables (`OPENROUTER_API_KEY`, etc.) in the Vercel project configuration.
4. Click **Deploy**.
