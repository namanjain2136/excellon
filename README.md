# Excellon

Convert numeric columns inside Excel files into Indian English words — entirely in your browser. No uploads, no backend, no tracking.

## Features

- 100% local processing (files never leave your device)
- Indian (Lakh / Crore) and International (Million / Billion) numbering systems
- Title Case, UPPERCASE, lowercase output
- Optional currency suffix (Rupees / Paise)
- Auto-detects numeric columns
- Inserts a new "In Words" column right next to the source column
- Preserves formulas, merges, and styles
- Works offline after first load

## Tech Stack

- React 19 + TanStack Start (Vite 7)
- Tailwind CSS v4
- SheetJS (`xlsx`) for workbook parsing
- `file-saver` for downloads
- Bun as the package manager / runtime

## Prerequisites

- [Bun](https://bun.sh) v1.1+ (recommended)
- Or Node.js 20+ with npm/pnpm/yarn

## Getting Started

### 1. Clone the repository

```bash
git clone <your-repo-url>
cd excellon
```

### 2. Install dependencies

```bash
bun install
```

Or with npm:

```bash
npm install
```

### 3. Run the dev server

```bash
bun run dev
```

The app will be available at **http://localhost:8080** (or the port shown in the terminal).

### 4. Build for production

```bash
bun run build
```

The production build is output to `.output/` (or `dist/` depending on the target).

### 5. Preview the production build

```bash
bun run start
```

## How to Use

1. Open the app in your browser.
2. Drag & drop an `.xlsx` / `.xls` file onto the upload area, or click to browse.
3. Select the sheet you want to process.
4. Pick the numeric column(s) to convert.
5. Choose your options:
   - Numbering system (Indian / International)
   - Letter case (Title / UPPER / lower)
   - Optional currency suffix
6. Click **Convert & Download** — the new file downloads instantly with an "In Words" column inserted beside each selected column.

## Project Structure

```
src/
├── routes/
│   ├── __root.tsx        # Root layout + SEO
│   └── index.tsx         # Main converter UI
├── components/
│   ├── excellon/         # App-specific components
│   └── ui/               # shadcn/ui components
├── lib/
│   ├── numberToWords.ts  # Number → words conversion logic
│   └── workbook.ts       # Excel parsing & column insertion
└── styles.css            # Theme tokens & Tailwind config
```

## Privacy

Excellon runs **entirely in your browser**. There is no server, no API, no analytics on your file contents. Your spreadsheets never leave your device.

## License

MIT
