````md id="gx3q4s"
# Excellon

Convert numeric columns inside Excel files into Indian English words — entirely in your browser.

No uploads. No backend. No tracking.

---

## Features

- 100% local processing — files never leave your device
- Indian (Lakh / Crore) and International (Million / Billion) numbering systems
- Title Case, UPPERCASE, and lowercase output formats
- Optional currency suffix support (Rupees / Paise)
- Auto-detects numeric columns intelligently
- Inserts a new “In Words” column beside the source column
- Preserves workbook structure, formulas, merges, and styles
- Works offline after first load
- Fast browser-based processing
- Responsive modern UI

---

## Tech Stack

- React 19
- TanStack Start (Vite 7)
- Tailwind CSS v4
- SheetJS (xlsx)
- file-saver
- npm

---

## Prerequisites

- Node.js 20+
- npm 10+

Download Node.js:
https://nodejs.org/

---

## Getting Started

### 1. Clone the Repository

```bash
git clone <your-repo-url>
cd excellon
````

---

### 2. Install Dependencies

```bash
npm install
```

---

### 3. Start Development Server

```bash
npm run dev
```

The application will run locally at:

```text
http://localhost:8080
```

(or the port shown in your terminal)

---

### 4. Build for Production

```bash
npm run build
```

---

### 5. Preview Production Build

```bash
npm run start
```

---

## How to Use

1. Open Excellon in your browser
2. Upload an `.xlsx` or `.xls` file
3. Select the worksheet
4. Choose numeric column(s) to convert
5. Configure conversion options:

   * Numbering System
   * Text Case
   * Currency Suffix
6. Click:

   ```text
   Convert & Download
   ```
7. Download the processed workbook instantly

New “In Words” columns are inserted automatically beside the selected columns.

---

## Project Structure

```text
src/
├── routes/
│   ├── __root.tsx
│   └── index.tsx
│
├── components/
│   ├── excellon/
│   └── ui/
│
├── lib/
│   ├── numberToWords.ts
│   └── workbook.ts
│
└── styles.css
```

---

## Privacy

Excellon processes files entirely inside your browser.

* No cloud uploads
* No server-side processing
* No analytics on spreadsheet contents
* No external APIs

Your Excel files never leave your device.

---

## Security Note

This project uses the `xlsx` library (SheetJS) for Excel parsing and workbook manipulation.

Some upstream security advisories may appear in `npm audit`, but the application processes files entirely locally in the browser and does not expose any backend or server-side upload endpoints.

---

## License

MIT License

```
```
