# DateShift AI

DateShift AI is a private browser-based document editor that updates dates and recalculates financial figures in Word and PDF files, previews the result, and downloads an updated copy.

[Open the GitHub Pages site](https://mlk2025.github.io/DateShift-AI/)

## Features

- Upload `.docx` and text-based `.pdf` documents.
- Detect numeric, named, weekday-prefixed, and timestamp date formats.
- Move dates forward or backward by a chosen number of months.
- Preserve month-end dates and exclude dates such as signatures or deadlines through instructions.
- Review and edit each detected date before export.
- Detect currency amounts and financial figures while excluding dates, percentages, and likely IDs.
- Rebalance selected figures to an exact target total proportionally, equally, or by changing only the final figure.
- Calculate with integer minor units so the exported figures cannot drift from the requested total through rounding.
- Preview the exact prepared DOCX or PDF used by the download action.
- Preserve Word formatting runs, tables, colors, headers, footers, and document structure.
- Process documents locally in the browser.

## Local development

Requirements:

- Node.js 22.13 or newer
- pnpm

```bash
pnpm install
pnpm dev
```

Open the local URL printed in the terminal.

## Validation

```bash
pnpm build
node --test tests/docx-preview.test.mjs
```

The document-export, date-format, and figure-calculation regression tests run through `tsx` and are located in the `tests` directory.

## Deploy to GitHub Pages

This repository includes a GitHub Actions workflow that builds and publishes the site automatically.

1. Create a GitHub repository and upload or push these files to its `main` branch.
2. Open the repository's **Settings → Pages** page.
3. Under **Build and deployment**, select **GitHub Actions** as the source.
4. Open the **Actions** tab and run **Deploy DateShift AI to GitHub Pages**, or push another commit to `main`.

The workflow automatically handles both project repositories such as `DateShift-AI` and account repositories such as `username.github.io`. The static GitHub Pages build keeps document processing in the browser and does not publish an OpenAI API key.

## Technology

- React
- vinext and Vite
- JSZip
- Mammoth
- PDF.js and pdf-lib
- docx-preview

## Document privacy

Uploaded documents are processed in the browser. They are not sent to the application server for date or figure detection, preview generation, or export.
