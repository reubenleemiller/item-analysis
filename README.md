# Item Analysis Builder

A Vercel-ready Next.js app that creates a Google Sheets-compatible item-analysis workbook from a roster CSV and an exam answer key

## Run locally

```bash
npm install
npm run clean
npm run dev
```

Open `http://localhost:3000`.

The development command uses Turbopack. `npm run clean` removes only Next.js's generated cache and is useful after switching between `next build` and `next dev`.

## Deploy to Vercel

Import this repository in Vercel and deploy. No server-side secrets are needed for the workbook download.

Set `NEXT_PUBLIC_GOOGLE_CLIENT_ID` to an OAuth web client ID configured with the production Vercel URL (and `http://localhost:3000` for development) as authorized JavaScript origins. The **Create Google Sheet** action signs the user in and requests the narrowly scoped `drive.file` permission only after they select it; the class list is then used to create a native Google Sheet in that user’s Drive.

## CSV format

The importer recognizes common headers such as `student number`, `student id`, `id`, `student name`, `name`, `first name`, and `last name`. If it cannot recognize headers, the first two columns are treated as student number and student name.
