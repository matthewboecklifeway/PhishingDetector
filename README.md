# Lifeway Phishing Detector

An Outlook Mail Add-in that scans the currently open email for phishing
indicators and displays a confidence score, red flags, and an AI-generated
analysis in a task pane.

The add-in is the front end only. It posts the email's subject, sender, and
body to a separate Flask analysis API which performs the actual scoring.

## Features

- One-click "Check for Phishing" button in the Outlook ribbon (Message Read
  surface).
- Confidence score (0-100%) with Low / Medium / High risk labelling.
- Red flag list covering sender issues, suspicious keywords, and link
  analysis.
- Toggle between a local model and Claude for the analysis backend.

## Repository Contents

| File | Purpose |
| --- | --- |
| `manifest.xml` | Office Add-in manifest registered with Outlook. |
| `taskpane.html` | UI shown inside the Outlook task pane. |
| `taskpane.js` | Extracts email data via Office.js and calls the analysis API. |
| `commands.html` | Required by the manifest; no visible UI. |

## Requirements

- Outlook (desktop or web) with support for Mailbox requirement set 1.3+.
- A running instance of the Flask analysis API exposing `POST /analyze`.
- The static files in this repo served over HTTPS at the URLs referenced in
  `manifest.xml` (the manifest points at
  `https://matthewboecklifeway.github.io/PhishingDetector/`).

## Configuration

1. **API endpoint** — set `API_BASE_URL` at the top of `taskpane.js` to the
   base URL of your Flask backend.
2. **App domain** — update the `<AppDomain>` entry in `manifest.xml` to match
   the backend's domain so Outlook permits the cross-origin request.
3. **Hosting URLs** — if you host the add-in somewhere other than the
   default GitHub Pages URL, update `IconUrl`, `HighResolutionIconUrl`,
   `SupportUrl`, `SourceLocation`, `Commands.Url`, `Taskpane.Url`, and the
   icon image URLs in `manifest.xml`.

## Installing the Add-in

In Outlook, go to **Get Add-ins → My add-ins → Add a custom add-in → Add
from file** and select `manifest.xml`. After installation, open any
message and click **Check for Phishing** in the ribbon to open the task
pane.

## Backend Contract

`taskpane.js` sends a `multipart/form-data` POST to `${API_BASE_URL}/analyze`
with two fields:

- `use_claude` — `true` or `false` depending on the model toggle.
- `pasted_content` — `From: …\nSubject: …\n\n<body>`.

The expected JSON response shape:

```json
{
  "confidence_score": 0,
  "analysis": "string",
  "sender_analysis": ["string"],
  "suspicious_keywords": ["string"],
  "link_analysis": ["string"],
  "links": [{ "url": "string" }],
  "email_details": { "subject": "string", "sender": "string" }
}
```
