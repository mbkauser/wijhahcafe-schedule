# Wijhah Café Schedule — Files Guide

You only need TWO files for this whole system:

## 1. index.html
**Where it goes:** Upload this to your GitHub repository, exactly as-is,
named `index.html`. This is your actual website.

**One thing you must edit before it works:** near the top of the
`<script>` section, find:

    const SCHEDULE_API_URL = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';

Replace the placeholder text (keep the quote marks) with the Web App URL
you get after deploying schedule-generator.gs (step 2 below).

## 2. schedule-generator.gs
**Where it goes:** This does NOT go on GitHub. It goes inside your
Google Sheet (the one linked to your Barista Availability Form):

  1. Open the Sheet → Extensions → Apps Script
  2. Delete any existing code in the editor
  3. Paste in the entire contents of schedule-generator.gs
  4. Save
  5. Reload the Sheet — a new menu "📅 Schedule Generator" will appear
  6. Click it → Generate & Publish Schedule (this builds the schedule)
  7. Deploy → New deployment → Web app → Execute as Me →
     Who has access: Anyone → Deploy → copy the URL it gives you
  8. Paste that URL into index.html (see above) before uploading to GitHub

## Ongoing use
- Whenever you want the schedule refreshed with new availability answers,
  reopen the Sheet and click "Generate & Publish Schedule" again.
- The website (index.html) doesn't need to be touched again after the
  URL is set — it always pulls the latest published schedule automatically.
