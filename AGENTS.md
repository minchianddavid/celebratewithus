# Wedding Website — Project Rules

## Change and Push Workflow

1. Make code changes step by step so each change can be reviewed independently.
2. Show the user what changed and wait for explicit confirmation before continuing to the next meaningful change.
3. Never push commits or changes to a remote repository without the user's explicit confirmation immediately before the push.
4. Treat the user's standalone command `cp` as explicit confirmation to commit all currently reviewed changes and push the resulting commit to the configured remote branch.

## Git and PR Rules

Do not include Codex branding or attribution in commits or pull requests.

## Google Apps Script Safety

The Apps Script (`docs/google-apps-script.js`) must only interact with the RSVP spreadsheet it is attached to:

1. Only use `SpreadsheetApp.getActiveSpreadsheet()`; never open another spreadsheet by ID or URL.
2. Only use `appendRow()`; never delete, clear, or overwrite existing rows.
3. Never access other Google services such as Drive, Gmail, or Calendar.
4. Never install automatic triggers.
5. Make no external HTTP calls except to Google's reCAPTCHA verification endpoint.

## Canonical Typography System

Typography is determined by semantic role, never by visual preference. Do not introduce another font role without updating this specification.

### 1. Brand Font — Display Serif

The Brand Font is the Display Serif and represents the couple's identity. Use it for all bride and groom names, including:

- `Minchi & David` in the Hero and invitation cover
- `Minchi Tsai` and `David Kuo` on Couple or profile sections
- An optional couple-name footer signature

Never use the Display Serif for section titles, dates, times, venues, Story body content, RSVP content, buttons, forms, or navigation.

### 2. Section Script — Emotional Headings

Use the decorative Section Script only for these section or profile headings:

- `Wedding Day`
- `RSVP`
- `Bride`
- `Groom`
- `Our Story`
- `A Few Things to Know`

Within `A Few Things to Know`, only the English section title uses the Section Script. FAQ questions, answers, links, and map text remain Sans-serif.

Do not use the Section Script for Hero announcements, names, information blocks, descriptions, modal copy, buttons, or navigation.

### 3. Sans-serif — Information and UI

Use the primary sans-serif for everything that is not a couple name or one of the approved Section Script headings, including:

- Hero announcements
- Dates, times, countdowns, schedules, and venues
- Maps links, buttons, and navigation
- RSVP fields, options, descriptions, and helper text
- Story body content
- Wedding Guide content, except the `A Few Things to Know` English section title
- Footer content

Sans-serif should account for more than 90% of the website typography.

### 4. Chinese Sans-serif

All Chinese text must use this dedicated Chinese sans-serif stack:

```css
--font-sans-zh: 'Noto Sans TC', 'PingFang TC', 'Microsoft JhengHei', 'Heiti TC', sans-serif;
```

Load `Noto Sans TC` explicitly and use `var(--font-sans-zh)` for Chinese text. Never place Chinese inside a Section Script or Brand Font element.

Do not apply a global `60%` size ratio or universal `0.1rem` letter-spacing to Chinese. Text with the same semantic role must consume the same shared Chinese typography token, such as:

- Section subtitle
- Field label
- Option text
- Venue information
- Body text
- Helper text
- Button text

Section-heading Chinese must remain on its own line below the English heading. Compact UI components are the exception: field labels, options, helper text, and placeholders may keep English and Chinese together when the component specification calls for a compact bilingual label. Do not use slash separators between languages.

### Required Hierarchy

- Cover: Display Serif couple name → date in sans-serif → button in sans-serif
- Hero: Sans-serif announcement → Chinese subtitle → Display Serif couple name → Chinese names → date → venue
- Wedding Day: Section Script → Chinese subtitle → date → weekday → schedule → venue → actions; only the section title uses script
- RSVP: Section Script → invitation sentence → form → buttons → helper text; only the section title uses script
- A Few Things to Know: Section Script → Chinese subtitle → compact FAQ questions and answers; only the section title uses script

### Consistency Rule

Before styling any text, classify it:

- Bride or groom name → Display Serif
- `Wedding Day`, `RSVP`, `Bride`, `Groom`, `Our Story`, or `A Few Things to Know` heading → Section Script
- Everything else → Sans-serif
- Chinese → `var(--font-sans-zh)` plus the shared Chinese semantic token for its role
