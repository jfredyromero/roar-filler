# rOar Filler

A Chrome extension that automatically fills daily hours in the [rOar](https://roar.marcel.ai) timesheet, so you don't have to enter the same recurring entries by hand every week.

## What it does

- Reads the current week directly from the timesheet page (dates, holidays, hours already logged).
- Reuses the active session's auth token from the MSAL cache — no separate login step.
- Fills each enabled weekday with a fixed set of entries:
  - **Regular days:** 1h Project Meeting + 7h Development & Engineering Collaboration.
  - **Holidays** (days marked "Public Holiday" on the timesheet): 8h Offtime, using a separate Holiday job row.
- Skips days that already have 8+ hours, and warns (without touching them) on days with a partial entry, so it never duplicates hours.
- Offers a **Preview** mode that shows exactly what would be submitted, without saving anything.

## Installation

1. Open `chrome://extensions` in Chrome.
2. Enable **Developer mode** (top right).
3. Click **Load unpacked** and select this project's folder.
4. Pin the rOar Filler icon to the toolbar if you want quick access.

## Usage

1. Open your timesheet at `roar.marcel.ai` and make sure the correct job(s) are already added for the week (the regular job, and the Holiday job only on weeks that actually have a holiday).
2. Click the extension icon.
3. **Preview** — shows the detected week and what would be filled, without saving.
4. **Fill week** — submits the entries to rOar.
5. **Reload page** — reloads the timesheet tab, useful after filling to verify the result before submitting for approval.

## Configuration

The job codes, task codes, and hour splits are hardcoded in `index.js` since they're specific to one employee/role setup:

- `EMPLOYEE` — employee code.
- `JOB_NORMAL` / `JOB_HOLIDAY` — job IDs, work codes, and PPM task metadata for regular and holiday weeks.
- `ENTRIES_NORMAL` / `ENTRIES_HOLIDAY` — the hour breakdown applied per day.

Update these constants if you fork this for a different employee, job, or hour allocation.

## How it works (technical notes)

- The timesheet grid lives inside an iframe, so the script is injected into `allFrames: true` and the correct frame's result is picked by content, not by index.
- It reads with `world: 'MAIN'` because the row's `userJobId` isn't available from an isolated content script — it has to be pulled from the React fiber tree (`__reactFiber$...`) attached to the DOM node, since rOar creates a new job assignment id every week.
- The auth token is pulled from `localStorage`/`sessionStorage` MSAL cache entries matching `timesheet_api_access`, picking the most recent non-expired one.

## Disclaimer

This is a personal productivity tool built against rOar's current (undocumented, unofficial) internal API and DOM structure. It can break silently if rOar changes its frontend or API contract — always use **Preview** first and verify the submitted timesheet before final approval.
