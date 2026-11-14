# Mari Takayama — website

Personal homepage for Mari Takayama, a professional ballet dancer based in NYC.
Two people edit this repo with different AI tools (Claude Code and Codex). These rules keep the work consistent.

## Structure

- One HTML file per page, each with its own CSS (in `<style>`) and small script. No build step, no framework.
  - `index.html` — Home: hero + a short overview card for each page, linking to the details
  - `about.html`, `private-lessons.html`, `radio-taiso.html`, `gallery.html`, `interview.html` (Press & Interviews: interviews plus articles about Mari's shows), `credits.html`, `booking.html`, `study-in-ny.html` (+ its Japanese twin `ballet-ryugaku.html`) — one page each
  - `lessons.html` — private lesson request form
- `editor.html` — photo/text editor; pick the page to edit at the top. Each page keeps its editor settings in `<script id="photo-editor-settings">`.
- `site-theme.css` — shared white-and-gold look (page titles, colors, buttons, menu), loaded by every page just before the editor styles. `site-motion.js` — scroll fade-ins, photo motion, soft page changes and the desktop cursor ring (off inside the editor).
- `site-shows.js` — upcoming performances (dates, times, ticket links). Home and Credits show them under Upcoming; the day after a production's last show it moves to the Credits "Ballet" list by itself. Add new shows there, not in the pages.
- `tools/booking-sheet.gs` — Google Apps Script for the lesson booking sheet (requests, availability, stats); its web app URL goes in `SHEET_APP` in `lessons.html`.
- `img/` — photos used by the page.
- Preview: open `index.html` in a browser.
- Published with GitHub Pages from the `main` branch.
- Google Analytics (G-5XCKST502L) is in each public page's `<head>`; it skips the editor, `/preview/` and local copies. Add the same snippet to any new page.

## Design (keep it)

- Menu order: Home → About Mari (About, Career [credits.html], Gallery, Press [interview.html]; all shown directly, no dropdown — on wide screens the menu is a second row under the name) → Work with Mari (Private Lessons, Book a Performance, For Ballet Students) → Community (Radio Taiso) → Contact. Group labels are `li.menu-group`, shown only in the phone menu. Every page shares the same nav and footer; keep them in sync.
- Colors are tokens on `:root` (`--slate`, `--mauve`, `--rose`, `--stage`, …). Change a token rather than adding new literal colors.
- Fonts: Cormorant Garamond (headings, nav), EB Garamond (body), from Google Fonts.
- Site copy is in English (the 日本語 button translates it via `site-language.js`). Exception: `ballet-ryugaku.html` is written in Japanese so Japanese searches find it; it is the Japanese twin of `study-in-ny.html`, so keep the two in step.
- Keep the layout working at phone width (about 400px) with no horizontal scroll.

## Photos

- Put new photos in `img/` as `.jpg`, long edge about 2000px at most, file names in lowercase-with-dashes.
- Every `<img>` that shows content gets an `alt` describing the photo; decorative banner images use `alt=""`.
- Remove photos from `img/` that the page no longer uses.

## Working together

- Before starting: `git pull`.
- Keep each change small and focused; commit with a message that says what changed (e.g. "Replace hero photo").
- When done: `git push` right away so the other person gets it.
- Don't rewrite or reformat parts of `index.html` you weren't asked to change; it makes the other person's changes conflict.
