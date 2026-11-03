# Mari Takayama — website

Personal homepage for Mari Takayama, a professional ballet dancer based in NYC.
Two people edit this repo with different AI tools (Claude Code and Codex). These rules keep the work consistent.

## Structure

- `index.html` — the whole site: HTML, CSS (in `<style>`) and a small script. No build step, no framework.
- `img/` — photos used by the page.
- Preview: open `index.html` in a browser.
- Published with GitHub Pages from the `main` branch.

## Design (keep it)

- Sections in order: nav → hero → About → Works (Lesson, Radio Taiso) → Gallery → Interview → footer/contact.
- Colors are tokens on `:root` (`--slate`, `--mauve`, `--rose`, `--stage`, …). Change a token rather than adding new literal colors.
- Fonts: Cormorant Garamond (headings, nav), EB Garamond (body), from Google Fonts.
- Site copy is in English.
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
