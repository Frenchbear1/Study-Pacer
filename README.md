# Study Pacer

Read your study material in one calm, themed reader that paces you through it.

Import PDFs, PowerPoints, Word documents, web pages, EPUBs, Markdown, text files or images, pick how much time you have, and read. Study Pacer tracks every page and slide across all the documents in your session, tells you how many are left, keeps each page on time, and interrupts you for breaks using the study method you choose.

## What it does

- **Reads everything in-app:** PDF (pdf.js), PowerPoint .pptx (drawn from the file's own layout: themes, masters, bullets, tables, pictures, speaker notes), Word .docx, HTML/web pages (cleaned to reader view), EPUB, Markdown, plain text and image sets. Text documents are cut into printed-page-sized pages.
- **Paces across several documents at once:** one session can hold many documents, each with its own page range, in any order (drag to reorder).
- **Three ways to set time:** a total ("90 minutes"), a pace ("45 s a page"), a deadline ("done by 4:30"), or open-ended.
- **Smart pacing:** pages with more text get more time and title slides get less (strength adjustable), or use even, front-loaded or back-loaded pacing. Time you gain or lose is shared over the remaining pages, carried to the next page (the original Study Pacer behaviour), or kept fixed.
- **Study methods as editable presets:** Pomodoro (25/5, long break every 4), Long Pomodoro (50/10), 52/17, Ultradian 90/20, Flowtime (break when you want; you earn a fifth of your focus time), Page sprints, or your own custom rhythm.
- **Interruptions:** when a focus block ends you get a full-screen animation and can take the break, ask for +2/+5/+10/+15/+25 minutes, finish the page first, or skip the break. A quiet heads-up comes a minute beforehand.
- **Breaks:** a breathing countdown, a rest tip and a **recall check** (write down what you remember from the pages you just read; it's saved with the session). The clock waits for you to say you're ready.
- **Pacer bar:** a page ring (with an outer focus-block arc), time left, pages to go, predicted finish and whether you're ahead or behind, based on your actual pace.
- **Learns your pace:** your words-per-minute and seconds-per-slide feed future time suggestions, and you get today/week totals and a streak.
- **Findly's reading themes:** the same colour themes, real paper textures, page tinting (PDFs, slides and images take on the theme; dark themes turn pages dark), and per-surface glass transparency and blur.
- **Reader:** page-by-page with Slide/Stack/Fade turns and swipe-to-turn that follows your finger, or continuous scroll; fit page or width; zoom; flag pages; a page grid overview; session stats with a time-per-page chart and log.
- Pauses automatically when you leave, keeps the screen awake during a session, resumes an unfinished session after a reload, and works offline as an installable app.

## Search

Every document is indexed by page as it comes in (headings and captions marked), and searched with the same engine as Findly: **Any word**, **All words**, **Exact phrase**, **Near** (within N words) and **Boolean** (AND, OR, NOT and parentheses), with match case, whole words, and **Look in** Anywhere / Headings / Captions / Body text. Search everything, the folder you're in, the document you're reading or the whole session. Results show each matching page with a highlighted snippet; opening one highlights the matches right on the page (PDFs included), and a hit bar steps between the matching pages. `Ctrl F` or `/` opens it.

## Changing a session

**⋯ → Edit session** (or `E`) opens the setup with the session's current settings: change the time, the study method and its numbers, the pacing, breaks, page ranges, or add and remove documents. Pages already read, time so far and where you are all carry over.

## Windows app

A lightweight Windows version (a few MB, using the WebView2 engine built into Windows) that reads your documents straight from folders on your computer. Nothing is copied:

- **Add folder** attaches any folder. Its subfolders become folders in the app, with a folder tree on the left, folder tiles and a breadcrumb.
- New, changed and deleted files are picked up when you come back to the app (or with the refresh button). Your progress on a file is kept as long as it's still there.
- Photos in the same folder read as one collection, in name order. To add photos, drop them into that folder.
- Select a folder to read or pace everything in it, in folder order.

Download it from the **Windows app** release on GitHub: `StudyPacer.exe` runs without installing; `StudyPacer-Setup.exe` installs it. It's rebuilt automatically whenever the app changes.

## Photo collections

Importing several photos asks whether they should be one collection (named, read in order like pages), separate photos, or added to a collection you already have. Add more any time from a collection's menu (**Add photos…**) or by dropping photos onto its card.

## Keys

`Ctrl F` or `/` search · `E` edit session · `←/→` pages · `Space` scrolls then turns · `P` pause · `B` break now · `F` flag · `G` all pages · `S` stats · `T` theme · `V` layout · `+/−` zoom · `H` hide bars · `Esc` back

## Running it

It's a static site with no build step: serve the folder (GitHub Pages, Vercel, or `python3 -m http.server`) and open `index.html`. Everything is stored on your device in IndexedDB.

## Credits

pdf.js (Apache 2.0), JSZip (MIT), mammoth.js (BSD-2). Paper textures adapted from Subtle Patterns (CC BY-SA 3.0, see `vendor/paper/LICENSE.txt`). Themes and glass system shared with Findly.
