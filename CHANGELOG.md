# Changelog

## 0.4.0 (2026-11-08)

- Slides look closer to PowerPoint. Body text now takes its size, bullets and spacing from the slide master. Before, bullets in decks made with python-pptx and similar tools were drawn at 18 pt instead of 32 pt, without bullet marks and too low on the slide.
- Text uses the deck's own fonts, including theme fonts such as Calibri and Aptos. A font that is not installed is replaced by a close match instead of the editor's UI font, and Latin text is no longer squeezed together.
- A content placeholder that shares a number with the master's date box (the right column of a two-column slide, for example) is no longer drawn grey, small and centred.
- SmartArt saved by PowerPoint is drawn, and the "cannot draw SmartArt" message shows only for SmartArt saved without a drawn copy.
- Copy a slide as a PNG picture, or Shift+click to save it as a `.png` file ("PPTX: Copy Slide as PNG Picture", "PPTX: Save Slide as PNG…").
- Live reload also catches decks saved by other programs, such as a script or an AI agent, within about a second.
- Slides are drawn when they are first shown, so the first slide of a long deck appears sooner, and one slide that cannot be drawn shows its text with a note instead of stopping the whole deck.
- An empty file says "This file is empty (0 bytes)". When a file cannot be opened, the slide number box reads "– / –" and the slide buttons, Present, search and copy are greyed out.
- Decks up to 300 MB open without asking (was 150 MB). A public 197 MB deck shows its first slide in about 5 seconds.
- "Fit" leaves room for the message above a slide, so the slide no longer gets scroll bars. The copy buttons sit in one "Copy: Markdown, PNG, All" group, and the toolbar stays on one row in a normal window.
- Shapes with broken outlines no longer fill the developer console with errors.
- The Marketplace name is now "PPTX Viewer & PPT Viewer, PowerPoint Viewer & Slideshow". The extension ID, settings and commands stay the same.

## 0.3.0 (2026-11-07)

- Present. The "Present" button (or "PPTX: Present Slides (Full Screen)") shows one slide at a time on a black screen, fitted to the window. Arrow keys, Space, Page Up and Page Down or a click move between slides, and Esc ends it. VS Code goes to Zen Mode (full screen, no side bars) while you present and comes back when you press Esc. Turn that off with `pptxViewer.presentFullScreen`.
- Old PowerPoint 97-2003 files (`.ppt`, `.pps`, `.pot`) open as text: slide titles, slide text, text boxes and speaker notes, in slide order, with search, the outline and Copy as Markdown. A line at the top says that pictures, shapes and layout are not shown. No LibreOffice or other app is needed.
- A `.pptx` deck saved with a `.ppt` name opens as a normal deck. A file that is damaged, protected with a password, or not PowerPoint at all shows a message with an "Open in Default App" button instead of an empty page.

## 0.2.4 (2026-11-07)

- "More free viewers" now starts with the Zovo extensions most related to this one. Picking one opens its Marketplace page, where you choose to install it.
- The README has a "Support Zovo" section with GitHub Sponsors and the developer's GitHub profile, and "More tools by Zovo" lists related extensions.

## 0.2.3 (2026-11-07)

- The footer under the viewer has a "Rate" link next to Sponsor and More free viewers. It opens the review page of this extension on the Marketplace. Nothing pops up and nothing asks you for a rating. "Hide" still hides the whole line.
- Marketplace search keywords updated.

## 0.2.2 (2026-11-07)

- The "More free viewers" list and the README list now show the 15 Zovo extensions on the Marketplace under their current names, and 12 that are coming soon.
- Extensions for related files now come first, starting with DOCX Viewer & Word Viewer with Dark Mode and CSV Viewer, Excel Viewer & Parquet Viewer.

## 0.2.1 (2026-11-07)

- The Marketplace name is now "PPTX Viewer & PowerPoint Viewer", so a search for PowerPoint viewer finds it. The extension ID, settings and commands stay the same.

## 0.2.0 (2026-11-06)

- Zoom. The toolbar has `-` and `+` buttons and a zoom list with Fit slide, Fit width and sizes from 50% to 400%. `Ctrl` or `Cmd` with the mouse wheel zooms around the mouse, and so does a trackpad pinch. `+`, `-` and `0` work in the viewer, and `Ctrl+=` and `Ctrl+-` (`Cmd` on macOS) too. A zoomed slide scrolls.
- Scroll view. The "Scroll" button shows every slide in one column, like a document. The slide in view becomes the current one, and the thumbnails, slide number and notes follow it. Slides are drawn only as they come near the view.
- Each file keeps its last view and zoom in the workspace. New settings `pptxViewer.defaultView` and `pptxViewer.defaultZoom` set how a deck opens the first time.
- New commands "PPTX: Zoom In", "Zoom Out", "Zoom to Fit Slide", "Zoom to Fit Width", "Zoom to 100%" and "Toggle Scroll View (All Slides in One Column)".
- The toolbar wraps onto a second line in a narrow editor, so no button is cut off.

## 0.1.4 (2026-11-06)

- The Marketplace page now starts with a screenshot of a real VS Code window with slides, thumbnails and speaker notes, and a short GIF of moving between slides and searching. Both were made from generated sample files. The images load from the GitHub repo and are not packed into the extension, so the download size stays the same.

## 0.1.3 (2026-11-06)

- A small line at the bottom right of the viewer has two links, "Sponsor" (GitHub Sponsors) and "More free viewers". The new setting `pptxViewer.showSupportLinks` (default `true`) hides it, and so does the "Hide" link next to them. There is no popup, and nothing is sent anywhere.
- New commands "PPTX: Support the developer" and "PPTX: More free viewers". The second one lists the other Zovo extensions and installs only the one you pick. Ones that are not on the Marketplace yet open their GitHub page.
- The listing gets a Sponsor button. The README has a short Support section and lists all Zovo extensions.

## 0.1.2 (2026-11-06)

- PPTX Viewer asks one time for a Marketplace rating. It counts each success (an opened deck). After 5 of them, and no sooner than 2 days after the first, it shows a small message with "Rate it", "Later" and "Don't ask again". "Later" waits for 20 more, at most two more times. It never shows in your first session, more than once per session, when the window is not focused, or in a remote or untrusted workspace.
- New setting `pptxViewer.askForRating` (default `true`) turns it off. Nothing is sent anywhere; "Rate it" only opens the review page in your browser.

## 0.1.1 (2026-11-06)

- Listing only, no change to the extension. The display name is now "PPTX Viewer for PowerPoint Slides" and the description is one sentence, so the listing uses the words people search for (powerpoint (42), pptx (7)).
- The 30 keywords start with the most searched terms.
- The README ends with links to the other Zovo extensions.

## 0.1.0

- First release. It opens .pptx, .ppsx, .potx and .pptm decks as slides with thumbnails.
- Speaker notes under each slide, a text outline, and search across titles, text, tables and notes.
- Copy the current slide or the whole deck as Markdown.
- Clear notice on slides with SmartArt, video, audio, embedded objects or charts.
- Size limit with "Open anyway", and a text-only mode for very large decks.
