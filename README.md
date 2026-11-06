# PPTX Viewer & PowerPoint Viewer

For developers who get slide decks in a repo or a pull request, PPTX Viewer is a PowerPoint viewer that opens your decks in VS Code with thumbnails, speaker notes and search.

<p><a href="https://marketplace.visualstudio.com/items?itemName=zovo.pptx-slide-viewer"><img src="https://img.shields.io/visual-studio-marketplace/v/zovo.pptx-slide-viewer" alt="Version"></a> <a href="https://marketplace.visualstudio.com/items?itemName=zovo.pptx-slide-viewer"><img src="https://img.shields.io/visual-studio-marketplace/i/zovo.pptx-slide-viewer" alt="Installs"></a> <a href="https://open-vsx.org/extension/zovo/pptx-slide-viewer"><img src="https://img.shields.io/open-vsx/v/zovo/pptx-slide-viewer" alt="Open VSX"></a></p>

![A 6 slide deck open in VS Code. Slide thumbnails on the left, slide 4 with a photo, a blue rounded box and a green circle in the middle, and its speaker notes underneath](https://raw.githubusercontent.com/theluckystrike/pptx-viewer-vscode/main/images/vscode-pptx.png)

*A real VS Code 1.142 window running this extension on [`test/fixtures/quarterly-review.pptx`](test/fixtures/quarterly-review.pptx), made by [`scripts/make-samples.py`](scripts/make-samples.py).*

- See each slide with thumbnails and its speaker notes. Zoom in on small print, or scroll through the whole deck like a document.
- Search titles, bullets, table cells and notes across the whole deck, then jump to the slide.
- Copy one slide or the whole deck as clean Markdown for a README, an issue or a chat with an AI assistant.

## Quick start

1. Install the extension.
2. Click any `.pptx`, `.ppsx`, `.potx` or `.pptm` file in the Explorer. Press `Ctrl+F` (`Cmd+F` on macOS) to search.

![A click on the .pptx file opens slide 1. Clicks on the thumbnails move to slides 2, 3 and 4, and a search for latency lists the matching slides in the Outline tab](https://raw.githubusercontent.com/theluckystrike/pptx-viewer-vscode/main/images/demo-pptx.gif)

*Recorded in the same window. The white arrow is drawn in to show where each click lands.*

Run "Get Started with PPTX Slide Viewer" from the Welcome page for a short tour.

## Search slides and notes

Type in the search box. The Outline tab lists every slide where the word appears, with the match marked, and says when the match is only in the notes. `Enter` goes to the next matching slide and `Shift+Enter` to the previous one. `Esc` clears the search.

<img src="images/screenshot-search.png" alt="Search for latency. The outline lists slides 1, 2, 3 and 6 with the word marked, and slide 3 shows a table of latency by region." width="100%">

Without a search, the Outline tab shows each slide title with its first lines of text, which is a quick way to skim a long deck.

## Zoom and scroll view

Use the `-` and `+` buttons or the zoom list in the toolbar. The list has Fit slide, Fit width and fixed sizes from 50% to 400%, and it shows the real percent. `Ctrl` or `Cmd` with the mouse wheel zooms around the mouse, and so does a trackpad pinch. A slide larger than the editor scrolls, so you can read a small table or chart label.

"Scroll" puts every slide in one column, like a document. The slide in view becomes the current one, so the thumbnails, the slide number and the speaker notes follow you as you scroll. Slides are drawn as they come near the view, so a long deck stays fast.

Each file keeps its last view and zoom in the workspace. The settings `pptxViewer.defaultView` and `pptxViewer.defaultZoom` choose how a deck opens the first time.

## Copy as Markdown

"Copy as Markdown" puts the current slide on the clipboard. The title becomes a `##` heading, bullets keep their indent levels, tables become Markdown tables, and the speaker notes become a quote. "Copy all" does the same for every slide. Text that looks like HTML is escaped, so it pastes as text.

## What it draws, and what it doesn't

Slides are drawn with their text boxes, pictures, tables and shapes. Charts are drawn from the numbers saved in the file, so they can look a little different from PowerPoint, and the slide says so.

SmartArt, video, audio, ink and embedded objects are not drawn. A slide that has any of them shows a short message that names them, so you know something is missing.

<img src="images/screenshot-light.png" alt="The table slide in a light VS Code theme." width="100%">

## Keys

| Key | Action |
| --- | --- |
| `←` `→`, `PageUp` `PageDown`, `Space` | Previous or next slide |
| `Home`, `End` | First or last slide |
| `Ctrl+F` / `Cmd+F` | Search slides and notes |
| `Enter`, `Shift+Enter` in the search box | Next or previous matching slide |
| `+`, `-` | Zoom in or out |
| `Ctrl+=`, `Ctrl+-` / `Cmd+=`, `Cmd+-` | Zoom in or out |
| `Ctrl` / `Cmd` + mouse wheel, trackpad pinch | Zoom around the mouse |
| `0` | Fit the slide in the editor |

## Settings

| Setting | Default | What it does |
| --- | --- | --- |
| `pptxViewer.maxFileSizeMB` | `150` | Ask before opening larger decks |
| `pptxViewer.showNotes` | `true` | Show speaker notes under the slide |
| `pptxViewer.renderSlides` | `true` | Draw the slides. Set to `false` to see only the outline, notes and search |
| `pptxViewer.defaultView` | `slide` | How a deck opens the first time. `slide` shows one slide at a time and `scroll` puts all slides in one column |
| `pptxViewer.defaultZoom` | `fit` | Zoom a deck opens with the first time, `fit` or `width` |
| `pptxViewer.askForRating` | `true` | Ask one time for a Marketplace rating, after 5 files and 2 days. Turn it off with `false` |

## Large decks

The extension reads the text of the deck first and skips images and video while it does, so the outline, notes and search are ready before any slide is drawn. A 120-slide test deck parses in about 20 ms. Thumbnails are drawn only as you scroll to them. Above 150 MB the viewer asks before it opens the file.

## Privacy and security

Everything runs on your machine. The extension makes no network calls and has no telemetry. The viewer runs under a strict Content-Security-Policy that allows only the extension's own scripts and blocks all network access. Text from the deck is inserted as plain text, never as HTML. Macros in `.pptm` files are never run, and links inside slides do nothing when clicked.

## Known limits

- It is a viewer. It doesn't edit, present full screen or export to PDF.
- Old binary `.ppt` files (PowerPoint 97 to 2003) are not supported. Save them as `.pptx` first.
- Fonts that are not installed on your machine are replaced, so line breaks can differ from PowerPoint.
- Table styles from the theme are not always drawn. The table text is always correct.
- Zoom changes the drawn slides. The text view, shown while slides are drawn or when drawing is off, keeps its normal size.
- If another PowerPoint extension is installed, VS Code may ask which one to use. Right-click the file and pick "Open With..." to switch.

## Credits

Slides are drawn by [pptx-preview](https://www.npmjs.com/package/pptx-preview) (ISC), which bundles ECharts (Apache-2.0), JSZip, lodash and uuid. The deck text is read with [fflate](https://github.com/101arrowz/fflate) (MIT). See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md). The extension itself is MIT licensed.

## Support

Zovo extensions have no ads and no tracking, and sponsors pay for the time that goes into fixes and new viewers. You can sponsor the work on [GitHub Sponsors](https://github.com/sponsors/theluckystrike).

## More viewers by Zovo

- [Context Drift](https://marketplace.visualstudio.com/items?itemName=zovo.context-drift) finds stale facts in AGENTS.md, CLAUDE.md and other agent rule files.
- [PDF Viewer & Reader with Dark Mode](https://marketplace.visualstudio.com/items?itemName=zovo.pdf-viewer-dark) reads PDFs with a real dark mode, search, an outline and thumbnails.
- [CSV, Excel & Parquet Viewer](https://marketplace.visualstudio.com/items?itemName=zovo.table-viewer-vscode) opens Parquet, Excel and CSV files as a sortable, filterable table.
- [DOCX Viewer & Word Viewer](https://marketplace.visualstudio.com/items?itemName=zovo.docx-viewer-vscode) previews Word files with dark mode, an outline and search.
- [HDF5 & NetCDF Explorer](https://marketplace.visualstudio.com/items?itemName=zovo.hdf5-netcdf-explorer) browses groups, variables and attributes in scientific data files.
- [Skill Shelf](https://marketplace.visualstudio.com/items?itemName=zovo.skill-shelf) lists, checks and toggles the Agent Skills that Claude Code, Codex and Cursor load.
- [Tensor Viewer](https://marketplace.visualstudio.com/items?itemName=zovo.tensor-viewer) shows NPY, NPZ, Safetensors and GGUF tensors with stats and heatmaps.

Coming soon to the Marketplace, with the code on GitHub now.

- [JSONL Viewer](https://github.com/theluckystrike/jsonl-viewer-vscode) opens JSONL and NDJSON files as a table, as records or as a chat view.
- [SQLite Table Viewer](https://github.com/theluckystrike/sqlite-table-viewer) browses SQLite tables read-only and runs SELECT queries.
- [Agent Session Viewer](https://github.com/theluckystrike/agent-session-viewer) browses and searches your local Claude Code and Codex sessions.
- [Audio Preview & Loudness Meter](https://github.com/theluckystrike/audio-preview-vscode) plays audio files with a waveform, a spectrogram and LUFS loudness.
- [Video Preview with Sound](https://github.com/theluckystrike/video-preview-vscode) plays MP4, MOV, MKV and WebM files with their sound.
- [Graphviz Preview & DOT Viewer](https://github.com/theluckystrike/graphviz-preview-vscode) shows a live preview of DOT files and exports SVG or PNG.
- [TIFF Viewer](https://github.com/theluckystrike/tiff-viewer-vscode) opens 16-bit, float, multi-page and GeoTIFF images.
- [Avro, Arrow & Feather Viewer](https://github.com/theluckystrike/arrow-avro-viewer-vscode) opens Avro, Arrow and Feather files as a paged table.
- [EML & MBOX Email Viewer](https://github.com/theluckystrike/eml-viewer-vscode) reads email files with remote images blocked.
- [3D Model Viewer](https://github.com/theluckystrike/model-viewer-vscode) views GLB, glTF, STL, OBJ, PLY and 3MF models.
- [File Viewers Pack](https://github.com/theluckystrike/zovo-file-viewers) installs the PDF, table, Word, PowerPoint, HDF5 and tensor viewers in one step.
