# Executable authoring contract

Use Python 3.10+ and install `requirements.txt`. Run `python scripts/forge.py doctor` to check libraries and renderer. All input/output paths are explicit; asset paths resolve relative to the spec. Tools do not fetch papers, verify references online or install external applications automatically.

## Commands

```text
python scripts/forge.py ingest paper.pdf source.json
python scripts/forge.py build deck.json deck.pptx
python scripts/forge.py audit deck.json deck.pptx --out audit.json
python scripts/render_review.py deck.pptx review
python scripts/render_review.py deck.pptx review --pdf deck.pdf
python scripts/render_review.py deck.pptx review --check
```

`ingest` extracts page text and image bounding boxes. It does not infer figure captions, OCR scanned pages or certify equations; inspect source pages. `audit` exits nonzero for structural failures, not for stylistic warnings. `render_review --check` requires all slides to have a recorded resolved review and matching PPTX/image hashes. It cannot establish that a person or agent actually inspected a slide; do the inspection before recording it.

## JSON spec

Top-level fields: `title`, `font` (default Times New Roman), `sources` and `slides`. The canvas is 13.333 × 7.5 inches. Every slide requires `title`, `role`, `seconds`, `notes`, `citations` (source IDs) and `elements`. Roles include title, context, question, method, result, proposal, conclusion, references and appendix. The builder automatically adds a short citation footer from the cited ledger entries. Add complete references as native text on a references slide.

Every element has `kind` and `[x,y,w,h]` as `box`, in inches, within the slide canvas. Reserve y=0.35–1.15 for titles and y>=6.9 for citation footer. Keep content inside x=0.5–12.8, y=1.3–6.7. Element kinds:

- `text`: `text`, optional `size`, `bold`, `color`. Keep no more than a few concise statements.
- `table`: `rows`, an array of equal-length string arrays; optional `size`.
- `chart`: `x` (finite, numerical), `series` (`name`, `values`), `x_label`, `y_label`, `source_ids`. Uses a native XY chart with true numerical spacing. All series share x; adapt the source for distinct sampling only with recorded scientific justification or extend the backend. Complex uncertainty plots use an external verified figure and retained source.
- `diagram`: `nodes` (`id`, `label`, `box` relative to the diagram origin), `edges` (`from`, `to`, optional `label`). Nodes and arrows are native shapes. Layout must be designed; the builder does not run graph layout.
- `image`: `path`, `source_ids`, optional `description`. Preserves aspect ratio and embeds original pixels. Retain editable external source. The raster component remains an embedded image.
- `equation`: `text` for a simple editable Unicode expression; or `omml_path` for a complete reviewed `m:oMathPara` XML document, plus `latex` source. Office Math rendering varies by renderer; inspect PowerPoint when using advanced math. Never treat unchecked automatic conversion as exact.

Sources require `id`, `citation`, `reference`, `locator`, `verified`. Source IDs for charts/images must also be included in the slide's citations. The program enforces evidence presence for result slides; the author remains responsible for verifying all other factual assertions and figure permissions. Build rejects `[MISSING]`, `[UNVERIFIED]` and `[TODO]` markers and empty slide lists. Omit unsupported research rather than silently adding placeholders. Work in progress may use placeholders outside the final build.

## Review report

Rendering creates numbered PNGs, a montage and `review.json` with the final PPTX SHA-256 and each PNG's SHA-256. For each page, inspect title hierarchy, text wrapping, minimum readable type, formulas, arrows, scientific interpretation, figure labels, citations and alignment. Record `findings` as an array of strings, then `checked: true` and `resolved: true` once all issues have been corrected. A clean slide can record `findings: []` after actual inspection. Rebuilding changes the PPTX hash and invalidates the review.
