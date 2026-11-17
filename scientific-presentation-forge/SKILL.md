---
name: scientific-presentation-forge
description: "Create evidence-grounded, editable scientific PowerPoint presentations for PhD seminars, academic conferences, paper reports, research proposals and thesis defenses. Integrate Academic PPTX, Scholar Slides, PaperBanana, Viz Skills, Scientific Slides and Slides capabilities: scholarly narrative, paper extraction, professional model diagrams, real-data visualization, exact equations and citations, native PPTX authoring and slide-by-slide visual inspection. Use for 科研PPT、博士Seminar、论文汇报、开题、毕业答辩 and research slide revision."
---

# Scientific Presentation Forge

Produce the presentation, its rebuildable source, and evidence of inspection. Treat this package as an integrated workflow with runnable offline tools; the six upstream projects need not be installed. Use optional upstream services only when available and appropriate. Read [origins.md](references/origins.md) for the six capability mappings and provenance.

## Execute

1. Inspect supplied papers, data, prior decks and institutional templates. Infer talk language independently of conversation language. Establish audience, duration and purpose from context; ask only for missing information that prevents useful work. Default to a 15-minute English research talk, 16:9, white background, navy text. Respect supplied typography; for this user's research decks prefer Times New Roman unless their template says otherwise.
2. Read [research.md](references/research.md). Extract full text, captions, tables and equations; make a source ledger with page/figure/table locations. Resolve citations against original publications or verified metadata. Separate the presenter's contribution from prior work. Never invent results to fill a slide.
3. Read [talks.md](references/talks.md). Build a slide plan with role, claim title, evidence IDs, exhibit, notes and seconds. Read titles sequentially to test the argument. Proceed autonomously when the request already authorizes production; do not impose recurring outline-approval checkpoints.
4. Read [visuals.md](references/visuals.md). Design model structures from a written node/edge contract, then critique scientific meaning and geometry. Plot supplied measured data or explicitly identified model simulation; keep their meanings separate. Retain training/test splits and perturbation times in dynamic modelling presentations.
5. Read [authoring.md](references/authoring.md). Create editable text, native tables/charts, connected shapes and equations. Use `scripts/forge.py build deck.json deck.pptx` for the bundled Python backend; use an available presentation runtime or PptxGenJS for richer layouts while keeping the same integrity and review gates. Full-slide screenshots cannot substitute for editable slides.
6. Run `python scripts/forge.py audit deck.json deck.pptx --out audit.json`. Correct missing citations/assets, unresolved placeholders, non-finite data, overflow and empty content. Audit is structural: independently verify scientific truth and inspect text for wrapping and readability.
7. Run `python scripts/render_review.py deck.pptx review` with LibreOffice available, or use `--pdf rendered-from-this-deck.pdf` after exporting the final PPTX in PowerPoint. Inspect **every** resulting PNG at readable size, plus the montage. Record each slide in `review/review.json` with `checked: true`, detailed `findings`, and `resolved: true` only after repair and reinspection. Never mechanically prefill a visual pass. Run `python scripts/render_review.py deck.pptx review --check` after recording review. Re-render and repeat review after any change to the PPTX.
8. Deliver the editable PPTX, rebuildable specification/source and data/figure assets, and concise inspection results. Include a PDF when requested or useful. Supply speaker notes matched to the visible slide wording and a references appendix. Keep conclusions as the last main-talk slide; put references and backup analyses after it. Report an unavailable rendering step accurately and leave visual status incomplete instead of claiming success.

## Run the bundled tools

Install declared dependencies in an available task environment with `python -m pip install -r requirements.txt`. Check with `python scripts/forge.py doctor`. No model API key is required. LibreOffice is an optional external renderer, required for automatic PPTX-to-PDF conversion; `--pdf` accepts a separately exported PDF. Never silently install system software.

Copy `assets/deck.template.json` into a task directory. Populate actual content; the empty template deliberately contains no invented research. Paths in the spec are relative to the spec. See [authoring.md](references/authoring.md) for the complete supported contract and commands. For source-PDF extraction run `python scripts/forge.py ingest paper.pdf source.json`; manually inspect caption/crop association, superscripts and OCR when needed.

## Completion gates

- Every quantitative claim and borrowed asset has a verified ledger entry and a visible short citation. Research not yet performed stays a proposed experiment in future tense.
- Equations preserve symbols, signs, subscripts, units and assumptions. Use native linear math for simple expressions; use reviewed Office Math XML for complex expressions. A LaTeX string by itself is a source, not proof of native equation editability.
- Text, tables, simple data charts and diagram nodes remain independently editable. Retain original artwork and editable source for imported complex figures; disclose embedded-image components.
- Every final slide has been rendered and inspected; structural audit and the hash-bound visual review both pass. Automated checks do not certify scientific accuracy or accessibility.
- Preserve real observations. Label digitized values and simulation outputs with their actual origin; never add artificial error bars, sample sizes or significance markers.
