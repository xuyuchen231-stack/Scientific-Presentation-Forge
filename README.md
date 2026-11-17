# Scientific Presentation Forge

An installable skill for evidence-grounded, editable scientific presentations: PhD seminars, academic conferences, paper reports, research proposals and thesis defenses.

**[Download skill.zip](https://github.com/xuyuchen231-stack/Scientific-Presentation-Forge/raw/refs/heads/main/skill.zip)**

## Six integrated capabilities

| Capability | Integration |
|---|---|
| Academic PPTX | Scholarly argument, claim titles and exhibit design |
| Scholar Slides | Paper extraction, source fidelity, equations and citations |
| PaperBanana | Scientific diagram planning and iterative critique |
| Viz Skills | Claim-driven visualization of real research data |
| Scientific Slides | Talk structure, pacing and visual inspection |
| Slides | Native editable PPTX authoring and render–inspect–fix workflow |

The package contains an independently authored implementation and source records. The six upstream projects do not need to be installed separately.

## Install

Download `skill.zip` and import it into an agent that supports custom skill ZIPs. For folder-based installation, extract it and place `scientific-presentation-forge` in your agent's skill directory. The entry point is `scientific-presentation-forge/SKILL.md`.

Install runtime dependencies in your Python environment:

```powershell
python -m pip install -r scientific-presentation-forge/requirements.txt
python scientific-presentation-forge/scripts/forge.py doctor
```

Python 3.10+ is required. Automatic PPTX rendering uses LibreOffice (`soffice` on PATH); alternatively export the final PPTX to PDF in PowerPoint and supply that PDF to the review tool. No model API key is required for the bundled tools.

## Use

Ask the agent: “Use Scientific Presentation Forge to turn my paper and real data into an editable 15-minute PhD seminar presentation, with model diagrams, accurate formulas, citations, speaker notes and inspection of every slide.”

The bundled tools provide PDF text/image inventory extraction, native PPTX generation, structural audits, numbered slide previews and a hash-bound inspection record. Charts retain an embedded editable workbook; tables, text and diagram nodes are native PowerPoint objects. Complex figures retain their original assets and editable external sources. Equations support simple editable text or reviewed Office Math XML.

```powershell
python scientific-presentation-forge/scripts/selftest.py
python scientific-presentation-forge/scripts/forge.py build deck.json deck.pptx
python scientific-presentation-forge/scripts/forge.py audit deck.json deck.pptx --out audit.json
python scientific-presentation-forge/scripts/render_review.py deck.pptx review
python scientific-presentation-forge/scripts/render_review.py deck.pptx review --check
```

After rendering, actually inspect each slide and record its findings in `review/review.json`. The review check rejects missing coverage and changed PPTX/preview files. Structural checks do not replace source verification or visual inspection.

## Validation

The package passed skill-format validation and executable software fixtures covering native charts, embedded workbooks, Office Math, diagrams, notes, invalid data and stale review detection. A three-slide fermentation research-proposal exercise was generated, rendered and visually inspected. See the authoring guide for supported specification fields and the origins guide for upstream links.

## License

The independently authored package is MIT licensed. Source projects and optional dependencies retain their respective licenses.
