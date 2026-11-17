# Capability integration and upstream provenance

This package is an independently authored implementation combining the following workflow capabilities. It does not redistribute the upstream projects' scripts, templates, model weights or private runtime modules, and does not require six separate installations. Names identify sources of design ideas, not affiliation. Source snapshot date: 2026-11-17.

| Capability | Source inspected | Integration |
|---|---|---|
| Academic PPTX | https://github.com/Gabberflast/academic-pptx-skill/blob/main/SKILL.md | Claim-title narrative, exhibit discipline and restrained scholarly layout |
| Scholar Slides | https://github.com/luwill/research-skills/blob/main/scholar-slides/SKILL.md | Paper digest, traceable evidence and exact editable scientific content |
| PaperBanana | https://github.com/dwzhu-pku/PaperBanana/blob/main/README.md | Reference-informed planning and iterative diagram critique |
| Viz Skills | https://github.com/cupcake777/viz-skills/blob/main/AGENTS.md | Claim-driven chart selection, real data, stable semantic styling |
| Scientific Slides | https://github.com/K-Dense-AI/scientific-agent-skills/blob/main/skills/scientific-slides/SKILL.md | Talk-specific pacing, scientific visuals and iterative visual review |
| Slides | https://github.com/JetBrains/skills/blob/main/slides/SKILL.md | Native editable PPTX elements, aspect-aware image placement and render/inspect/fix cycle |

The Slides mirror attributes OpenAI's source at https://github.com/openai/skills/tree/e6afb0d74cc75d220df2faf3dd6c635c2dc6a108/skills/.curated/slides. The current original path was unavailable when checked; the inspected mirror was used for workflow understanding.

Resolve differences in favor of this task's requirements: native editable PPTX over image-slide defaults; traceable data over decorative plots; readable argument over fixed visual-density targets; authorized autonomous production over mandatory approval pauses. The included backend uses python-pptx for portability and offline operation rather than copying an upstream JavaScript stack. Optional PaperBanana/model services are not bundled or required. Specialized diagrams may need a domain renderer or manual native construction using the documented semantic contract.
