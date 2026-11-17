# Model diagrams and data visualization

## Structure diagram contract

Write a node table (ID, label, scientific meaning, input/output, learned/mechanistic role) and edge table (source, target, quantity, direction, role). Distinguish measurement, model computation, estimation, optimization and actuation. Represent feedback only where the implemented system contains it. Verify dimensions and names against equations and code. Use grouped native PowerPoint shapes, connected arrows, consistent spacing and minimal color semantics. Label arrows carrying states, rates or control inputs. Avoid crossing connectors and arrows hidden beneath nodes.

Adapt the PaperBanana workflow into five passes: inspect relevant examples; plan semantics and layout; select restrained scientific styling; construct the visual; critique scientific correctness and readability, then revise. These passes can be performed by one agent; external model pipelines are optional. Generative imagery is appropriate for conceptual illustrations, never exact model connectivity, quantitative plots, formulas or citations. Rebuild precise diagrams deterministically.

For fermentation digital twins, distinguish process, sensors, state estimator, mechanistic balances, learned component and controller when actually present. Indicate whether learning predicts rates, parameters, dynamic residuals or output residuals. Include the inputs needed to explain changing feed/induction responses. Keep CFD/compartment, GEM/FBA or NMPC blocks only when supported by the actual project.

## Chart selection

| Claim | Preferred display | Required checks |
|---|---|---|
| Dynamic response | XY scatter with connected numerical x coordinates, observations as markers, simulations as lines | Time units, perturbations, irregular sampling, prediction horizon, uncertainty definition |
| Model comparison | Paired outcomes, held-out metrics or per-batch plots | Same test data, metric direction, units, replicate definition |
| Distribution | Dots/box/violin with observed values | Sample size, grouping and uncertainty |
| Relationship | Scatter with appropriate fitted curve | Confounding, sample units, fit assumptions |
| Ablation | Paired differences or interval plot | Comparable evaluation protocol |
| Dense matrix | Heatmap with meaningful scale | Colorbar units, ordering, missing values |

Use native charts for simple supported plots so data remain editable. Use a domain plotting library for complex statistical displays, retaining CSV/code plus SVG/PDF/PNG. Preserve original values and numerical x spacing. Do not use an equally spaced category line chart for irregular time samples. Avoid 3D decoration, dual axes unless clearly justified, rainbow palettes and unjustified truncated bar baselines. Keep one color per scientific role throughout the deck. Use line styles and markers in addition to color.
