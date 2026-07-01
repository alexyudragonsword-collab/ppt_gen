# Style And Prompt Reference

## Visual Style

- Canvas: 16:9 landscape, preferably 3840 x 2160 in prompt intent; generated tool may return a smaller 16:9 image.
- Palette: white background, Huawei red `#C7000B` accents, deep blue `#003B8F` diagrams, light gray-blue panels, black titles.
- Layout: title top-left, red underline, HUAWEI logo top-right, summary strip bottom, internal label bottom-left, page number bottom-right.
- Background: subtle pale-gray circuit pattern, very low contrast.
- Visual language: chips, SC circuits, waveforms, clock phases, spectrum charts, trade-off triangles, roadmaps, risk matrices.
- Audience: Chinese technical management. Emphasize decision value, risk, roadmap, and quantified examples.

## Prompt Template

Use one prompt per slide:

```text
Use case: productivity-visual / infographic-diagram.
Asset type: 16:9 Chinese technical management presentation slide image.
Create slide <PAGE> of <TOTAL> in a Huawei internal communication style: clean white background, subtle pale-gray circuit pattern, Huawei-red accent lines, deep-blue technical diagrams, polished corporate engineering aesthetic.

MANDATORY fixed layout: large bold title at top-left, same position and size across all slides; red underline under title; Huawei logo mark and word HUAWEI at top-right; bottom-left red outlined label with exact Chinese text “内部交流”; bottom-right page number “<PAGE>/<TOTAL>”; bottom summary box across bottom with red border and target icon.

Title text, verbatim: “<TITLE>”
Main visual: <VISUAL DESCRIPTION>
Content: <3-6 SHORT BULLETS OR KPI CARDS>
Bottom summary text, verbatim: “一句话总结：<ONE SENTENCE>”

Constraints: professional Chinese typography, all Chinese text legible and correctly spelled, no extra logos, no watermark, no HTML/SVG look, raster presentation slide, unified red-blue-white palette.
```

## Recommended 18-Slide Structure

1. Cover: topic, audience, scope, three value badges.
2. One-sentence conclusion: why the technology matters.
3. Technical background: application scenarios and trends.
4. Mechanism: sampling, charge sharing, historical charge memory.
5. Advantages and costs: benefit/risk balance.
6. Topology roadmap: major PSC-IIR families and application bands.
7. Key model: `R_eq = 1/(C_S f_S)` and multi-phase CT equivalent.
8. Trade-off triangle: selectivity, noise, power.
9. Non-ideal factors: risk matrix and mitigations.
10. Clocking: non-overlap, phase sequence, jitter, power.
11. Noise and linearity: kT/C, folding, IIP3/SFDR.
12. Verification flow: z-domain, CT model, PSS/PAC, PNOISE, FFT, MC, PEX.
13. Low-frequency example: audio/sensor sizing.
14. Silicon example: measured low-IF prototype metrics.
15. Frequency expansion: baseband to N-path RF.
16. Alternatives comparison: PSC-IIR vs active SC, FIR, RC/gm-C.
17. Project landing plan: phases, gates, deliverables.
18. Final recommendation: decision checklist and next step.

## Slide Text Limits

- Title: 8-22 Chinese characters or bilingual equivalent.
- Bullets: max 4 lines per side; each line under 22 Chinese characters when possible.
- KPI cards: max 8 cards per slide.
- Summary sentence: one sentence, under 34 Chinese characters when possible.

## QA Heuristics

- Check logo, title, page number, and `内部交流` at thumbnail size.
- Check the summary sentence at full size.
- Regenerate any slide where page number, title, or mandatory label is malformed.
- Prefer correcting slide content in the prompt rather than post-editing the bitmap.
