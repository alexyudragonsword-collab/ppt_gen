---
name: huawei-internal-pptx-from-doc
description: Use when a user provides a report, PDF, DOCX, Markdown, or technical document and asks Codex to generate a Chinese Huawei internal communication style PPTX, slide images, executive technical presentation, one-click deck, or a similar red-blue-white image-based management report deck.
---

# Huawei Internal PPTX From Doc

## Overview

Turn a source document into a Chinese technical-management presentation in a Huawei internal交流 visual style: summarize the document, plan a deck under 20 slides, generate one full-slide bitmap per slide, then package those images into a widescreen PPTX.

Use this skill together with the relevant source-reading skill (`pdf`, `documents`, or `spreadsheets`), the built-in `image_gen` tool, and the `presentations` skill for final PPTX validation.

## Workflow

1. Read the source document completely enough to extract claims, metrics, examples, constraints, and references. For PDFs, use `pdfplumber` or the `pdf` skill; for DOCX, use the `documents` skill.
2. Write a Chinese executive technical summary: value proposition, architecture, mechanisms, risks, validation method, examples, and adoption recommendation.
3. Build a slide outline with a cover page and no more than 20 slides. Prefer 16-18 slides for dense technical reports.
4. Read `references/style-and-prompts.md` and use its fixed layout contract.
5. Lock the full outline, total page count, and one prompt per slide before generating images. Keep each page visual, concise, and self-contained.
6. Generate slide bitmaps with maximum practical parallelism: launch as many independent built-in `image_gen` calls at once as the current platform/tooling allows, with exactly one final `image_gen` output selected per slide. Do not intentionally wait for slide 01 to finish before starting slide 02 when the prompts are already ready.
7. After a parallel batch returns, copy the selected generated PNGs into a workspace folder named like `output/imagegen/<topic>-slides/slide-01.png` through `slide-NN.png`. Record the original generated image path, copied path, slide number, and prompt title in `output/imagegen/<topic>-slides/imagegen-ledger.txt`.
8. If any slide visibly fails QA or generation fails, regenerate only the failed slide numbers in the next parallel repair batch, then update the ledger so each final `slide-NN.png` maps to the selected successful `image_gen` output.
9. Run `scripts/build_image_deck.mjs` to merge images into a PPTX, one full-frame image per slide.
10. Verify the PPTX and the generation lineage: expected slide count, expected media count, `imagegen-ledger.txt` has one entry per slide, final rendered contact sheet has no blank pages, no visible borders, and page numbers match `x/N`.

## Non-Negotiable Generation Contract

For this skill, a "slide image" means a direct output from the built-in `image_gen` tool copied into the workspace. Do not substitute locally drawn full-slide images for any final slide.

- Do not create final slide images with PIL, Python drawing, HTML/CSS screenshots, SVG, canvas, PowerPoint native shapes, or any other deterministic renderer, even to improve Chinese text fidelity.
- Do not post-process generated slide images by drawing or replacing titles, page numbers, labels, logos, diagrams, or summary strips. Filesystem copy/rename, metadata inspection, and PPTX `fit=cover` placement are allowed.
- If generated Chinese text, page numbers, or required chrome are wrong, regenerate that slide with shorter text and stronger prompt constraints. Do not repair the bitmap with local drawing.
- If repeated regeneration cannot satisfy the mandatory chrome or text, stop and report the limitation instead of switching execution strategy.
- If the user explicitly asks for editable PPTX or deterministic text rendering instead of image-generated slides, use a different presentation workflow; do not claim this image-generation skill was followed.

## Required Slide Contract

Every generated slide must include:

- Top-left title in the same size and position on every slide.
- Red underline beneath the title.
- Top-right Huawei logo plus `HUAWEI` wordmark. If no official logo asset is supplied, note that the generated mark is illustrative and should be replaced for production distribution.
- Bottom-left red outlined label: `内部交流`.
- Bottom-right page number: `当前页码/总页码`.
- Bottom red-bordered summary strip with a target icon and a sentence starting with `一句话总结：`.
- White background, pale circuit texture, Huawei-red accents, deep-blue technical line art, clean corporate engineering layout.

## Image Generation Rules

- Use the built-in `image_gen` tool for every final slide image. This is required for this skill, not optional.
- Use maximum available parallelism for the initial slide image pass. Prepare all prompts first, then issue concurrent `image_gen` calls in batches up to the platform's current parallel tool-call limit. If the runtime serializes calls internally, continue with the available behavior, but do not add an avoidable slide-by-slide wait loop.
- Keep slide text short and large. Dense Chinese text often degrades in generated images.
- Prefer diagrams, chips, flowcharts, spectra, roadmaps, KPI cards, and matrices over paragraphs.
- Generate exactly one final image per slide; do not use variants unless a page visibly fails QA.
- After each batch completes, copy outputs into the workspace. Do not leave final assets only under `$CODEX_HOME/generated_images`.
- Keep a lineage ledger. Each final `slide-NN.png` must correspond to one copied `image_gen` output path.
- Keep slide numbering deterministic under parallel generation: prompts, copied filenames, ledger rows, and page numbers must all use the locked outline order, not completion order.

## PPTX Assembly

Use the bundled script after images are named sequentially:

```powershell
& "<node.exe>" "<skill>/scripts/build_image_deck.mjs" `
  --images-dir "C:\path\to\slide-images" `
  --out "C:\path\to\deck.pptx" `
  --expected-count 18
```

The script creates a 1280 x 720 deck, inserts each image at `left=0, top=0, width=1280, height=720`, uses `fit=cover`, exports final previews, and writes QA artifacts next to the PPTX.

If the image aspect ratio is not 16:9, inspect the contact sheet carefully. Use `fit=cover` to avoid borders; regenerate misframed source images instead of changing the PPTX to contain mode.

## QA Checklist

- Source document has been summarized in Chinese before slide generation.
- Outline has a cover page and is `<= 20` slides.
- Each slide image contains title, logo area, `内部交流`, page number, and `一句话总结：`.
- All slide images are sequentially named and non-empty.
- `imagegen-ledger.txt` exists and has exactly `N` entries mapping generated source paths to `slide-NN.png`.
- No final slide image was produced by PIL, Python drawing, HTML/CSS screenshots, SVG, canvas, PowerPoint native rendering, or other local drawing pipelines.
- PPTX contains exactly `N` slides and `N` embedded media files.
- Final imported render/contact sheet shows no blank pages, no unintended borders, and no clipped page chrome.

## Common Mistakes

| Mistake | Fix |
| --- | --- |
| Replacing `image_gen` with local drawing to stabilize Chinese text | Regenerate with shorter text; if still unacceptable, report the limitation. |
| Generating too much text per image | Rewrite as KPI cards, bullets, and diagrams before regenerating. |
| Page count changes mid-process | Lock the outline first, then use the same total page count in every prompt. |
| Generating slide images one by one after all prompts are ready | Launch the initial pass as parallel `image_gen` batches up to the platform limit, then repair only failed slides. |
| Relying on generated tiny legal/reference text | Keep citations in notes or appendix text outside generated bitmaps if needed. |
| Leaving images in default generated folder | Copy final PNGs into the project before assembling the PPTX. |
| Using `contain` in PPTX assembly | Use `cover` for full bleed; fix crop by regenerating the slide image. |
