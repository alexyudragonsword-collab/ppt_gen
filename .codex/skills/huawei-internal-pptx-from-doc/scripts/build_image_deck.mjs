import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

function parseArgs(argv) {
  const args = {
    width: 1280,
    height: 720,
    fit: "cover",
    expectedCount: null,
    nodeModules: process.env.NODE_MODULES_DIR ?? null,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const key = argv[i];
    const value = argv[i + 1];
    if (!key.startsWith("--")) continue;
    i += 1;
    switch (key) {
      case "--images-dir":
        args.imagesDir = value;
        break;
      case "--out":
        args.out = value;
        break;
      case "--expected-count":
        args.expectedCount = Number(value);
        break;
      case "--width":
        args.width = Number(value);
        break;
      case "--height":
        args.height = Number(value);
        break;
      case "--fit":
        args.fit = value;
        break;
      case "--node-modules":
        args.nodeModules = value;
        break;
      default:
        throw new Error(`Unknown argument: ${key}`);
    }
  }
  if (!args.imagesDir || !args.out) {
    throw new Error("Usage: build_image_deck.mjs --images-dir <dir> --out <deck.pptx> [--expected-count N]");
  }
  if (!["cover", "contain"].includes(args.fit)) {
    throw new Error("--fit must be cover or contain");
  }
  return args;
}

async function loadArtifactTool(nodeModules) {
  try {
    return await import("@oai/artifact-tool");
  } catch {
    // Fall through to explicit runtime paths.
  }

  const candidates = [
    nodeModules,
    process.env.CODEX_NODE_MODULES,
    process.env.NODE_REPL_NODE_MODULE_DIRS?.split(path.delimiter)[0],
    path.join(process.cwd(), "node_modules"),
    "C:\\Users\\buadmin\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules",
  ].filter(Boolean);

  for (const candidate of candidates) {
    const artifactToolPath = path.join(
      candidate,
      "@oai",
      "artifact-tool",
      "dist",
      "artifact_tool.mjs",
    );
    try {
      await fs.access(artifactToolPath);
      return await import(pathToFileURL(artifactToolPath).href);
    } catch {
      // Try the next candidate.
    }
  }
  throw new Error("Cannot locate @oai/artifact-tool. Pass --node-modules <path> or run inside the Codex primary runtime.");
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function readImageBytes(filePath) {
  const bytes = await fs.readFile(filePath);
  return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
}

async function verifyImagegenLedger(imagesDir, imageNames) {
  const ledgerPath = path.join(imagesDir, "imagegen-ledger.txt");
  let ledgerText;
  try {
    ledgerText = await fs.readFile(ledgerPath, "utf8");
  } catch {
    throw new Error(
      `Missing imagegen-ledger.txt in ${imagesDir}. Each final slide image must be copied from a built-in image_gen output and recorded in the ledger.`,
    );
  }

  const lines = ledgerText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const missing = [];
  for (const imageName of imageNames) {
    const matching = lines.filter((line) => line.includes(imageName));
    if (matching.length === 0) {
      missing.push(`${imageName}: no ledger entry`);
      continue;
    }
    if (!matching.some((line) => /generated_images/i.test(line))) {
      missing.push(`${imageName}: ledger entry does not reference generated_images`);
    }
  }
  if (missing.length > 0) {
    throw new Error(
      [
        "Invalid imagegen-ledger.txt. Final slide images must be traceable to built-in image_gen outputs.",
        ...missing,
      ].join("\n"),
    );
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const { FileBlob, Presentation, PresentationFile } = await loadArtifactTool(args.nodeModules);

  const imageNames = (await fs.readdir(args.imagesDir))
    .filter((name) => /^slide-\d{2,3}\.png$/i.test(name))
    .sort();
  if (args.expectedCount !== null && imageNames.length !== args.expectedCount) {
    throw new Error(`Expected ${args.expectedCount} images, found ${imageNames.length}`);
  }
  if (imageNames.length === 0) {
    throw new Error(`No slide-XX.png images found in ${args.imagesDir}`);
  }
  await verifyImagegenLedger(args.imagesDir, imageNames);

  const outputDir = path.dirname(args.out);
  const stem = path.basename(args.out, path.extname(args.out));
  const qaDir = path.join(outputDir, `${stem}-qa`);
  const previewDir = path.join(qaDir, "final-preview");
  await fs.mkdir(outputDir, { recursive: true });
  await fs.mkdir(previewDir, { recursive: true });

  const presentation = Presentation.create({
    slideSize: { width: args.width, height: args.height },
  });

  for (const [index, imageName] of imageNames.entries()) {
    const slide = presentation.slides.add();
    slide.background.fill = "#ffffff";
    slide.images.add({
      blob: await readImageBytes(path.join(args.imagesDir, imageName)),
      contentType: "image/png",
      alt: `Generated full-slide image ${index + 1}`,
      fit: args.fit,
      position: { left: 0, top: 0, width: args.width, height: args.height },
    });
  }

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(args.out);

  const imported = await PresentationFile.importPptx(await FileBlob.load(args.out));
  const snapshot = await imported.inspect({ kind: "slide,image", maxChars: 12000 });
  await fs.writeFile(path.join(qaDir, "inspect-final.txt"), snapshot.ndjson, "utf8");

  for (const [index, slide] of imported.slides.items.entries()) {
    const png = await imported.export({ slide, format: "png", scale: 1 });
    await writeBlob(
      path.join(previewDir, `final-slide-${String(index + 1).padStart(2, "0")}.png`),
      png,
    );
  }
  const montage = await imported.export({
    format: "webp",
    montage: {
      format: "webp",
      columns: 3,
      slideWidth: 480,
      padding: 12,
      gap: 16,
      background: "#f5f5f5",
    },
  });
  await writeBlob(path.join(qaDir, "final-contact-sheet.webp"), montage);

  await fs.writeFile(
    path.join(qaDir, "visual-qa.txt"),
    [
      `PPTX: ${args.out}`,
      `Images: ${imageNames.length}`,
      `Slide size: ${args.width} x ${args.height}`,
      `Fit: ${args.fit}`,
      "Verify: inspect-final.txt, final-preview/*.png, final-contact-sheet.webp.",
    ].join("\n"),
    "utf8",
  );

  console.log(JSON.stringify({
    pptx: args.out,
    slides: imageNames.length,
    qaDir,
    temp: os.tmpdir(),
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
