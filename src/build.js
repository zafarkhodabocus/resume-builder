#!/usr/bin/env node
/**
 * resume-builder
 * Generates a clean, ATS-friendly .docx resume from a JSON config.
 *
 * Usage:
 *   node src/build.js path/to/resume.config.json [output.docx]
 *
 * The config format is documented in README.md and example/resume.config.json.
 */

const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, BorderStyle,
  AlignmentType, LevelFormat, convertInchesToTwip,
} = require("docx");

function buildDocument(config) {
  const theme = Object.assign(
    { accent: "1F3864", muted: "444444", font: "Calibri" },
    config.theme || {}
  );

  const sectionHeading = (text) => new Paragraph({
    spacing: { before: 240, after: 100 },
    border: { bottom: { color: theme.accent, space: 2, style: BorderStyle.SINGLE, size: 6 } },
    children: [new TextRun({ text, bold: true, color: theme.accent, size: 22, font: theme.font })],
  });

  const bullet = (text) => new Paragraph({
    numbering: { reference: "main-bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: 20, font: theme.font })],
  });

  // Supports a leading "**bold label**" segment, e.g. "**Stack:** Python, Node.js"
  const bulletRich = (text) => {
    const match = /^\*\*(.+?)\*\*(.*)$/.exec(text);
    if (!match) return bullet(text);
    const [, boldPart, rest] = match;
    return new Paragraph({
      numbering: { reference: "main-bullets", level: 0 },
      spacing: { after: 60 },
      children: [
        new TextRun({ text: boldPart, bold: true, size: 20, font: theme.font }),
        new TextRun({ text: rest, size: 20, font: theme.font }),
      ],
    });
  };

  const jobHeader = (title) => new Paragraph({
    spacing: { before: 180, after: 20 },
    children: [new TextRun({ text: title, bold: true, size: 21, font: theme.font })],
  });

  const jobDates = (dates) => new Paragraph({
    spacing: { after: 20 },
    children: [new TextRun({ text: dates, italics: true, size: 19, font: theme.font })],
  });

  const jobContext = (text) => new Paragraph({
    spacing: { after: 80 },
    children: [new TextRun({ text, italics: true, size: 19, color: theme.muted, font: theme.font })],
  });

  const children = [];

  // Header block
  children.push(new Paragraph({
    spacing: { after: 20 },
    children: [new TextRun({ text: config.name || "", bold: true, size: 34, color: theme.accent, font: theme.font })],
  }));
  if (config.title) {
    children.push(new Paragraph({
      spacing: { after: 20 },
      children: [new TextRun({ text: config.title, size: 22, font: theme.font })],
    }));
  }
  if (config.tagline) {
    children.push(new Paragraph({
      spacing: { after: 20 },
      children: [new TextRun({ text: config.tagline, size: 17, color: theme.muted, font: theme.font })],
    }));
  }
  if (config.contact) {
    children.push(new Paragraph({
      spacing: { after: 0 },
      children: [new TextRun({ text: config.contact.join("  •  "), size: 17, color: theme.muted, font: theme.font })],
    }));
  }

  // Arbitrary ordered sections
  for (const section of config.sections || []) {
    children.push(sectionHeading(section.heading));

    if (section.type === "bullets") {
      for (const item of section.items || []) children.push(bulletRich(item));
    }

    if (section.type === "experience") {
      for (const job of section.jobs || []) {
        children.push(jobHeader(job.title));
        if (job.dates) children.push(jobDates(job.dates));
        if (job.context) children.push(jobContext(job.context));
        for (const item of job.bullets || []) children.push(bullet(item));
      }
    }

    if (section.type === "text") {
      for (const line of section.lines || []) {
        children.push(new Paragraph({
          spacing: { after: 10 },
          children: [new TextRun({ text: line.text, bold: !!line.bold, size: line.size || 20, font: theme.font })],
        }));
      }
    }
  }

  return new Document({
    numbering: {
      config: [{
        reference: "main-bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: convertInchesToTwip(0.25), hanging: convertInchesToTwip(0.15) } } },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 720, bottom: 720, left: 810, right: 810 },
        },
      },
      children,
    }],
  });
}

async function main() {
  const [, , configPath, outPath] = process.argv;
  if (!configPath) {
    console.error("Usage: node src/build.js <config.json> [output.docx]");
    process.exit(1);
  }

  const config = JSON.parse(fs.readFileSync(configPath, "utf8"));
  const doc = buildDocument(config);
  const buf = await Packer.toBuffer(doc);

  const out = outPath || `${(config.name || "resume").replace(/\s+/g, "_")}.docx`;
  fs.writeFileSync(out, buf);
  console.log(`Wrote ${path.resolve(out)}`);
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { buildDocument };
