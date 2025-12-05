#!/usr/bin/env node

/**
 * Process a PPTX file:
 *   1. Extract notes into <outputPath>/notes.json
 *   2. Convert PPTX to PDF into <outputPath>/slides.pdf
 *
 * Usage:
 *   node convert.js file.pptx [outputPath]
 */

const fs = require("fs");
const { exec } = require("child_process");
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");
const path = require("path");

if (process.argv.length < 3) {
  console.error("Usage: node process-ppt.js file.pptx [outputPath]");
  process.exit(1);
}

const inputPath = process.argv[2];
const outputPath = process.argv[3]
  ? path.resolve(process.argv[3])
  : process.cwd();

// Ensure output directory exists
if (!fs.existsSync(outputPath)) {
  fs.mkdirSync(outputPath, { recursive: true });
}

const notesFile = path.join(outputPath, "notes.json");
const pdfFile = path.join(outputPath, "slides.pdf");

// ---------------------------------------------------------------------------
// Helpers – Extract Notes
// ---------------------------------------------------------------------------

function extractTextFromNode(node) {
  let text = "";
  if (!node) return text;

  function walk(n) {
    if (n.nodeType === 3) {
      text += n.nodeValue;
    }
    if (n.childNodes) {
      for (let i = 0; i < n.childNodes.length; i++) {
        walk(n.childNodes[i]);
      }
    }
  }

  walk(node);
  return text.trim();
}

/**
 * @param {Element} span <a:r> element containing span information
 * @returns {string} formatted span as html
 */
function formatSpan(span) {
  let begin = "";
  let end = "";

  const addTag = (tag) => {
    begin += `<${tag}>`;
    end = `</${tag}>` + end;
  };

  const formatting = span.getElementsByTagName("a:rPr")[0];
  if (formatting) {
    if (parseInt(formatting.getAttribute("b"))) addTag("b");
    if (parseInt(formatting.getAttribute("i"))) addTag("i");
    if (formatting.getAttribute("u") && formatting.getAttribute("u") !== "none")
      addTag("u");
    if (
      formatting.getAttribute("strike") &&
      formatting.getAttribute("strike") !== "noStrike"
    ) {
      addTag("s");
    }
  }

  const textContent = span.getElementsByTagName("a:t")[0]?.textContent;

  return begin + textContent + end;
}

/**
 * @param {Element} line <a:p> element containing line information
 * @returns {string} formatted line as html
 */
function formatLine(line) {
  let str = "";
  const formatting = line.getElementsByTagName("a:pPr")[0];

  if (formatting) {
    const indent = parseInt(formatting.getAttribute("lvl"));
    if (indent) str += "&#9;".repeat(indent) + "&#x2022;";
  }

  const spans = line.getElementsByTagName("a:r");

  for (let i = 0; i < spans.length; ++i) {
    str += " " + formatSpan(spans[i]);
  }

  return "<p>" + str + "</p>";
}

/**
 * @param {JSZip} zip
 * @returns {{slide: number, notes: string}[]}
 */
async function extractNotes(zip) {
  const notes = [];

  for (const filename of Object.keys(zip.files)) {
    if (/notesSlide[0-9]+/i.test(filename)) {
      const xml = await zip.file(filename).async("string");
      const dom = new DOMParser().parseFromString(xml, "text/xml");
      const textContent = dom.getElementsByTagName("p:txBody")[0];
      if (!textContent) continue;

      let slideText = "";

      const lines = textContent.getElementsByTagName("a:p");
      for (let i = 0; i < lines.length; ++i) {
        slideText += formatLine(lines[i]);
      }

      const match = filename.match(/notesSlide(\d+)\.xml$/);
      const slideNumber = match ? parseInt(match[1], 10) : -1;

      notes.push({
        slide: slideNumber,
        notes: slideText.trim(),
      });
    }
  }

  notes.sort((a, b) => a.slide - b.slide);
  return notes;
}

// ---------------------------------------------------------------------------
// LibreOffice PDF Conversion
// ---------------------------------------------------------------------------

function convertToPdf(filePath, outputDir) {
  return new Promise((resolve, reject) => {
    const cmd = `libreoffice --headless --convert-to pdf "${filePath}" --outdir "${outputDir}"`;

    console.log("Converting to PDF...");
    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        return reject(`LibreOffice error: ${error.message}`);
      }
      console.log(stdout);
      resolve();
    });
  });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

(async () => {
  try {
    const data = fs.readFileSync(inputPath);
    const zip = await JSZip.loadAsync(data);

    // Extract Notes
    const notes = await extractNotes(zip);
    fs.writeFileSync(notesFile, JSON.stringify(notes, null, 2), "utf8");
    console.log(`✔ Notes written to ${notesFile}`);

    // Convert to PDF
    await convertToPdf(inputPath, outputPath);

    // LibreOffice outputs something like input.pdf
    const inputBase = path.basename(inputPath, path.extname(inputPath));
    const generatedPdf = path.join(outputPath, `${inputBase}.pdf`);

    if (fs.existsSync(generatedPdf)) {
      fs.renameSync(generatedPdf, pdfFile);
      console.log(`✔ PDF saved as ${pdfFile}`);
    } else {
      console.error("⚠ PDF conversion failed: output file not found.");
    }
  } catch (err) {
    console.error("Error:", err);
  }
})();
