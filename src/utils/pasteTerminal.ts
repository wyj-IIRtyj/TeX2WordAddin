/* global Word */
/* eslint-disable prettier/prettier */
/**
 * @file: pasteTerminal.ts
 * @description: TeX2WordAddin 智能粘贴中转区逻辑 (Scheme B)
 */

import { MathConversionEngine, MathJaxShim, prefixMathElements } from "./MathConversionEngine";

/**
 * Escapes special XML characters in a string.
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Parses mixed text (mixed plain text and LaTeX math) into segments and generates a Word OOXML package.
 */
export function generateOxmlFromMixedText(engine: MathConversionEngine, text: string): string {
  // Split into lines (paragraphs)
  const lines = text.split(/\r?\n/);
  const paragraphXmls: string[] = [];

  // Match $$...$$, \[...\], \(...\), literal \$, or $...$
  const segmentRegex = /(\$\$[\s\S]*?\$\$|\\\[[\s\S]*?\\\]|\\\([\s\S]*?\\\)|\\\$|\$[\s\S]*?\$)/g;

  for (const line of lines) {
    if (!line.trim()) {
      // Empty paragraph
      paragraphXmls.push(`<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`);
      continue;
    }

    const segments = line.split(segmentRegex);
    const runsXml: string[] = [];

    for (let i = 0; i < segments.length; i++) {
      const segment = segments[i];
      if (!segment) continue;

      if (segment === "\\$") {
        // Literal escaped dollar sign
        runsXml.push(`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve">$</w:t></w:r>`);
        continue;
      }

      // Reset regex index
      segmentRegex.lastIndex = 0;
      if (segmentRegex.test(segment)) {
        // This is a math segment
        let latex = segment.trim();
        const isDisplay = latex.startsWith("$$") || latex.startsWith("\\[");

        if (latex.startsWith("$$")) {
          latex = latex.replace(/^\$\$/g, "").replace(/\$\$$/g, "").trim();
        } else if (latex.startsWith("\\[")) {
          latex = latex.replace(/^\\\[/g, "").replace(/\\\]$/g, "").trim();
        } else if (latex.startsWith("\\(")) {
          latex = latex.replace(/^\\\(/g, "").replace(/\\\)$/g, "").trim();
        } else if (latex.startsWith("$")) {
          latex = latex.replace(/^\$/g, "").replace(/\$$/g, "").trim();
        }

        try {
          // Convert LaTeX to MathML
          const mathml = MathJaxShim.tex2mml(latex);
          
          // Convert MathML to OMML snippet
          const ommlSnippet = (engine as any).convertMathmlToOmmlViaXslt(mathml);
          
          // Prefix the math elements
          const prefixedOmml = prefixMathElements(ommlSnippet);
          
          // If display math inside a paragraph, wrap in m:oMathPara to center
          const mathOutput = isDisplay ? `<m:oMathPara>${prefixedOmml}</m:oMathPara>` : prefixedOmml;
          
          runsXml.push(mathOutput);
        } catch {
          // If conversion fails, render as plain text error run in red
          runsXml.push(`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">[Math Error: ${escapeXml(latex)}]</w:t></w:r>`);
        }
      } else {
        // This is a plain text segment
        runsXml.push(`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t xml:space="preserve">${escapeXml(segment)}</w:t></w:r>`);
      }
    }

    paragraphXmls.push(`
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        ${runsXml.join("\n")}
      </w:p>
    `);
  }

  // Wrap everything in a Flat OPC Package
  const fullOxml = `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
        <w:body>
          ${paragraphXmls.join("\n")}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`.trim();

  return fullOxml;
}

/**
 * Parses mixed text pasted in the terminal, converts math, and inserts it at the current selection.
 */
export async function pasteTerminalMixedText(engine: MathConversionEngine, text: string): Promise<void> {
  await Word.run(async (context) => {
    const oxmlPayload = generateOxmlFromMixedText(engine, text);
    const selection = context.document.getSelection();
    selection.insertOoxml(oxmlPayload, Word.InsertLocation.replace);
    await context.sync();
  });
}
