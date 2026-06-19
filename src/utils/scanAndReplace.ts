/* global Word */
/* eslint-disable no-useless-escape */
/* eslint-disable prettier/prettier */
/**
 * @file: scanAndReplace.ts
 * @description: TeX2WordAddin 整页扫描与一键自动化转换逻辑 (Scheme A)
 */

import { MathConversionEngine } from "./MathConversionEngine";

export interface ScanProgress {
  total: number;
  current: number;
  batchIndex: number;
  batchTotal: number;
  status: string;
}

export interface ScannedFormula {
  id: string;
  type: "display" | "inline";
  rawText: string;
  latex: string;
  isDisplay: boolean;
  occurrenceIndex: number; // 0-based index of this specific rawText in the document
}

function isLikelyLatex(text: string, isDisplay: boolean): boolean {
  if (isDisplay) return true;

  let inner = text.trim();
  if (inner.startsWith("$") && inner.endsWith("$")) {
    inner = inner.substring(1, inner.length - 1).trim();
  } else if (inner.startsWith("\\(") && inner.endsWith("\\)")) {
    inner = inner.substring(2, inner.length - 2).trim();
  }

  if (!inner) return false;

  // Filter out common english sentences matched by currency ranges, e.g. "$10 to $20"
  if (/\b(and|or|to|in|for|with|the|a|of|is|are)\b/i.test(inner)) {
    return false;
  }

  // Filter out plain numbers or currency values, e.g. "$100" or "$1,000.50"
  if (/^[0-9\s.,]+$/.test(inner)) {
    return false;
  }

  return true;
}

/**
 * Scans the document for LaTeX formulas and returns them as a structured list.
 */
export async function scanDocument(
  context: Word.RequestContext
): Promise<ScannedFormula[]> {
  // Search for display formulas using properly escaped dollar signs and brackets
  const displayRanges = context.document.body.search("\\$\\$[!^13]@\\$\\$", { matchWildcards: true });
  const displayBracketRanges = context.document.body.search("\\\\\\\\[[!^13]@\\\\\\\]", { matchWildcards: true });

  // Search for inline formulas using properly escaped dollar signs and paren delimiters
  const inlineDollarRanges = context.document.body.search("\\$[!^13]@\\$", { matchWildcards: true });
  const inlineParenRanges = context.document.body.search("\\\\\\([!^13]@\\\\\\)", { matchWildcards: true });

  displayRanges.load("items");
  displayBracketRanges.load("items");
  inlineDollarRanges.load("items");
  inlineParenRanges.load("items");
  
  await context.sync();

  const scannedItems: ScannedFormula[] = [];
  const occurrenceCounts: { [rawText: string]: number } = {};

  const addItem = (itemRange: Word.Range, isDisplay: boolean, isBracketOrParen: boolean, type: "display" | "inline") => {
    const rawText = itemRange.text;
    if (occurrenceCounts[rawText] === undefined) {
      occurrenceCounts[rawText] = 0;
    } else {
      occurrenceCounts[rawText]++;
    }

    let latex = rawText.trim();
    if (isDisplay) {
      if (isBracketOrParen) {
        latex = latex.replace(/^\\\[/g, "").replace(/\\\]$/g, "").trim();
      } else {
        latex = latex.replace(/^\$\$/g, "").replace(/\$\$$/g, "").trim();
      }
    } else {
      if (isBracketOrParen) {
        latex = latex.replace(/^\\\(/g, "").replace(/\\\)$/g, "").trim();
      } else {
        latex = latex.replace(/^\$/g, "").replace(/\$$/g, "").trim();
      }
    }

    // Generate unique ID starting with t2w- prefix
    const id = `t2w-${type}-${scannedItems.length}-${Date.now()}`;
    
    // Insert temporary Content Control to track this range
    const cc = itemRange.insertContentControl();
    cc.tag = id;
    cc.title = "LaTeX Formula";

    scannedItems.push({
      id,
      type,
      rawText,
      latex,
      isDisplay,
      occurrenceIndex: occurrenceCounts[rawText]
    });
  };

  // Compare location of inline dollar ranges with display ranges to filter out duplicates
  const inlineComparisons: {
    item: Word.Range;
    relations: any[];
  }[] = [];

  for (const inlineItem of inlineDollarRanges.items) {
    const relations: any[] = [];
    for (const displayItem of displayRanges.items) {
      relations.push(inlineItem.compareLocationWith(displayItem));
    }
    for (const displayBracketItem of displayBracketRanges.items) {
      relations.push(inlineItem.compareLocationWith(displayBracketItem));
    }
    inlineComparisons.push({ item: inlineItem, relations });
  }

  // Execute single sync to resolve all queued location comparisons in a single batch
  await context.sync();

  // Add display items
  for (const item of displayRanges.items) {
    addItem(item, true, false, "display");
  }
  for (const item of displayBracketRanges.items) {
    addItem(item, true, true, "display");
  }

  // Add inline items, ensuring they don't lie inside display formulas
  for (const comp of inlineComparisons) {
    const isInsideOrEqual = comp.relations.some(r => {
      const val = r.value ? r.value.toLowerCase() : "";
      return val === "inside" || val === "insidestart" || val === "insideend" || val === "equal";
    });
    if (!isInsideOrEqual && isLikelyLatex(comp.item.text, false)) {
      addItem(comp.item, false, false, "inline");
    }
  }
  for (const item of inlineParenRanges.items) {
    if (isLikelyLatex(item.text, false)) {
      addItem(item, false, true, "inline");
    }
  }

  return scannedItems;
}

/**
 * Converts a single scanned formula item in the document.
 */
export async function convertSingleFormula(
  engine: MathConversionEngine,
  item: ScannedFormula
): Promise<void> {
  await Word.run(async (context) => {
    // 1. Try to find the range using Content Control tag
    const ccs = context.document.contentControls.getByTag(item.id);
    /* eslint-disable-next-line office-addins/no-navigational-load */
    ccs.load("items/length");
    await context.sync();

    let targetRange: Word.Range | null = null;
    if (ccs.items.length > 0) {
      targetRange = ccs.items[0].getRange();
    } else {
      // Fallback: search by text, limiting length to 250 characters to prevent SearchStringInvalidOrTooLong
      const truncatedText = item.rawText.substring(0, 250);
      const escapedText = truncatedText.replace(/\\/g, "\\\\");
      const searchResults = context.document.body.search(escapedText, { matchWildcards: false });
      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length > item.occurrenceIndex) {
        targetRange = searchResults.items[item.occurrenceIndex];
      }
    }

    if (targetRange) {
      const oxml = engine.transformTexToOxml(item.latex, item.isDisplay);
      targetRange.insertOoxml(oxml, Word.InsertLocation.replace);
      await context.sync();
    } else {
      throw new Error(`Formula range not found in document.`);
    }
  });
}

/**
 * Converts all scanned formulas at once in a single transaction.
 */
export async function convertAllScannedFormulas(
  engine: MathConversionEngine,
  items: ScannedFormula[],
  onProgress: (current: number, total: number) => void
): Promise<void> {
  await Word.run(async (context) => {
    // 1. Load all content controls
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    // Map tag to content control
    const ccMap: { [tag: string]: Word.ContentControl } = {};
    for (const cc of ccs.items) {
      if (cc.tag) {
        ccMap[cc.tag] = cc;
      }
    }

    // 2. Load fallback searches for missing controls
    const fallbackItems = items.filter(item => !ccMap[item.id]);
    const uniqueRawTexts = Array.from(new Set(fallbackItems.map(item => item.rawText)));
    const searchMap: { [rawText: string]: Word.RangeCollection } = {};

    for (const rawText of uniqueRawTexts) {
      const truncatedText = rawText.substring(0, 250);
      const escapedText = truncatedText.replace(/\\/g, "\\\\");
      searchMap[rawText] = context.document.body.search(escapedText, { matchWildcards: false });
      searchMap[rawText].load("items");
    }

    if (uniqueRawTexts.length > 0) {
      await context.sync();
    }

    // 3. Process replacements sequentially
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      let targetRange: Word.Range | null = null;

      const cc = ccMap[item.id];
      if (cc) {
        targetRange = cc.getRange();
      } else {
        const searchResults = searchMap[item.rawText];
        if (searchResults && searchResults.items.length > item.occurrenceIndex) {
          targetRange = searchResults.items[item.occurrenceIndex];
        }
      }

      if (targetRange) {
        const oxml = engine.transformTexToOxml(item.latex, item.isDisplay);
        targetRange.insertOoxml(oxml, Word.InsertLocation.replace);
      }
      onProgress(i + 1, items.length);
    }
    await context.sync();
  });
}

/**
 * Deletes all temporary LaTeX formula content controls from the document while keeping their text contents.
 */
export async function clearFormulaContentControls(): Promise<void> {
  await Word.run(async (context) => {
    const ccs = context.document.contentControls;
    ccs.load("items/tag");
    await context.sync();

    for (const cc of ccs.items) {
      if (cc.tag && cc.tag.startsWith("t2w-")) {
        cc.cannotDelete = false;
        cc.delete(true); // true = keep contents
      }
    }
    await context.sync();
  });
}

/**
 * Legacy wrapper: Scans the document for LaTeX formulas, converts them to OMML, and inserts them in-place.
 */
export async function scanAndReplaceDocument(
  engine: MathConversionEngine,
  onProgress: (progress: ScanProgress) => void
): Promise<void> {
  engine.clearLogs();

  await Word.run(async (context) => {
    onProgress({
      total: 100,
      current: 0,
      batchIndex: 1,
      batchTotal: 2,
      status: "Searching document for display and inline formulas..."
    });

    await clearFormulaContentControls();
    const items = await scanDocument(context);
    if (items.length === 0) {
      onProgress({
        total: 0,
        current: 0,
        batchIndex: 2,
        batchTotal: 2,
        status: "Scan completed. No formulas found."
      });
      return;
    }

    onProgress({
      total: items.length,
      current: 0,
      batchIndex: 2,
      batchTotal: 2,
      status: `Converting ${items.length} formulas...`
    });

    await convertAllScannedFormulas(engine, items, (current, total) => {
      onProgress({
        total,
        current,
        batchIndex: 2,
        batchTotal: 2,
        status: `Converting formula ${current}/${total}...`
      });
    });

    onProgress({
      total: items.length,
      current: items.length,
      batchIndex: 2,
      batchTotal: 2,
      status: `Scan and Replace completed. Total converted: ${items.length} formulas.`
    });
  });
}
