/* global window, document, console, fetch, XSLTProcessor, DOMParser, XMLSerializer, Word */
/* eslint-disable prettier/prettier */
/**
 * @file: MathConversionEngine.ts
 * @description: TeX2WordAddin 核心公式转换与可观测性管道控制类
 */

// Helper to convert TeX to MathML string using the global MathJax object loaded via script tag
export const MathJaxShim = {
  tex2mml(latexCode: string): string {
    const globalMathJax = (window as any).MathJax;
    if (!globalMathJax || !globalMathJax.tex2mml) {
      throw new Error("MathJax has not finished loading. Please wait.");
    }
    return globalMathJax.tex2mml(latexCode, { display: true });
  },
};

/**
 * Helper to prefix math elements in XSLT output to be strictly compatible with Word's OOXML schema
 */
export function prefixMathElements(xml: string): string {
  // Replace xmlns="...math" with xmlns:m="...math" using robust regex
  let result = xml.replace(
    /xmlns\s*=\s*['"]http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/math['"]/gi,
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
  );
  // Add m: prefix to all start tags, matching tags with optional hyphens/underscores/colons
  result = result.replace(/<([a-zA-Z0-9:\-_]+)([^>]*?)>/g, (match, tagName, attrs) => {
    if (tagName.includes(":") || tagName === "xml") {
      return match;
    }
    return `<m:${tagName}${attrs}>`;
  });
  // Add m: prefix to all end tags
  result = result.replace(/<\/([a-zA-Z0-9:\-_]+)>/g, (match, tagName) => {
    if (tagName.includes(":")) {
      return match;
    }
    return `</m:${tagName}>`;
  });
  return result;
}

export class MathConversionEngine {
  private loggerName = "TeX2WordAddin.Engine";
  private logs: string[] = [];
  private onLogCallback: ((log: string) => void) | null = null;

  public static xsltStylesheetText: string = "";

  public setOnLog(callback: (log: string) => void) {
    this.onLogCallback = callback;
  }

  private log(level: "INFO" | "DEBUG" | "WARN" | "ERROR", message: string, details?: any) {
    const timestamp = new Date().toISOString();
    const detailStr = details ? " " + JSON.stringify(details) : "";
    const logLine = `[${timestamp}] [${level}] [${this.loggerName}] ${message}${detailStr}`;
    console.log(logLine);
    this.logs.push(logLine);
    if (this.onLogCallback) {
      this.onLogCallback(logLine);
    }
  }

  public getLogs(): string[] {
    return this.logs;
  }

  public clearLogs() {
    this.logs = [];
  }

  public static async loadStylesheet(): Promise<void> {
    if (this.xsltStylesheetText) return;
    const response = await fetch("/assets/mml2omml.xsl");
    if (!response.ok) {
      throw new Error(`Failed to load XSLT stylesheet from server: ${response.statusText}`);
    }
    this.xsltStylesheetText = await response.text();
  }

  /**
   * MathML to OMML using XSLT
   */
  private convertMathmlToOmmlViaXslt(mathmlString: string): string {
    this.log("DEBUG", "状态 2: 正在执行 MML2OMML XSLT 转换结构映射...");
    try {
      const xsltProcessor = new XSLTProcessor();

      if (!MathConversionEngine.xsltStylesheetText) {
        throw new Error("XSLT stylesheet has not been loaded. Call MathConversionEngine.loadStylesheet() first.");
      }

      const parser = new DOMParser();
      const serializer = new XMLSerializer();

      const xslDoc = parser.parseFromString(MathConversionEngine.xsltStylesheetText, "text/xml");
      const xslStr = serializer.serializeToString(xslDoc);
      if (xslStr.includes("parsererror")) {
        throw new Error("XSLT Stylesheet XML Parser Error: " + xslStr);
      }

      xsltProcessor.importStylesheet(xslDoc);

      const xmlDoc = parser.parseFromString(mathmlString, "text/xml");
      const xmlStr = serializer.serializeToString(xmlDoc);
      if (xmlStr.includes("parsererror")) {
        throw new Error("MathML XML Parser Error: " + xmlStr);
      }

      let ommlSnippet = "";
      let transformError: any = null;

      try {
        const resultDoc = xsltProcessor.transformToDocument(xmlDoc);
        if (resultDoc) {
          const serialized = serializer.serializeToString(resultDoc);
          if (!serialized.includes("parsererror") && serialized.trim() !== "") {
            ommlSnippet = serialized;
          } else if (serialized.includes("parsererror")) {
            transformError = new Error("transformToDocument returned parsererror: " + serialized);
          }
        }
      } catch (e) {
        transformError = e;
      }

      // Fallback to transformToFragment using an XML Document as owner document
      if (!ommlSnippet) {
        this.log("WARN", "transformToDocument 失败或返回空值，正在尝试 transformToFragment 回退方案...", { error: transformError });
        try {
          const xmlOwnerDoc = document.implementation.createDocument(null, "root", null);
          const fragment = xsltProcessor.transformToFragment(xmlDoc, xmlOwnerDoc);
          if (fragment) {
            const serialized = serializer.serializeToString(fragment);
            if (!serialized.includes("parsererror") && serialized.trim() !== "") {
              ommlSnippet = serialized;
            } else if (serialized.includes("parsererror")) {
              throw new Error("transformToFragment returned parsererror: " + serialized);
            }
          }
        } catch (fallbackErr) {
          this.log("ERROR", "transformToFragment 回退方案也发生错误", { error: fallbackErr });
          throw transformError || fallbackErr;
        }
      }

      if (!ommlSnippet || ommlSnippet.trim() === "") {
        throw new Error("XSLT transformation returned empty result in both transformToDocument and transformToFragment.");
      }

      // Clean up XML namespaces or declaration added by the transform
      ommlSnippet = ommlSnippet.replace(/<\?xml[\s\S]*?\?>/g, "").trim();

      this.log("DEBUG", "状态 2 成功: OMML 转换映射完毕", { omml: ommlSnippet });
      return ommlSnippet;
    } catch (e) {
      this.log("ERROR", "XSLT 转换结构映射发生错误", e);
      throw e;
    }
  }

  /**
   * 核心转换入口：将 LaTeX 转换为完整的 Word OXML 字符串
   * @param latexCode 传入的纯 LaTeX 字符串
   * @param isDisplay 是否为独立段落公式 (true) 还是行内公式 (false)
   */
  public transformTexToOxml(latexCode: string, isDisplay: boolean = true): string {
    const optimized = optimizeLatexFormula(latexCode);
    this.log("INFO", "开始执行 LaTeX 转换管道", { original: latexCode, optimized, isDisplay });
    try {
      // 1. 调用 MathJax 生成 standard MathML
      this.log("DEBUG", "状态 1: 正在调用 MathJax.tex2mml...");
      const mathml = MathJaxShim.tex2mml(optimized);
      this.log("DEBUG", "状态 1 成功: MathML 字符生成完毕", { mathml: mathml });

      // 2. 将 MathML 通过 XSLT 脚本转换为 Word 的 OMML 格式
      const ommlSnippet = this.convertMathmlToOmmlViaXslt(mathml);

      // Prefix math elements to comply with Word's strict OOXML parser
      const prefixedOmml = prefixMathElements(ommlSnippet);
      this.log("DEBUG", "状态 2.5: 前缀化后的 OMML 字符", { prefixedOmml: prefixedOmml });

      // 3. 构建符合 Word Open XML 命名规范的标准外层包裹
      this.log("DEBUG", "状态 3: 正在进行 Open XML 命名空间与段落节点封装...");

      const innerContent = isDisplay 
        ? `<w:p><m:oMathPara>${prefixedOmml}</m:oMathPara></w:p>` 
        : `<w:p>${prefixedOmml}</w:p>`;

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
          ${innerContent}
        </w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`.trim();

      this.log("INFO", "转换管道执行完毕，成功生成高质量 OXML Payload", { fullOxml: fullOxml });
      return fullOxml;
    } catch (error) {
      this.log("ERROR", "转换管道核心步骤发生崩溃", { error });
      throw error;
    }
  }

  /**
   * 驱动 Word API 执行物理文档插入事务
   */
  public async insertFormulaIntoWord(latexCode: string, isDisplay: boolean = true): Promise<void> {
    this.log("INFO", "准备调用 Office.js 驱动物理文档事务");
    await Word.run(async (context) => {
      const oxmlPayload = this.transformTexToOxml(latexCode, isDisplay);
      const selection = context.document.getSelection();
      this.log("DEBUG", "正在执行选区物理替换 (insertOoxml)...");
      selection.insertOoxml(oxmlPayload, Word.InsertLocation.replace);
      await context.sync();
      this.log("INFO", "Office.js 事务同步提交成功，Word 原生公式渲染就绪");
    });
  }
}

/**
 * LaTeX helper to detect normal parentheses, brackets, vertical bars wrapping tall structures,
 * and automatically convert them into stretchy delimiters (\left and \right) to adapt Word height.
 */
export function convertToStretchyDelimiters(latex: string): string {
  let result = latex;
  const hasTall = (str: string) => {
    return str.includes("\\frac") || str.includes("\\sum") || str.includes("\\int") || str.includes("\\prod") || str.includes("\\partial") || str.includes("/");
  };

  let modified = "";
  let i = 0;
  while (i < result.length) {
    // 1. Parentheses ( ... )
    if (result[i] === '(' && (i === 0 || !result.substring(0, i).trim().endsWith("\\left"))) {
      let depth = 1;
      let j = i + 1;
      while (j < result.length) {
        if (result[j] === '(') depth++;
        if (result[j] === ')') depth--;
        if (depth === 0) break;
        j++;
      }
      if (j < result.length) {
        const inner = result.substring(i + 1, j);
        if (hasTall(inner)) {
          modified += "\\left(" + convertToStretchyDelimiters(inner) + "\\right)";
          i = j + 1;
          continue;
        }
      }
    }
    
    // 2. Square brackets [ ... ]
    if (result[i] === '[' && (i === 0 || !result.substring(0, i).trim().endsWith("\\left"))) {
      let depth = 1;
      let j = i + 1;
      while (j < result.length) {
        if (result[j] === '[') depth++;
        if (result[j] === ']') depth--;
        if (depth === 0) break;
        j++;
      }
      if (j < result.length) {
        const inner = result.substring(i + 1, j);
        if (hasTall(inner)) {
          modified += "\\left[" + convertToStretchyDelimiters(inner) + "\\right]";
          i = j + 1;
          continue;
        }
      }
    }

    // 3. Norm double vertical bars \| ... \|
    if (result.substring(i, i + 2) === "\\|" && !result.substring(0, i).trim().endsWith("\\left")) {
      let j = i + 2;
      while (j < result.length - 1) {
        if (result.substring(j, j + 2) === "\\|") break;
        j++;
      }
      if (j < result.length - 1) {
        const inner = result.substring(i + 2, j);
        if (hasTall(inner)) {
          modified += "\\left\\|" + convertToStretchyDelimiters(inner) + "\\right\\|";
          i = j + 2;
          continue;
        }
      }
    }
    
    // 4. Absolute value single vertical bars | ... |
    if (result[i] === '|' && (i === 0 || !result.substring(0, i).trim().endsWith("\\left")) && result[i - 1] !== '\\') {
      let j = i + 1;
      while (j < result.length) {
        if (result[j] === '|' && result[j - 1] !== '\\') break;
        j++;
      }
      if (j < result.length) {
        const inner = result.substring(i + 1, j);
        if (hasTall(inner)) {
          modified += "\\left|" + convertToStretchyDelimiters(inner) + "\\right|";
          i = j + 1;
          continue;
        }
      }
    }

    modified += result[i];
    i++;
  }
  return modified;
}

/**
 * LaTeX helper to automatically wrap the argument/body of summation/product/integrals in {...}
 * to prevent Word from rendering empty operand dotted boxes.
 */
export function wrapNaryArguments(latex: string): string {
  let result = latex;
  // Match n-ary operators like \sum, \prod, \int with optional limit scripts
  // and ensure any unwrapped remaining expression is grouped in {...}
  result = result.replace(/(\\(sum|prod|int|coprod|bigcup|bigcap|oint|iint|iiint)(?:_[a-zA-Z0-9]+|_\{[^{}]*\}|\^[a-zA-Z0-9]+|\^\{[^{}]*\})*)\s*([^{}\s][^$]*)$/g, (match, op, _name, rest) => {
    if (rest && rest.trim()) {
      return `${op} {${rest.trim()}}`;
    }
    return match;
  });
  return result;
}

/**
 * Consolidates all formula structural optimization adjustments.
 */
export function optimizeLatexFormula(latex: string): string {
  let optimized = latex.trim();
  optimized = convertToStretchyDelimiters(optimized);
  optimized = wrapNaryArguments(optimized);
  return optimized;
}

