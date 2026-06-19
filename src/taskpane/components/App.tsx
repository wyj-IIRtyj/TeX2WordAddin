import * as React from "react";
import { useState, useEffect } from "react";
import { MathConversionEngine } from "../../utils/MathConversionEngine";
import {
  ScanProgress,
  ScannedFormula,
  scanDocument,
  convertSingleFormula,
  convertAllScannedFormulas,
  clearFormulaContentControls
} from "../../utils/scanAndReplace";
import { pasteTerminalMixedText } from "../../utils/pasteTerminal";

/**
 * Robust copy to clipboard helper supporting older WebKit contexts in Word Desktop.
 */
function copyToClipboard(text: string): boolean {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text);
      return true;
    }
  } catch (e) {
    console.error("navigator.clipboard failed, trying fallback", e);
  }

  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.top = "0";
    textarea.style.left = "0";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    const success = document.execCommand("copy");
    document.body.removeChild(textarea);
    return success;
  } catch (err) {
    console.error("Fallback clipboard copy failed", err);
    return false;
  }
}

// Create a single instance of the conversion engine
const engine = new MathConversionEngine();

interface AppProps {
  title: string;
}

export const App: React.FC<AppProps> = ({ title }) => {
  const [activeTab, setActiveTab] = useState<"scan" | "paste" | "logs">("scan");
  
  // Stylesheet loading state
  const [stylesheetLoaded, setStylesheetLoaded] = useState<boolean>(false);
  const [loadingError, setLoadingError] = useState<string | null>(null);

  // Scheme A (Scan) state
  const [scanProgress, setScanProgress] = useState<ScanProgress | null>(null);
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [scanStats, setScanStats] = useState<{ total: number; processed: number; failed: number } | null>(null);
  const [scannedFormulas, setScannedFormulas] = useState<ScannedFormula[]>([]);
  const [selectedFormulaId, setSelectedFormulaId] = useState<string | null>(null);
  const [isConvertingSingle, setIsConvertingSingle] = useState<boolean>(false);

  // Scheme B (Paste) state
  const [pasteText, setPasteText] = useState<string>("");
  const [isPasting, setIsPasting] = useState<boolean>(false);

  // Logs state
  const [logs, setLogs] = useState<string[]>([]);
  const [selectedLog, setSelectedLog] = useState<string | null>(null);

  // Load the stylesheet on mount
  useEffect(() => {
    async function load() {
      try {
        // Derive the base path from the current page URL so asset paths work
        // under both dev (localhost:3000/) and prod (github.io/TeX2WordAddin/)
        const pagePath = window.location.pathname; // e.g. "/TeX2WordAddin/taskpane.html"
        const basePath = pagePath.substring(0, pagePath.lastIndexOf("/") + 1); // e.g. "/TeX2WordAddin/"

        // Load stylesheet
        await MathConversionEngine.loadStylesheet(basePath);
        
        // Dynamically load MathJax if not present
        if (!(window as any).MathJax) {
          const localFontUrl = window.location.origin + basePath + "assets/output/chtml/fonts/woff-v2";
          // Configure MathJax to load font WOFFs from local server to avoid 404 errors and cross-origin blocking
          (window as any).MathJax = {
            chtml: {
              fontURL: localFontUrl
            },
            startup: {
              ready: () => {
                const mj = (window as any).MathJax;
                if (mj && mj.config && mj.config.chtml) {
                  mj.config.chtml.fontURL = localFontUrl;
                }
                mj.startup.defaultReady();
              }
            }
          };
          await new Promise<void>((resolve, reject) => {
            const script = document.createElement("script");
            script.src = basePath + "assets/tex-mml-chtml.js";
            script.type = "text/javascript";
            script.onload = () => resolve();
            script.onerror = () => reject(new Error("Failed to load MathJax script"));
            document.head.appendChild(script);
          });
        }
        
        setStylesheetLoaded(true);
      } catch (err: any) {
        setLoadingError(err.message || String(err));
      }
    }
    load();
  }, []);

  // Sync engine logs to state
  useEffect(() => {
    engine.setOnLog((logLine) => {
      setLogs((prev) => [...prev, logLine]);
    });
  }, []);

  // Clear log screen
  const handleClearLogs = () => {
    engine.clearLogs();
    setLogs([]);
    setSelectedLog(null);
  };

  // One-click copy only error/warning logs without duplicate patterns
  const handleCopyErrorLogs = () => {
    const errorLogs = logs.filter(
      (line) => line.includes("[ERROR]") || line.includes("[WARN]")
    );
    // Strip timestamp and details to filter duplicates
    const uniqueErrors = Array.from(
      new Set(errorLogs.map((line) => line.replace(/^\[[^\]]+\]\s*/, "")))
    );
    const textToCopy = uniqueErrors.join("\n");
    
    if (textToCopy.trim()) {
      const success = copyToClipboard(textToCopy);
      if (success) {
        alert("关键调试错误日志已成功复制到剪贴板！");
      } else {
        alert("复制失败，请手动在日志列表中查看并复制错误信息。");
      }
    } else {
      alert("当前日志中没有检测到任何调试警告或错误记录。");
    }
  };

  // Run Scheme A - Step 1: Scan and list all formulas in the document
  const handleScanSearch = async () => {
    if (!stylesheetLoaded) return;
    setIsScanning(true);
    setScanStats(null);
    setSelectedFormulaId(null);
    setScanProgress({
      total: 100,
      current: 0,
      batchIndex: 1,
      batchTotal: 1,
      status: "正在扫描检索文档中的 LaTeX 公式..."
    });
    try {
      // Clear any leftover temporary Content Controls first
      await clearFormulaContentControls();

      const formulas = await Word.run(async (context) => {
        return await scanDocument(context);
      });
      setScannedFormulas(formulas);
      if (formulas.length === 0) {
        setScanProgress({
          total: 0,
          current: 0,
          batchIndex: 1,
          batchTotal: 1,
          status: "未在文档中检索到任何 LaTeX 公式。"
        });
      } else {
        setScanProgress({
          total: formulas.length,
          current: 0,
          batchIndex: 1,
          batchTotal: 1,
          status: `检索完成！共发现 ${formulas.length} 个待转换公式。`
        });
      }
    } catch (err: any) {
      console.error(err);
      setScanProgress({
        total: 0,
        current: 0,
        batchIndex: 1,
        batchTotal: 1,
        status: `公式检索失败: ${err.message || String(err)}`
      });
    } finally {
      setIsScanning(false);
    }
  };

  // Convert selected single formula
  const handleConvertSelected = async () => {
    if (!selectedFormulaId) return;
    const item = scannedFormulas.find((f) => f.id === selectedFormulaId);
    if (!item) return;

    setIsConvertingSingle(true);
    setScanProgress({
      total: 1,
      current: 0,
      batchIndex: 1,
      batchTotal: 1,
      status: `正在原位转换所选公式: ${item.latex}`
    });
    try {
      await convertSingleFormula(engine, item);
      setScannedFormulas((prev) => prev.filter((f) => f.id !== item.id));
      setSelectedFormulaId(null);
      setScanProgress({
        total: 1,
        current: 1,
        batchIndex: 1,
        batchTotal: 1,
        status: "公式转换成功！"
      });
    } catch (err: any) {
      console.error(err);
      setScanProgress({
        total: 1,
        current: 0,
        batchIndex: 1,
        batchTotal: 1,
        status: `公式转换失败: ${err.message || String(err)}`
      });
    } finally {
      setIsConvertingSingle(false);
    }
  };

  // Convert one specific formula directly from the list item play button
  const handleConvertSpecific = async (item: ScannedFormula, e: React.MouseEvent) => {
    e.stopPropagation(); // Avoid triggering selection
    setIsConvertingSingle(true);
    setScanProgress({
      total: 1,
      current: 0,
      batchIndex: 1,
      batchTotal: 1,
      status: `正在原位转换公式: ${item.latex}`
    });
    try {
      await convertSingleFormula(engine, item);
      setScannedFormulas((prev) => prev.filter((f) => f.id !== item.id));
      if (selectedFormulaId === item.id) {
        setSelectedFormulaId(null);
      }
      setScanProgress({
        total: 1,
        current: 1,
        batchIndex: 1,
        batchTotal: 1,
        status: "公式转换成功！"
      });
    } catch (err: any) {
      console.error(err);
      setScanProgress({
        total: 1,
        current: 0,
        batchIndex: 1,
        batchTotal: 1,
        status: `公式转换失败: ${err.message || String(err)}`
      });
    } finally {
      setIsConvertingSingle(false);
    }
  };

  // Convert all scanned formulas at once
  const handleConvertAll = async () => {
    if (scannedFormulas.length === 0) return;
    setIsScanning(true);
    setScanStats(null);
    const totalCount = scannedFormulas.length;
    setScanProgress({
      total: totalCount,
      current: 0,
      batchIndex: 1,
      batchTotal: 1,
      status: `正在执行一键转换全部 ${totalCount} 个公式...`
    });
    try {
      await convertAllScannedFormulas(engine, scannedFormulas, (current, total) => {
        setScanProgress({
          total,
          current,
          batchIndex: 1,
          batchTotal: 1,
          status: `正在转换公式 ${current}/${total}...`
        });
      });
      setScanStats({
        total: totalCount,
        processed: totalCount,
        failed: 0
      });
      setScannedFormulas([]);
      setSelectedFormulaId(null);
    } catch (err: any) {
      console.error(err);
      setScanProgress({
        total: totalCount,
        current: 0,
        batchIndex: 1,
        batchTotal: 1,
        status: `一键转换失败: ${err.message || String(err)}`
      });
    } finally {
      setIsScanning(false);
    }
  };

  // Cancel / Reset scan list
  const handleCancelScan = async () => {
    try {
      await clearFormulaContentControls();
    } catch (err) {
      console.error("Failed to clear content controls on cancel", err);
    }
    setScannedFormulas([]);
    setSelectedFormulaId(null);
    setScanProgress(null);
    setScanStats(null);
  };

  // Run Scheme B (Smart Paste & Inject)
  const handlePasteInject = async () => {
    if (!stylesheetLoaded || !pasteText.trim()) return;
    setIsPasting(true);
    try {
      await pasteTerminalMixedText(engine, pasteText);
      setPasteText("");
    } catch (err: any) {
      console.error(err);
    } finally {
      setIsPasting(false);
    }
  };

  // Helper to render live MathJax preview for scanned list items
  const renderMathJaxPreview = (latex: string) => {
    try {
      const mathjax = (window as any).MathJax;
      if (mathjax && mathjax.tex2chtml) {
        const node = mathjax.tex2chtml(latex, { display: false });
        return { __html: node.outerHTML };
      }
    } catch (e) {
      console.error("Preview render error", e);
    }
    return { __html: `<code style="color:#8b949e">${latex}</code>` };
  };

  // Filter logs to display only WARN and ERROR entries, deduplicated by message text
  const getUniqueDisplayLogs = () => {
    const errorLogs = logs.filter((log) => log.includes("[ERROR]") || log.includes("[WARN]"));
    const seen = new Set<string>();
    const uniqueLogs: string[] = [];
    for (const log of errorLogs) {
      const msg = log.replace(/^\[[^\]]+\]\s*/, "");
      if (!seen.has(msg)) {
        seen.add(msg);
        uniqueLogs.push(log);
      }
    }
    return uniqueLogs;
  };
  const displayLogs = getUniqueDisplayLogs();

  return (
    <div style={styles.appContainer}>
      {/* Header Panel */}
      <div style={styles.header}>
        <div style={styles.logoRow}>
          <span style={styles.logoIcon}>∑</span>
          <span style={styles.titleText}>{title}</span>
        </div>
        <div style={styles.subtitleText}>LaTeX ➔ Word Native Equation Converter</div>
      </div>

      {loadingError && (
        <div style={styles.errorBanner}>
          <strong>环境初始化失败:</strong> {loadingError}
          <div style={{ marginTop: "4px", fontSize: "10px" }}>请刷新页面或检查网络连接后重试。</div>
        </div>
      )}

      {/* Tabs Selector */}
      <div style={styles.tabBar}>
        <button
          onClick={() => setActiveTab("scan")}
          style={{ ...styles.tabButton, ...(activeTab === "scan" ? styles.tabButtonActive : {}) }}
        >
          整页扫描 (A)
        </button>
        <button
          onClick={() => setActiveTab("paste")}
          style={{ ...styles.tabButton, ...(activeTab === "paste" ? styles.tabButtonActive : {}) }}
        >
          智能粘贴 (B)
        </button>
        <button
          onClick={() => setActiveTab("logs")}
          style={{ ...styles.tabButton, ...(activeTab === "logs" ? styles.tabButtonActive : {}) }}
        >
          日志记录
        </button>
      </div>

      {/* Tab Body */}
      <div style={styles.bodyContainer}>
        
        {/* Tab 1: Scan & Replace (Scheme A) */}
        {activeTab === "scan" && (
          <div style={styles.panel}>
            <div style={styles.description}>
              自动检索文档中所有的 LaTeX 公式 (支持 <code>$$...$$</code>, <code>$...$</code>, <code>\(...\)</code>, <code>\[...\]</code>) 并原位替换成 Word 的原生公式。
            </div>

            {scannedFormulas.length === 0 && !scanStats ? (
              <button
                onClick={handleScanSearch}
                disabled={isScanning || !stylesheetLoaded}
                style={{
                  ...styles.scanBtn,
                  ...(!stylesheetLoaded ? styles.btnDisabled : {}),
                  ...(isScanning ? styles.btnPulse : {}),
                  margin: "8px 0",
                  flexShrink: 0,
                }}
              >
                {isScanning ? "正在检索公式..." : "检索文档公式"}
              </button>
            ) : (
              <div>
                {/* Scan List Actions Header */}
                {scannedFormulas.length > 0 && (
                  <div style={styles.topActionsRow}>
                    <button
                      onClick={handleCancelScan}
                      style={{ ...styles.topActionBtn, ...styles.topActionBtnDanger }}
                      disabled={isScanning || isConvertingSingle}
                    >
                      取消
                    </button>
                    <button
                      onClick={handleConvertSelected}
                      style={{
                        ...styles.topActionBtn,
                        ...(selectedFormulaId ? styles.topActionBtnPrimary : {}),
                      }}
                      disabled={!selectedFormulaId || isScanning || isConvertingSingle}
                    >
                      转换所选 (1个)
                    </button>
                    <button
                      onClick={handleConvertAll}
                      style={{ ...styles.topActionBtn, ...styles.topActionBtnPrimary }}
                      disabled={isScanning || isConvertingSingle}
                    >
                      转换全部 ({scannedFormulas.length}个)
                    </button>
                  </div>
                )}

                {/* Scanned Formulas List */}
                {scannedFormulas.length > 0 && (
                  <div style={styles.formulaList}>
                    {scannedFormulas.map((item) => (
                      <div
                        key={item.id}
                        onClick={() => setSelectedFormulaId(item.id)}
                        style={{
                          ...styles.formulaItem,
                          ...(selectedFormulaId === item.id ? styles.formulaItemActive : {}),
                        }}
                      >
                        <div style={styles.formulaMeta}>
                          <span style={styles.formulaIndex}>#{item.occurrenceIndex + 1}</span>
                          <span style={styles.formulaType}>{item.type === "display" ? "块级" : "行内"}</span>
                        </div>
                        <div
                          style={styles.formulaPreview}
                          dangerouslySetInnerHTML={renderMathJaxPreview(item.latex)}
                        />
                        <button
                          onClick={(e) => handleConvertSpecific(item, e)}
                          style={styles.playBtn}
                          disabled={isScanning || isConvertingSingle}
                        >
                          转换
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {scanProgress && (
              <div style={styles.progressContainer}>
                <div style={styles.progressHeader}>
                  <span>进度: {scanProgress.current}/{scanProgress.total}</span>
                </div>
                <div style={styles.progressBarBg}>
                  <div
                    style={{
                      ...styles.progressBarFill,
                      width: `${scanProgress.total > 0 ? (scanProgress.current / scanProgress.total) * 100 : 0}%`,
                    }}
                  />
                </div>
                <div style={styles.progressStatus}>{scanProgress.status}</div>
              </div>
            )}

            {scanStats && (
              <div style={styles.statsCard}>
                <div style={styles.statsTitle}>扫描统计结果:</div>
                <div style={styles.statsGrid}>
                  <div style={styles.statBox}>
                    <div style={styles.statVal}>{scanStats.total}</div>
                    <div style={styles.statLabel}>检索公式总数</div>
                  </div>
                  <div style={{ ...styles.statBox, color: "#3fb950" }}>
                    <div style={styles.statVal}>{scanStats.processed}</div>
                    <div style={styles.statLabel}>成功转换数</div>
                  </div>
                  <div style={{ ...styles.statBox, color: scanStats.failed > 0 ? "#f85149" : "#8b949e" }}>
                    <div style={styles.statVal}>{scanStats.failed}</div>
                    <div style={styles.statLabel}>转换失败数</div>
                  </div>
                </div>
                <button
                  onClick={handleCancelScan}
                  style={{ ...styles.primaryBtn, marginTop: "12px", width: "100%" }}
                >
                  返回重新扫描
                </button>
              </div>
            )}
          </div>
        )}

        {/* Tab 2: Smart Paste (Scheme B) */}
        {activeTab === "paste" && (
          <div style={styles.panel}>
            <div style={styles.description}>
              在此处输入或直接粘贴(Ctrl+V)包含 LaTeX 混合文本，系统会自动分离出 LaTeX 渲染为 Word 完美公式一次性输入光标所在位置。
            </div>

            <textarea
              value={pasteText}
              onChange={(e) => setPasteText(e.target.value)}
              placeholder="在这里粘贴混合文本，例如：\n这里是一个公式 $$E=mc^2$$ 还有 inline 公式 \\(a^2+b^2=c^2\\)。"
              style={styles.textArea}
              disabled={isPasting || !stylesheetLoaded}
            />

            <button
              onClick={handlePasteInject}
              disabled={isPasting || !pasteText.trim() || !stylesheetLoaded}
              style={{
                ...styles.primaryBtn,
                ...(!pasteText.trim() || isPasting || !stylesheetLoaded ? styles.btnDisabled : {}),
              }}
            >
              {isPasting ? "正在分析并注入..." : "解析并混合注入当前光标"}
            </button>
          </div>
        )}

        {/* Tab 3: Logs */}
        {activeTab === "logs" && (
          <div style={styles.panel}>
            <div style={styles.logHeader}>
              <span style={styles.logTitleText}>日志记录控制面板 ({displayLogs.length} 条有效记录)</span>
              <div style={{ display: "flex", gap: "6px" }}>
                <button onClick={handleCopyErrorLogs} style={styles.clearBtn}>
                  复制关键错误日志
                </button>
                <button onClick={handleClearLogs} style={styles.clearBtn}>
                  清除日志
                </button>
              </div>
            </div>

            <div style={styles.logList}>
              {displayLogs.length === 0 ? (
                <div style={styles.logEmpty}>当前没有警告或错误日志。若公式转换异常，相关排查数据会在此处显示。</div>
              ) : (
                displayLogs.map((log, idx) => {
                  const isError = log.includes("[ERROR]");
                  const isWarn = log.includes("[WARN]");
                  let logColor = "#c9d1d9";
                  if (isError) logColor = "#f85149";
                  else if (isWarn) logColor = "#e3b341";
                  
                  return (
                    <div
                      key={idx}
                      onClick={() => setSelectedLog(log)}
                      style={{
                        ...styles.logItem,
                        color: logColor,
                        borderLeft: isError ? "3px solid #f85149" : isWarn ? "3px solid #e3b341" : "3px solid #58a6ff",
                      }}
                    >
                      {log.substring(0, 120)}{log.length > 120 ? "..." : ""}
                    </div>
                  );
                })
              )}
            </div>

            {selectedLog && (
              <div style={styles.logModal}>
                <div style={styles.logModalHeader}>
                  <span>详细日志数据</span>
                  <button onClick={() => setSelectedLog(null)} style={styles.closeBtn}>✕</button>
                </div>
                <pre style={styles.logPre}>{selectedLog}</pre>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

// CSS styles
const styles = {
  appContainer: {
    display: "flex",
    flexDirection: "column" as const,
    height: "100vh",
    backgroundColor: "#0d1117",
    color: "#c9d1d9",
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
  },
  header: {
    padding: "16px",
    backgroundColor: "#161b22",
    borderBottom: "1px solid #30363d",
    display: "flex",
    flexDirection: "column" as const,
    gap: "4px",
    flexShrink: 0,
  },
  logoRow: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  logoIcon: {
    fontSize: "22px",
    fontWeight: "bold",
    color: "#58a6ff",
  },
  titleText: {
    fontSize: "16px",
    fontWeight: "700",
    color: "#f0f6fc",
  },
  subtitleText: {
    fontSize: "11px",
    color: "#8b949e",
  },
  errorBanner: {
    margin: "12px 16px 0 16px",
    padding: "10px 12px",
    backgroundColor: "rgba(248, 81, 73, 0.1)",
    border: "1px solid rgba(248, 81, 73, 0.4)",
    borderRadius: "6px",
    color: "#ff7b72",
    fontSize: "11px",
    lineHeight: "1.4",
    flexShrink: 0,
  },
  tabBar: {
    display: "flex",
    borderBottom: "1px solid #30363d",
    backgroundColor: "#161b22",
    padding: "0 8px",
    flexShrink: 0,
  },
  tabButton: {
    flex: 1,
    padding: "12px 4px",
    backgroundColor: "transparent",
    border: "none",
    borderBottom: "2px solid transparent",
    color: "#8b949e",
    fontSize: "13px",
    fontWeight: "600",
    cursor: "pointer",
    transition: "color 0.2s, border-bottom 0.2s",
    textAlign: "center" as const,
  },
  tabButtonActive: {
    color: "#f0f6fc",
    borderBottom: "2px solid #f78166",
  },
  bodyContainer: {
    flex: 1,
    padding: "16px",
    overflowY: "auto" as const,
    display: "flex",
    flexDirection: "column" as const,
  },
  panel: {
    display: "flex",
    flexDirection: "column" as const,
    gap: "12px",
    flex: 1,
  },
  description: {
    fontSize: "12px",
    lineHeight: "1.5",
    color: "#8b949e",
    backgroundColor: "#161b22",
    padding: "10px 12px",
    borderRadius: "6px",
    border: "1px solid #30363d",
    flexShrink: 0,
  },
  actionCard: {
    backgroundColor: "#161b22",
    border: "1px solid #30363d",
    borderRadius: "8px",
    padding: "16px",
    display: "flex",
    flexDirection: "column" as const,
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    minHeight: "150px",
  },
  scanBtn: {
    width: "100%",
    padding: "14px 20px",
    backgroundColor: "#238636",
    color: "#ffffff",
    border: "1px solid #2ea44f",
    borderRadius: "6px",
    fontSize: "14px",
    fontWeight: "700",
    cursor: "pointer",
    transition: "background-color 0.2s",
    boxShadow: "0 1px 0 rgba(27, 31, 36, 0.1)",
  },
  primaryBtn: {
    padding: "10px 16px",
    backgroundColor: "#21262d",
    border: "1px solid #30363d",
    borderRadius: "6px",
    color: "#c9d1d9",
    fontSize: "13px",
    fontWeight: "600",
    cursor: "pointer",
    transition: "background-color 0.2s",
    textAlign: "center" as const,
  },
  btnDisabled: {
    opacity: 0.5,
    cursor: "not-allowed",
  },
  btnPulse: {
    animation: "pulse 1.5s infinite",
  },
  textArea: {
    flex: 1,
    minHeight: "120px",
    backgroundColor: "#0d1117",
    color: "#c9d1d9",
    border: "1px solid #30363d",
    borderRadius: "6px",
    padding: "12px",
    fontSize: "12px",
    fontFamily: "inherit",
    resize: "none" as const,
    lineHeight: "1.5",
    outline: "none",
  },
  progressContainer: {
    backgroundColor: "#161b22",
    border: "1px solid #30363d",
    borderRadius: "8px",
    padding: "12px",
    display: "flex",
    flexDirection: "column" as const,
    gap: "8px",
    flexShrink: 0,
  },
  progressHeader: {
    display: "flex",
    justifyContent: "space-between",
    fontSize: "12px",
    color: "#8b949e",
    fontWeight: "500",
  },
  progressBarBg: {
    height: "6px",
    backgroundColor: "#30363d",
    borderRadius: "3px",
    overflow: "hidden",
  },
  progressBarFill: {
    height: "100%",
    backgroundColor: "#58a6ff",
    borderRadius: "3px",
    transition: "width 0.3s ease",
  },
  progressStatus: {
    fontSize: "11px",
    color: "#8b949e",
  },
  statsCard: {
    backgroundColor: "rgba(22, 27, 34, 0.4)",
    border: "1px solid #30363d",
    borderRadius: "8px",
    padding: "12px",
    flexShrink: 0,
  },
  statsTitle: {
    fontSize: "13px",
    fontWeight: "600",
    color: "#f0f6fc",
    marginBottom: "10px",
  },
  statsGrid: {
    display: "flex",
    gap: "8px",
  },
  statBox: {
    flex: 1,
    backgroundColor: "#0d1117",
    border: "1px solid #21262d",
    borderRadius: "6px",
    padding: "10px 4px",
    textAlign: "center" as const,
  },
  statVal: {
    fontSize: "18px",
    fontWeight: "700",
  },
  statLabel: {
    fontSize: "10px",
    color: "#8b949e",
    marginTop: "2px",
  },
  logHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexShrink: 0,
  },
  logTitleText: {
    fontSize: "13px",
    fontWeight: "600",
    color: "#f0f6fc",
  },
  clearBtn: {
    padding: "4px 8px",
    backgroundColor: "transparent",
    border: "1px solid #30363d",
    color: "#c9d1d9",
    borderRadius: "4px",
    fontSize: "11px",
    cursor: "pointer",
  },
  logList: {
    backgroundColor: "#0d1117",
    border: "1px solid #30363d",
    borderRadius: "6px",
    padding: "8px",
    flex: 1,
    overflowY: "auto" as const,
    display: "flex",
    flexDirection: "column" as const,
    gap: "4px",
    fontFamily: "'JetBrains Mono', monospace",
    fontSize: "11px",
  },
  logItem: {
    padding: "6px",
    backgroundColor: "#161b22",
    borderRadius: "4px",
    cursor: "pointer",
    whiteSpace: "nowrap" as const,
    overflow: "hidden",
    textOverflow: "ellipsis",
    transition: "background-color 0.1s",
    flexShrink: 0,
  },
  logEmpty: {
    color: "#8b949e",
    textAlign: "center" as const,
    padding: "20px 0",
  },
  logModal: {
    backgroundColor: "#161b22",
    border: "1px solid #30363d",
    borderRadius: "8px",
    padding: "12px",
    display: "flex",
    flexDirection: "column" as const,
    gap: "8px",
    flexShrink: 0,
    marginTop: "8px",
  },
  logModalHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    fontSize: "12px",
    fontWeight: "600",
    color: "#f0f6fc",
  },
  closeBtn: {
    background: "none",
    border: "none",
    color: "#8b949e",
    cursor: "pointer",
    fontSize: "14px",
  },
  logPre: {
    margin: 0,
    padding: "8px",
    backgroundColor: "#0d1117",
    border: "1px solid #30363d",
    borderRadius: "4px",
    maxHeight: "180px",
    overflow: "auto" as const,
    fontSize: "11px",
    fontFamily: "'JetBrains Mono', monospace",
    whiteSpace: "pre-wrap" as const,
    color: "#58a6ff",
  },
  // Scanned Formula List styles
  topActionsRow: {
    display: "flex",
    justifyContent: "space-between",
    gap: "6px",
    marginBottom: "10px",
    flexShrink: 0,
  },
  topActionBtn: {
    flex: 1,
    padding: "8px 4px",
    fontSize: "11px",
    fontWeight: "bold" as const,
    borderRadius: "4px",
    cursor: "pointer",
    border: "1px solid #30363d",
    backgroundColor: "#21262d",
    color: "#c9d1d9",
    textAlign: "center" as const,
    transition: "background-color 0.1s",
  },
  topActionBtnPrimary: {
    backgroundColor: "#238636",
    borderColor: "#2ea44f",
    color: "#ffffff",
  },
  topActionBtnDanger: {
    backgroundColor: "#da3637",
    borderColor: "#f85149",
    color: "#ffffff",
  },
  formulaList: {
    backgroundColor: "#0d1117",
    border: "1px solid #30363d",
    borderRadius: "6px",
    padding: "6px",
    flex: 1,
    overflowY: "auto" as const,
    display: "flex",
    flexDirection: "column" as const,
    gap: "6px",
  },
  formulaItem: {
    padding: "8px 10px",
    backgroundColor: "#161b22",
    border: "1px solid #30363d",
    borderRadius: "4px",
    cursor: "pointer",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "10px",
    transition: "background-color 0.1s, border-color 0.1s",
    flexShrink: 0,
  },
  formulaItemActive: {
    backgroundColor: "#1f2937",
    borderColor: "#58a6ff",
  },
  formulaMeta: {
    display: "flex",
    flexDirection: "column" as const,
    gap: "2px",
    alignItems: "flex-start",
    flexShrink: 0,
  },
  formulaIndex: {
    fontSize: "11px",
    fontWeight: "bold",
    color: "#58a6ff",
  },
  formulaType: {
    fontSize: "9px",
    color: "#8b949e",
    backgroundColor: "#21262d",
    padding: "1px 4px",
    borderRadius: "2px",
  },
  formulaPreview: {
    flexGrow: 1,
    overflowX: "auto" as const,
    maxWidth: "160px",
    padding: "4px",
    backgroundColor: "#0d1117",
    border: "1px solid #21262d",
    borderRadius: "4px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  playBtn: {
    backgroundColor: "#238636",
    color: "#ffffff",
    border: "none",
    borderRadius: "4px",
    padding: "4px 8px",
    fontSize: "10px",
    fontWeight: "bold" as const,
    cursor: "pointer",
    flexShrink: 0,
  },
};

export default App;
