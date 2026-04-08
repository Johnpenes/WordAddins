import * as React from "react";
import { createRoot } from "react-dom/client";

// We now rely on Word selection rather than persistent font highlight.

const App = () => {
  const [status, setStatus] = React.useState("Idle");
  const [data, setData] = React.useState<any[]>([]);
  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [loading, setLoading] = React.useState(false);
  const [searchValue, setSearchValue] = React.useState("");

  // Maps display index (0-based) → raw index in body.footnotes.items.
  // Needed because some documents have non-auto-numbered footnotes (e.g. author
  // asterisk footnotes) that appear in body.footnotes.items but should be skipped.
  const itemIndicesRef = React.useRef<number[]>([]);

  const goToFootnote = async (index: number, dataOverride?: any[]) => {
    const sourceData = Array.isArray(dataOverride) ? dataOverride : data;
    const safeIndex = Math.max(0, Math.min(index, sourceData.length - 1));
    setCurrentIndex(safeIndex);

    try {
      const supportsWordApi15 = Office.context.requirements.isSetSupported("WordApi", "1.5");
      if (!supportsWordApi15) return;

      const rawBodyText = String(sourceData[safeIndex]?.bodyText ?? "").trim();
      const normalizedBodyText = rawBodyText.replace(/\s+/g, " ").trim();
      const targetBodyText = normalizedBodyText.slice(0, 80);
      const tailBodyText = normalizedBodyText.length > 80 ? normalizedBodyText.slice(-80) : normalizedBodyText;
      const searchOpts = { matchCase: false as const, matchWholeWord: false as const, ignorePunct: true as const, ignoreSpace: true as const };

      await Word.run(async (context) => {
        const body = context.document.body;
        const footnotes = body.footnotes;
        footnotes.load("items");
        await context.sync();

        // Use the stored raw index to skip non-auto-numbered footnotes (e.g. author asterisk
        // footnotes in Ho.docx that appear in body.footnotes.items but are not display-numbered).
        const rawIndex =
          itemIndicesRef.current.length > safeIndex ? itemIndicesRef.current[safeIndex] : safeIndex;

        if (rawIndex >= footnotes.items.length) return;

        const ref = footnotes.items[rawIndex].reference;
        let refParagraph: Word.Paragraph | null = null;

        try {
          refParagraph = ref.paragraphs.getFirst();
          refParagraph.load("text");
          await context.sync();
        } catch {
          refParagraph = null;
        }

        // Helper: search for a short anchor and select the first result.
        const trySearch = async (scope: Word.Paragraph | Word.Range | Word.Body, str: string): Promise<boolean> => {
          if (!str) return false;
          try {
            const matches = scope.search(str, searchOpts);
            matches.load("items");
            await context.sync();
            if (matches.items.length > 0) {
              matches.items[0].select();
              await context.sync();
              return true;
            }
          } catch {
            /* search threw — treat as no match */
          }
          return false;
        };

        // Helper: select a full sentence by finding a short start anchor and a short tail anchor,
        // then expanding the range between them.
        const tryExpandBetweenAnchors = async (
          scope: Word.Paragraph | Word.Range | Word.Body,
          startAnchor: string,
          endAnchor: string
        ): Promise<boolean> => {
          if (!startAnchor) return false;

          try {
            const startMatches = scope.search(startAnchor, searchOpts);
            startMatches.load("items");

            const endMatches = scope.search(endAnchor || startAnchor, searchOpts);
            endMatches.load("items");

            await context.sync();

            if (startMatches.items.length === 0 || endMatches.items.length === 0) {
              return false;
            }

            const startRange = startMatches.items[0];
            const endRange = endMatches.items[endMatches.items.length - 1];
            const expanded = startRange.expandTo(endRange);
            expanded.select();
            await context.sync();
            return true;
          } catch {
            return false;
          }
        };

        // In the containing paragraph, try to span the full sentence using short anchors.
        if (normalizedBodyText && refParagraph) {
          if (await tryExpandBetweenAnchors(refParagraph, targetBodyText, tailBodyText)) return;
          for (const str of [targetBodyText, tailBodyText]) {
            if (await trySearch(refParagraph, str)) return;
          }
        }

        // Then try document-wide searches. Prefer spanning the full sentence via anchors.
        if (normalizedBodyText) {
          if (await tryExpandBetweenAnchors(body, targetBodyText, tailBodyText)) return;
          for (const str of [targetBodyText, tailBodyText]) {
            if (await trySearch(body, str)) return;
          }
        }

        // Final guaranteed fallback: select the footnote reference superscript so Word scrolls there.
        try {
          ref.select();
          await context.sync();
        } catch { /* nothing more we can do */ }
      });
    } catch (err) {
      console.warn("Unable to navigate to body sentence / footnote reference", err);
    }
  };

  const analyzeCurrentDocument = async () => {
    try {
      setLoading(true);
      setStatus("Reading current Word document");

      let documentOoxml = "";
      let extractedFootnotes: {
        number: number;
        text: string;
        ooxml: string;
        bodyText: string;
        referenceOoxml: string;
      }[] = [];

      const supportsWordApi15 = Office.context.requirements.isSetSupported("WordApi", "1.5");

      await Word.run(async (context) => {
        if (!supportsWordApi15) {
          throw new Error("This version of Word does not support footnote APIs (WordApi 1.5).");
        }

        try {
          const body = context.document.body;
          const footnotes = body.footnotes;
          footnotes.load("items");
          await context.sync();

          const documentOoxmlResult = body.getOoxml();
          const ooxmlResults = footnotes.items.map((fn) => fn.body.getOoxml());
          const referenceOoxmlResults = footnotes.items.map((fn) => fn.reference.getOoxml());
          const refParagraphs = footnotes.items.map((fn) => fn.reference.paragraphs.getFirst());

          footnotes.items.forEach((fn) => {
            fn.body.load("text");
          });
          refParagraphs.forEach((p) => {
            p.load("text");
          });
          await context.sync();

          documentOoxml = documentOoxmlResult.value || "";

          // Filter to auto-numbered footnotes only (their body OOXML contains "footnoteRef").
          // Some documents have non-auto footnotes (e.g. an author asterisk footnote with type=normal
          // but no <w:footnoteRef/>) that appear in body.footnotes.items and shift all indices by +1.
          const autoIndices: number[] = [];
          footnotes.items.forEach((_, i) => {
            if ((ooxmlResults[i]?.value || "").includes("footnoteRef")) {
              autoIndices.push(i);
            }
          });
          itemIndicesRef.current = autoIndices;

          extractedFootnotes = autoIndices.map((rawIndex, displayIdx) => ({
            number: displayIdx + 1,
            text: footnotes.items[rawIndex].body.text || "",
            ooxml: ooxmlResults[rawIndex]?.value || "",
            bodyText: refParagraphs[rawIndex]?.text || "",
            referenceOoxml: referenceOoxmlResults[rawIndex]?.value || "",
          }));
        } catch (primaryErr) {
          console.warn("body.footnotes path failed; falling back to document.getFootnoteBody()", primaryErr);

          const footnoteBody = context.document.getFootnoteBody();
          footnoteBody.load("text");
          const footnoteOoxml = footnoteBody.getOoxml();
          await context.sync();

          const combinedText = footnoteBody.text || "";
          const combinedOoxml = footnoteOoxml.value || "";
          if (combinedText.trim()) {
            extractedFootnotes = [
              {
                number: 1,
                text: combinedText,
                ooxml: combinedOoxml,
                bodyText: "",
                referenceOoxml: "",
              },
            ];
          } else {
            extractedFootnotes = [];
          }
        }
      });

      setStatus(`Read ${extractedFootnotes.length} footnotes from Word`);

      if (extractedFootnotes.length === 0) {
        setData([]);
        setStatus("No footnotes found in this document");
        return;
      }

      const response = await fetch("https://localhost:3000/api/process-footnotes", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ documentOoxml, footnotes: extractedFootnotes }),
      });

      setStatus(`Received HTTP ${response.status}`);

      const result = await response.json();

      const footnotes = Array.isArray(result?.footnotes) ? result.footnotes : [];
      setData(footnotes);
      setCurrentIndex(0);
      setStatus(`Loaded ${footnotes.length} footnotes`);

      if (footnotes.length > 0) {
        await goToFootnote(0, footnotes);
      }
    } catch (err: any) {
      console.error("Word processing error:", err);
      alert(`Error: ${String(err?.message || err)}`);
      setStatus(`Error: ${String(err?.message || err)}`);
    } finally {
      setLoading(false);
    }
  };

  const current = data[currentIndex];

  const renderSearchLink = (src: any) => {
    const name: string = src?.displayName;
    if (!name) return null;
    const isUrl = name.startsWith("http://") || name.startsWith("https://");
    const href = isUrl ? name : `https://www.google.com/search?q=${encodeURIComponent(name)}`;
    const label = isUrl ? name : `Search: ${name}`;
    return (
      <div style={{ marginTop: 8, fontSize: 12 }}>
        <a href={href} target="_blank" rel="noopener noreferrer">
          {label}
        </a>
      </div>
    );
  };

  const isBackReference = (src: any): boolean =>
    Array.isArray(src?.rules) &&
    src.rules.some((r: string) => r.includes("Rule 4.1") || r.includes("Rule 4.2"));

  const renderFileLabel = (src: any) => {
    if (!src?.fileLabel) return null;
    if (isBackReference(src)) {
      return (
        <div style={{ marginTop: 8, fontSize: 12, color: "#888" }}>
          <strong>Suggested File Name:</strong> N/A — check root footnote
        </div>
      );
    }
    return (
      <div style={{ marginTop: 8, fontSize: 12 }}>
        <strong>Suggested File Name:</strong> {String(src.fileLabel)}
      </div>
    );
  };

  const renderRuns = (runs: any[]) => {
    if (!Array.isArray(runs) || runs.length === 0) return null;

    return runs.map((run, i) => (
      <span
        key={i}
        style={{
          fontStyle: run?.italic ? "italic" : "normal",
          fontWeight: run?.bold ? "bold" : "normal",
          textDecoration: run?.underline ? "underline" : "none",
          fontVariant: run?.small_caps ? "small-caps" : "normal",
          fontFamily: "Times New Roman, serif",
        }}
      >
        {String(run?.text ?? "")}
      </span>
    ));
  };
  return (
    <div style={{ padding: 16, fontFamily: "Arial" }}>
      <h2>Footnote Checker</h2>

      <button onClick={analyzeCurrentDocument}>Check Footnotes</button>

      <div style={{ marginTop: 12 }}>
        <strong>Status:</strong> {status}
      </div>

      {loading && <div style={{ marginTop: 8 }}>Loading...</div>}

      {data.length > 0 && current && (
        <div style={{ marginTop: 16 }}>
          <div style={{ marginBottom: 12 }}>
            <button
              onClick={() => {
                void goToFootnote(currentIndex - 1, data);
              }}
              disabled={currentIndex === 0}
            >
              ← Prev
            </button>
            <button
              onClick={() => {
                void goToFootnote(currentIndex + 1, data);
              }}
              disabled={currentIndex === data.length - 1}
              style={{ marginLeft: 8 }}
            >
              Next →
            </button>
            <span style={{ marginLeft: 12, fontSize: 12 }}>
              {currentIndex + 1} / {data.length}
            </span>
            <input
              type="text"
              value={searchValue}
              onChange={(e) => setSearchValue(e.target.value)}
              placeholder="Footnote #"
              style={{ marginLeft: 12, width: 90 }}
            />
            <button
              onClick={() => {
                const n = parseInt(searchValue, 10);
                if (!Number.isNaN(n) && n >= 1 && n <= data.length) {
                  void goToFootnote(n - 1, data);
                }
              }}
              style={{ marginLeft: 6 }}
            >
              Go
            </button>
          </div>

          <h3>Footnote {current.number}</h3>

          <div style={{ marginTop: 12, marginBottom: 8, fontSize: 12 }}>
            <strong>Sources:</strong> {Array.isArray(current.sources) ? current.sources.length : 0}
          </div>

          {Array.isArray(current.sources) && current.sources.length > 1 ? (
            current.sources.map((src: any, idx: number) => (
              <div
                key={idx}
                style={{
                  border: "1px solid #ccc",
                  borderRadius: 6,
                  padding: 10,
                  marginBottom: 10,
                }}
              >
                <div>
                  <strong>Source {idx + 1}:</strong>
                </div>
                <div style={{ marginTop: 6, fontFamily: "Times New Roman, serif" }}>
                  {Array.isArray(src?.runs) && src.runs.length > 0 ? renderRuns(src.runs) : String(src?.text ?? "")}
                </div>

                {Array.isArray(src?.rules) && src.rules.length > 0 && (
                  <div style={{ marginTop: 8, fontSize: 12 }}>
                    <strong>Rules:</strong> {src.rules.join(", ")}
                  </div>
                )}

                {Array.isArray(src?.warnings) && src.warnings.length > 0 && (
                  <div style={{ marginTop: 8, fontSize: 12, color: "#b00020" }}>
                    <strong>Warnings:</strong> {src.warnings.join(", ")}
                  </div>
                )}

                {renderFileLabel(src)}

                {renderSearchLink(src)}
              </div>
            ))
          ) : Array.isArray(current.sources) && current.sources.length === 1 ? (
            <div style={{ marginTop: 6 }}>
              <div style={{ fontFamily: "Times New Roman, serif" }}>
                {Array.isArray(current.sources[0]?.runs) && current.sources[0].runs.length > 0
                  ? renderRuns(current.sources[0].runs)
                  : String(current.sources[0]?.text ?? "")}
              </div>

              {Array.isArray(current.sources[0]?.rules) && current.sources[0].rules.length > 0 && (
                <div style={{ marginTop: 8, fontSize: 12 }}>
                  <strong>Rules:</strong> {current.sources[0].rules.join(", ")}
                </div>
              )}

              {Array.isArray(current.sources[0]?.warnings) && current.sources[0].warnings.length > 0 && (
                <div style={{ marginTop: 8, fontSize: 12, color: "#b00020" }}>
                  <strong>Warnings:</strong> {current.sources[0].warnings.join(", ")}
                </div>
              )}

              {renderFileLabel(current.sources[0])}

              {renderSearchLink(current.sources[0])}
            </div>
          ) : (
            <div style={{ fontSize: 12, color: "#666" }}>No split sources for this footnote yet.</div>
          )}
        </div>
      )}
    </div>
  );
};

// Ensure Office is ready before rendering React
Office.onReady(() => {
  const rootElement = document.getElementById("container");
  if (!rootElement) return;

  const root = createRoot(rootElement);
  root.render(<App />);
});