import * as React from "react";
import { createRoot } from "react-dom/client";

const App = () => {
  const [status, setStatus] = React.useState("Idle");
  const [data, setData] = React.useState<any[]>([]);
  const [currentIndex, setCurrentIndex] = React.useState(0);
  const [loading, setLoading] = React.useState(false);

  const analyzeCurrentDocument = async () => {
    try {
      setLoading(true);
      setStatus("Reading current Word document");

      let extractedFootnotes: { number: number; text: string; ooxml: string }[] = [];

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

          const ooxmlResults = footnotes.items.map((fn) => fn.body.getOoxml());

          footnotes.items.forEach((fn) => {
            fn.body.load("text");
          });
          await context.sync();
          await context.sync();

          extractedFootnotes = footnotes.items.map((fn, i) => ({
            number: i + 1,
            text: fn.body.text || "",
            ooxml: ooxmlResults[i]?.value || "",
          }));
        } catch (primaryErr) {
          console.warn("body.footnotes path failed; falling back to document.getFootnoteBody()", primaryErr);

          const footnoteBody = context.document.getFootnoteBody();
          footnoteBody.load("text");
          const footnoteOoxml = footnoteBody.getOoxml();
          await context.sync();
          await context.sync();

          const combinedText = footnoteBody.text || "";
          const combinedOoxml = footnoteOoxml.value || "";
          if (combinedText.trim()) {
            extractedFootnotes = [
              {
                number: 1,
                text: combinedText,
                ooxml: combinedOoxml,
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
        body: JSON.stringify({ footnotes: extractedFootnotes }),
      });

      setStatus(`Received HTTP ${response.status}`);

      const result = await response.json();

      const footnotes = Array.isArray(result?.footnotes) ? result.footnotes : [];
      setData(footnotes);
      setCurrentIndex(0);
      setStatus(`Loaded ${footnotes.length} footnotes`);
    } catch (err: any) {
      console.error("Word processing error:", err);
      alert(`Error: ${String(err?.message || err)}`);
      setStatus(`Error: ${String(err?.message || err)}`);
    } finally {
      setLoading(false);
    }
  };

  const current = data[currentIndex];
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
              onClick={() => setCurrentIndex((i) => Math.max(0, i - 1))}
              disabled={currentIndex === 0}
            >
              ← Prev
            </button>
            <button
              onClick={() => setCurrentIndex((i) => Math.min(data.length - 1, i + 1))}
              disabled={currentIndex === data.length - 1}
              style={{ marginLeft: 8 }}
            >
              Next →
            </button>
            <span style={{ marginLeft: 12, fontSize: 12 }}>
              {currentIndex + 1} / {data.length}
            </span>
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

                {src?.fileLabel && (
                  <div style={{ marginTop: 8, fontSize: 12 }}>
                    <strong>Suggested File:</strong> {String(src.fileLabel)}
                  </div>
                )}
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

              {current.sources[0]?.fileLabel && (
                <div style={{ marginTop: 8, fontSize: 12 }}>
                  <strong>Suggested File:</strong> {String(current.sources[0].fileLabel)}
                </div>
              )}
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