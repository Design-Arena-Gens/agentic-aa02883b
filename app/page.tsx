"use client";

import { useRef, useState } from "react";

export default function Page() {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [status, setStatus] = useState<string>("");
  const [downloading, setDownloading] = useState<boolean>(false);

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    const file = fileInputRef.current?.files?.[0];
    if (!file) {
      setStatus("Please choose a .docx file.");
      return;
    }
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setStatus("File must be a .docx.");
      return;
    }

    setStatus("Formatting? This can take a few seconds.");
    setDownloading(false);

    const form = new FormData();
    form.append("file", file);

    const res = await fetch("/api/format", {
      method: "POST",
      body: form
    });

    if (!res.ok) {
      const text = await res.text();
      setStatus(`Failed: ${text || res.statusText}`);
      return;
    }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);

    // Trigger download
    const a = document.createElement("a");
    a.href = url;
    const base = file.name.replace(/\.docx$/i, "");
    a.download = `${base}.formatted.docx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    setStatus("Done. Your formatted file has been downloaded.");
    setDownloading(true);
  }

  return (
    <main style={{ maxWidth: 720, margin: "0 auto", padding: 24 }}>
      <h1 style={{ margin: 0, fontSize: 28 }}>DOCX Formatter</h1>
      <p style={{ color: "#555" }}>Normalizes fonts, spacing, headings, and margins.</p>

      <form onSubmit={handleSubmit} style={{ display: "grid", gap: 12 }}>
        <input ref={fileInputRef} type="file" accept=".docx" />
        <button type="submit" style={{
          background: "#111827",
          color: "white",
          border: 0,
          borderRadius: 8,
          padding: "10px 14px",
          cursor: "pointer",
          fontWeight: 600
        }}>Format .docx</button>
      </form>

      <div style={{ marginTop: 12, minHeight: 24 }}>
        {status && <span>{status}</span>}
      </div>

      <section style={{ marginTop: 24, color: "#333" }}>
        <h2 style={{ fontSize: 18, marginBottom: 8 }}>What it enforces</h2>
        <ul>
          <li>Default font: Times New Roman, 12 pt</li>
          <li>Paragraph spacing: 0 before/after, 1.15 line spacing</li>
          <li>Heading sizes for H1/H2/H3 normalized</li>
          <li>1 inch page margins</li>
        </ul>
      </section>
    </main>
  );
}
