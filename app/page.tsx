"use client";

import { ChangeEvent, useMemo, useState } from "react";

type ResponseType = "MC" | "NR" | "WR";
type Question = { id: string; part: string; responseType: ResponseType; maxPoints: number; label: string };
type Student = { number: string; name: string };

declare global {
  interface Window {
    google?: {
      accounts: {
        oauth2: {
          initTokenClient: (config: {
            client_id: string;
            scope: string;
            callback: (response: { access_token?: string; error?: string }) => void;
          }) => { requestAccessToken: (config?: { prompt?: string }) => void };
        };
      };
    };
  }
}

const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID;
const types: ResponseType[] = ["MC", "NR", "WR"];
const starterQuestion = (id = "question-1", part = "A", responseType: ResponseType = "MC"): Question => ({ id, part, responseType, maxPoints: 1, label: "" });

function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let value = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (character === '"' && quoted && text[index + 1] === '"') {
      value += '"';
      index += 1;
    } else if (character === '"') {
      quoted = !quoted;
    } else if (character === "," && !quoted) {
      row.push(value.trim());
      value = "";
    } else if ((character === "\n" || character === "\r") && !quoted) {
      if (character === "\r" && text[index + 1] === "\n") index += 1;
      row.push(value.trim());
      if (row.some(Boolean)) rows.push(row);
      row = [];
      value = "";
    } else {
      value += character;
    }
  }
  row.push(value.trim());
  if (row.some(Boolean)) rows.push(row);
  return rows;
}

function parseStudents(text: string): Student[] {
  const rows = parseCsv(text);
  if (!rows.length) return [];
  const header = rows[0].map((cell) => cell.toLowerCase().replace(/[^a-z0-9]/g, ""));
  const numberIndex = header.findIndex((cell) => ["studentnumber", "studentid", "studentno", "id", "number"].includes(cell));
  const nameIndex = header.findIndex((cell) => ["studentname", "name", "fullname"].includes(cell));
  const firstIndex = header.findIndex((cell) => ["firstname", "first"].includes(cell));
  const lastIndex = header.findIndex((cell) => ["lastname", "last"].includes(cell));
  const data = numberIndex >= 0 || nameIndex >= 0 || firstIndex >= 0 ? rows.slice(1) : rows;
  return data
    .map((row) => ({
      number: row[numberIndex >= 0 ? numberIndex : 0] ?? "",
      name: nameIndex >= 0 ? row[nameIndex] ?? "" : [row[firstIndex >= 0 ? firstIndex : 1], row[lastIndex >= 0 ? lastIndex : -1]].filter(Boolean).join(" "),
    }))
    .filter((student) => student.number || student.name);
}

function loadGoogleIdentity() {
  return new Promise<void>((resolve, reject) => {
    if (window.google?.accounts.oauth2) return resolve();
    const existing = document.querySelector<HTMLScriptElement>('script[src="https://accounts.google.com/gsi/client"]');
    if (existing) {
      existing.addEventListener("load", () => resolve(), { once: true });
      existing.addEventListener("error", () => reject(new Error("Google sign-in could not load.")), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Google sign-in could not load."));
    document.head.appendChild(script);
  });
}

async function getAccessToken(prompt = "") {
  if (!GOOGLE_CLIENT_ID) throw new Error("Google Drive connection is not configured for this deployment.");
  await loadGoogleIdentity();
  return new Promise<string>((resolve, reject) => {
    const client = window.google!.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: "https://www.googleapis.com/auth/drive.file",
      callback: (response) => response.access_token ? resolve(response.access_token) : reject(new Error(response.error || "Google sign-in was cancelled.")),
    });
    client.requestAccessToken({ prompt });
  });
}

export default function Home() {
  const [examTitle, setExamTitle] = useState("Item Analysis");
  const [parts, setParts] = useState<string[]>(["A", "B"]);
  const [enabledTypes, setEnabledTypes] = useState<ResponseType[]>(types);
  const [questions, setQuestions] = useState<Question[]>([starterQuestion()]);
  const [selectedQuestionIds, setSelectedQuestionIds] = useState<Set<string>>(new Set());
  const [newPart, setNewPart] = useState("");
  const [bulkCount, setBulkCount] = useState(5);
  const [bulkPart, setBulkPart] = useState("A");
  const [bulkType, setBulkType] = useState<ResponseType>("MC");
  const [bulkMaxPoints, setBulkMaxPoints] = useState(1);
  const [includePartTotals, setIncludePartTotals] = useState(true);
  const [calculateCorrelationByPart, setCalculateCorrelationByPart] = useState(false);
  const [calculateStatisticsByPart, setCalculateStatisticsByPart] = useState(false);
  const [students, setStudents] = useState<Student[]>([]);
  const [csvName, setCsvName] = useState("");
  const [status, setStatus] = useState("");
  const [working, setWorking] = useState(false);
  const [authenticating, setAuthenticating] = useState(false);
  const [googleAccessToken, setGoogleAccessToken] = useState<string | null>(null);
  const [createdSheetUrl, setCreatedSheetUrl] = useState("");

  const totalPoints = useMemo(() => questions.reduce((total, question) => total + (Number(question.maxPoints) || 0), 0), [questions]);
  const activeQuestions = questions.filter((question) => parts.includes(question.part) && enabledTypes.includes(question.responseType));

  const updateQuestion = (id: string, updates: Partial<Question>) => setQuestions((current) => current.map((question) => question.id === id ? { ...question, ...updates } : question));
  const removeQuestion = (id: string) => {
    setQuestions((current) => current.filter((question) => question.id !== id));
    setSelectedQuestionIds((current) => {
      const next = new Set(current);
      next.delete(id);
      return next;
    });
  };
  const addPart = () => {
    const normalized = newPart.trim();
    if (!normalized) return;
    if (parts.some((part) => part.toLowerCase() === normalized.toLowerCase())) {
      setStatus(`A part named “${normalized}” already exists.`);
      return;
    }
    setParts((current) => [...current, normalized]);
    setBulkPart(normalized);
    setNewPart("");
  };
  const removePart = (part: string) => {
    if (parts.length === 1) return;
    const fallback = parts.find((candidate) => candidate !== part)!;
    setQuestions((current) => current.map((question) => question.part === part ? { ...question, part: fallback } : question));
    setParts((current) => current.filter((candidate) => candidate !== part));
    if (bulkPart === part) setBulkPart(fallback);
  };
  const toggleType = (type: ResponseType) => {
    setEnabledTypes((current) => {
      if (current.includes(type)) {
        if (current.length === 1) return current;
        const fallback = current.find((candidate) => candidate !== type)!;
        setQuestions((questions) => questions.map((question) => question.responseType === type ? { ...question, responseType: fallback } : question));
        return current.filter((candidate) => candidate !== type);
      }
      return [...current, type];
    });
  };
  const addQuestions = (count = 1, part = parts[0], responseType = enabledTypes[0], maxPoints = 1) => {
    const safeCount = Math.min(150 - questions.length, Math.max(1, Math.floor(Number(count) || 1)));
    if (safeCount < 1) {
      setStatus("A workbook can include up to 150 questions.");
      return;
    }
    const stamp = Date.now();
    setQuestions((current) => [...current, ...Array.from({ length: safeCount }, (_, index) => ({ ...starterQuestion(`question-${stamp}-${index}`, part, responseType), maxPoints: Number(maxPoints) || 1 }))]);
    if (safeCount < count) setStatus("Added the remaining questions up to the 150-question workbook limit.");
  };
  const toggleQuestionSelection = (id: string) => setSelectedQuestionIds((current) => {
    const next = new Set(current);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });
  const toggleAllQuestions = () => setSelectedQuestionIds((current) => current.size === questions.length ? new Set() : new Set(questions.map((question) => question.id)));
  const deleteSelectedQuestions = () => {
    if (!selectedQuestionIds.size) return;
    setQuestions((current) => current.filter((question) => !selectedQuestionIds.has(question.id)));
    setSelectedQuestionIds(new Set());
  };
  const importRoster = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const parsed = parseStudents(await file.text());
    setStudents(parsed);
    setCsvName(file.name);
    setStatus(parsed.length ? `${parsed.length} students imported from ${file.name}.` : "No student rows were found in that CSV.");
  };
  const connectGoogleDrive = async () => {
    setAuthenticating(true);
    setStatus("Opening Google sign-in…");
    try {
      setGoogleAccessToken(await getAccessToken("select_account"));
      setStatus("Google Drive connected. You can now create the sheet.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Unable to connect to Google Drive.");
    } finally {
      setAuthenticating(false);
    }
  };
  const createGoogleSheet = async () => {
    const exportQuestions = activeQuestions.filter((question) => Number(question.maxPoints) > 0);
    if (!exportQuestions.length) return setStatus("Add at least one question with a positive maximum score.");
    if (!students.length) return setStatus("Import the student class list CSV before creating the sheet.");
    if (!googleAccessToken) return setStatus("Log in with Google before creating the sheet.");
    const spinnerStartedAt = Date.now();
    setWorking(true);
    setCreatedSheetUrl("");
    setStatus("Building the Item Analysis workbook…");
    try {
      const response = await fetch("/api/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ examTitle, questions: exportQuestions, students, options: { includePartTotals, calculateCorrelationByPart, calculateStatisticsByPart } }),
      });
      if (!response.ok) throw new Error((await response.json()).error || "The workbook could not be built.");
      const workbook = await response.blob();
      setStatus("Creating Google Sheet in Drive…");
      const form = new FormData();
      form.append("metadata", new Blob([JSON.stringify({ name: examTitle.trim() || "Item Analysis", mimeType: "application/vnd.google-apps.spreadsheet" })], { type: "application/json" }));
      form.append("file", workbook, "item-analysis.xlsx");
      const upload = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink", {
        method: "POST",
        headers: { Authorization: `Bearer ${googleAccessToken}` },
        body: form,
      });
      const result = await upload.json() as { webViewLink?: string; error?: { message?: string } };
      if (!upload.ok || !result.webViewLink) throw new Error(result.error?.message || "Google Drive did not return a file link.");
      setCreatedSheetUrl(result.webViewLink);
      setStatus("Google Sheet created. Open it in a new tab below.");
    } catch (error) {
      setStatus(error instanceof Error ? error.message : "Unable to create the Google Sheet.");
    } finally {
      const remainingSpinnerTime = 1000 - (Date.now() - spinnerStartedAt);
      if (remainingSpinnerTime > 0) await new Promise((resolve) => window.setTimeout(resolve, remainingSpinnerTime));
      setWorking(false);
    }
  };

  return (
    <>
      <header className="site-header">
        <div className="site-header-inner">
          <a className="brand-link" href="https://rmtutoringservices.com" aria-label="Visit the RM Tutoring Services homepage">
            <span className="logo-shell">
              <img className="brand-logo" src="/rm-tutoring-logo.png" alt="" />
            </span>
            <span className="brand-copy">
              <strong>RM Tutoring Item Analysis Tool</strong>
              <span>Smarter Math Starts Here</span>
            </span>
          </a>
          <div className="header-actions">
            <button className={googleAccessToken ? "google-login connected" : "google-login"} type="button" onClick={connectGoogleDrive} disabled={authenticating || !GOOGLE_CLIENT_ID}>
              {authenticating ? "Connecting…" : googleAccessToken ? "Google Drive connected" : "Log in with Google"}
            </button>
            <a className="home-link" href="https://rmtutoringservices.com">Visit our homepage <span aria-hidden="true">→</span></a>
          </div>
        </div>
      </header>
      <main>
        <section className="hero">
          <p className="eyebrow">Assessment setup</p>
          <h1>Item Analysis Builder</h1>
          <p>Create a native Google Sheet in your Drive with the same Item Analysis layout, sized to this test and its class list.</p>
        </section>

        <section className="panel basics">
          <label className="field wide"><span>Exam title</span><input value={examTitle} onChange={(event) => setExamTitle(event.target.value)} placeholder="e.g. Unit 2 Assessment" /></label>
          <div className="control-group"><span>Response types</span><div className="toggles">{types.map((type) => <button key={type} className={enabledTypes.includes(type) ? "toggle selected" : "toggle"} onClick={() => toggleType(type)} type="button">{type}</button>)}</div></div>
          <div className="control-group export-options"><span>Workbook and calculation options</span><label><input type="checkbox" checked={includePartTotals} onChange={(event) => setIncludePartTotals(event.target.checked)} /> Add separate part-total columns</label><label><input type="checkbox" checked={calculateCorrelationByPart} onChange={(event) => setCalculateCorrelationByPart(event.target.checked)} /> Calculate item correlation per part</label><label><input type="checkbox" checked={calculateStatisticsByPart} onChange={(event) => setCalculateStatisticsByPart(event.target.checked)} /> Calculate item statistics per part</label></div>
          <div className="part-manager"><span>Test parts</span><div className="part-chips">{parts.map((part) => <span className="part-chip" key={part}>{part}<button type="button" onClick={() => removePart(part)} disabled={parts.length === 1} aria-label={`Remove ${part}`}>×</button></span>)}</div><div className="add-part"><input value={newPart} onChange={(event) => setNewPart(event.target.value)} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); addPart(); } }} placeholder="e.g. Part C" aria-label="New part name" /><button type="button" onClick={addPart}>Add part</button></div></div>
        </section>

        <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">1. Answer key setup</p><h2>Questions and maximum points</h2></div><div className="summary">{activeQuestions.length} items · {totalPoints} points</div></div>
        <p className="hint">Add as many named parts as your test needs. Labels are optional; blank labels automatically become Part A-MC1, Part C-WR2, and so on.</p>
        <div className="bulk-actions" aria-label="Bulk question actions"><div className="bulk-fields"><label>Questions<input type="number" min="1" max="150" value={bulkCount} onChange={(event) => setBulkCount(Number(event.target.value))} /></label><label>Part<select value={bulkPart} onChange={(event) => setBulkPart(event.target.value)}>{parts.map((part) => <option key={part} value={part}>{part}</option>)}</select></label><label>Type<select value={bulkType} onChange={(event) => setBulkType(event.target.value as ResponseType)}>{enabledTypes.map((type) => <option key={type}>{type}</option>)}</select></label><label>Max points<input type="number" min="0.01" step="0.25" value={bulkMaxPoints} onChange={(event) => setBulkMaxPoints(Number(event.target.value))} /></label></div><button className="add-button" type="button" onClick={() => addQuestions(bulkCount, bulkPart, bulkType, bulkMaxPoints)}>+ Add questions</button></div>
        <div className="selection-actions"><label><input type="checkbox" checked={questions.length > 0 && selectedQuestionIds.size === questions.length} onChange={toggleAllQuestions} /> Select all</label><button type="button" className="delete-selected" onClick={deleteSelectedQuestions} disabled={!selectedQuestionIds.size}>Delete selected{selectedQuestionIds.size ? ` (${selectedQuestionIds.size})` : ""}</button></div>
        <div className="question-grid question-head"><span>Select</span><span>Part</span><span>Type</span><span>Max points</span><span>Question label</span><span /></div>
        {questions.map((question) => <div className="question-grid" key={question.id}>
          <label className="question-select"><input type="checkbox" checked={selectedQuestionIds.has(question.id)} onChange={() => toggleQuestionSelection(question.id)} aria-label={`Select ${question.label || "question"}`} /></label>
          <select value={question.part} onChange={(event) => updateQuestion(question.id, { part: event.target.value })}>{parts.map((part) => <option key={part}>{part}</option>)}</select>
          <select value={question.responseType} onChange={(event) => updateQuestion(question.id, { responseType: event.target.value as ResponseType })}>{enabledTypes.map((type) => <option key={type}>{type}</option>)}</select>
          <input type="number" min="0.01" step="0.25" value={question.maxPoints} onChange={(event) => updateQuestion(question.id, { maxPoints: Number(event.target.value) })} />
          <input value={question.label} onChange={(event) => updateQuestion(question.id, { label: event.target.value })} placeholder="Optional, e.g. Q1" />
          <button className="icon-button" type="button" onClick={() => removeQuestion(question.id)} aria-label="Remove question">×</button>
        </div>)}
        <button className="add-button single-add" type="button" onClick={() => addQuestions(1)}>+ Add one question</button>
        </section>

        <section className="panel roster">
        <div><p className="eyebrow">2. Class list</p><h2>Import students before export</h2><p className="hint">Upload a CSV with student number and student name. Common header names are recognized automatically.</p></div>
        <label className="upload"><input type="file" accept=".csv,text/csv" onChange={importRoster} /><span>Choose CSV</span></label>
        <div className="roster-result"><strong>{students.length || "No"}</strong> student{students.length === 1 ? "" : "s"}{csvName ? <small>{csvName}</small> : null}</div>
        </section>

        <section className="create-panel">
        <div><h2>Ready to create the Google Sheet?</h2><p>The file is created directly in the signed-in Google account’s Drive. Student data is used only to create that file.</p></div>
        <button className={working ? "create-button creating" : "create-button"} type="button" onClick={createGoogleSheet} disabled={working || !googleAccessToken} aria-label={working ? "Creating Google Sheet" : undefined}>
          {working ? <span className="button-spinner" aria-hidden="true" /> : googleAccessToken ? "Create Google Sheet" : "Log in to create"}
        </button>
        {status ? <p className="status" role="status">{status}</p> : null}
        {createdSheetUrl ? <a className="created-sheet-link" href={createdSheetUrl} target="_blank" rel="noopener noreferrer">Open the new Google Sheet</a> : null}
        {!GOOGLE_CLIENT_ID ? <p className="config-note">Google Drive sign-in needs <code>NEXT_PUBLIC_GOOGLE_CLIENT_ID</code> configured in Vercel.</p> : null}
        </section>
      </main>
      <footer className="site-footer">
        <nav aria-label="Legal">
          <a href="https://www.rmtutoringservices.com/pages/privacy-policy">Privacy Policy</a>
          <a href="https://www.rmtutoringservices.com/pages/terms-of-service">Terms of Service</a>
        </nav>
      </footer>
    </>
  );
}
