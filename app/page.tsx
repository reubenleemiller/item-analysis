"use client";

import { ChangeEvent, useMemo, useState } from "react";

type Part = "A" | "B";
type ResponseType = "MC" | "NR" | "WR";
type Question = { id: string; part: Part; responseType: ResponseType; maxPoints: number; label: string };
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
const starterQuestion = (id = "question-1"): Question => ({ id, part: "A", responseType: "MC", maxPoints: 1, label: "" });

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
  const [enabledParts, setEnabledParts] = useState<Part[]>(["A", "B"]);
  const [enabledTypes, setEnabledTypes] = useState<ResponseType[]>(types);
  const [questions, setQuestions] = useState<Question[]>([starterQuestion()]);
  const [students, setStudents] = useState<Student[]>([]);
  const [csvName, setCsvName] = useState("");
  const [status, setStatus] = useState("");
  const [working, setWorking] = useState(false);
  const [authenticating, setAuthenticating] = useState(false);
  const [googleAccessToken, setGoogleAccessToken] = useState<string | null>(null);
  const [createdSheetUrl, setCreatedSheetUrl] = useState("");

  const totalPoints = useMemo(() => questions.reduce((total, question) => total + (Number(question.maxPoints) || 0), 0), [questions]);
  const activeQuestions = questions.filter((question) => enabledParts.includes(question.part) && enabledTypes.includes(question.responseType));

  const updateQuestion = (id: string, updates: Partial<Question>) => setQuestions((current) => current.map((question) => question.id === id ? { ...question, ...updates } : question));
  const removeQuestion = (id: string) => setQuestions((current) => current.filter((question) => question.id !== id));
  const togglePart = (part: Part) => {
    setEnabledParts((current) => {
      if (current.includes(part)) {
        if (current.length === 1) return current;
        const fallback = current.find((candidate) => candidate !== part)!;
        setQuestions((questions) => questions.map((question) => question.part === part ? { ...question, part: fallback } : question));
        return current.filter((candidate) => candidate !== part);
      }
      return [...current, part];
    });
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
        body: JSON.stringify({ examTitle, questions: exportQuestions, students }),
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
        <div className="control-group"><span>Test sections</span><div className="toggles">{(["A", "B"] as Part[]).map((part) => <button key={part} className={enabledParts.includes(part) ? "toggle selected" : "toggle"} onClick={() => togglePart(part)} type="button">Part {part}</button>)}</div></div>
        <div className="control-group"><span>Response types</span><div className="toggles">{types.map((type) => <button key={type} className={enabledTypes.includes(type) ? "toggle selected" : "toggle"} onClick={() => toggleType(type)} type="button">{type}</button>)}</div></div>
        </section>

        <section className="panel">
        <div className="section-heading"><div><p className="eyebrow">1. Answer key setup</p><h2>Questions and maximum points</h2></div><div className="summary">{activeQuestions.length} items · {totalPoints} points</div></div>
        <p className="hint">Use the Part and response-type toggles to match the test. Labels are optional; blank labels automatically become A-MC1, B-WR2, and so on.</p>
        <div className="question-grid question-head"><span>Part</span><span>Type</span><span>Max points</span><span>Question label</span><span /></div>
        {questions.map((question) => <div className="question-grid" key={question.id}>
          <select value={question.part} onChange={(event) => updateQuestion(question.id, { part: event.target.value as Part })}>{enabledParts.map((part) => <option key={part}>{part}</option>)}</select>
          <select value={question.responseType} onChange={(event) => updateQuestion(question.id, { responseType: event.target.value as ResponseType })}>{enabledTypes.map((type) => <option key={type}>{type}</option>)}</select>
          <input type="number" min="0.01" step="0.25" value={question.maxPoints} onChange={(event) => updateQuestion(question.id, { maxPoints: Number(event.target.value) })} />
          <input value={question.label} onChange={(event) => updateQuestion(question.id, { label: event.target.value })} placeholder="Optional, e.g. Q1" />
          <button className="icon-button" type="button" onClick={() => removeQuestion(question.id)} aria-label="Remove question">×</button>
        </div>)}
        <button className="add-button" type="button" onClick={() => setQuestions((current) => [...current, { ...starterQuestion(`question-${Date.now()}`), part: enabledParts[0], responseType: enabledTypes[0] }])}>+ Add question</button>
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
