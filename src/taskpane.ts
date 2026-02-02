/**
 * Document Redaction Challenge
 * - Redact emails, phones, SSNs, credit card numbers, Employee IDs, MRNs, INSs in the document body
 * - Add "CONFIDENTIAL DOCUMENT" header (Primary header, first section)
 * - Enable Track Changes
 */

type SensitiveType = "email" | "phone" | "ssn" | "creditCard" | "employeeId" | "mrn" | "ins";
type RedactionCounts = Record<SensitiveType, number>;

type RunSummary = {
  trackingRequested: boolean;
  trackingEnabled: boolean;
  headerRequested: boolean;
  headerInserted: boolean;
  uniqueSensitiveStringsFound: number;
  replacements: number;
  byType: RedactionCounts;
  notes: string[];
};

const CONFIDENTIAL_TEXT = "CONFIDENTIAL DOCUMENT";

const REDACTION_MARKERS: Record<SensitiveType, string> = {
  email: "[REDACTED EMAIL]",
  phone: "[REDACTED PHONE]",
  ssn: "[REDACTED SSN]",
  creditCard: "[REDACTED CARD]",
  employeeId: "[REDACTED EMPLOYEE ID]",
  mrn: "[REDACTED MRN]",
  ins: "[REDACTED INS]",
};

let lastSummaryText = "";

// Helpers
function setStatus(msg: string, kind: "info" | "ok" | "warn" | "error" = "info") {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = msg;
  el.setAttribute("data-kind", kind);
}

function setBusy(busy: boolean) {
  const btn = document.getElementById("run") as HTMLButtonElement | null;
  if (!btn) return;
  btn.disabled = busy;
  btn.setAttribute("aria-busy", String(busy));
}

function enableRunButton(enabled: boolean) {
  const btn = document.getElementById("run") as HTMLButtonElement | null;
  if (!btn) return;
  btn.disabled = !enabled;
  btn.textContent = enabled ? "Redact document" : "Loading Office…";
}

function wordApi15Supported(): boolean {
  try {
    return (
      typeof Office !== "undefined" &&
      !!Office.context?.requirements?.isSetSupported &&
      Office.context.requirements.isSetSupported("WordApi", "1.5") &&
      typeof Word !== "undefined"
    );
  } catch {
    return false;
  }
}

function wordApiDesktop13Supported(): boolean {
  try {
    return (
      typeof Office !== "undefined" &&
      !!Office.context?.requirements?.isSetSupported &&
      Office.context.requirements.isSetSupported("WordApiDesktop", "1.3") &&
      typeof Word !== "undefined"
    );
  } catch {
    return false;
  }
}

async function copyToClipboardFallback(text: string): Promise<boolean> {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.top = "0";
  ta.style.left = "0";
  ta.style.width = "1px";
  ta.style.height = "1px";
  ta.style.opacity = "0";
  ta.setAttribute("readonly", "");

  document.body.appendChild(ta);
  ta.focus();
  ta.select();

  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch {
    ok = false;
  } finally {
    document.body.removeChild(ta);
  }

  return ok;
}

async function copyText(text: string): Promise<boolean> {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // ignore and fallback
  }
  return copyToClipboardFallback(text);
}

function renderUI() {
  const root = document.getElementById("app");
  if (!root) throw new Error("Missing #app in index.html");

  root.innerHTML = `
    <div class="shell">
      <div class="hero">
        <div class="title">Add-In for Redaction</div>
        <div class="subtitle">
          Redact sensitive info, add a confidentiality header, and track changes.
        </div>
      </div>

      <div class="card">
        <div class="controls">
          <label class="check">
            <input id="optHeader" type="checkbox" checked />
            <span>Add "${CONFIDENTIAL_TEXT}" header</span>
          </label>

          <label class="check">
            <input id="optTrack" type="checkbox" checked />
            <span>Enable Track Changes</span>
          </label>
        </div>

        <button id="run" class="primary" disabled>Loading Office…</button>
        <div id="status" class="status" data-kind="info">Loading…</div>

        <details class="details">
          <summary>What gets redacted?</summary>
          <ul>
            <li>Emails → ${REDACTION_MARKERS.email}</li>
            <li>Phone numbers → ${REDACTION_MARKERS.phone}</li>
            <li>SSNs → ${REDACTION_MARKERS.ssn}</li>
            <li>Credit cards → ${REDACTION_MARKERS.creditCard}</li>
            <li>Employee IDs → ${REDACTION_MARKERS.employeeId}</li>
            <li>MRNs → ${REDACTION_MARKERS.mrn}</li>
            <li>INSs → ${REDACTION_MARKERS.ins}</li>
          </ul>
        </details>
      </div>

      <div class="footer">
        Tip: In Word, go to <b>Review → Track Changes</b> to confirm edits were tracked.
        <br/>
        <a id="copySummary" href="#" style="color: rgba(255,255,255,0.8); text-decoration: underline;">
          Copy last run summary
        </a>
      </div>
    </div>
  `;

  const runBtn = document.getElementById("run") as HTMLButtonElement;
  runBtn.addEventListener("click", () => void runRedaction());

  const copyLink = document.getElementById("copySummary") as HTMLAnchorElement;
  copyLink.addEventListener("click", async (e) => {
    e.preventDefault();

    if (!lastSummaryText) {
      setStatus("No summary yet — run the redaction once.", "warn");
      return;
    }

    const ok = await copyText(lastSummaryText);
    setStatus(
      ok ? "Copied summary to clipboard." : "Copy is blocked here. Copy manually from the status box.",
      ok ? "ok" : "warn"
    );
  });
}

// Track changes
async function enableTrackChangesIfPossible(
  context: Word.RequestContext,
  requested: boolean,
  notes: string[]
): Promise<boolean> {
  if (!requested) return false;

  if (!wordApi15Supported()) {
    notes.push("Track Changes not enabled (WordApi 1.5 not available).");
    return false;
  }

  try {
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    notes.push("Track Changes enabled (trackAll).");
    return true;
  } catch (err) {
    notes.push("Track Changes failed to enable (continuing without it).");
    console.error("Track changes enable failed:", err);
    return false;
  }
}


// Confidential Header or a fallback Banner if Header not possible
async function insertBodyBannerIfNeeded(context: Word.RequestContext, notes: string[]): Promise<boolean> {
  const body = context.document.body;

  const existing = body.search(CONFIDENTIAL_TEXT, { matchCase: false });
  existing.load("items");
  await context.sync();

  if (existing.items.length > 0) return false;

  const p = body.insertParagraph(CONFIDENTIAL_TEXT, Word.InsertLocation.start);
  p.font.bold = true;
  p.font.size = 14;

  try {
    p.font.color = "#B91C1C";
  } catch {
    // ignore
  }

  p.alignment = Word.Alignment.centered;

  notes.push(`Inserted visible top-of-body "${CONFIDENTIAL_TEXT}" banner.`);
  return true;
}

async function ensureConfidentialHeader(
  context: Word.RequestContext,
  requested: boolean,
  notes: string[]
): Promise<boolean> {
  if (!requested) return false;

  try {
    const firstSection = context.document.sections.getFirst();
    const headerBody = firstSection.getHeader(Word.HeaderFooterType.primary);

    const existing = headerBody.search(CONFIDENTIAL_TEXT, { matchCase: false });
    existing.load("items");
    await context.sync();

    if (existing.items.length > 0) {
      notes.push("Header already present (skipped inserting again).");
      return false;
    }

    const p = headerBody.insertParagraph(CONFIDENTIAL_TEXT, Word.InsertLocation.start);
    p.font.bold = true;
    p.font.size = 12;
    p.alignment = Word.Alignment.centered;

    notes.push(`Inserted "${CONFIDENTIAL_TEXT}" header.`);
    return true;
  } catch (err) {
    notes.push("Header insert failed in this context; using visible banner.");
    console.error("Header insert failed:", err);
    return insertBodyBannerIfNeeded(context, notes);
  }
}

// Token utilities + replacement
function normalizeToken(s: string): string {
  return s
    .replace(/\u00A0/g, " ")
    .replace(/[‐-‒–—−]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

function digitsOnly(s: string): string {
  return s.replace(/\D/g, "");
}

function alnumOnly(s: string): string {
  return s.replace(/[^A-Za-z0-9]/g, "");
}

async function replaceTokens(
  context: Word.RequestContext,
  body: Word.Body,
  tokens: string[],
  replacement: string,
  simplify?: (s: string) => string
): Promise<number> {
  const unique = Array.from(new Set(tokens.map((t) => t.trim()))).filter(Boolean);

  let replaced = 0;

  for (const raw of unique) {
    const exact = raw;
    const norm = normalizeToken(raw);
    const simple = simplify ? simplify(raw) : "";

    // A) exact-ish
    let ranges = body.search(exact, { matchCase: false, matchWholeWord: false });
    ranges.load("items");
    await context.sync();

    // B) normalized
    if (ranges.items.length === 0 && norm !== exact) {
      ranges = body.search(norm, {
        matchCase: false,
        matchWholeWord: false,
        ignorePunct: true,
        ignoreSpace: true,
      });
      ranges.load("items");
      await context.sync();
    }

    // C) simplified
    if (ranges.items.length === 0 && simple) {
      ranges = body.search(simple, {
        matchCase: false,
        matchWholeWord: false,
        ignorePunct: true,
        ignoreSpace: true,
      });
      ranges.load("items");
      await context.sync();
    }

    if (ranges.items.length === 0) continue;

    for (const r of ranges.items) {
      r.insertText(replacement, Word.InsertLocation.replace);
      replaced += 1;
    }

    await context.sync();
  }

  return replaced;
}

// Redaction functions
async function redactEmailsByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const emailRegex = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
  const tokens = Array.from(text.matchAll(emailRegex)).map((m) => (m[0] ?? "").trim());

  return replaceTokens(context, body, tokens, REDACTION_MARKERS.email);
}

async function redactPhonesByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const phoneRegex = /\(\d{3}\)[\s\u00A0]\d{3}[‐-‒–—−-]\d{4}\b/g; // (555) 445-6677

  const tokens = Array.from(text.matchAll(phoneRegex)).map((m) => m[0] ?? "");
  return replaceTokens(context, body, tokens, REDACTION_MARKERS.phone, digitsOnly);
}

async function redactEmployeeIdsByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const empRegex = /\bEMP[‐-‒–—−-]\d{4}[‐-‒–—−-]\d{4}\b/gi;

  const tokens = Array.from(text.matchAll(empRegex)).map((m) => m[0] ?? "");
  return replaceTokens(context, body, tokens, REDACTION_MARKERS.employeeId, alnumOnly);
}

async function redactMrnsByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const mrnRegex = /\bMRN[‐-‒–—−-]\d{6}\b/gi;

  const tokens = Array.from(text.matchAll(mrnRegex)).map((m) => m[0] ?? "");
  return replaceTokens(context, body, tokens, REDACTION_MARKERS.mrn, alnumOnly);
}

async function redactCreditCardsByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const ccRegex = /\b\d{4}[‐-‒–—−-]\d{4}[‐-‒–—−-]\d{4}[‐-‒–—−-]\d{4}\b/g;

  const tokens = Array.from(text.matchAll(ccRegex)).map((m) => m[0] ?? "");
  return replaceTokens(context, body, tokens, REDACTION_MARKERS.creditCard, digitsOnly);
}

async function redactSsnsByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const ssnRegex = /\b(?!000|666|9\d\d)\d{3}[- ]?(?!00)\d{2}[- ]?(?!0000)\d{4}\b/g;

  const tokens = Array.from(text.matchAll(ssnRegex)).map((m) => m[0] ?? "");
  return replaceTokens(context, body, tokens, REDACTION_MARKERS.ssn, digitsOnly);
}

async function redactInsByRegex(context: Word.RequestContext, body: Word.Body): Promise<number> {
  body.load("text");
  await context.sync();

  const text = body.text ?? "";
  const insRegex = /\bINS[‐-‒–—−-]\d{8}\b/gi;

  const tokens = Array.from(text.matchAll(insRegex)).map((m) => m[0] ?? "");
  return replaceTokens(context, body, tokens, REDACTION_MARKERS.ins, alnumOnly);
}

// EMail hyperlink cleanup
async function removeMailtoHyperlinksInBody(
  context: Word.RequestContext,
  body: Word.Body,
  notes: string[]
): Promise<number> {
  if (!wordApiDesktop13Supported()) {
    notes.push("Mailto hyperlink might be visible as cleanup skipped (WordApiDesktop 1.3 not available).");
    return 0;
  }

  try {
    const full = body.getRange();
    const links = full.hyperlinks;
    links.load("items/address,textToDisplay");
    await context.sync();

    let removed = 0;
    for (const h of links.items) {
      const addr = (h.address ?? "").toLowerCase();
      if (addr.startsWith("mailto:")) {
        h.delete();
        removed += 1;
      }
    }

    if (removed > 0) await context.sync();
    notes.push(removed > 0 ? `Removed ${removed} mailto hyperlink(s).` : "No mailto hyperlinks found.");
    return removed;
  } catch (err) {
    notes.push("Mailto hyperlink cleanup failed (continuing).");
    console.error("Mailto hyperlink cleanup failed:", err);
    return 0;
  }
}

// Redact body
async function redactBody(context: Word.RequestContext, notes: string[]): Promise<{
  replacements: number;
  byType: RedactionCounts;
  unique: number;
}> {
  const body = context.document.body;

  const byType: RedactionCounts = {
    email: 0,
    phone: 0,
    ssn: 0,
    creditCard: 0,
    employeeId: 0,
    mrn: 0,
    ins: 0,
  };

  byType.email = await redactEmailsByRegex(context, body);
  byType.phone = await redactPhonesByRegex(context, body);
  byType.employeeId = await redactEmployeeIdsByRegex(context, body);
  byType.mrn = await redactMrnsByRegex(context, body);
  byType.creditCard = await redactCreditCardsByRegex(context, body);
  byType.ssn = await redactSsnsByRegex(context, body);
  byType.ins = await redactInsByRegex(context, body);

  await removeMailtoHyperlinksInBody(context, body, notes);

  const total =
    byType.email +
    byType.phone +
    byType.employeeId +
    byType.mrn +
    byType.creditCard +
    byType.ins +
    byType.ssn;

  return { replacements: total, byType, unique: total };
}

// Summary
function formatSummary(s: RunSummary): string {
  const lines: string[] = [];
  lines.push("Document Redaction Challenge Run Summary");
  lines.push(`Track Changes requested: ${s.trackingRequested ? "Yes" : "No"}`);
  lines.push(`Track Changes enabled:   ${s.trackingEnabled ? "Yes" : "No"}`);
  lines.push(`Header requested:        ${s.headerRequested ? "Yes" : "No"}`);
  lines.push(`Header inserted:         ${s.headerInserted ? "Yes" : "No"}`);
  lines.push(
    `Replacements: ${s.replacements} (emails ${s.byType.email}, phones ${s.byType.phone}, SSNs ${s.byType.ssn}, cards ${s.byType.creditCard}, empIDs ${s.byType.employeeId}, INSs ${s.byType.ins}, MRNs ${s.byType.mrn})`
  );
  lines.push(`Unique sensitive strings found: ${s.uniqueSensitiveStringsFound}`);
  if (s.notes.length) {
    lines.push("Notes:");
    for (const n of s.notes) lines.push(`- ${n}`);
  }
  return lines.join("\n");
}

// Main run
async function runRedaction() {
  setBusy(true);
  setStatus("Working…", "info");

  try {
    const optHeader = (document.getElementById("optHeader") as HTMLInputElement).checked;
    const optTrack = (document.getElementById("optTrack") as HTMLInputElement).checked;

    if (typeof Word === "undefined") {
      setStatus("Open this add-in inside Word to run redaction.", "warn");
      return;
    }

    await Word.run(async (context) => {
      const notes: string[] = [];

      setStatus("Step 1/3: Enabling Track Changes…", "info");
      const trackingEnabled = await enableTrackChangesIfPossible(context, optTrack, notes);

      setStatus("Step 2/3: Adding confidentiality header…", "info");
      const headerInserted = await ensureConfidentialHeader(context, optHeader, notes);

      await context.sync();

      setStatus("Step 3/3: Redacting…", "info");
      const r = await redactBody(context, notes);

      const summary: RunSummary = {
        trackingRequested: optTrack,
        trackingEnabled,
        headerRequested: optHeader,
        headerInserted,
        uniqueSensitiveStringsFound: r.unique,
        replacements: r.replacements,
        byType: r.byType,
        notes,
      };

      lastSummaryText = formatSummary(summary);

      setStatus(
        `Done. Replacements: ${summary.replacements} (emails ${summary.byType.email}, phones ${summary.byType.phone}, SSNs ${summary.byType.ssn}, cards ${summary.byType.creditCard}, empIDs ${summary.byType.employeeId}, INSs ${summary.byType.ins}, MRNs ${summary.byType.mrn}).`,
        "ok"
      );
    });
  } catch (e) {
    const anyErr = e as any;
    const message = anyErr?.message ?? (typeof anyErr === "string" ? anyErr : "Unknown error");
    const debug = anyErr?.debugInfo ? JSON.stringify(anyErr.debugInfo, null, 2) : "(no debugInfo)";
    setStatus(`Error: ${message}\n\nOfficeExtension.Error.debugInfo:\n${debug}`, "error");
    console.error("Office add-in error:", e);
  } finally {
    setBusy(false);
  }
}

// Boot
renderUI();

if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      enableRunButton(true);
      setStatus(
        wordApi15Supported()
          ? "Ready. Click “Redact document” button."
          : "Ready. Click “Redact document”. (WordApi 1.5 NOT supported → no tracking)",
        wordApi15Supported() ? "info" : "warn"
      );
    } else {
      enableRunButton(false);
      setStatus("This add-in is intended for Microsoft Word.", "warn");
    }
  });
} else {
  enableRunButton(false);
  setStatus("UI loaded. Open inside Word to run redaction.", "warn");
}
