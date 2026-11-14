/**
 * Environmental Science Chapters 8–9 qualifier score receiver.
 * Bound to "ES Chapters 8–9 Qualifier Results" spreadsheet.
 * Deploy as: Web app > Execute as Me > Anyone.
 * This is a study-practice logger, not a tamper-proof testing system.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = JSON.parse((e.postData && e.postData.contents) || "{}");
    const name = safeText_(p.name, 100);
    const block = safeText_(p.block, 60);
    const essay = safeText_(p.essay, 10000);
    const correct = Number(p.correct);
    const total = Number(p.total);
    const used = Number(p.durationSeconds);
    if (!name || !block || essay.length < 40 || total !== 60 ||
        !Number.isInteger(correct) || correct < 0 || correct > total ||
        !Number.isFinite(used) || used < 0) {
      return json_({ok:false,error:"Invalid or incomplete attempt"});
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tab = ss.getSheetByName("All Attempts");
    if (!tab) throw new Error("All Attempts tab not found");
    const last = tab.getLastRow();
    const prior = last > 1 ? tab.getRange(2, 2, last - 1, 2).getDisplayValues() : [];
    const attempt = prior.filter(r => r[0].trim().toLowerCase() === name.toLowerCase() &&
        r[1].trim().toLowerCase() === block.toLowerCase()).length + 1;
    const missed = Array.isArray(p.missed)
      ? p.missed.slice(0, 60).map(x => safeText_(x, 150)).join("; ") : "";
    const now = new Date(); // trusted server-side submission timestamp
    // Date, student, class, attempt, correct, total, percentage, missed, essay, seconds, status, late flag.
    // Set the late flag at the time of submission; no full-column ARRAYFORMULA is needed.
    const cutoff = new Date("2026-11-21T00:00:00-05:00");
    const deadlineStatus = now >= cutoff ? "LATE" : "ON TIME";
    const row = last + 1;
    tab.getRange(row, 1, 1, 12).setValues([[now, name, block, attempt, correct, total,
        Math.round(correct / total * 100), missed, essay,
        Math.round(used), "Completed", deadlineStatus]]);
    tab.getRange(row, 1).setNumberFormat("mmm d, yyyy h:mm AM/PM");
    return json_({ok:true,attempt:attempt});
  } catch (err) {
    return json_({ok:false,error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
function doGet() {
  return json_({ok:true,message:"ES qualifier score receiver is online"});
}
function safeText_(v,max) {
  // Treat user text as literal sheet content, never as a formula.
  let s=String(v == null ? "" : v).trim().slice(0,max);
  if (/^[=+@\-]/.test(s)) s="'"+s;
  return s;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
