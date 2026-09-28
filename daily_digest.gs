/**
 * Daily Newsletter Digest — Google Apps Script
 * Collects the day's newsletters from chosen senders and emails you one combined digest.
 *
 * SETUP (one time):
 *  1. Go to https://script.google.com → New project → paste this whole file in place of Code.gs.
 *  2. Project Settings (gear icon) → set Time zone to "America/New_York".
 *  3. Pick "testDigest" in the function dropdown → Run → approve the Gmail permissions.
 *     (Google will warn the app is unverified: Advanced → Go to project → Allow. It's your own script.)
 *  4. Check your inbox for the test digest. If it looks right, run "installDailyTrigger" once.
 *  To stop it: run "removeTriggers".
 */

// ───────────── SETTINGS ─────────────
const CONFIG = {
  sendTo: Session.getActiveUser().getEmail(), // or 'benewm@gmail.com'
  sendHour: 7,                // hour of day (0–23) the digest is sent
  lookbackHours: 26,          // window to scan (a little over 24h so nothing slips between runs)
  snippetChars: 700,          // how much of each newsletter to include (0 = subject + link only)
  processedLabel: 'Digest-Processed', // applied to emails already included, so nothing repeats
  archiveAfter: false,        // true = also move the originals out of your inbox
  markReadAfter: false,       // true = also mark the originals as read

  // Each sender: a display name for the digest, plus text to match in the "From" line
  // (sender name or email address, case-insensitive). Add or edit freely.
  senders: [
    { name: 'History Facts',          match: ['history facts', 'historyfacts'] },
    { name: 'The New York Times',     match: ['new york times', 'nytimes.com'] },
    { name: 'Daily Passport',         match: ['daily passport', 'dailypassport'] },
    { name: 'Fast Company Daily',     match: ['fast company', 'fastcompany.com'] },
    { name: '1440',                   match: ['1440', 'join1440.com'] },
    { name: 'TLDR',                   match: ['tldr', 'tldrnewsletter.com'] },
  ],
};

// ───────────── MAIN ─────────────
function sendDailyDigest() {
  const label = getOrCreateLabel_(CONFIG.processedLabel);
  const cutoff = new Date(Date.now() - CONFIG.lookbackHours * 3600 * 1000);

  // Broad Gmail search; exact sender matching is done below.
  const fromTerms = CONFIG.senders
    .flatMap(s => s.match)
    .map(m => `from:"${m}"`)
    .join(' OR ');
  const query = `newer_than:2d -label:${CONFIG.processedLabel} -in:sent {${fromTerms}}`;
  const threads = GmailApp.search(query, 0, 300);

  const bySender = {};   // sender name -> array of items
  const usedThreads = [];

  threads.forEach(thread => {
    let used = false;
    thread.getMessages().forEach(msg => {
      if (msg.getDate() < cutoff) return;
      const sender = matchSender_(msg.getFrom());
      if (!sender) return;
      (bySender[sender.name] = bySender[sender.name] || []).push({
        subject: msg.getSubject() || '(no subject)',
        date: msg.getDate(),
        snippet: CONFIG.snippetChars > 0 ? cleanBody_(msg.getPlainBody(), CONFIG.snippetChars) : '',
        link: 'https://mail.google.com/mail/u/0/#all/' + thread.getId(),
      });
      used = true;
    });
    if (used) usedThreads.push(thread);
  });

  const total = Object.values(bySender).reduce((n, a) => n + a.length, 0);
  if (total === 0) {
    Logger.log('No matching newsletters in the window — no digest sent.');
    return;
  }

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEE, MMM d');
  const subject = `Daily Digest — ${today} (${total} item${total === 1 ? '' : 's'})`;
  const { html, text } = buildDigest_(bySender, today);

  GmailApp.sendEmail(CONFIG.sendTo, subject, text, { htmlBody: html, name: 'Daily Digest' });

  // Mark originals so they are never included twice.
  usedThreads.forEach(t => {
    t.addLabel(label);
    if (CONFIG.markReadAfter) t.markRead();
    if (CONFIG.archiveAfter) t.moveToArchive();
  });
  Logger.log(`Digest sent with ${total} items.`);
}

/** Sends a digest now WITHOUT labeling/archiving anything — safe for testing. */
function testDigest() {
  const saved = { a: CONFIG.archiveAfter, r: CONFIG.markReadAfter, l: CONFIG.processedLabel };
  CONFIG.archiveAfter = false;
  CONFIG.markReadAfter = false;
  CONFIG.processedLabel = 'Digest-Test-Unused';
  try {
    sendDailyDigest();
  } finally {
    const test = GmailApp.getUserLabelByName('Digest-Test-Unused');
    if (test) { test.getThreads().forEach(t => t.removeLabel(test)); test.deleteLabel(); }
    CONFIG.archiveAfter = saved.a; CONFIG.markReadAfter = saved.r; CONFIG.processedLabel = saved.l;
  }
}

// ───────────── SCHEDULING ─────────────
function installDailyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger('sendDailyDigest').timeBased().everyDays(1).atHour(CONFIG.sendHour).create();
  Logger.log(`Daily trigger installed for about ${CONFIG.sendHour}:00 (${Session.getScriptTimeZone()}).`);
}

function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyDigest')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ───────────── HELPERS ─────────────
function matchSender_(from) {
  const f = (from || '').toLowerCase();
  return CONFIG.senders.find(s => s.match.some(m => f.includes(m.toLowerCase()))) || null;
}

function cleanBody_(body, maxChars) {
  let t = (body || '')
    .replace(/\[?(https?:\/\/|mailto:)\S+\]?/g, '')   // strip raw URLs
    .replace(/^.*(view (this email )?in (your )?browser|unsubscribe|manage preferences).*$/gim, '')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n\s*\n+/g, '\n\n')
    .trim();
  if (t.length > maxChars) t = t.slice(0, maxChars).replace(/\s+\S*$/, '') + '…';
  return t;
}

function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function buildDigest_(bySender, today) {
  const tz = Session.getScriptTimeZone();
  let html = `<div style="font-family:Arial,Helvetica,sans-serif;max-width:680px;color:#222">
    <h2 style="margin:0 0 4px">Daily Digest</h2>
    <div style="color:#777;margin-bottom:18px">${esc_(today)}</div>`;
  let text = `DAILY DIGEST — ${today}\n\n`;

  CONFIG.senders.forEach(s => {
    const items = bySender[s.name];
    if (!items || !items.length) return;
    items.sort((a, b) => a.date - b.date);
    html += `<h3 style="border-bottom:2px solid #eee;padding-bottom:4px;margin:22px 0 8px">${esc_(s.name)}</h3>`;
    text += `=== ${s.name} ===\n`;
    items.forEach(it => {
      const time = Utilities.formatDate(it.date, tz, 'h:mm a');
      html += `<div style="margin:0 0 14px">
        <a href="${it.link}" style="font-weight:bold;color:#1a4fa0;text-decoration:none">${esc_(it.subject)}</a>
        <span style="color:#999;font-size:12px"> · ${time}</span>
        ${it.snippet ? `<div style="white-space:pre-line;font-size:13px;color:#444;margin-top:4px">${esc_(it.snippet)}</div>` : ''}
      </div>`;
      text += `- ${it.subject} (${time})\n  ${it.link}\n${it.snippet ? it.snippet + '\n' : ''}\n`;
    });
  });

  html += `<div style="color:#aaa;font-size:11px;margin-top:24px">Click a headline to open the full newsletter in Gmail.</div></div>`;
  return { html, text };
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}
