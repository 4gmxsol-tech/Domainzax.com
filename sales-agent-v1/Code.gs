/**
 * DomainZax Sales Agent v1
 * Google Sheets + Apps Script. Safe-by-default: sending is disabled until explicitly enabled.
 * No third-party APIs required. GmailApp and spreadsheet services only.
 */

const SHEETS = {
  domains: 'Domains',
  prospects: 'Prospects',
  outreach: 'Outreach',
  suppression: 'Suppression',
  config: 'Config',
  log: 'ActivityLog'
};

const HEADERS = {
  Domains: ['domain','owned_verified','status','category','asking_price','floor_price','marketplace_url','notes','last_verified_at'],
  Prospects: ['prospect_id','company','website','contact_name','role','email','source_url','evidence','domain_match','relevance_score','compliance_checked','status','created_at'],
  Outreach: ['outreach_id','prospect_id','domain','email','subject','body','status','sent_at','gmail_thread_id','followup1_at','followup2_at','last_action_at','notes'],
  Suppression: ['email_normalized','company_or_person','reason','date_added','source'],
  Config: ['key','value','description'],
  ActivityLog: ['timestamp','action','result','details']
};

const DEFAULT_CONFIG = [
  ['SEND_ENABLED','FALSE','Must be TRUE to allow sending. Keep FALSE during setup and testing.'],
  ['DAILY_SEND_CAP','5','Hard cap for new first-contact emails per day.'],
  ['MIN_RELEVANCE_SCORE','80','Minimum prospect relevance score (0-100) for sending.'],
  ['FOLLOWUP1_BUSINESS_DAYS','5','First follow-up after this many business days, only if no reply.'],
  ['FOLLOWUP2_BUSINESS_DAYS','10','Final follow-up after this many business days, only if no reply.'],
  ['MAX_FOLLOWUPS','2','Maximum follow-ups per prospect.'],
  ['SENDER_NAME','Zax | DomainZax','Visible sender name.'],
  ['REPLY_TO','sales@domainzax.com','Must be a working mailbox that you control; update if needed.'],
  ['PUBLIC_CONTACT_EMAIL','sales@domainzax.com','Contact address shown in message footer.'],
  ['TIMEZONE','Europe/Paris','Timezone used for scheduling and date calculations.'],
  ['DAILY_RUN_HOUR','10','Preferred local hour for the daily trigger.'],
  ['CAMPAIGN_SIGNATURE','Zax\\nDomainZax','Signature appended to emails.']
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('DomainZax Agent')
    .addItem('1. Set up / repair sheets', 'setupAgent')
    .addItem('2. Seed candidate domain list (review required)', 'seedCandidateDomains')
    .addItem('3. Generate email drafts', 'generateDrafts')
    .addItem('4. Run sending + follow-up cycle', 'runDailyCycle')
    .addItem('5. Install daily trigger', 'installDailyTrigger')
    .addItem('6. Disable all sending', 'disableSending')
    .addToUi();
}

function setupAgent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,HEADERS[name].length).setValues([HEADERS[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,HEADERS[name].length).setFontWeight('bold');
    }
  });
  const cfg = ss.getSheetByName(SHEETS.config);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2,1,DEFAULT_CONFIG.length,3).setValues(DEFAULT_CONFIG);
  }
  log_('SETUP','OK','Sheets and safe defaults created. Sending is disabled.');
  SpreadsheetApp.getUi().alert('Setup complete. SEND_ENABLED is FALSE. Verify domain ownership, mailbox, prospect sources, and compliance before enabling sending.');
}

function seedCandidateDomains() {
  setupIfNeeded_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.domains);
  const existing = new Set(sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat().map(v => String(v).toLowerCase().trim()) : []);
  // Candidate names surfaced in prior work. All are deliberately unverified and cannot be marketed until reviewed.
  const candidates = [
    ['Domainzax.com','NO','REVIEW_REQUIRED','Domain sales brand','','','https://domainzax.com','Brand domain; verify registrar ownership and availability before activation',''],
    ['Rexation.com','NO','REVIEW_REQUIRED','Brandable','','','','Previously reported acquired; verify current ownership',''],
    ['TheCompute.si','NO','REVIEW_REQUIRED','AI / computing','','','','Prior portfolio discussion; verify current ownership',''],
    ['TheCortex.si','NO','REVIEW_REQUIRED','AI / cognition','','','','Previously reported acquired; verify current ownership',''],
    ['TheSpark.si','NO','REVIEW_REQUIRED','Brandable / technology','','','','Prior portfolio discussion; verify current ownership',''],
    ['Quantification.si','NO','REVIEW_REQUIRED','Data / analytics','','','','Previously reported acquired; verify current ownership',''],
    ['Cognitive.si','NO','REVIEW_REQUIRED','AI / cognition','','','','Previously reported acquired; verify current ownership',''],
    ['PhysicalManipulation.com','NO','REVIEW_REQUIRED','Robotics','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['HumanoidBehavior.com','NO','REVIEW_REQUIRED','Humanoid robotics','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['RobotIntelligenceAI.com','NO','REVIEW_REQUIRED','Robotics / AI','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['HumanoidPlanning.com','NO','REVIEW_REQUIRED','Humanoid robotics','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['RobotEmbodiment.com','NO','REVIEW_REQUIRED','Embodied AI','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['HumanoidContext.com','NO','REVIEW_REQUIRED','Humanoid robotics','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['RobotStack.co','NO','REVIEW_REQUIRED','Robotics software','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['HumanoidUI.com','NO','REVIEW_REQUIRED','Robotics UI','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['RoboticsBehavior.com','NO','REVIEW_REQUIRED','Robotics','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['HumanoidIntelligenceAI.com','NO','REVIEW_REQUIRED','Humanoid AI','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['HumanoidIntelligenceLab.com','NO','REVIEW_REQUIRED','Humanoid AI','','','','Previously mentioned portfolio candidate; verify ownership',''],
    ['WorldAgents.co','NO','REVIEW_REQUIRED','AI agents','','','','Previously mentioned portfolio candidate; verify ownership','']
  ];
  const rows = candidates.filter(r => !existing.has(r[0].toLowerCase()));
  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,HEADERS.Domains.length).setValues(rows);
  log_('SEED_DOMAINS','OK',rows.length + ' candidate rows added; all set to NO / REVIEW_REQUIRED.');
  SpreadsheetApp.getUi().alert(rows.length + ' candidate rows added. Every row is unverified and blocked from outreach until you confirm ownership and change owned_verified to YES and status to ACTIVE.');
}

function generateDrafts() {
  setupIfNeeded_();
  const ss = SpreadsheetApp.getActive();
  const domains = table_(ss.getSheetByName(SHEETS.domains));
  const prospects = table_(ss.getSheetByName(SHEETS.prospects));
  const outreachSh = ss.getSheetByName(SHEETS.outreach);
  const existing = table_(outreachSh).map(r => String(r.outreach_id || ''));
  let created = 0;
  prospects.forEach(p => {
    if (String(p.status).toUpperCase() !== 'READY') return;
    const domain = String(p.domain_match || '').trim().toLowerCase();
    const d = domains.find(x => String(x.domain).trim().toLowerCase() === domain);
    if (!d || String(d.owned_verified).toUpperCase() !== 'YES' || String(d.status).toUpperCase() !== 'ACTIVE') return;
    const score = Number(p.relevance_score || 0);
    if (score < Number(config_('MIN_RELEVANCE_SCORE','80'))) return;
    if (!validEmail_(p.email) || !p.source_url || !p.evidence || String(p.compliance_checked).toUpperCase() !== 'YES') return;
    if (isSuppressed_(p.email)) return;
    const id = String(p.prospect_id || '').trim();
    if (!id) continue;
    const oid = id + '|' + domain;
    if (existing.includes(oid)) continue;
    const subject = 'A domain that may fit ' + String(p.company || 'your team');
    const body = buildEmail_(p, d);
    outreachSh.appendRow([oid,id,d.domain,p.email,subject,body,'DRAFT','','','','','', 'Draft only; review before enabling sending.']);
    created++;
  });
  log_('GENERATE_DRAFTS','OK',created + ' draft(s) created; no email sent.');
  SpreadsheetApp.getUi().alert(created + ' draft(s) created. No messages were sent.');
}

function buildEmail_(p, d) {
  const first = String(p.contact_name || '').trim().split(/\s+/)[0] || 'there';
  const company = String(p.company || 'your team').trim();
  const domain = String(d.domain || '').trim();
  const evidence = String(p.evidence || '').trim();
  const source = String(p.source_url || '').trim();
  const replyTo = config_('PUBLIC_CONTACT_EMAIL','sales@domainzax.com');
  const sig = config_('CAMPAIGN_SIGNATURE','Zax\nDomainZax').replace(/\\n/g,'\n');
  return 'Hi ' + first + ',\n\n' +
    'I came across ' + company + ' and noticed: ' + evidence + '\n\n' +
    'I’m reaching out because ' + domain + ' may be relevant to your work in this space. If securing a domain like this is useful for your roadmap or brand, I can share the asking price and transfer details. No pressure if it is not a fit.\n\n' +
    'Context: ' + source + '\n\n' +
    'If this is not relevant, reply “no thanks” and I will not follow up. You can also contact us at ' + replyTo + '.\n\n' +
    sig + '\n\n' +
    'If you prefer not to receive further messages from DomainZax, simply reply “unsubscribe”.';
}

function runDailyCycle() {
  setupIfNeeded_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Another agent cycle is already running.');
  try {
    processRepliesAndSuppressions_();
    processFollowups_();
    processFirstContacts_();
  } finally {
    lock.releaseLock();
  }
}

function processFirstContacts_() {
  if (String(config_('SEND_ENABLED','FALSE')).toUpperCase() !== 'TRUE') {
    log_('FIRST_CONTACTS','SAFE_STOP','SEND_ENABLED is not TRUE. No messages sent.');
    return;
  }
  const ss = SpreadsheetApp.getActive();
  const outSh = ss.getSheetByName(SHEETS.outreach);
  const rows = table_(outSh);
  const sentToday = rows.filter(r => String(r.status).toUpperCase() === 'SENT' && sameLocalDay_(r.sent_at)).length;
  let remaining = Math.max(0, Number(config_('DAILY_SEND_CAP','5')) - sentToday);
  if (!remaining) { log_('FIRST_CONTACTS','CAP_REACHED','Daily first-contact cap reached.'); return; }

  const domainRows = table_(ss.getSheetByName(SHEETS.domains));
  const prospectRows = table_(ss.getSheetByName(SHEETS.prospects));
  const prospectById = Object.fromEntries(prospectRows.map(p => [String(p.prospect_id),p]));
  const domainByName = Object.fromEntries(domainRows.map(d => [String(d.domain).toLowerCase(),d]));
  for (let i=0; i<rows.length && remaining>0; i++) {
    const r = rows[i];
    if (String(r.status).toUpperCase() !== 'DRAFT') continue;
    const p = prospectById[String(r.prospect_id)];
    const d = domainByName[String(r.domain).toLowerCase()];
    if (!p || !d || String(d.owned_verified).toUpperCase() !== 'YES' || String(d.status).toUpperCase() !== 'ACTIVE') { updateOutreach_(i+2,7,'BLOCKED_DOMAIN_NOT_VERIFIED','Domain ownership/status not verified.'); continue; }
    if (Number(p.relevance_score || 0) < Number(config_('MIN_RELEVANCE_SCORE','80')) || String(p.compliance_checked).toUpperCase() !== 'YES' || !p.source_url || !p.evidence) { updateOutreach_(i+2,7,'BLOCKED_REVIEW','Missing score, evidence, source, or compliance check.'); continue; }
    if (!validEmail_(r.email) || isSuppressed_(r.email)) { updateOutreach_(i+2,7,'BLOCKED_EMAIL_OR_SUPPRESSION','Invalid email or suppressed recipient.'); continue; }
    if (hasPriorOutreach_(r.email, r.domain, r.outreach_id)) { updateOutreach_(i+2,7,'BLOCKED_DUPLICATE','Existing outreach found for recipient/domain.'); continue; }
    const body = String(r.body || '');
    if (!body || body.indexOf('Context: https://') < 0 || body.indexOf('unsubscribe') < 0) { updateOutreach_(i+2,7,'BLOCKED_MESSAGE_CHECK','Required source or opt-out wording missing.'); continue; }
    try {
      const opts = {name:config_('SENDER_NAME','Zax | DomainZax'), replyTo:config_('REPLY_TO','sales@domainzax.com')};
      GmailApp.sendEmail(String(r.email), String(r.subject), body, opts);
      Utilities.sleep(800);
      const thread = findSentThread_(r.email, r.subject);
      updateOutreach_(i+2,7,'SENT','');
      updateOutreach_(i+2,8,new Date(),'');
      if (thread) updateOutreach_(i+2,9,thread.getId(),'');
      updateOutreach_(i+2,12,new Date(),'');
      log_('SEND_FIRST','SENT',r.email+' | '+r.domain);
      remaining--;
    } catch (e) {
      updateOutreach_(i+2,7,'SEND_ERROR',String(e));
      log_('SEND_FIRST','ERROR',r.email+' | '+String(e));
    }
  }
}


function processRepliesAndSuppressions_() {
  const ss=SpreadsheetApp.getActive();
  const rows=table_(ss.getSheetByName(SHEETS.outreach));
  for (const r of rows) {
    if (!r.gmail_thread_id || !['SENT','FOLLOWUP1_SENT','FOLLOWUP2_SENT'].includes(String(r.status).toUpperCase())) continue;
    let thread=null;
    try { thread=GmailApp.getThreadById(String(r.gmail_thread_id)); } catch(e) {}
    if (!thread) continue;
    const me=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
    const messages=thread.getMessages();
    const replies=messages.filter(m => String(m.getFrom()).toLowerCase().indexOf(me)<0 && asDate_(r.sent_at) && m.getDate()>asDate_(r.sent_at));
    if (!replies.length) continue;
    const latest=replies[replies.length-1];
    const text=String(latest.getPlainBody()||'').toLowerCase();
    const optOut=/(unsubscribe|opt.?out|remove me|stop emailing|stop contacting|do not contact|don't contact|no more emails|no thanks|not interested|désinscri|ne me contactez plus|stop sending)/i.test(text);
    if (optOut && !isSuppressed_(r.email)) {
      addSuppression(r.email,'Opt-out or refusal detected in reply',r.prospect_id,'Gmail thread '+r.gmail_thread_id);
      log_('REPLY_SCAN','SUPPRESSED',r.email+' | detected opt-out/refusal');
    }
    const sh=ss.getSheetByName(SHEETS.outreach);
    const all=table_(sh);
    const idx=all.findIndex(x=>String(x.outreach_id)===String(r.outreach_id));
    if (idx>=0) {
      sh.getRange(idx+2,7).setValue(optOut?'SUPPRESSED':'REPLY_RECEIVED');
      sh.getRange(idx+2,13).setValue(optOut?'Reply opt-out/refusal detected; suppressed.':'Reply detected; automatic follow-ups stopped. Review reply manually.');
    }
  }
}

function processFollowups_() {
  if (String(config_('SEND_ENABLED','FALSE')).toUpperCase() !== 'TRUE') return;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.outreach);
  const rows = table_(sh);
  const cap = Math.max(0, Number(config_('DAILY_SEND_CAP','5')));
  let sentToday = rows.filter(r => ['SENT','FOLLOWUP1_SENT','FOLLOWUP2_SENT'].includes(String(r.status).toUpperCase()) && sameLocalDay_(r.last_action_at)).length;
  const pRows = table_(ss.getSheetByName(SHEETS.prospects));
  const pById = Object.fromEntries(pRows.map(p => [String(p.prospect_id),p]));
  for (let i=0;i<rows.length && sentToday<cap;i++) {
    const r=rows[i], status=String(r.status).toUpperCase();
    if (!['SENT','FOLLOWUP1_SENT'].includes(status)) continue;
    const p=pById[String(r.prospect_id)];
    if (!p || isSuppressed_(r.email) || hasReply_(r)) {
      if (hasReply_(r)) updateOutreach_(i+2,7,'REPLY_RECEIVED','Reply detected; automated follow-up stopped.');
      continue;
    }
    const sentAt = asDate_(r.sent_at); // Both follow-ups are timed from the original first-contact date.
    if (!sentAt) continue;
    const age = businessDaysBetween_(sentAt,new Date());
    const maxFollowups=Number(config_('MAX_FOLLOWUPS','2'));
    if (status==='SENT' && maxFollowups>=1 && age>=Number(config_('FOLLOWUP1_BUSINESS_DAYS','5'))) {
      if (sendFollowup_(r,1)) { updateOutreach_(i+2,7,'FOLLOWUP1_SENT',''); updateOutreach_(i+2,10,new Date(),''); updateOutreach_(i+2,12,new Date(),''); sentToday++; }
    } else if (status==='FOLLOWUP1_SENT' && maxFollowups>=2 && age>=Number(config_('FOLLOWUP2_BUSINESS_DAYS','10'))) {
      if (sendFollowup_(r,2)) { updateOutreach_(i+2,7,'FOLLOWUP2_SENT',''); updateOutreach_(i+2,11,new Date(),''); updateOutreach_(i+2,12,new Date(),''); sentToday++; }
    }
  }
}

function sendFollowup_(r, n) {
  if (!r.gmail_thread_id) {
    const thread=findSentThread_(r.email,r.subject);
    if (!thread) return false;
    updateOutreachById_(r.outreach_id,9,thread.getId());
    r.gmail_thread_id=thread.getId();
  }
  if (hasReply_(r) || isSuppressed_(r.email)) return false;
  const subject = (String(r.subject).toLowerCase().startsWith('re:') ? r.subject : 'Re: '+r.subject);
  const body = n===1
    ? 'Hi,\n\nJust following up once on my note about '+r.domain+'. If it is not relevant, no action is needed and I will not keep chasing.\n\nZax\nDomainZax\n\nReply “unsubscribe” to stop further contact.'
    : 'Hi,\n\nI’ll close the loop here so I do not clutter your inbox. If '+r.domain+' becomes relevant later, you can reach us at '+config_('PUBLIC_CONTACT_EMAIL','sales@domainzax.com')+'.\n\nZax\nDomainZax\n\nReply “unsubscribe” to stop further contact.';
  GmailApp.sendEmail(String(r.email),subject,body,{name:config_('SENDER_NAME','Zax | DomainZax'),replyTo:config_('REPLY_TO','sales@domainzax.com')});
  log_('FOLLOWUP'+n,'SENT',r.email+' | '+r.domain);
  return true;
}

function hasReply_(r) {
  let thread=null;
  if (r.gmail_thread_id) {
    try { thread=GmailApp.getThreadById(String(r.gmail_thread_id)); } catch(e) {}
  }
  if (!thread) thread=findSentThread_(r.email,r.subject);
  if (!thread) return false;
  const me=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
  return thread.getMessages().some(m => String(m.getFrom()).toLowerCase().indexOf(me)<0 && m.getDate()>asDate_(r.sent_at));
}

function findSentThread_(email, subject) {
  const safeEmail=String(email).replace(/["\\]/g,'');
  const safeSubject=String(subject).replace(/["\\]/g,'');
  const threads=GmailApp.search('in:sent to:'+safeEmail+' subject:"'+safeSubject+'" newer_than:30d',0,10);
  return threads.length ? threads[0] : null;
}

function installDailyTrigger() {
  const triggers=ScriptApp.getProjectTriggers();
  triggers.filter(t=>t.getHandlerFunction()==='runDailyCycle').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('runDailyCycle').timeBased().everyDays(1).atHour(Number(config_('DAILY_RUN_HOUR','10'))).inTimezone(config_('TIMEZONE','Europe/Paris')).create();
  log_('TRIGGER','OK','Daily run trigger installed. Sending still depends on SEND_ENABLED.');
  SpreadsheetApp.getUi().alert('Daily trigger installed. Sending is still controlled by SEND_ENABLED.');
}

function disableSending() {
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(SHEETS.config);
  if (!sh) throw new Error('Run setupAgent first.');
  const rows=sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) if (rows[i][0]==='SEND_ENABLED') sh.getRange(i+1,2).setValue('FALSE');
  log_('KILL_SWITCH','OK','SEND_ENABLED set to FALSE.');
  SpreadsheetApp.getUi().alert('Sending disabled.');
}

function addSuppression(email, reason, companyOrPerson, source) {
  setupIfNeeded_();
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.suppression);
  const normalized=String(email||'').trim().toLowerCase();
  if (!normalized) throw new Error('Email is required.');
  if (!isSuppressed_(normalized)) sh.appendRow([normalized,companyOrPerson||'',reason||'Opt-out',new Date(),source||'Manual']);
  log_('SUPPRESSION','ADDED',normalized+' | '+(reason||'Opt-out'));
}

function setupIfNeeded_() {
  const ss=SpreadsheetApp.getActive();
  if (!ss.getSheetByName(SHEETS.config)) setupAgent();
}

function table_(sh) {
  if (!sh || sh.getLastRow()<2) return [];
  const values=sh.getDataRange().getValues();
  const headers=values.shift().map(String);
  return values.filter(r=>r.some(v=>v!=='' && v!==null)).map(r=>{
    const obj={}; headers.forEach((h,i)=>obj[h]=r[i]); return obj;
  });
}

function config_(key, fallback) {
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.config);
  if (!sh || sh.getLastRow()<2) return fallback;
  const rows=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  for (const r of rows) if (String(r[0])===key) return String(r[1]);
  return fallback;
}

function isSuppressed_(email) {
  const normalized=String(email||'').trim().toLowerCase();
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.suppression);
  return !!sh && table_(sh).some(r=>String(r.email_normalized).toLowerCase()===normalized);
}

function hasPriorOutreach_(email, domain, currentId) {
  return table_(SpreadsheetApp.getActive().getSheetByName(SHEETS.outreach)).some(r =>
    String(r.outreach_id)!==String(currentId) &&
    String(r.email).toLowerCase()===String(email).toLowerCase() &&
    String(r.domain).toLowerCase()===String(domain).toLowerCase() &&
    !['DRAFT','BLOCKED_DOMAIN_NOT_VERIFIED','BLOCKED_REVIEW','BLOCKED_EMAIL_OR_SUPPRESSION','BLOCKED_DUPLICATE','SEND_ERROR'].includes(String(r.status).toUpperCase())
  );
}

function updateOutreach_(row, col, value, note) {
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.outreach);
  sh.getRange(row,col).setValue(value);
  if (note!==undefined && note!=='') sh.getRange(row,13).setValue(note);
}
function updateOutreachById_(id,col,value) {
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.outreach);
  const rows=table_(sh);
  const i=rows.findIndex(r=>String(r.outreach_id)===String(id));
  if(i>=0) sh.getRange(i+2,col).setValue(value);
}
function validEmail_(email) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email||'').trim()); }
function sameLocalDay_(v) {
  const d=asDate_(v); if(!d) return false;
  const tz=config_('TIMEZONE','Europe/Paris');
  return Utilities.formatDate(d,tz,'yyyy-MM-dd')===Utilities.formatDate(new Date(),tz,'yyyy-MM-dd');
}
function asDate_(v) { if(v instanceof Date && !isNaN(v.getTime())) return v; const d=new Date(v); return isNaN(d.getTime())?null:d; }
function businessDaysBetween_(start,end) {
  let count=0, d=new Date(start); d.setHours(0,0,0,0);
  const e=new Date(end); e.setHours(0,0,0,0);
  while(d<e) { d.setDate(d.getDate()+1); const day=d.getDay(); if(day!==0 && day!==6) count++; }
  return count;
}
function log_(action,result,details) {
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(SHEETS.log);
  if(!sh) { sh=ss.insertSheet(SHEETS.log); sh.appendRow(HEADERS.ActivityLog); }
  sh.appendRow([new Date(),action,result,details]);
}
