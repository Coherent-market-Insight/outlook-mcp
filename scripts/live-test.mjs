// Live integration test against the real Outlook profile on this machine.
//
//   OUTLOOK_MCP_TEST_SEND_TO=you@example.com [OUTLOOK_MCP_TEST_ACCOUNT=other@account] node scripts/live-test.mjs
//
// Safety rules:
//  - everything created is prefixed "[outlook-mcp test]" and removed again in `finally`
//  - mail is only ever sent to OUTLOOK_MCP_TEST_SEND_TO (skipped when unset)
//  - meetings are saved without sending invitations
//  - the cleanup removes only items created by this run, by id; it never searches by subject
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { connect } from './mcp-call.mjs';

const TAG = '[outlook-mcp test]';
const sendTo = process.env.OUTLOOK_MCP_TEST_SEND_TO;
const stamp = new Date().toISOString().slice(0, 19).replace('T', ' ');

let pass = 0; let fail = 0;
const ok = (name, cond, detail = '') => {
  if (cond) { pass++; console.log(`  ok    ${name}`); }
  else { fail++; console.log(`  FAIL  ${name}${detail ? '\n        ' + String(detail).split('\n').slice(0, 4).join('\n        ') : ''}`); }
};
const idOf = (text) => /([0-9a-f]{6}:[~0-9A-Za-z]+)/.exec(text)?.[1];

const c = connect();
await c.init();
const call = async (n, a) => c.call(n, a);
const cleanup = { mail: [], events: [], contacts: [], tasks: [] };

try {
  console.log('\n# system');
  const tools = await c.list();
  ok('tools/list exposes 31 tools', tools.length === 31, tools.length);
  const chars = JSON.stringify(tools).length;
  ok(`schema size sane (${chars} chars ~${Math.round(chars / 4)} tokens)`, chars < 40000);
  const st = await call('status');
  ok('status lists accounts', /accounts/.test(st.text) && /@/.test(st.text), st.text);

  console.log('\n# mail: read');
  const folders = await call('mail_folders', { depth: 1 });
  ok('mail_folders shows Inbox with counts', /Inbox\s+\d+\/\d+/.test(folders.text), folders.text);
  const list = await call('mail_list', { limit: 3 });
  const firstId = idOf(list.text.split('\n')[1] ?? '');
  ok('mail_list returns ids', !!firstId && /shown/.test(list.text), list.text);
  const unread = await call('mail_list', { unread: true, limit: 2 });
  ok('mail_list unread filter', !unread.error, unread.text);
  const q = await call('mail_list', { query: 'the', limit: 2 });
  ok('mail_list full-text query', !q.error && /match/.test(q.text), q.text);
  const sent = await call('mail_list', { folder: 'sent', limit: 2 });
  ok('mail_list sent folder uses recipient column', !sent.error && /→/.test(sent.text), sent.text);
  const since = await call('mail_list', { since: '2020-01-01', limit: 2 });
  ok('mail_list since filter', !since.error, since.text);
  const paged = await call('mail_list', { limit: 2, offset: 2 });
  ok('mail_list paging', !paged.error && !paged.text.includes(firstId ?? '#'), paged.text);
  const get = await call('mail_get', { id: firstId });
  ok('mail_get returns headers and body', /^subject:/m.test(get.text) && /^--$/m.test(get.text), get.text);
  const getFull = await call('mail_get', { id: firstId, headers: true, max_chars: 300 });
  ok('mail_get headers=true lists addresses', /<[^>]+@[^>]+>/.test(getFull.text), getFull.text);
  const thread = await call('mail_thread', { id: firstId });
  ok('mail_thread returns messages', !thread.error && thread.text.length > 20, thread.text);
  const att = await call('mail_list', { attachments: true, limit: 1 });
  const attId = idOf(att.text.split('\n')[1] ?? '');
  if (attId) {
    const l = await call('mail_attachments', { id: attId });
    ok('mail_attachments lists', /\d+\. /.test(l.text), l.text);
    const dir = path.join(os.tmpdir(), 'outlook-mcp-live-test');
    const s = await call('mail_attachments', { id: attId, save: true, directory: dir });
    const file = /→ (.+)$/m.exec(s.text)?.[1];
    ok('mail_attachments saves to disk', !!file && fs.existsSync(file), s.text);
    fs.rmSync(dir, { recursive: true, force: true });
  }

  console.log('\n# mail: drafts, reply, forward, flags, delete');
  const to = sendTo ?? 'outlook-mcp-test@example.invalid';
  const d1 = await call('mail_draft_save', { to: [to], subject: `${TAG} draft ${stamp}`, body: 'Line one.\nLine two.' });
  const draftId = idOf(d1.text);
  ok('mail_draft_save creates', !!draftId, d1.text);
  if (draftId) cleanup.mail.push(draftId);
  const d2 = await call('mail_draft_save', { id: draftId, subject: `${TAG} draft edited ${stamp}`, cc: [to] });
  ok('mail_draft_save updates', /edited/.test(d2.text) && /cc:/.test(d2.text), d2.text);
  const dl = await call('mail_list', { folder: 'drafts', query: 'draft edited', limit: 5 });
  ok('draft found via search', dl.text.includes(draftId ?? '#'), dl.text);
  const dget = await call('mail_get', { id: draftId });
  ok('mail_get reads draft body', dget.text.includes('Line two.'), dget.text);
  const upd = await call('mail_update', { ids: [draftId], add_categories: ['Test'], importance: 'high', read: false });
  ok('mail_update categories/importance on a draft', /updated 1\/1/.test(upd.text), upd.text);
  const chk = await call('mail_get', { id: draftId });
  ok('category and importance persisted', /categories: Test/.test(chk.text) && /\[.*!.*\]/.test(chk.text), chk.text);
  const upd2 = await call('mail_update', { ids: [draftId], remove_categories: ['Test'] });
  ok('mail_update removes category', /updated 1\/1/.test(upd2.text), upd2.text);
  const dflag = await call('mail_update', { ids: [draftId], flag: 'flagged' });
  ok('flagging a draft fails with a clear message', /Drafts cannot be flagged/.test(dflag.text), dflag.text);

  // real flag round-trip on an existing inbox message; original state restored in `finally`
  const before = await call('mail_list', { limit: 1, flagged: true });
  const wasFlagged = firstId && before.text.includes(firstId);
  if (!wasFlagged) {
    const fl = await call('mail_update', { ids: [firstId], flag: 'flagged' });
    ok('mail_update flag=flagged on received mail', /updated 1\/1/.test(fl.text), fl.text);
    const fl2 = await call('mail_list', { flagged: true, limit: 50 });
    ok('flagged filter finds it', fl2.text.includes(firstId), fl2.text);
    const fl3 = await call('mail_update', { ids: [firstId], flag: 'none' });
    ok('mail_update flag=none restores', /updated 1\/1/.test(fl3.text), fl3.text);
  }

  const reply = await call('mail_reply', { id: firstId, body: `${TAG} reply draft`, all: true });
  const replyId = idOf(reply.text.replace(/^reply draft: /, ''));
  ok('mail_reply saves draft (not sent)', /reply draft:/.test(reply.text), reply.text);
  if (replyId) cleanup.mail.push(replyId);
  const rget = await call('mail_get', { id: replyId, quoted: true });
  ok('reply has our text above quoted original', rget.text.includes('reply draft'), rget.text);

  const fwd = await call('mail_forward', { id: firstId, to: [to], body: `${TAG} fwd note` });
  const fwdId = idOf(fwd.text.replace(/^forward draft: /, ''));
  ok('mail_forward saves draft', /forward draft:/.test(fwd.text), fwd.text);
  if (fwdId) cleanup.mail.push(fwdId);

  const bad = await call('mail_draft_save', { to: [to] });
  ok('mail_draft_save validates required fields', bad.error, bad.text);
  const badId = await call('mail_get', { id: 'zzzzzz:nope' });
  ok('bad id gives a clear error', badId.error && /id|store/i.test(badId.text), badId.text);

  const del = await call('mail_delete', { ids: cleanup.mail.splice(0) });
  ok('mail_delete moves drafts to Deleted Items', /deleted \d\/\d/.test(del.text), del.text);

  console.log('\n# mail: folders and move');
  const fname = `outlook-mcp-test-${Date.now()}`;
  const fc = await call('mail_folder_create', { name: fname });
  ok('mail_folder_create', /created/.test(fc.text), fc.text);
  const d3 = await call('mail_draft_save', { to: [to], subject: `${TAG} move me`, body: 'x' });
  const d3id = idOf(d3.text);
  const mv = await call('mail_move', { ids: [d3id], folder: `Inbox/${fname}` });
  ok('mail_move to new folder', /moved 1\/1/.test(mv.text), mv.text);
  const inNew = await call('mail_list', { folder: `Inbox/${fname}` });
  ok('moved item listed in new folder', /1 shown/.test(inNew.text), inNew.text);
  const lid = idOf(inNew.text.split('\n')[1] ?? '');
  const pd = await call('mail_delete', { ids: [lid], permanent: true });
  ok('mail_delete permanent', /permanently deleted 1\/1/.test(pd.text), pd.text);
  fs.writeFileSync(path.join(os.tmpdir(), 'outlook-mcp-folder-to-remove.txt'), fname);

  console.log('\n# mail: policy gates');
  const ro = connect(['--read-only']); await ro.init();
  const roTools = (await ro.list()).map((t) => t.name);
  ok('--read-only hides every write tool', !roTools.some((n) => /send|save|delete|update|move|create|reply|forward|respond/.test(n)) && roTools.includes('mail_list'), roTools.join(','));
  ro.close();
  const ns = connect(['--disable-send']); await ns.init();
  const nsTools = (await ns.list()).map((t) => t.name);
  ok('--disable-send hides send tools only', !nsTools.includes('mail_send') && nsTools.includes('mail_draft_save'), nsTools.join(','));
  ns.close();
  const ts = connect(['--toolsets', 'calendar']); await ts.init();
  const tsTools = (await ts.list()).map((t) => t.name);
  ok('--toolsets limits tools', tsTools.every((n) => n === 'status' || n.startsWith('calendar_')) && tsTools.length > 3, tsTools.join(','));
  ts.close();
  const al = connect(['--allowed-domains', 'example.invalid']); await al.init();
  const blocked = await al.call('mail_send', { to: [sendTo ?? 'someone@gmail.com'], subject: `${TAG} must not send`, body: 'blocked' });
  ok('--allowed-domains refuses other domains', blocked.error && /not in OUTLOOK_MCP_ALLOWED_DOMAINS/.test(blocked.text), blocked.text);
  al.close();

  console.log('\n# mail: send (test inbox only)');
  if (!sendTo) console.log('  skip  set OUTLOOK_MCP_TEST_SEND_TO to test mail_send');
  else {
    const f = path.join(os.tmpdir(), 'outlook-mcp-attach.txt');
    fs.writeFileSync(f, 'attachment from outlook-mcp live test\n');
    const s1 = await call('mail_send', { to: [sendTo], subject: `${TAG} send ${stamp}`, body: 'Plain body from live test.', attachments: [f] });
    ok('mail_send with attachment', /^sent →/.test(s1.text), s1.text);
    const s2 = await call('mail_send', { to: [sendTo], subject: `${TAG} html ${stamp}`, body: '<p><b>HTML</b> body</p>', html: true, importance: 'high' });
    ok('mail_send html', /^sent →/.test(s2.text), s2.text);
    const d4 = await call('mail_draft_save', { to: [sendTo], subject: `${TAG} draft-send ${stamp}`, body: 'sent from a saved draft' });
    const d4id = idOf(d4.text);
    const s3 = await call('mail_draft_send', { id: d4id });
    ok('mail_draft_send', /^sent →/.test(s3.text), s3.text);
    const otherAccount = process.env.OUTLOOK_MCP_TEST_ACCOUNT; // a non-default account in your profile
    if (otherAccount) {
      const s4 = await call('mail_send', { to: [sendTo], subject: `${TAG} from-account ${stamp}`, body: `sent with account=${otherAccount}`, account: otherAccount });
      console.log(`  info  account selection → ${s4.text.split('\n')[0]}`);
      ok('mail_send account either honoured or failed closed', /^sent →/.test(s4.text) || /Nothing was sent/.test(s4.text), s4.text);
    } else console.log('  skip  set OUTLOOK_MCP_TEST_ACCOUNT to test sending from a non-default account');
    const unresolved = await call('mail_send', { to: ['definitely-not-a-real-person-xyz'], subject: `${TAG} unresolved`, body: 'x' });
    ok('mail_send refuses unresolved recipients', unresolved.error && /Unresolved|Nothing was sent/.test(unresolved.text), unresolved.text);
    fs.rmSync(f, { force: true });
  }

  console.log('\n# calendar');
  const cals = await call('calendar_list');
  ok('calendar_list', /items/.test(cals.text), cals.text);
  const day = new Date(Date.now() + 14 * 864e5).toISOString().slice(0, 10);
  const ev = await call('calendar_event_create', { subject: `${TAG} meeting`, start: `${day}T10:00`, end: `${day}T10:30`, location: 'Room 1', body: 'agenda', show_as: 'tentative', reminder_minutes: 15, categories: ['Test'] });
  const evId = idOf(ev.text);
  ok('calendar_event_create', !!evId && /event created/.test(ev.text), ev.text);
  if (evId) cleanup.events.push(evId);
  const evs = await call('calendar_events', { start: day, end: day });
  ok('calendar_events lists it with show-as', evs.text.includes(evId) && /\(tentative\)/.test(evs.text), evs.text);
  const evq = await call('calendar_events', { start: day, end: day, query: 'Room 1' });
  ok('calendar_events query filter', evq.text.includes(evId), evq.text);
  const eg = await call('calendar_event_get', { id: evId });
  ok('calendar_event_get shows details', /reminder: 15 min/.test(eg.text) && /agenda/.test(eg.text), eg.text);
  const eu = await call('calendar_event_update', { id: evId, start: `${day}T11:00`, location: 'Room 2', subject: `${TAG} meeting moved`, show_as: 'free' });
  ok('calendar_event_update keeps duration', /11:00-11:30/.test(eu.text) && /Room 2/.test(eu.text) && /\(free\)/.test(eu.text), eu.text);
  const eu2 = await call('calendar_event_update', { id: evId, end: `${day}T12:00` });
  ok('calendar_event_update end only', /11:00-12:00/.test(eu2.text), eu2.text);

  const ad = await call('calendar_event_create', { subject: `${TAG} all-day`, start: day, all_day: true });
  const adId = idOf(ad.text);
  ok('all-day event', /all-day/.test(ad.text), ad.text);
  if (adId) cleanup.events.push(adId);

  const rec = await call('calendar_event_create', { subject: `${TAG} recurring`, start: `${day}T15:00`, duration_minutes: 30, repeat: 'weekly', repeat_count: 3 });
  const recId = idOf(rec.text);
  ok('recurring event created', /\[R\]|R\]/.test(rec.text) || /\[.*R.*\]/.test(rec.text), rec.text);
  if (recId) cleanup.events.push(recId);
  const wk = new Date(Date.now() + 14 * 864e5);
  const until = new Date(wk.getTime() + 22 * 864e5).toISOString().slice(0, 10);
  const occ = await call('calendar_events', { start: day, end: until, query: 'recurring' });
  const occLines = occ.text.split('\n').filter((l) => /recurring/.test(l) && l.includes(':'));
  ok('recurrence expands to 3 occurrences', occLines.length >= 3, occ.text);
  const second = occLines[1]?.match(/(\d{4}-\d{2}-\d{2}) \w{3} (\d{2}:\d{2})/);
  if (second) {
    const startIso = `${second[1]}T${second[2]}`;
    const ou = await call('calendar_event_update', { id: recId, occurrence_start: startIso, location: 'Moved occurrence' });
    ok('update a single occurrence', /Moved occurrence/.test(ou.text), ou.text);
    const od = await call('calendar_event_delete', { id: recId, occurrence_start: startIso });
    ok('delete a single occurrence', /deleted/.test(od.text), od.text);
    const occ2 = await call('calendar_events', { start: day, end: until, query: 'recurring' });
    ok('series now has 2 occurrences', occ2.text.split('\n').filter((l) => /recurring/.test(l) && l.includes(':')).length === 2, occ2.text);
  }

  const mtg = await call('calendar_event_create', { subject: `${TAG} meeting with attendee`, start: `${day}T16:00`, attendees: [sendTo ?? 'someone@example.com'] });
  const mtgId = idOf(mtg.text);
  ok('meeting saved without sending invites', /invites NOT sent/.test(mtg.text), mtg.text);
  if (mtgId) cleanup.events.push(mtgId);
  const ma = await call('calendar_event_get', { id: mtgId });
  ok('attendees listed with response status', /attendees:/.test(ma.text), ma.text);

  const fb = await call('calendar_free_busy', { start: day, end: day, duration_minutes: 30 });
  ok('calendar_free_busy lists windows', /free windows/.test(fb.text), fb.text);
  const fbBad = await call('calendar_free_busy', { start: day, work_start: 17, work_end: 9 });
  ok('calendar_free_busy validates hours', fbBad.error, fbBad.text);
  const cal404 = await call('calendar_events', { start: day, calendar: 'no-such-calendar' });
  ok('unknown calendar gives clear error', cal404.error && /not found/i.test(cal404.text), cal404.text);

  for (const id of cleanup.events.splice(0)) {
    const r = await call('calendar_event_delete', { id });
    ok(`cleanup event ${id}`, /deleted/.test(r.text), r.text);
  }

  console.log('\n# contacts');
  const cs = await call('contact_save', { full_name: `${TAG} Jane Tester`, company: 'Test Co', job_title: 'QA', email: 'jane.tester@example.invalid', mobile_phone: '+10000000', categories: ['Test'] });
  const cid = idOf(cs.text);
  ok('contact_save creates', !!cid, cs.text);
  if (cid) cleanup.contacts.push(cid);
  const csr = await call('contacts_search', { query: 'Jane Tester' });
  ok('contacts_search finds it', csr.text.includes(cid), csr.text);
  const cg = await call('contact_get', { id: cid });
  ok('contact_get shows fields', /company: Test Co/.test(cg.text) && /jane\.tester@example\.invalid/.test(cg.text), cg.text);
  const cu = await call('contact_save', { id: cid, job_title: 'Senior QA' });
  ok('contact_save updates', /Senior QA/.test(cu.text), cu.text);
  const rr = await call('recipient_resolve', { names: [sendTo ?? 'jane.tester@example.invalid', 'zzzz-nobody-zzzz'] });
  ok('recipient_resolve resolves and reports failures', /not resolved/.test(rr.text), rr.text);
  const cd = await call('contact_delete', { ids: cleanup.contacts.splice(0) });
  ok('contact_delete', /deleted 1/.test(cd.text), cd.text);

  console.log('\n# tasks');
  const ts1 = await call('task_save', { subject: `${TAG} task`, due: day, body: 'notes', importance: 'high' });
  const tid = idOf(ts1.text);
  ok('task_save creates', !!tid && /due/.test(ts1.text), ts1.text);
  if (tid) cleanup.tasks.push(tid);
  const tl = await call('tasks_list', { query: 'outlook-mcp test' });
  ok('tasks_list finds it', tl.text.includes(tid), tl.text);
  const tu = await call('task_save', { id: tid, status: 'in_progress', percent: 40 });
  ok('task_save updates status', /in_progress/.test(tu.text), tu.text);
  const tc = await call('task_save', { id: tid, status: 'completed' });
  ok('task_save completes', /completed/.test(tc.text), tc.text);
  const tlo = await call('tasks_list', { query: 'outlook-mcp test' });
  ok('completed task leaves the open list', !tlo.text.includes(tid), tlo.text);
  const tld = await call('tasks_list', { status: 'completed', query: 'outlook-mcp test' });
  ok('completed list contains it', tld.text.includes(tid), tld.text);
  const td = await call('task_delete', { ids: cleanup.tasks.splice(0) });
  ok('task_delete', /deleted 1/.test(td.text), td.text);
} catch (e) {
  fail++;
  console.log('  FAIL  unexpected exception', e);
} finally {
  // best-effort cleanup of anything a failed step left behind
  if (cleanup.mail.length) await call('mail_delete', { ids: cleanup.mail });
  for (const id of cleanup.events) await call('calendar_event_delete', { id });
  if (cleanup.contacts.length) await call('contact_delete', { ids: cleanup.contacts });
  if (cleanup.tasks.length) await call('task_delete', { ids: cleanup.tasks });
  c.close();
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
