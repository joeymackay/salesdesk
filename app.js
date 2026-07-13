const { useState, useEffect, useRef, useMemo } = React;
const idbOpen = () => new Promise((res, rej) => {
  const r = indexedDB.open("salesdesk", 1);
  r.onupgradeneeded = () => r.result.createObjectStore("kv");
  r.onsuccess = () => res(r.result);
  r.onerror = () => rej(r.error);
});
const idbGet = async (k) => {
  const d = await idbOpen();
  return new Promise((res, rej) => {
    const t = d.transaction("kv").objectStore("kv").get(k);
    t.onsuccess = () => res(t.result);
    t.onerror = () => rej(t.error);
  });
};
const idbSet = async (k, v) => {
  const d = await idbOpen();
  return new Promise((res, rej) => {
    const t = d.transaction("kv", "readwrite").objectStore("kv").put(v, k);
    t.onsuccess = () => res();
    t.onerror = () => rej(t.error);
  });
};
const idbDel = async (k) => {
  const d = await idbOpen();
  return new Promise((res, rej) => {
    const t = d.transaction("kv", "readwrite").objectStore("kv").delete(k);
    t.onsuccess = () => res();
    t.onerror = () => rej(t.error);
  });
};
const CFG_KEY = "salesdesk_cfg";
const CACHE_KEY = "cache";
const DIRTY_KEY = "dirty";
const loadCfg = () => {
  try {
    return JSON.parse(localStorage.getItem(CFG_KEY) || "null") || {};
  } catch (e) {
    return {};
  }
};
const saveCfg = (c) => {
  try {
    localStorage.setItem(CFG_KEY, JSON.stringify(c));
  } catch (e) {
  }
};
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Files.ReadWrite"];
let msalApp = null;
async function msalInit(clientId) {
  if (msalApp) return msalApp;
  msalApp = new msal.PublicClientApplication({
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
  });
  await msalApp.initialize();
  await msalApp.handleRedirectPromise();
  return msalApp;
}
function currentAccount() {
  if (!msalApp) return null;
  const all = msalApp.getAllAccounts();
  return all && all.length ? all[0] : null;
}
async function getToken() {
  const account = currentAccount();
  if (!account) throw new Error("no-account");
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
    throw new Error("redirecting");
  }
}
const filePathUrl = (path) => {
  const clean = String(path || "").replace(/^\/+/, "");
  return `${GRAPH}/me/drive/root:/${clean.split("/").map(encodeURIComponent).join("/")}`;
};
async function graphReadFile(path) {
  const token = await getToken();
  const res = await fetch(filePathUrl(path) + ":/content", { headers: { Authorization: `Bearer ${token}` } });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Read failed (${res.status})`);
  return await res.text();
}
async function graphWriteFile(path, text) {
  const token = await getToken();
  const res = await fetch(filePathUrl(path) + ":/content", {
    method: "PUT",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
    body: text
  });
  if (!res.ok) throw new Error(`Save failed (${res.status})`);
  return true;
}
const LS_KEY = "salesdesk_fallback_v1";
const C = {
  bg: "#EEF1EC",
  surface: "#FFFFFF",
  ink: "#16211B",
  inkSoft: "#5B6862",
  line: "#E2E8E1",
  green: "#16624A",
  greenDeep: "#0F4736",
  greenSoft: "#E2F0E9",
  brass: "#9A6E14",
  brassSoft: "#F5EBD5",
  violet: "#5B4B8A",
  violetSoft: "#ECE8F6",
  amber: "#B4650A",
  amberSoft: "#FBF0DF",
  red: "#A8321F",
  redSoft: "#F7E7E3",
  chip: "#EAEFE9"
};
const TIERS = {
  a: { bar: "#16624A", bg: "#E2F0E9", fg: "#0F4736", label: "A", name: "Priority A" },
  b: { bar: "#9A6E14", bg: "#F5EBD5", fg: "#7A560F", label: "B", name: "Priority B" },
  c: { bar: "#9BA79F", bg: "#EAEFE9", fg: "#5B6862", label: "C", name: "Priority C" }
};
const tierOf = (a) => TIERS[a && a.priority] || TIERS.b;
const TIER_RANK = { a: 0, b: 1, c: 2 };
const STAGES = {
  customer: { label: "Customer", bg: "#E2F0E9", fg: "#0F4736" },
  prospect: { label: "Prospect", bg: "#F5EBD5", fg: "#7A560F" }
};
const stageOf = (a) => STAGES[a && a.stage] || STAGES.customer;
const FONT_HEAD = "'Space Grotesk', sans-serif";
const FONT_BODY = "'Inter', sans-serif";
const FONT_MONO = "'IBM Plex Mono', monospace";
const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const todayStr = () => {
  const d = /* @__PURE__ */ new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const fmtDate = (iso) => {
  if (!iso) return "no date";
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  const t = todayStr();
  if (iso === t) return "Today";
  const tm = /* @__PURE__ */ new Date();
  tm.setDate(tm.getDate() + 1);
  const tmStr = `${tm.getFullYear()}-${String(tm.getMonth() + 1).padStart(2, "0")}-${String(tm.getDate()).padStart(2, "0")}`;
  if (iso === tmStr) return "Tomorrow";
  return dt.toLocaleDateString("en-US", { month: "short", day: "numeric" });
};
const daysFromToday = (iso) => {
  if (!iso) return null;
  const [y, m, d] = iso.split("-").map(Number);
  const a = new Date(y, m - 1, d);
  const n = /* @__PURE__ */ new Date();
  const b = new Date(n.getFullYear(), n.getMonth(), n.getDate());
  return Math.round((a - b) / 864e5);
};
const addDaysIso = (iso, n) => {
  const [y, m, d] = (iso || todayStr()).split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + n);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}-${String(dt.getDate()).padStart(2, "0")}`;
};
const lastTouched = (account, tasks) => {
  let latest = account.lastContact || null;
  tasks.forEach((t) => {
    if (t.accountId === account.id && t.done && t.completedAt) {
      if (!latest || t.completedAt > latest) latest = t.completedAt;
    }
  });
  return latest;
};
const touchStatus = (account, tasks) => {
  const cad = account.cadenceDays;
  if (!cad) return { state: "off" };
  const last = lastTouched(account, tasks);
  if (!last) return { state: "overdue", since: null, cad };
  const since = -daysFromToday(last);
  if (since >= cad) return { state: "overdue", since, cad };
  if (since >= cad - Math.max(2, Math.round(cad * 0.2))) return { state: "due", since, cad };
  return { state: "ok", since, cad };
};
function Icon({ name, size = 16 }) {
  const p = { width: size, height: size, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 1.9, strokeLinecap: "round", strokeLinejoin: "round", className: "sd-icon" };
  const paths = {
    today: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("circle", { cx: "12", cy: "12", r: "9" }), /* @__PURE__ */ React.createElement("path", { d: "M12 7v5l3 2" })),
    dump: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M9 18h6" }), /* @__PURE__ */ React.createElement("path", { d: "M10 21h4" }), /* @__PURE__ */ React.createElement("path", { d: "M12 3a6 6 0 0 0-4 10.5c.6.6 1 1.4 1 2.5h6c0-1.1.4-1.9 1-2.5A6 6 0 0 0 12 3Z" })),
    accounts: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M3 9.5 5 4h14l2 5.5" }), /* @__PURE__ */ React.createElement("path", { d: "M4 9.5h16V20H4z" }), /* @__PURE__ */ React.createElement("path", { d: "M3 9.5a3 3 0 0 0 6 0 3 3 0 0 0 6 0 3 3 0 0 0 6 0" }), /* @__PURE__ */ React.createElement("path", { d: "M10 20v-5h4v5" })),
    strategies: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("circle", { cx: "12", cy: "12", r: "8" }), /* @__PURE__ */ React.createElement("circle", { cx: "12", cy: "12", r: "4" }), /* @__PURE__ */ React.createElement("circle", { cx: "12", cy: "12", r: "0.6", fill: "currentColor" })),
    team: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("circle", { cx: "9", cy: "8", r: "3" }), /* @__PURE__ */ React.createElement("path", { d: "M3.5 20a5.5 5.5 0 0 1 11 0" }), /* @__PURE__ */ React.createElement("path", { d: "M16 5.2a3 3 0 0 1 0 5.6" }), /* @__PURE__ */ React.createElement("path", { d: "M17.5 14.3A5.5 5.5 0 0 1 20.5 19" })),
    mail: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("rect", { x: "3", y: "5", width: "18", height: "14", rx: "2" }), /* @__PURE__ */ React.createElement("path", { d: "m3.5 6.5 8.5 6 8.5-6" })),
    repeat: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M17 2.5 20 5.5 17 8.5" }), /* @__PURE__ */ React.createElement("path", { d: "M20 5.5H9a5 5 0 0 0-5 5" }), /* @__PURE__ */ React.createElement("path", { d: "M7 21.5 4 18.5 7 15.5" }), /* @__PURE__ */ React.createElement("path", { d: "M4 18.5h11a5 5 0 0 0 5-5" })),
    bell: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M18 8a6 6 0 0 0-12 0c0 7-3 9-3 9h18s-3-2-3-9" }), /* @__PURE__ */ React.createElement("path", { d: "M13.5 21a1.9 1.9 0 0 1-3 0" })),
    plus: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M12 5v14M5 12h14" })),
    clock: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("circle", { cx: "12", cy: "12", r: "9" }), /* @__PURE__ */ React.createElement("path", { d: "M12 7v5l3 2" })),
    search: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("circle", { cx: "11", cy: "11", r: "7" }), /* @__PURE__ */ React.createElement("path", { d: "m20 20-3.5-3.5" })),
    project: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M3 7a2 2 0 0 1 2-2h4l2 2.5h8a2 2 0 0 1 2 2V18a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z" })),
    spark: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M12 3l1.8 4.9L18.5 9.7 13.8 11.5 12 16.4 10.2 11.5 5.5 9.7 10.2 7.9Z" }), /* @__PURE__ */ React.createElement("path", { d: "M18 15l.7 1.9 1.9.7-1.9.7-.7 1.9-.7-1.9-1.9-.7 1.9-.7Z" })),
    send: /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("path", { d: "M22 2 11 13" }), /* @__PURE__ */ React.createElement("path", { d: "M22 2 15 22l-4-9-9-4Z" }))
  };
  return /* @__PURE__ */ React.createElement("svg", { ...p }, paths[name] || null);
}
const EMPTY = { accounts: [], tasks: [], strategies: [], dumps: [], people: [], roles: [], team: [], ownerPin: "", templates: [], projects: [] };
const STORE_KEY = "salesdesk_v1";
async function copyText(t) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(t);
      return true;
    }
  } catch (e) {
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = t;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    ta.style.top = "0";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    const ok = document.execCommand("copy");
    document.body.removeChild(ta);
    return ok;
  } catch (e) {
    return false;
  }
}
function buildOutbox(db) {
  const team = db.team || [];
  const owner = team.find((m) => m.isOwner) || null;
  const ownerEmail = owner && owner.email ? owner.email.trim() : "";
  const accName = {}, accPrio = {};
  (db.accounts || []).forEach((a) => {
    accName[a.id] = a.name;
    accPrio[a.id] = a.priority || "b";
  });
  const memById = {};
  team.forEach((m) => {
    memById[m.id] = m;
  });
  const firstEmail = (team.find((m) => m.email && m.email.trim()) || {}).email || "";
  const fallback = ownerEmail || (firstEmail ? firstEmail.trim() : "");
  const projName = {};
  (db.projects || []).forEach((p) => {
    projName[p.id] = p.name;
  });
  const doneIds = new Set((db.tasks || []).filter((t) => t.done).map((t) => t.id));
  const openTasks = (db.tasks || []).filter((t) => !t.done && t.due && !(t.dependsOn && !doneIds.has(t.dependsOn) && (db.tasks || []).some((x) => x.id === t.dependsOn))).map((t) => {
    const asg = t.assigneeId ? memById[t.assigneeId] : null;
    const asgEmail = asg && asg.email ? asg.email.trim() : "";
    const context = accName[t.accountId] || (t.projectId ? projName[t.projectId] : null) || "(none)";
    return {
      title: t.title,
      due: t.due,
      accountName: context,
      priority: t.accountId ? accPrio[t.accountId] || "b" : "b",
      assigneeName: asg ? asg.name : "",
      assigneeEmail: asgEmail,
      routeTo: asgEmail || fallback
    };
  });
  const recipients = [];
  const seen = /* @__PURE__ */ new Set();
  team.forEach((m) => {
    const e = m.email && m.email.trim();
    if (e && !seen.has(e.toLowerCase())) {
      recipients.push({ name: m.name, email: e, isOwner: !!m.isOwner });
      seen.add(e.toLowerCase());
    }
  });
  return {
    note: "Auto-generated for the Power Automate email flow. Do not edit by hand \u2014 SalesDesk overwrites this on every save.",
    generatedAt: todayStr(),
    recipients,
    openTasks
  };
}
function serialize(db) {
  const clone = { ...db };
  delete clone.outbox;
  clone.outbox = buildOutbox(db);
  return JSON.stringify(clone, null, 2);
}
function hydrate(text) {
  if (!text || !text.trim()) return { ...EMPTY };
  const parsed = JSON.parse(text);
  delete parsed.outbox;
  return { ...EMPTY, ...parsed };
}
function Btn({ children, onClick, kind = "primary", small, style, title, type, icon }) {
  const base = {
    fontFamily: FONT_BODY,
    fontWeight: 600,
    fontSize: small ? 12.5 : 13.5,
    padding: small ? "5px 10px" : "8px 14px",
    borderRadius: 8,
    cursor: "pointer",
    border: "1px solid transparent",
    whiteSpace: "nowrap",
    display: "inline-flex",
    alignItems: "center",
    gap: 6
  };
  const kinds = {
    primary: { background: C.green, color: "#fff" },
    ghost: { background: "transparent", color: C.inkSoft, border: `1px solid ${C.line}` },
    danger: { background: "transparent", color: C.red, border: `1px solid ${C.line}` },
    subtle: { background: C.chip, color: C.ink }
  };
  return /* @__PURE__ */ React.createElement("button", { className: "sd-btn", type: type || "button", title, onClick, style: { ...base, ...kinds[kind], ...style } }, icon && /* @__PURE__ */ React.createElement(Icon, { name: icon, size: small ? 13 : 15 }), children);
}
function Check({ done, onToggle }) {
  return /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: onToggle,
      "aria-label": done ? "Mark as not done" : "Mark as done",
      style: {
        width: 20,
        height: 20,
        minWidth: 20,
        borderRadius: 6,
        cursor: "pointer",
        border: `2px solid ${done ? C.green : "#B9C4BC"}`,
        background: done ? C.green : "transparent",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        transition: "all 0.15s"
      }
    },
    done && /* @__PURE__ */ React.createElement("svg", { width: "11", height: "11", viewBox: "0 0 12 12", fill: "none" }, /* @__PURE__ */ React.createElement("path", { d: "M2 6.5L4.8 9.2L10 3.5", stroke: "#fff", strokeWidth: "2.2", strokeLinecap: "round", strokeLinejoin: "round" }))
  );
}
function DuePill({ due, done }) {
  const diff = daysFromToday(due);
  let bg = C.chip, fg = C.inkSoft;
  if (!done && due) {
    if (diff < 0) {
      bg = C.redSoft;
      fg = C.red;
    } else if (diff === 0) {
      bg = C.amberSoft;
      fg = C.amber;
    } else if (diff <= 7) {
      bg = C.greenSoft;
      fg = C.green;
    }
  }
  return /* @__PURE__ */ React.createElement("span", { style: {
    fontFamily: FONT_MONO,
    fontSize: 11,
    fontWeight: 500,
    padding: "3px 8px",
    borderRadius: 6,
    background: bg,
    color: fg,
    whiteSpace: "nowrap"
  } }, done ? "done" : diff !== null && diff < 0 ? `${fmtDate(due)} \xB7 ${-diff}d late` : fmtDate(due));
}
function Empty({ title, body }) {
  return /* @__PURE__ */ React.createElement("div", { style: { padding: "42px 24px", textAlign: "center", color: C.inkSoft } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontSize: 16, fontWeight: 600, color: C.ink, marginBottom: 6 } }, title), /* @__PURE__ */ React.createElement("div", { style: { fontSize: 13.5, maxWidth: 380, margin: "0 auto", lineHeight: 1.5 } }, body));
}
const inputStyle = {
  fontFamily: FONT_BODY,
  fontSize: 13.5,
  padding: "8px 11px",
  borderRadius: 8,
  border: `1px solid ${C.line}`,
  background: "#fff",
  color: C.ink,
  outline: "none",
  width: "100%"
};
const REPEAT_OPTS = [
  { v: 0, label: "Doesn't repeat" },
  { v: 7, label: "Every week" },
  { v: 14, label: "Every 2 weeks" },
  { v: 30, label: "Every month" },
  { v: 60, label: "Every 2 months" },
  { v: 90, label: "Every quarter" }
];
const repeatLabel = (d) => (REPEAT_OPTS.find((o) => o.v === d) || { label: `Every ${d}d` }).label;
function blockedByTask(task, taskById) {
  if (!task || !task.dependsOn) return null;
  const pre = taskById[task.dependsOn];
  return pre && !pre.done ? pre : null;
}
function TaskRow({ task, account, project, strategy, onToggle, onUpdate, onDelete, onSnooze, onGoAccount, onGoProject, showAccount = true, team = [], canAssign = false, siblings = null, blockedBy = null }) {
  const [open, setOpen] = useState(false);
  const assignee = team.find((m) => m.id === task.assigneeId);
  const clickableAccount = showAccount && account && onGoAccount;
  const clickableProject = showAccount && project && onGoProject;
  const subs = task.subtasks || [];
  const subsDone = subs.filter((s) => s.done).length;
  return /* @__PURE__ */ React.createElement("div", { className: "sd-row", style: { borderBottom: `1px solid ${C.line}`, borderRadius: 6, opacity: blockedBy ? 0.62 : 1 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 10, padding: "10px 6px" } }, /* @__PURE__ */ React.createElement(Check, { done: task.done, onToggle }), /* @__PURE__ */ React.createElement(
    "div",
    {
      onClick: () => setOpen(!open),
      style: { flex: 1, cursor: "pointer", minWidth: 0 }
    },
    /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 6 } }, task.repeatDays ? /* @__PURE__ */ React.createElement("span", { style: { color: C.brass, display: "flex", flex: "none" }, title: repeatLabel(task.repeatDays) }, /* @__PURE__ */ React.createElement(Icon, { name: "repeat", size: 13 })) : null, /* @__PURE__ */ React.createElement("span", { style: {
      fontSize: 14,
      fontWeight: 500,
      color: task.done ? C.inkSoft : C.ink,
      textDecoration: task.done ? "line-through" : "none",
      overflow: "hidden",
      textOverflow: "ellipsis",
      whiteSpace: "nowrap"
    } }, task.title)),
    /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center", marginTop: 3, flexWrap: "wrap" } }, blockedBy && /* @__PURE__ */ React.createElement(
      "span",
      {
        style: { fontSize: 11, fontWeight: 600, color: C.brass, background: C.brassSoft, padding: "2px 8px", borderRadius: 6, display: "inline-flex", alignItems: "center", gap: 4 },
        title: `Starts after: ${blockedBy.title}`
      },
      /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 11 }),
      " waiting on: ",
      blockedBy.title
    ), showAccount && account && /* @__PURE__ */ React.createElement(
      "span",
      {
        className: clickableAccount ? "sd-link" : void 0,
        onClick: clickableAccount ? (e) => {
          e.stopPropagation();
          onGoAccount(account.id);
        } : void 0,
        style: {
          fontSize: 11.5,
          fontWeight: 600,
          color: C.green,
          background: C.greenSoft,
          padding: "2px 8px",
          borderRadius: 6,
          cursor: clickableAccount ? "pointer" : "default"
        }
      },
      account.name
    ), showAccount && project && /* @__PURE__ */ React.createElement(
      "span",
      {
        className: clickableProject ? "sd-link" : void 0,
        onClick: clickableProject ? (e) => {
          e.stopPropagation();
          onGoProject(project.id);
        } : void 0,
        style: {
          fontSize: 11.5,
          fontWeight: 600,
          color: C.violet,
          background: C.violetSoft,
          padding: "2px 8px",
          borderRadius: 6,
          cursor: clickableProject ? "pointer" : "default",
          display: "inline-flex",
          alignItems: "center",
          gap: 4
        }
      },
      /* @__PURE__ */ React.createElement(Icon, { name: "project", size: 11 }),
      " ",
      project.name
    ), assignee && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11.5, fontWeight: 600, color: C.brass, background: C.brassSoft, padding: "2px 8px", borderRadius: 6 } }, "@", assignee.name), strategy && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11, color: C.inkSoft, fontFamily: FONT_MONO } }, "\u25B8 ", strategy.name), subs.length > 0 && !open && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11, fontWeight: 600, color: subsDone === subs.length ? C.green : C.inkSoft, background: C.chip, padding: "2px 7px", borderRadius: 6, display: "inline-flex", alignItems: "center", gap: 3 } }, "\u2611 ", subsDone, "/", subs.length), task.links && task.links.length > 0 && !open && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11, fontWeight: 600, color: C.green, background: C.greenSoft, padding: "2px 7px", borderRadius: 6, display: "inline-flex", alignItems: "center", gap: 3 } }, /* @__PURE__ */ React.createElement(Icon, { name: "mail", size: 11 }), " ", task.links.length), task.table && (task.table.rows || []).length > 0 && !open && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11, fontWeight: 600, color: C.inkSoft, background: C.chip, padding: "2px 7px", borderRadius: 6 } }, "\u25A6 ", (task.table.rows || []).length, "\xD7", (task.table.columns || []).length), task.notes && !open && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11.5, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: 260 } }, "\u{1F4DD} ", task.notes))
  ), !task.done && onSnooze && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 3 } }, /* @__PURE__ */ React.createElement(SnoozeBtn, { label: "+1d", onClick: () => onSnooze(task.id, 1) }), /* @__PURE__ */ React.createElement(SnoozeBtn, { label: "+3d", onClick: () => onSnooze(task.id, 3) }), /* @__PURE__ */ React.createElement(SnoozeBtn, { label: "+1w", onClick: () => onSnooze(task.id, 7) })), /* @__PURE__ */ React.createElement(DuePill, { due: task.due, done: task.done }), /* @__PURE__ */ React.createElement("button", { onClick: () => setOpen(!open), "aria-label": "Expand task", style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 12, padding: 4 } }, open ? "\u25B4" : "\u25BE")), open && /* @__PURE__ */ React.createElement("div", { style: { padding: "0 6px 14px 34px", display: "flex", flexDirection: "column", gap: 8 } }, /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: task.notes || "",
      onChange: (e) => onUpdate({ notes: e.target.value }),
      placeholder: "Notes / outcome \u2014 what happened when you reached out?",
      rows: 2,
      style: { ...inputStyle, resize: "vertical", fontSize: 13 }
    }
  ), /* @__PURE__ */ React.createElement(SubtaskList, { subtasks: subs, onChange: (subtasks) => onUpdate({ subtasks }) }), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, color: C.inkSoft } }, "Due"), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "date",
      value: task.due || "",
      onChange: (e) => onUpdate({ due: e.target.value }),
      style: { ...inputStyle, width: 150, padding: "5px 8px", fontSize: 12.5 }
    }
  ), /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, color: C.inkSoft, marginLeft: 4, display: "inline-flex", alignItems: "center", gap: 4 } }, /* @__PURE__ */ React.createElement(Icon, { name: "repeat", size: 13 }), " Repeat"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: task.repeatDays || 0,
      onChange: (e) => onUpdate({ repeatDays: Number(e.target.value) || null }),
      style: { ...inputStyle, width: "auto", minWidth: 130, padding: "5px 8px", fontSize: 12.5 }
    },
    REPEAT_OPTS.map((o) => /* @__PURE__ */ React.createElement("option", { key: o.v, value: o.v }, o.label))
  ), canAssign && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, color: C.inkSoft, marginLeft: 4 } }, "Assignee"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: task.assigneeId || "",
      onChange: (e) => onUpdate({ assigneeId: e.target.value || null }),
      style: { ...inputStyle, width: "auto", minWidth: 130, padding: "5px 8px", fontSize: 12.5 }
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "Unassigned"),
    team.map((m) => /* @__PURE__ */ React.createElement("option", { key: m.id, value: m.id }, m.name))
  )), /* @__PURE__ */ React.createElement("div", { style: { flex: 1 } }), onDelete && /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: onDelete }, "Delete")), siblings && siblings.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, color: C.inkSoft, display: "inline-flex", alignItems: "center", gap: 4 } }, /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 13 }), " Start only after"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: task.dependsOn || "",
      onChange: (e) => onUpdate({ dependsOn: e.target.value || null }),
      style: { ...inputStyle, width: "auto", minWidth: 170, padding: "5px 8px", fontSize: 12.5 }
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "Nothing \u2014 always active"),
    siblings.filter((s) => s.id !== task.id).map((s) => /* @__PURE__ */ React.createElement("option", { key: s.id, value: s.id }, s.title, s.done ? " \u2713" : ""))
  )), blockedBy && /* @__PURE__ */ React.createElement("div", { style: { fontSize: 11.5, color: C.brass, fontFamily: FONT_MONO } }, '\u23F3 Dormant until "', blockedBy.title, '" is completed \u2014 then it activates and appears on Today.'), task.repeatDays ? /* @__PURE__ */ React.createElement("div", { style: { fontSize: 11.5, color: C.brass, fontFamily: FONT_MONO } }, "\u21BB When you complete this, the next one is created automatically for ", repeatLabel(task.repeatDays).toLowerCase(), ".") : null, /* @__PURE__ */ React.createElement(TaskLinks, { links: task.links, onChange: (links) => onUpdate({ links }) }), /* @__PURE__ */ React.createElement(TaskTable, { table: task.table, onChange: (table) => onUpdate({ table }) })));
}
function SubtaskList({ subtasks, onChange }) {
  const [title, setTitle] = useState("");
  const list = subtasks || [];
  const add = () => {
    if (!title.trim()) return;
    onChange([...list, { id: Math.random().toString(36).slice(2, 9), title: title.trim(), done: false }]);
    setTitle("");
  };
  const toggle = (id) => onChange(list.map((s) => s.id === id ? { ...s, done: !s.done } : s));
  const edit = (id, t) => onChange(list.map((s) => s.id === id ? { ...s, title: t } : s));
  const remove = (id) => onChange(list.filter((s) => s.id !== id));
  return /* @__PURE__ */ React.createElement("div", { style: { marginTop: 2 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 6, marginBottom: 6 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11.5, fontWeight: 600, color: C.inkSoft, textTransform: "uppercase", letterSpacing: "0.04em" } }, "Subtasks"), list.length > 0 && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft } }, list.filter((s) => s.done).length, "/", list.length)), list.map((s) => /* @__PURE__ */ React.createElement("div", { key: s.id, style: { display: "flex", alignItems: "center", gap: 8, padding: "3px 0" } }, /* @__PURE__ */ React.createElement(Check, { done: s.done, onToggle: () => toggle(s.id) }), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: s.title,
      onChange: (e) => edit(s.id, e.target.value),
      style: {
        ...inputStyle,
        border: "none",
        background: "transparent",
        padding: "2px 0",
        fontSize: 13,
        flex: 1,
        color: s.done ? C.inkSoft : C.ink,
        textDecoration: s.done ? "line-through" : "none"
      }
    }
  ), /* @__PURE__ */ React.createElement("button", { onClick: () => remove(s.id), title: "Remove subtask", style: { background: "none", border: "none", cursor: "pointer", color: "#B9C4BC", fontSize: 14, padding: "0 4px" } }, "\xD7"))), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 6, marginTop: 4 } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: title,
      onChange: (e) => setTitle(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && add(),
      placeholder: "Add a subtask\u2026",
      style: { ...inputStyle, fontSize: 12.5, padding: "6px 9px" }
    }
  ), /* @__PURE__ */ React.createElement(Btn, { small: true, onClick: add }, "Add")));
}
function SnoozeBtn({ label, onClick }) {
  return /* @__PURE__ */ React.createElement(
    "button",
    {
      className: "sd-btn",
      onClick,
      title: `Push due date ${label}`,
      style: {
        fontFamily: FONT_MONO,
        fontSize: 10.5,
        fontWeight: 500,
        color: C.inkSoft,
        background: "transparent",
        border: `1px solid ${C.line}`,
        borderRadius: 5,
        padding: "2px 6px",
        cursor: "pointer"
      }
    },
    label
  );
}
function autoWidth(v, min = 4, max = 44) {
  return `${Math.min(max, Math.max(min, (v || "").length + 1))}ch`;
}
function TaskTable({ table, onChange }) {
  const cellStyle = {
    border: "none",
    background: "transparent",
    outline: "none",
    fontFamily: FONT_BODY,
    fontSize: 13,
    color: C.ink,
    padding: "6px 9px",
    boxSizing: "content-box"
  };
  const cellBox = { border: `1px solid ${C.line}`, padding: 0, verticalAlign: "top", background: "#fff" };
  if (!table) {
    return /* @__PURE__ */ React.createElement(
      "button",
      {
        className: "sd-link",
        onClick: () => onChange({ columns: ["Item", "Qty"], rows: [["", ""]] }),
        style: { background: "none", border: "none", cursor: "pointer", color: C.green, fontSize: 12.5, fontWeight: 600, padding: 0, alignSelf: "flex-start" }
      },
      "+ Add a table"
    );
  }
  const cols = table.columns || [];
  const rows = table.rows || [];
  const setColumn = (ci, name) => {
    const columns = cols.slice();
    columns[ci] = name;
    onChange({ columns, rows });
  };
  const setCell = (ri, ci, val) => {
    const nrows = rows.map((r) => r.slice());
    nrows[ri][ci] = val;
    onChange({ columns: cols, rows: nrows });
  };
  const addColumn = () => onChange({ columns: [...cols, `Column ${cols.length + 1}`], rows: rows.map((r) => [...r, ""]) });
  const removeColumn = (ci) => {
    if (cols.length <= 1) {
      onChange(null);
      return;
    }
    onChange({ columns: cols.filter((_, i) => i !== ci), rows: rows.map((r) => r.filter((_, i) => i !== ci)) });
  };
  const addRow = () => onChange({ columns: cols, rows: [...rows, cols.map(() => "")] });
  const removeRow = (ri) => {
    if (rows.length <= 1) {
      onChange(null);
      return;
    }
    onChange({ columns: cols, rows: rows.filter((_, i) => i !== ri) });
  };
  const colWidth = (ci) => {
    let longest = (cols[ci] || "").length;
    rows.forEach((r) => {
      longest = Math.max(longest, (r[ci] || "").length);
    });
    return autoWidth("x".repeat(longest));
  };
  return /* @__PURE__ */ React.createElement("div", { style: { marginTop: 4 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 6 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11.5, fontWeight: 600, color: C.inkSoft, textTransform: "uppercase", letterSpacing: "0.04em" } }, "Table"), /* @__PURE__ */ React.createElement("button", { className: "sd-link", onClick: () => onChange(null), style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 11.5, padding: 0 } }, "remove")), /* @__PURE__ */ React.createElement("div", { style: { overflowX: "auto", paddingBottom: 2 } }, /* @__PURE__ */ React.createElement("table", { style: { borderCollapse: "collapse", tableLayout: "auto" } }, /* @__PURE__ */ React.createElement("thead", null, /* @__PURE__ */ React.createElement("tr", null, cols.map((c, ci) => /* @__PURE__ */ React.createElement("th", { key: ci, style: { ...cellBox, background: C.chip, position: "relative" } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: c,
      onChange: (e) => setColumn(ci, e.target.value),
      placeholder: `Column ${ci + 1}`,
      style: { ...cellStyle, fontWeight: 600, fontSize: 12.5, color: C.ink, width: colWidth(ci) }
    }
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => removeColumn(ci),
      title: "Delete column",
      style: { background: "none", border: "none", cursor: "pointer", color: "#B9C4BC", fontSize: 12, padding: "0 5px 0 0" }
    },
    "\xD7"
  )))), /* @__PURE__ */ React.createElement("th", { style: { border: "none", padding: "0 0 0 4px" } }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: addColumn,
      title: "Add column",
      className: "sd-btn",
      style: { background: "transparent", border: `1px dashed ${C.line}`, borderRadius: 6, cursor: "pointer", color: C.inkSoft, fontSize: 13, padding: "5px 9px", whiteSpace: "nowrap" }
    },
    "+ col"
  )))), /* @__PURE__ */ React.createElement("tbody", null, rows.map((r, ri) => /* @__PURE__ */ React.createElement("tr", { key: ri }, cols.map((_, ci) => /* @__PURE__ */ React.createElement("td", { key: ci, style: cellBox }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: r[ci] || "",
      onChange: (e) => setCell(ri, ci, e.target.value),
      style: { ...cellStyle, width: colWidth(ci) }
    }
  ))), /* @__PURE__ */ React.createElement("td", { style: { border: "none", padding: "0 0 0 4px" } }, /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => removeRow(ri),
      title: "Delete row",
      style: { background: "none", border: "none", cursor: "pointer", color: "#B9C4BC", fontSize: 14, padding: "2px 4px" }
    },
    "\xD7"
  ))))))), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: addRow,
      className: "sd-btn",
      style: { marginTop: 6, background: "transparent", border: `1px dashed ${C.line}`, borderRadius: 6, cursor: "pointer", color: C.inkSoft, fontSize: 12.5, fontWeight: 600, padding: "5px 11px" }
    },
    "+ Add row"
  ));
}
function normUrl(u) {
  u = (u || "").trim();
  if (!u) return "";
  return /^https?:\/\//i.test(u) ? u : "https://" + u;
}
function outlookSearchQuery(ref) {
  const parts = [];
  if (ref.subject) parts.push(`subject:"${ref.subject.replace(/^(re|fw|fwd):\s*/i, "").trim()}"`);
  if (ref.sender) parts.push(`from:${ref.sender}`);
  return parts.join(" ");
}
function outlookSearchUrl(ref) {
  return `https://outlook.office.com/mail/deeplink/search?query=${encodeURIComponent(outlookSearchQuery(ref))}`;
}
function TaskLinks({ links, onChange }) {
  const [subject, setSubject] = useState("");
  const [sender, setSender] = useState("");
  const [url, setUrl] = useState("");
  const [copiedId, setCopiedId] = useState(null);
  const list = links || [];
  const add = () => {
    const s = subject.trim(), f = sender.trim(), u = normUrl(url);
    if (!s && !u) return;
    onChange([...list, { id: Math.random().toString(36).slice(2, 9), subject: s, sender: f, url: u }]);
    setSubject("");
    setSender("");
    setUrl("");
  };
  const remove = (id) => onChange(list.filter((l) => l.id !== id));
  const openUrl = (u) => {
    try {
      window.open(u, "_blank", "noopener,noreferrer");
    } catch (e) {
    }
  };
  const copySearch = async (l) => {
    const ok = await copyText(outlookSearchQuery(l));
    if (ok) {
      setCopiedId(l.id);
      setTimeout(() => setCopiedId((c) => c === l.id ? null : c), 1800);
    }
  };
  const display = (l) => l.subject || l.label || "Email thread";
  return /* @__PURE__ */ React.createElement("div", { style: { marginTop: 4 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 6, marginBottom: 6 } }, /* @__PURE__ */ React.createElement("span", { style: { color: C.inkSoft, display: "flex" } }, /* @__PURE__ */ React.createElement(Icon, { name: "mail", size: 13 })), /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11.5, fontWeight: 600, color: C.inkSoft, textTransform: "uppercase", letterSpacing: "0.04em" } }, "Email threads")), list.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: 7, marginBottom: 9 } }, list.map((l) => /* @__PURE__ */ React.createElement("div", { key: l.id, style: { display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "span",
    {
      style: { fontSize: 12.5, fontWeight: 600, color: C.ink, maxWidth: 300, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
      title: l.sender ? `${display(l)} \xB7 from ${l.sender}` : display(l)
    },
    display(l)
  ), l.subject && /* @__PURE__ */ React.createElement(
    "button",
    {
      className: "sd-btn",
      onClick: () => copySearch(l),
      title: "Copy the search terms, then paste into the Outlook search box",
      style: {
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        fontFamily: FONT_BODY,
        fontSize: 12,
        fontWeight: 600,
        color: copiedId === l.id ? "#fff" : C.green,
        background: copiedId === l.id ? C.green : C.greenSoft,
        border: "1px solid transparent",
        borderRadius: 7,
        padding: "4px 10px",
        cursor: "pointer"
      }
    },
    /* @__PURE__ */ React.createElement(Icon, { name: copiedId === l.id ? "today" : "search", size: 12 }),
    " ",
    copiedId === l.id ? "Copied \u2014 paste in Outlook" : "Copy search"
  ), l.url && /* @__PURE__ */ React.createElement(
    "button",
    {
      className: "sd-btn",
      onClick: () => openUrl(l.url),
      title: l.url,
      style: {
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        fontFamily: FONT_BODY,
        fontSize: 12,
        fontWeight: 600,
        color: C.inkSoft,
        background: "transparent",
        border: `1px solid ${C.line}`,
        borderRadius: 7,
        padding: "4px 10px",
        cursor: "pointer"
      }
    },
    /* @__PURE__ */ React.createElement(Icon, { name: "mail", size: 12 }),
    " Open link"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => remove(l.id),
      title: "Remove",
      style: { background: "none", border: "none", cursor: "pointer", color: "#B9C4BC", fontSize: 15, padding: "0 4px" }
    },
    "\xD7"
  )))), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 6, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: subject,
      onChange: (e) => setSubject(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && add(),
      placeholder: "Email subject (e.g. New Product Wishlist for A2Z)",
      style: { ...inputStyle, flex: "1 1 240px", fontSize: 12.5, padding: "6px 9px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: sender,
      onChange: (e) => setSender(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && add(),
      placeholder: "From (optional)",
      style: { ...inputStyle, flex: "0 1 150px", fontSize: 12.5, padding: "6px 9px" }
    }
  ), /* @__PURE__ */ React.createElement(Btn, { small: true, onClick: add }, "+ Thread")), /* @__PURE__ */ React.createElement("details", { style: { marginTop: 6 } }, /* @__PURE__ */ React.createElement("summary", { style: { fontSize: 11, color: C.inkSoft, cursor: "pointer" } }, "or paste a direct message link instead"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 6, marginTop: 6, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: url,
      onChange: (e) => setUrl(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && add(),
      placeholder: "https://outlook.office.com/mail/\u2026/id/\u2026",
      style: { ...inputStyle, flex: "1 1 260px", fontSize: 12.5, padding: "6px 9px" }
    }
  ))), /* @__PURE__ */ React.createElement("div", { style: { fontSize: 11, color: C.inkSoft, marginTop: 6, lineHeight: 1.45 } }, /* @__PURE__ */ React.createElement("b", null, "Copy search"), " puts the subject search on your clipboard \u2014 click into Outlook's search box and paste. It scopes to the subject line so it's fast and lands on the current thread."));
}
function SalesDesk() {
  const [db, setDb] = useState(EMPTY);
  const [phase, setPhase] = useState("boot");
  const [cfg, setCfg] = useState(loadCfg());
  const [userName, setUserName] = useState("");
  const [saveState, setSaveState] = useState("idle");
  const [online, setOnline] = useState(navigator.onLine);
  const [errMsg, setErrMsg] = useState("");
  const [page, setPage] = useState("today");
  const [accountId, setAccountId] = useState(null);
  const [projectSel, setProjectSel] = useState(null);
  const [meId, setMeId] = useState(void 0);
  const [identityLoaded, setIdentityLoaded] = useState(false);
  const skipSave = useRef(true);
  const importInputRef = useRef(null);
  const fsMode = "graph";
  const fileName = (cfg.path || "").split("/").pop() || "salesdesk-data.json";
  useEffect(() => {
    const on = () => setOnline(true), off = () => setOnline(false);
    window.addEventListener("online", on);
    window.addEventListener("offline", off);
    return () => {
      window.removeEventListener("online", on);
      window.removeEventListener("offline", off);
    };
  }, []);
  useEffect(() => {
    (async () => {
      const c = loadCfg();
      if (!c.clientId || !c.path) {
        setPhase("config");
        return;
      }
      try {
        await msalInit(c.clientId);
      } catch (e) {
        setErrMsg("Couldn't start Microsoft sign-in: " + e.message);
        setPhase("config");
        return;
      }
      if (!currentAccount()) {
        setPhase("signin");
        return;
      }
      await loadFromCloud(c);
    })();
  }, []);
  const loadFromCloud = async (c) => {
    const acct = currentAccount();
    setUserName(acct ? acct.name || acct.username || "" : "");
    let cached = null;
    try {
      cached = await idbGet(CACHE_KEY);
    } catch (e) {
    }
    if (cached) {
      skipSave.current = true;
      setDb(hydrate(cached));
      setPhase("ready");
    }
    if (!navigator.onLine) {
      if (!cached) {
        skipSave.current = true;
        setDb(EMPTY);
      }
      setSaveState("offline");
      setPhase("ready");
      return;
    }
    try {
      const text = await graphReadFile(c.path);
      const dirty = await idbGet(DIRTY_KEY).catch(() => false);
      if (dirty && cached) {
        setPhase("ready");
        try {
          await graphWriteFile(c.path, cached);
          await idbSet(DIRTY_KEY, false);
          setSaveState("saved");
        } catch (e) {
        }
        return;
      }
      skipSave.current = true;
      setDb(text ? hydrate(text) : EMPTY);
      try {
        await idbSet(CACHE_KEY, text || JSON.stringify(EMPTY));
      } catch (e) {
      }
      setPhase("ready");
    } catch (e) {
      if (e.message === "redirecting") return;
      setErrMsg(e.message);
      if (!cached) {
        skipSave.current = true;
        setDb(EMPTY);
      }
      setPhase("ready");
    }
  };
  const signIn = async () => {
    try {
      await msalInit(cfg.clientId);
      await msalApp.loginRedirect({ scopes: SCOPES });
    } catch (e) {
      setErrMsg(e.message);
    }
  };
  const signOut = async () => {
    try {
      await idbDel(CACHE_KEY);
    } catch (e) {
    }
    try {
      await msalApp.logoutRedirect();
    } catch (e) {
      setPhase("signin");
    }
  };
  useEffect(() => {
    if (phase !== "ready") return;
    if (skipSave.current) {
      skipSave.current = false;
      return;
    }
    setSaveState("saving");
    const t = setTimeout(async () => {
      const text = serialize(db);
      try {
        await idbSet(CACHE_KEY, text);
      } catch (e) {
      }
      if (!navigator.onLine) {
        try {
          await idbSet(DIRTY_KEY, true);
        } catch (e) {
        }
        setSaveState("offline");
        return;
      }
      try {
        await graphWriteFile(cfg.path, text);
        try {
          await idbSet(DIRTY_KEY, false);
        } catch (e) {
        }
        setSaveState("saved");
        setTimeout(() => setSaveState((s) => s === "saved" ? "idle" : s), 1500);
      } catch (e) {
        if (e.message === "redirecting") return;
        try {
          await idbSet(DIRTY_KEY, true);
        } catch (e2) {
        }
        setSaveState("error");
        setErrMsg(e.message);
      }
    }, 900);
    return () => clearTimeout(t);
  }, [db, phase]);
  useEffect(() => {
    if (phase !== "ready" || !online) return;
    (async () => {
      const dirty = await idbGet(DIRTY_KEY).catch(() => false);
      if (!dirty) return;
      try {
        const cached = await idbGet(CACHE_KEY);
        if (cached) {
          await graphWriteFile(cfg.path, cached);
          await idbSet(DIRTY_KEY, false);
          setSaveState("saved");
        }
      } catch (e) {
      }
    })();
  }, [online, phase]);
  const refresh = async () => {
    setSaveState("saving");
    try {
      const text = await graphReadFile(cfg.path);
      skipSave.current = true;
      setDb(text ? hydrate(text) : EMPTY);
      try {
        await idbSet(CACHE_KEY, text || JSON.stringify(EMPTY));
      } catch (e) {
      }
      setSaveState("saved");
      setTimeout(() => setSaveState("idle"), 1200);
    } catch (e) {
      if (e.message !== "redirecting") {
        setSaveState("error");
        setErrMsg(e.message);
      }
    }
  };
  const exportJson = () => {
    const blob = new Blob([serialize(db)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "salesdesk-data.json";
    a.click();
    URL.revokeObjectURL(url);
  };
  const importJson = (file) => {
    const r = new FileReader();
    r.onload = () => {
      try {
        setDb(hydrate(r.result));
      } catch (e) {
        alert("That file doesn't look like SalesDesk data.");
      }
    };
    r.readAsText(file);
  };
  const switchFile = () => {
    setPhase("config");
  };
  const mut = (fn) => setDb((d) => fn(JSON.parse(JSON.stringify(d))));
  const addAccount = (name, contact, priority = "b", stage = "customer") => {
    const a = { id: uid(), name: name.trim(), contact: (contact || "").trim(), notes: "", priority, stage, cadenceDays: null, lastContact: null, createdAt: todayStr() };
    mut((d) => {
      d.accounts.push(a);
      return d;
    });
    return a.id;
  };
  const updateAccount = (id, patch) => mut((d) => {
    const a = d.accounts.find((x) => x.id === id);
    if (a) Object.assign(a, patch);
    return d;
  });
  const deleteAccount = (id) => {
    if (!window.confirm("Delete this account and all its tasks?")) return;
    mut((d) => {
      d.accounts = d.accounts.filter((a) => a.id !== id);
      d.tasks = d.tasks.filter((t) => t.accountId !== id);
      d.people = (d.people || []).filter((p) => p.accountId !== id);
      d.strategies.forEach((s) => {
        s.accountIds = s.accountIds.filter((x) => x !== id);
      });
      return d;
    });
    setAccountId(null);
  };
  const rememberRole = (d, role) => {
    if (role && !(d.roles || []).some((r) => r.toLowerCase() === role.toLowerCase())) {
      d.roles = [...d.roles || [], role].sort((a, b) => a.localeCompare(b));
    }
  };
  const addPerson = (p) => mut((d) => {
    const role = (p.role || "").trim();
    rememberRole(d, role);
    d.people = d.people || [];
    d.people.push({ id: uid(), accountId: p.accountId, name: p.name.trim(), role, email: (p.email || "").trim(), notes: "", met: !!p.met, managerId: p.managerId || null });
    return d;
  });
  const updatePerson = (id, patch) => mut((d) => {
    const p = (d.people || []).find((x) => x.id === id);
    if (!p) return d;
    if (patch.role !== void 0) {
      patch = { ...patch, role: patch.role.trim() };
      rememberRole(d, patch.role);
    }
    Object.assign(p, patch);
    return d;
  });
  const deletePerson = (id) => mut((d) => {
    const p = (d.people || []).find((x) => x.id === id);
    (d.people || []).forEach((c) => {
      if (c.managerId === id) c.managerId = p ? p.managerId : null;
    });
    d.people = (d.people || []).filter((x) => x.id !== id);
    return d;
  });
  const addTask = (t) => mut((d) => {
    d.tasks.push({ id: uid(), done: false, notes: "", createdAt: todayStr(), ...t });
    return d;
  });
  const updateTask = (id, patch) => mut((d) => {
    const t = d.tasks.find((x) => x.id === id);
    if (t) Object.assign(t, patch);
    return d;
  });
  const toggleTask = (id) => mut((d) => {
    const t = d.tasks.find((x) => x.id === id);
    if (!t) return d;
    t.done = !t.done;
    t.completedAt = t.done ? todayStr() : null;
    if (t.done && t.repeatDays) {
      const already = d.tasks.some((x) => x.spawnedFrom === t.id && !x.done);
      if (!already) {
        d.tasks.push({
          id: uid(),
          accountId: t.accountId || null,
          projectId: t.projectId || null,
          strategyId: t.strategyId || null,
          title: t.title,
          notes: "",
          done: false,
          assigneeId: t.assigneeId || null,
          repeatDays: t.repeatDays,
          spawnedFrom: t.id,
          table: t.table || null,
          subtasks: (t.subtasks || []).map((s) => ({ id: uid(), title: s.title, done: false })),
          due: addDaysIso(todayStr(), t.repeatDays),
          createdAt: todayStr()
        });
      }
    }
    return d;
  });
  const snoozeTask = (id, days) => mut((d) => {
    const t = d.tasks.find((x) => x.id === id);
    if (t) t.due = addDaysIso(t.due && daysFromToday(t.due) > 0 ? t.due : todayStr(), days);
    return d;
  });
  const deleteTask = (id) => mut((d) => {
    d.tasks = d.tasks.filter((t) => t.id !== id);
    return d;
  });
  const markContacted = (accountId2) => mut((d) => {
    const a = d.accounts.find((x) => x.id === accountId2);
    if (a) a.lastContact = todayStr();
    return d;
  });
  const addTemplate = (tpl) => {
    const id = uid();
    mut((d) => {
      d.templates = d.templates || [];
      d.templates.push({ id, name: tpl.name.trim() || "Untitled", subject: tpl.subject || "", body: tpl.body || "" });
      return d;
    });
    return id;
  };
  const updateTemplate = (id, patch) => mut((d) => {
    const t = (d.templates || []).find((x) => x.id === id);
    if (t) Object.assign(t, patch);
    return d;
  });
  const deleteTemplate = (id) => mut((d) => {
    d.templates = (d.templates || []).filter((t) => t.id !== id);
    return d;
  });
  const addDump = (text) => mut((d) => {
    d.dumps.unshift({ id: uid(), text: text.trim(), createdAt: todayStr() });
    return d;
  });
  const deleteDump = (id) => mut((d) => {
    d.dumps = d.dumps.filter((x) => x.id !== id);
    return d;
  });
  const addProject = (name) => {
    const id = uid();
    mut((d) => {
      d.projects = d.projects || [];
      d.projects.push({ id, name: name.trim(), notes: "", createdAt: todayStr() });
      return d;
    });
    return id;
  };
  const updateProject = (id, patch) => mut((d) => {
    const p = (d.projects || []).find((x) => x.id === id);
    if (p) Object.assign(p, patch);
    return d;
  });
  const deleteProject = (id) => {
    if (!window.confirm("Delete this project and all its tasks?")) return;
    mut((d) => {
      d.projects = (d.projects || []).filter((p) => p.id !== id);
      d.tasks = d.tasks.filter((t) => t.projectId !== id);
      return d;
    });
    setProjectSel(null);
  };
  const addMember = (name, isOwner2 = false, email = "") => {
    const id = uid();
    mut((d) => {
      d.team = d.team || [];
      d.team.push({ id, name: name.trim(), isOwner: !!isOwner2, email: (email || "").trim() });
      return d;
    });
    return id;
  };
  const updateMember = (id, patch) => mut((d) => {
    const m = (d.team || []).find((x) => x.id === id);
    if (m) Object.assign(m, patch);
    return d;
  });
  const deleteMember = (id) => mut((d) => {
    d.team = (d.team || []).filter((m) => m.id !== id);
    d.tasks.forEach((t) => {
      if (t.assigneeId === id) t.assigneeId = null;
    });
    return d;
  });
  const setOwnerPin = (pin) => mut((d) => {
    d.ownerPin = pin || "";
    return d;
  });
  const addStrategy = ({ name, taskText, due, accountIds, assigneeId }) => {
    const sid = uid();
    mut((d) => {
      d.strategies.unshift({ id: sid, name: name.trim(), createdAt: todayStr(), accountIds: [...accountIds] });
      accountIds.forEach((aid) => {
        d.tasks.push({ id: uid(), accountId: aid, strategyId: sid, title: taskText.trim(), due, done: false, notes: "", assigneeId: assigneeId || null, createdAt: todayStr() });
      });
      return d;
    });
  };
  const deleteStrategy = (id) => {
    if (!window.confirm("Delete this strategy and its remaining open tasks? Completed tasks are kept in account history.")) return;
    mut((d) => {
      d.strategies = d.strategies.filter((s) => s.id !== id);
      d.tasks = d.tasks.filter((t) => !(t.strategyId === id && !t.done));
      return d;
    });
  };
  useEffect(() => {
    if (phase !== "ready") return;
    (async () => {
      try {
        const stored = await idbGet("identity");
        if (stored) setMeId(stored);
      } catch (e) {
      }
      setIdentityLoaded(true);
    })();
  }, [phase]);
  const chooseIdentity = async (id) => {
    setMeId(id);
    setPage("today");
    setAccountId(null);
    try {
      await idbSet("identity", id);
    } catch (e) {
    }
  };
  const clearIdentity = async () => {
    setMeId(void 0);
    try {
      await idbDel("identity");
    } catch (e) {
    }
  };
  const team = db.team || [];
  const meMember = team.find((m) => m.id === meId) || null;
  const isOwner = team.length === 0 || meMember && meMember.isOwner;
  const accById = useMemo(() => Object.fromEntries(db.accounts.map((a) => [a.id, a])), [db.accounts]);
  const projById = useMemo(() => Object.fromEntries((db.projects || []).map((p) => [p.id, p])), [db.projects]);
  const stratById = useMemo(() => Object.fromEntries(db.strategies.map((s) => [s.id, s])), [db.strategies]);
  const openTasks = db.tasks.filter((t) => !t.done);
  const goAccount = (id) => {
    setAccountId(id);
    setPage("accounts");
  };
  const goProject = (id) => {
    setProjectSel(id);
    setPage("projects");
  };
  const fontCss = `
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');
    * { box-sizing: border-box; }
    body { margin: 0; }
    input:focus, textarea:focus, select:focus { border-color: ${C.green} !important; box-shadow: 0 0 0 3px ${C.greenSoft}; }
    button:focus-visible, a:focus-visible { outline: 2px solid ${C.green}; outline-offset: 2px; }
    ::placeholder { color: #9AA69E; }
    .sd-icon { display: block; }
    .sd-card { transition: transform .15s ease, box-shadow .15s ease, border-color .15s ease; }
    .sd-card:hover { transform: translateY(-1px); box-shadow: 0 8px 22px rgba(18,40,30,.07); border-color: #D3DDD5; }
    .sd-btn { transition: filter .14s ease, background .14s ease, border-color .14s ease; }
    .sd-btn:hover { filter: brightness(.95); }
    .sd-nav { transition: background .12s ease, color .12s ease; }
    .sd-nav:hover { background: ${C.chip}; }
    .sd-row { transition: background .12s ease; }
    .sd-row:hover { background: #FAFBFA; }
    .sd-link { transition: color .12s ease; }
    @media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } .sd-card:hover { transform: none; } }
    @media (max-width: 840px) {
      .sd-shell { flex-direction: column !important; }
      .sd-rail { width: auto !important; min-width: 0 !important; height: auto !important; position: static !important; border-right: none !important; border-bottom: 1px solid ${C.line} !important; padding: 14px 14px 10px !important; }
      .sd-navwrap { flex-direction: row !important; flex-wrap: nowrap !important; overflow-x: auto !important; gap: 6px !important; -webkit-overflow-scrolling: touch; scrollbar-width: none; }
      .sd-navwrap::-webkit-scrollbar { display: none; }
      .sd-navwrap button { flex: 0 0 auto !important; padding: 8px 12px !important; }
      .sd-main { padding: 18px 16px 96px !important; }
      .sd-footer { flex-wrap: wrap !important; }
      input, textarea, select { font-size: 16px !important; }  /* stops iOS/Android zoom-on-focus */
    }
    @media (display-mode: standalone) {
      .sd-rail { padding-top: max(14px, env(safe-area-inset-top)) !important; }
      .sd-main { padding-bottom: calc(96px + env(safe-area-inset-bottom)) !important; }
    }
  `;
  if (phase === "boot") {
    return /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_BODY, padding: 60, textAlign: "center", color: C.inkSoft } }, "Loading your desk\u2026");
  }
  if (phase === "config") {
    return /* @__PURE__ */ React.createElement(
      ConfigScreen,
      {
        cfg,
        err: errMsg,
        fontCss,
        onSave: (c) => {
          saveCfg(c);
          setCfg(c);
          setErrMsg("");
          msalApp = null;
          setPhase("boot");
          setTimeout(() => window.location.reload(), 50);
        }
      }
    );
  }
  if (phase === "signin") {
    return /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_BODY, background: C.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 } }, /* @__PURE__ */ React.createElement("style", null, fontCss), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 16, padding: "34px 30px", maxWidth: 420, width: "100%", textAlign: "center" } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontWeight: 700, fontSize: 26, letterSpacing: "-0.02em", marginBottom: 8 } }, "Sales", /* @__PURE__ */ React.createElement("span", { style: { color: C.green } }, "Desk")), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 14, lineHeight: 1.55, margin: "0 0 22px" } }, "Sign in with your Microsoft account to sync with ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 12 } }, fileName), " in OneDrive."), /* @__PURE__ */ React.createElement(Btn, { onClick: signIn, style: { fontSize: 15, padding: "12px 22px", width: "100%", justifyContent: "center" } }, "Sign in with Microsoft"), errMsg && /* @__PURE__ */ React.createElement("p", { style: { color: C.red, fontSize: 12, marginTop: 14, marginBottom: 0 } }, errMsg), /* @__PURE__ */ React.createElement(
      "button",
      {
        className: "sd-link",
        onClick: () => setPhase("config"),
        style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 12, marginTop: 16 }
      },
      "Change settings"
    )));
  }
  if (phase === "ready" && identityLoaded && team.length > 0 && !meMember) {
    return /* @__PURE__ */ React.createElement(IdentityPicker, { team, ownerPin: db.ownerPin, onPick: chooseIdentity, fontCss });
  }
  const navItems = isOwner ? [
    { key: "today", label: "Today", icon: "today", count: openTasks.filter((t) => t.due && daysFromToday(t.due) <= 0).length },
    { key: "accounts", label: "Accounts", icon: "accounts", count: db.accounts.length },
    { key: "projects", label: "Projects", icon: "project", count: (db.projects || []).length },
    { key: "strategies", label: "Strategies", icon: "strategies", count: db.strategies.length },
    { key: "dump", label: "Brain dump", icon: "dump", count: db.dumps.length },
    { key: "team", label: "Team", icon: "team", count: team.length }
  ] : [
    { key: "today", label: "My tasks", icon: "today", count: db.tasks.filter((t) => t.assigneeId === meId && !t.done && t.due && daysFromToday(t.due) <= 0).length }
  ];
  return /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_BODY, background: C.bg, minHeight: "100vh", color: C.ink } }, /* @__PURE__ */ React.createElement("style", null, fontCss), /* @__PURE__ */ React.createElement("div", { className: "sd-shell", style: { display: "flex", minHeight: "100vh", maxWidth: 1180, margin: "0 auto" } }, /* @__PURE__ */ React.createElement("aside", { className: "sd-rail", style: {
    width: 216,
    minWidth: 216,
    borderRight: `1px solid ${C.line}`,
    padding: "24px 16px",
    display: "flex",
    flexDirection: "column",
    gap: 3,
    position: "sticky",
    top: 0,
    height: "100vh"
  } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 2 } }, /* @__PURE__ */ React.createElement("span", { style: { width: 26, height: 26, borderRadius: 7, background: C.green, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", flex: "none" } }, /* @__PURE__ */ React.createElement(Icon, { name: "strategies", size: 15 })), /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontWeight: 700, fontSize: 19, letterSpacing: "-0.02em" } }, "Sales", /* @__PURE__ */ React.createElement("span", { style: { color: C.green } }, "Desk"))), /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_MONO, fontSize: 10.5, color: C.inkSoft, marginBottom: 16, paddingLeft: 2 } }, (/* @__PURE__ */ new Date()).toLocaleDateString("en-US", { weekday: "long", month: "short", day: "numeric" })), team.length > 0 && meMember && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6, background: C.chip, borderRadius: 8, padding: "7px 11px", marginBottom: 14 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, fontWeight: 600, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, meMember.name, isOwner ? " \xB7 owner" : ""), /* @__PURE__ */ React.createElement("button", { className: "sd-link", onClick: clearIdentity, title: "Switch user", style: { background: "none", border: "none", cursor: "pointer", color: C.green, fontSize: 11, fontWeight: 700, padding: 0, whiteSpace: "nowrap" } }, "switch")), /* @__PURE__ */ React.createElement("div", { className: "sd-navwrap", style: { display: "flex", flexDirection: "column", gap: 3 } }, navItems.map((n) => {
    const active = page === n.key;
    return /* @__PURE__ */ React.createElement(
      "button",
      {
        key: n.key,
        className: "sd-nav",
        onClick: () => {
          setPage(n.key);
          if (n.key !== "accounts") setAccountId(null);
          if (n.key !== "projects") setProjectSel(null);
        },
        style: {
          display: "flex",
          alignItems: "center",
          gap: 10,
          padding: "9px 12px",
          borderRadius: 9,
          border: "none",
          cursor: "pointer",
          background: active ? C.greenSoft : "transparent",
          color: active ? C.green : C.inkSoft,
          fontFamily: FONT_BODY,
          fontWeight: 600,
          fontSize: 13.5,
          textAlign: "left"
        }
      },
      /* @__PURE__ */ React.createElement(Icon, { name: n.icon, size: 17 }),
      /* @__PURE__ */ React.createElement("span", { style: { flex: 1 } }, n.label),
      n.count ? /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, fontWeight: 500, color: active ? C.green : C.inkSoft, background: active ? "#fff" : C.chip, borderRadius: 20, padding: "1px 7px", minWidth: 20, textAlign: "center" } }, n.count) : null
    );
  })), /* @__PURE__ */ React.createElement("div", { style: { flex: 1, minHeight: 20 } }), /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_MONO, fontSize: 10.5, color: saveState === "error" ? C.red : saveState === "offline" ? C.brass : C.inkSoft, height: 16, paddingLeft: 2 } }, saveState === "saving" && "syncing\u2026", saveState === "saved" && "\u2713 synced", saveState === "offline" && "offline \u2014 will sync", saveState === "error" && "\u26A0 sync failed"), /* @__PURE__ */ React.createElement(
    "div",
    {
      style: { fontFamily: FONT_MONO, fontSize: 10.5, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", marginTop: 4, paddingLeft: 2 },
      title: `${fileName}${userName ? " \xB7 " + userName : ""}`
    },
    "\u2601 ",
    fileName
  ), /* @__PURE__ */ React.createElement("div", { className: "sd-footer", style: { display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: refresh, title: "Pull the latest from OneDrive" }, "Refresh"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: exportJson, title: "Download a backup copy" }, "Export"), isOwner && /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: signOut, title: "Sign out of Microsoft" }, "Sign out"))), /* @__PURE__ */ React.createElement("main", { className: "sd-main", style: { flex: 1, padding: "26px 34px 64px", minWidth: 0 } }, !isOwner && /* @__PURE__ */ React.createElement(
    TodayView,
    {
      db,
      accById,
      projById,
      stratById,
      onToggle: toggleTask,
      onUpdate: updateTask,
      onDelete: deleteTask,
      onSnooze: snoozeTask,
      team,
      isOwner: false,
      meId
    }
  ), isOwner && page === "today" && /* @__PURE__ */ React.createElement(
    TodayView,
    {
      db,
      accById,
      projById,
      stratById,
      onToggle: toggleTask,
      onUpdate: updateTask,
      onDelete: deleteTask,
      onSnooze: snoozeTask,
      onGoAccount: goAccount,
      onGoProject: goProject,
      team,
      isOwner: true,
      meId
    }
  ), isOwner && page === "dump" && /* @__PURE__ */ React.createElement(
    DumpView,
    {
      db,
      addDump,
      deleteDump,
      addTask,
      addAccount
    }
  ), isOwner && page === "accounts" && !accountId && /* @__PURE__ */ React.createElement(AccountsView, { db, addAccount, onOpen: goAccount }), isOwner && page === "accounts" && accountId && accById[accountId] && /* @__PURE__ */ React.createElement(
    AccountDetail,
    {
      account: accById[accountId],
      db,
      stratById,
      team,
      onBack: () => setAccountId(null),
      updateAccount,
      deleteAccount,
      markContacted,
      addTask,
      toggleTask,
      updateTask,
      deleteTask,
      snoozeTask,
      addPerson,
      updatePerson,
      deletePerson,
      addTemplate,
      updateTemplate,
      deleteTemplate
    }
  ), isOwner && page === "projects" && !projectSel && /* @__PURE__ */ React.createElement(ProjectsView, { db, addProject, onOpen: goProject }), isOwner && page === "projects" && projectSel && projById[projectSel] && /* @__PURE__ */ React.createElement(
    ProjectDetail,
    {
      project: projById[projectSel],
      db,
      stratById,
      team,
      onBack: () => setProjectSel(null),
      updateProject,
      deleteProject,
      addTask,
      toggleTask,
      updateTask,
      deleteTask,
      snoozeTask
    }
  ), isOwner && page === "strategies" && /* @__PURE__ */ React.createElement(
    StrategiesView,
    {
      db,
      accById,
      team,
      addStrategy,
      deleteStrategy,
      onGoAccount: goAccount,
      toggleTask,
      updateTask
    }
  ), isOwner && page === "team" && /* @__PURE__ */ React.createElement(
    TeamView,
    {
      db,
      meId,
      addMember,
      updateMember,
      deleteMember,
      setOwnerPin,
      chooseIdentity
    }
  ))), /* @__PURE__ */ React.createElement(AIAssistant, null));
}
function Section({ label, count, color, children }) {
  if (!count) return null;
  return /* @__PURE__ */ React.createElement("div", { style: { marginBottom: 26 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "baseline", gap: 8, marginBottom: 4 } }, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_HEAD, fontSize: 13, fontWeight: 600, letterSpacing: "0.04em", textTransform: "uppercase", color: color || C.inkSoft } }, label), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft } }, count)), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 14px" } }, children));
}
function TodayView({ db, accById, projById = {}, stratById, onToggle, onUpdate, onDelete, onSnooze, onGoAccount, onGoProject, team = [], isOwner = true, meId = null }) {
  const [showDone, setShowDone] = useState(false);
  const [showWaiting, setShowWaiting] = useState(false);
  const [who, setWho] = useState("all");
  const byDue = (a, b) => (a.due || "9999").localeCompare(b.due || "9999");
  const scope = (t) => {
    if (!isOwner) return t.assigneeId === meId;
    if (who === "all") return true;
    if (who === "__unassigned") return !t.assigneeId;
    return t.assigneeId === who;
  };
  const pool = db.tasks.filter(scope);
  const taskById = Object.fromEntries(db.tasks.map((t) => [t.id, t]));
  const allOpen = pool.filter((t) => !t.done);
  const waiting = allOpen.filter((t) => blockedByTask(t, taskById)).sort(byDue);
  const open = allOpen.filter((t) => !blockedByTask(t, taskById));
  const overdue = open.filter((t) => t.due && daysFromToday(t.due) < 0).sort(byDue);
  const today = open.filter((t) => t.due && daysFromToday(t.due) === 0).sort(byDue);
  const week = open.filter((t) => t.due && daysFromToday(t.due) > 0 && daysFromToday(t.due) <= 7).sort(byDue);
  const later = open.filter((t) => !t.due || daysFromToday(t.due) > 7).sort(byDue);
  const done = pool.filter((t) => t.done).sort((a, b) => (b.completedAt || "").localeCompare(a.completedAt || "")).slice(0, 15);
  const touchDue = isOwner ? db.accounts.map((a) => ({ a, s: touchStatus(a, db.tasks) })).filter((x) => x.s.state === "overdue" || x.s.state === "due").sort((x, y) => TIER_RANK[x.a.priority || "b"] - TIER_RANK[y.a.priority || "b"]) : [];
  const row = (t) => /* @__PURE__ */ React.createElement(
    TaskRow,
    {
      key: t.id,
      task: t,
      account: accById[t.accountId],
      project: projById[t.projectId],
      strategy: stratById[t.strategyId],
      team,
      canAssign: isOwner,
      blockedBy: blockedByTask(t, taskById),
      onToggle: () => onToggle(t.id),
      onUpdate: (p) => onUpdate(t.id, p),
      onSnooze,
      onDelete: isOwner ? () => onDelete(t.id) : null,
      onGoAccount: isOwner ? onGoAccount : void 0,
      onGoProject: isOwner ? onGoProject : void 0
    }
  );
  const accts = new Set(open.map((t) => t.accountId || t.projectId).filter(Boolean)).size;
  const hi = (/* @__PURE__ */ new Date()).getHours();
  const greeting = hi < 12 ? "Good morning" : hi < 18 ? "Good afternoon" : "Good evening";
  return /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_MONO, fontSize: 11, letterSpacing: "0.08em", textTransform: "uppercase", color: C.inkSoft, marginBottom: 6 } }, greeting), /* @__PURE__ */ React.createElement("h1", { style: { fontFamily: FONT_HEAD, fontSize: 26, fontWeight: 700, margin: "0 0 4px", letterSpacing: "-0.025em" } }, isOwner ? "Where you left off" : "My tasks"), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, margin: "0 0 22px" } }, open.length === 0 ? isOwner ? "Nothing open here. Add tasks from an account or launch a strategy." : "You're all caught up \u2014 no open tasks assigned to you." : `${open.length} open task${open.length === 1 ? "" : "s"}${isOwner ? ` across ${accts} account${accts === 1 ? "" : "s"}` : ""}.`)), isOwner && team.length > 0 && /* @__PURE__ */ React.createElement("select", { value: who, onChange: (e) => setWho(e.target.value), style: { ...inputStyle, width: "auto", minWidth: 150 } }, /* @__PURE__ */ React.createElement("option", { value: "all" }, "Everyone's tasks"), /* @__PURE__ */ React.createElement("option", { value: "__unassigned" }, "Unassigned"), team.map((m) => /* @__PURE__ */ React.createElement("option", { key: m.id, value: m.id }, m.name, m.isOwner ? " (you)" : "")))), (overdue.length > 0 || touchDue.length > 0) && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 } }, overdue.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "sd-card", style: { flex: "1 1 200px", background: C.surface, border: `1px solid ${C.redSoft}`, borderLeft: `3px solid ${C.red}`, borderRadius: 12, padding: "14px 16px" } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 7, color: C.red, marginBottom: 4 } }, /* @__PURE__ */ React.createElement(Icon, { name: "bell", size: 15 }), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_HEAD, fontWeight: 600, fontSize: 13 } }, "Overdue")), /* @__PURE__ */ React.createElement("div", { style: { fontSize: 13, color: C.inkSoft } }, overdue.length, " task", overdue.length === 1 ? "" : "s", " past due \u2014 the oldest is ", -daysFromToday(overdue[0].due), " day", -daysFromToday(overdue[0].due) === 1 ? "" : "s", " late.")), touchDue.length > 0 && /* @__PURE__ */ React.createElement("div", { className: "sd-card", style: { flex: "1 1 200px", background: C.surface, border: `1px solid ${C.brassSoft}`, borderLeft: `3px solid ${C.brass}`, borderRadius: 12, padding: "14px 16px" } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 7, color: C.brass, marginBottom: 4 } }, /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 15 }), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_HEAD, fontWeight: 600, fontSize: 13 } }, "Time to reach out")), /* @__PURE__ */ React.createElement("div", { style: { fontSize: 13, color: C.inkSoft, display: "flex", flexWrap: "wrap", gap: 6, marginTop: 4 } }, touchDue.slice(0, 6).map(({ a }) => /* @__PURE__ */ React.createElement(
    "span",
    {
      key: a.id,
      className: "sd-link",
      onClick: () => onGoAccount && onGoAccount(a.id),
      style: { cursor: "pointer", fontSize: 12, fontWeight: 600, color: tierOf(a).fg, background: tierOf(a).bg, padding: "3px 9px", borderRadius: 20 }
    },
    a.name
  )), touchDue.length > 6 && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft, alignSelf: "center" } }, "+", touchDue.length - 6, " more")))), open.length === 0 && done.length === 0 && (isOwner ? /* @__PURE__ */ React.createElement(Empty, { title: "A clear desk", body: "Start by adding your accounts, then create tasks per account \u2014 or define a strategy to fan tasks out to everyone at once." }) : /* @__PURE__ */ React.createElement(Empty, { title: "Nothing on your plate", body: "When new tasks are assigned to you, they'll show up here grouped by due date." })), /* @__PURE__ */ React.createElement(Section, { label: "Overdue", count: overdue.length, color: C.red }, overdue.map(row)), /* @__PURE__ */ React.createElement(Section, { label: "Due today", count: today.length, color: C.amber }, today.map(row)), /* @__PURE__ */ React.createElement(Section, { label: "Next 7 days", count: week.length, color: C.green }, week.map(row)), /* @__PURE__ */ React.createElement(Section, { label: "Later / unscheduled", count: later.length }, later.map(row)), waiting.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { marginBottom: 26 } }, /* @__PURE__ */ React.createElement("button", { onClick: () => setShowWaiting(!showWaiting), style: { background: "none", border: "none", cursor: "pointer", color: C.brass, fontSize: 12.5, fontWeight: 600, padding: "4px 0", display: "inline-flex", alignItems: "center", gap: 5 } }, /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 13 }), " ", showWaiting ? "\u25B4 Hide" : "\u25BE Show", " waiting on an earlier step (", waiting.length, ")"), showWaiting && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 14px", marginTop: 8 } }, waiting.map(row))), done.length > 0 && /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("button", { onClick: () => setShowDone(!showDone), style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 12.5, fontWeight: 600, padding: "4px 0" } }, showDone ? "\u25B4 Hide" : "\u25BE Show", " recently completed (", done.length, ")"), showDone && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 14px", marginTop: 8 } }, done.map(row))));
}
function DumpView({ db, addDump, deleteDump, addTask, addAccount }) {
  const [text, setText] = useState("");
  const [convertId, setConvertId] = useState(null);
  const submit = () => {
    if (text.trim()) {
      addDump(text);
      setText("");
    }
  };
  return /* @__PURE__ */ React.createElement("div", { style: { maxWidth: 640 } }, /* @__PURE__ */ React.createElement("h1", { style: { fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, margin: "0 0 4px", letterSpacing: "-0.02em" } }, "Brain dump"), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, margin: "0 0 20px" } }, "Get it out of your head. Turn items into real tasks when you're ready."), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginBottom: 22 } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: text,
      onChange: (e) => setText(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "What's on your mind? Press Enter to capture.",
      style: { ...inputStyle, fontSize: 14, padding: "11px 14px" },
      autoFocus: true
    }
  ), /* @__PURE__ */ React.createElement(Btn, { onClick: submit }, "Capture")), db.dumps.length === 0 && /* @__PURE__ */ React.createElement(Empty, { title: "Head's clear", body: "Anything you capture here waits until you turn it into a task or let it go." }), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: 8 } }, db.dumps.map((d) => /* @__PURE__ */ React.createElement("div", { key: d.id, style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 10, padding: "12px 14px" } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 10 } }, /* @__PURE__ */ React.createElement("span", { style: { flex: 1, fontSize: 14 } }, d.text), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 10.5, color: C.inkSoft } }, fmtDate(d.createdAt)), /* @__PURE__ */ React.createElement(Btn, { kind: "subtle", small: true, onClick: () => setConvertId(convertId === d.id ? null : d.id) }, convertId === d.id ? "Cancel" : "\u2192 Task"), /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: () => deleteDump(d.id) }, "\u2715")), convertId === d.id && /* @__PURE__ */ React.createElement(
    ConvertForm,
    {
      dump: d,
      accounts: db.accounts,
      addAccount,
      onConvert: (taskData) => {
        addTask(taskData);
        deleteDump(d.id);
        setConvertId(null);
      }
    }
  )))));
}
function ConvertForm({ dump, accounts, addAccount, onConvert }) {
  const [acc, setAcc] = useState(accounts[0]?.id || "__new");
  const [newName, setNewName] = useState("");
  const [due, setDue] = useState("");
  const go = () => {
    let accountId = acc;
    if (acc === "__new") {
      if (!newName.trim()) return;
      accountId = addAccount(newName);
    }
    onConvert({ accountId, title: dump.text, due: due || null });
  };
  return /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginTop: 10, alignItems: "center", flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("select", { value: acc, onChange: (e) => setAcc(e.target.value), style: { ...inputStyle, width: "auto", minWidth: 160 } }, accounts.map((a) => /* @__PURE__ */ React.createElement("option", { key: a.id, value: a.id }, a.name)), /* @__PURE__ */ React.createElement("option", { value: "__new" }, "+ New account\u2026")), acc === "__new" && /* @__PURE__ */ React.createElement("input", { value: newName, onChange: (e) => setNewName(e.target.value), placeholder: "Account name", style: { ...inputStyle, width: 170 } }), /* @__PURE__ */ React.createElement("input", { type: "date", value: due, onChange: (e) => setDue(e.target.value), style: { ...inputStyle, width: 150 } }), /* @__PURE__ */ React.createElement(Btn, { small: true, onClick: go }, "Create task"));
}
function AccountsView({ db, addAccount, onOpen }) {
  const [name, setName] = useState("");
  const [contact, setContact] = useState("");
  const [priority, setPriority] = useState("b");
  const [stage, setStage] = useState("customer");
  const [sortBy, setSortBy] = useState("priority");
  const [stageFilter, setStageFilter] = useState("all");
  const [query, setQuery] = useState("");
  const submit = () => {
    if (!name.trim()) return;
    addAccount(name, contact, priority, stage);
    setName("");
    setContact("");
  };
  const stats = (id) => {
    const ts = db.tasks.filter((t) => t.accountId === id && !t.done);
    const next = ts.filter((t) => t.due).sort((a, b) => a.due.localeCompare(b.due))[0];
    const overdue = ts.filter((t) => t.due && daysFromToday(t.due) < 0).length;
    return { open: ts.length, next, overdue };
  };
  const q = query.trim().toLowerCase();
  const filtered = db.accounts.filter((a) => {
    if (stageFilter !== "all" && (a.stage || "customer") !== stageFilter) return false;
    if (q && !((a.name || "").toLowerCase().includes(q) || (a.contact || "").toLowerCase().includes(q))) return false;
    return true;
  });
  const sorted = [...filtered].sort((a, b) => {
    if (sortBy === "name") return a.name.localeCompare(b.name);
    if (sortBy === "attention") {
      const rank = (x) => {
        const s = touchStatus(x, db.tasks);
        const od = db.tasks.some((t) => t.accountId === x.id && !t.done && t.due && daysFromToday(t.due) < 0);
        return (od ? 0 : 10) + (s.state === "overdue" ? 0 : s.state === "due" ? 1 : 5);
      };
      return rank(a) - rank(b) || a.name.localeCompare(b.name);
    }
    return TIER_RANK[a.priority || "b"] - TIER_RANK[b.priority || "b"] || a.name.localeCompare(b.name);
  });
  const custCount = db.accounts.filter((a) => (a.stage || "customer") === "customer").length;
  const prospCount = db.accounts.filter((a) => a.stage === "prospect").length;
  return /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("h1", { style: { fontFamily: FONT_HEAD, fontSize: 26, fontWeight: 700, margin: "0 0 4px", letterSpacing: "-0.025em" } }, "Accounts"), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, margin: "0 0 20px" } }, "The retailers you sell to, ranked by priority. Open one for its contacts, tasks, and history."), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap", alignItems: "center" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: name,
      onChange: (e) => setName(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "Account name (e.g. Hudson Toys)",
      style: { ...inputStyle, flex: "1 1 200px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: contact,
      onChange: (e) => setContact(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "Contact (optional)",
      style: { ...inputStyle, flex: "0 1 160px" }
    }
  ), /* @__PURE__ */ React.createElement("select", { value: stage, onChange: (e) => setStage(e.target.value), style: { ...inputStyle, width: "auto" }, title: "Relationship stage" }, /* @__PURE__ */ React.createElement("option", { value: "customer" }, "Customer"), /* @__PURE__ */ React.createElement("option", { value: "prospect" }, "Prospect")), /* @__PURE__ */ React.createElement("select", { value: priority, onChange: (e) => setPriority(e.target.value), style: { ...inputStyle, width: "auto" }, title: "Priority tier" }, /* @__PURE__ */ React.createElement("option", { value: "a" }, "Priority A"), /* @__PURE__ */ React.createElement("option", { value: "b" }, "Priority B"), /* @__PURE__ */ React.createElement("option", { value: "c" }, "Priority C")), /* @__PURE__ */ React.createElement(Btn, { icon: "plus", onClick: submit }, "Add")), db.accounts.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 14, marginBottom: 18, alignItems: "center", flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("div", { style: { position: "relative", flex: "1 1 240px", maxWidth: 340 } }, /* @__PURE__ */ React.createElement("span", { style: { position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: C.inkSoft, pointerEvents: "none", display: "flex" } }, /* @__PURE__ */ React.createElement(Icon, { name: "search", size: 15 })), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: query,
      onChange: (e) => setQuery(e.target.value),
      placeholder: "Search accounts or contacts\u2026",
      style: { ...inputStyle, paddingLeft: 33, paddingRight: query ? 30 : 11 }
    }
  ), query && /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => setQuery(""),
      title: "Clear search",
      "aria-label": "Clear search",
      style: { position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 16, lineHeight: 1, padding: 2 }
    },
    "\xD7"
  )), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 6, alignItems: "center" } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft, marginRight: 2 } }, "Show"), [["all", `All (${db.accounts.length})`], ["customer", `Customers (${custCount})`], ["prospect", `Prospects (${prospCount})`]].map(([k, lbl]) => /* @__PURE__ */ React.createElement("button", { key: k, onClick: () => setStageFilter(k), className: "sd-btn", style: {
    fontFamily: FONT_BODY,
    fontSize: 12,
    fontWeight: 600,
    padding: "4px 11px",
    borderRadius: 20,
    cursor: "pointer",
    border: `1px solid ${stageFilter === k ? C.green : C.line}`,
    background: stageFilter === k ? C.greenSoft : "transparent",
    color: stageFilter === k ? C.green : C.inkSoft
  } }, lbl))), db.accounts.length > 1 && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 6, alignItems: "center" } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft, marginRight: 2 } }, "Sort by"), [["priority", "Priority"], ["attention", "Needs attention"], ["name", "Name"]].map(([k, lbl]) => /* @__PURE__ */ React.createElement("button", { key: k, onClick: () => setSortBy(k), className: "sd-btn", style: {
    fontFamily: FONT_BODY,
    fontSize: 12,
    fontWeight: 600,
    padding: "4px 11px",
    borderRadius: 20,
    cursor: "pointer",
    border: `1px solid ${sortBy === k ? C.green : C.line}`,
    background: sortBy === k ? C.greenSoft : "transparent",
    color: sortBy === k ? C.green : C.inkSoft
  } }, lbl)))), q && db.accounts.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { fontSize: 12.5, color: C.inkSoft, marginBottom: 12, marginTop: -6 } }, sorted.length, " ", sorted.length === 1 ? "match" : "matches", ' for "', query.trim(), '"'), db.accounts.length === 0 && /* @__PURE__ */ React.createElement(Empty, { title: "No accounts yet", body: "Add the retailers you sell to, tag each as a customer or prospect, and give it a priority. Strategies can then fan tasks out to all of them at once." }), db.accounts.length > 0 && sorted.length === 0 && /* @__PURE__ */ React.createElement(Empty, { title: "No matches", body: q ? `No accounts match "${query.trim()}". Try a different spelling or clear the search.` : "No accounts match this filter yet." }), /* @__PURE__ */ React.createElement("div", { style: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(268px, 1fr))", gap: 12 } }, sorted.map((a) => {
    const s = stats(a.id);
    const tier = tierOf(a);
    const stg = stageOf(a);
    const touch = touchStatus(a, db.tasks);
    return /* @__PURE__ */ React.createElement("button", { key: a.id, className: "sd-card", onClick: () => onOpen(a.id), style: {
      textAlign: "left",
      background: C.surface,
      border: `1px solid ${C.line}`,
      borderRadius: 12,
      padding: 0,
      cursor: "pointer",
      fontFamily: FONT_BODY,
      overflow: "hidden",
      display: "flex"
    } }, /* @__PURE__ */ React.createElement("div", { style: { width: 4, background: tier.bar, flex: "none" } }), /* @__PURE__ */ React.createElement("div", { style: { padding: "14px 16px", flex: 1, minWidth: 0 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 4 } }, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_HEAD, fontWeight: 600, fontSize: 15.5, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, a.name), /* @__PURE__ */ React.createElement("span", { style: { marginLeft: "auto", fontFamily: FONT_MONO, fontSize: 11, fontWeight: 600, color: tier.fg, background: tier.bg, borderRadius: 5, padding: "1px 7px", flex: "none" } }, tier.label)), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 6, marginBottom: 8, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 10.5, fontWeight: 700, letterSpacing: "0.04em", textTransform: "uppercase", color: stg.fg, background: stg.bg, borderRadius: 5, padding: "2px 7px" } }, stg.label), a.contact && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, a.contact)), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 6 } }, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11.5, color: s.open ? C.green : C.inkSoft } }, s.open ? `${s.open} open` : "no open tasks"), s.overdue > 0 && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11.5, color: C.red } }, s.overdue, " overdue"), s.next && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11.5, color: C.inkSoft } }, "next ", fmtDate(s.next.due))), touch.state === "overdue" && /* @__PURE__ */ React.createElement("div", { style: { fontSize: 11.5, color: C.brass, marginTop: 7, display: "flex", alignItems: "center", gap: 5 } }, /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 12 }), " ", touch.since == null ? "never contacted" : `${touch.since}d since contact \u2014 time to reach out`), touch.state === "due" && /* @__PURE__ */ React.createElement("div", { style: { fontSize: 11.5, color: C.inkSoft, marginTop: 7, display: "flex", alignItems: "center", gap: 5 } }, /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 12 }), " touch base soon (", touch.since, "d / ", touch.cad, "d)")));
  })));
}
function AccountDetail({ account, db, stratById, team = [], onBack, updateAccount, deleteAccount, markContacted, addTask, toggleTask, updateTask, deleteTask, snoozeTask, addPerson, updatePerson, deletePerson, addTemplate, updateTemplate, deleteTemplate }) {
  const [title, setTitle] = useState("");
  const [due, setDue] = useState("");
  const [assigneeId, setAssigneeId] = useState("");
  const [repeatDays, setRepeatDays] = useState(0);
  const [compose, setCompose] = useState(null);
  const tasks = db.tasks.filter((t) => t.accountId === account.id);
  const open = tasks.filter((t) => !t.done).sort((a, b) => (a.due || "9999").localeCompare(b.due || "9999"));
  const done = tasks.filter((t) => t.done).sort((a, b) => (b.completedAt || "").localeCompare(a.completedAt || ""));
  const tById = Object.fromEntries(tasks.map((t) => [t.id, t]));
  const siblings = tasks.map((t) => ({ id: t.id, title: t.title, done: t.done }));
  const tier = tierOf(account);
  const touch = touchStatus(account, db.tasks);
  const last = lastTouched(account, db.tasks);
  const submit = () => {
    if (!title.trim()) return;
    addTask({ accountId: account.id, title, due: due || null, assigneeId: assigneeId || null, repeatDays: repeatDays || null });
    setTitle("");
    setDue("");
    setRepeatDays(0);
  };
  const CADENCE_OPTS = [[0, "No reminder"], [14, "Every 2 weeks"], [30, "Monthly"], [45, "Every 6 weeks"], [60, "Every 2 months"], [90, "Quarterly"]];
  return /* @__PURE__ */ React.createElement("div", { style: { maxWidth: 740 } }, /* @__PURE__ */ React.createElement("button", { className: "sd-link", onClick: onBack, style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 13, fontWeight: 600, padding: 0, marginBottom: 16, display: "inline-flex", alignItems: "center", gap: 4 } }, "\u2190 All accounts"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 10, marginBottom: 4 } }, /* @__PURE__ */ React.createElement("span", { style: { width: 5, height: 26, borderRadius: 3, background: tier.bar, flex: "none" } }), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: account.name,
      onChange: (e) => updateAccount(account.id, { name: e.target.value }),
      style: { fontFamily: FONT_HEAD, fontSize: 25, fontWeight: 700, border: "none", background: "transparent", color: C.ink, outline: "none", flex: 1, letterSpacing: "-0.025em", padding: 0 }
    }
  ), /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: () => deleteAccount(account.id) }, "Delete")), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: account.contact || "",
      onChange: (e) => updateAccount(account.id, { contact: e.target.value }),
      placeholder: "Contact person / email / phone",
      style: { ...inputStyle, border: "none", background: "transparent", padding: "0 0 0 15px", fontSize: 13, color: C.inkSoft, marginBottom: 14 }
    }
  ), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center", background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "12px 14px", marginBottom: 16 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 7 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft } }, "Priority"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: account.priority || "b",
      onChange: (e) => updateAccount(account.id, { priority: e.target.value }),
      style: { ...inputStyle, width: "auto", padding: "5px 8px", fontSize: 12.5, fontWeight: 600, color: tier.fg }
    },
    /* @__PURE__ */ React.createElement("option", { value: "a" }, "A \u2014 top"),
    /* @__PURE__ */ React.createElement("option", { value: "b" }, "B \u2014 mid"),
    /* @__PURE__ */ React.createElement("option", { value: "c" }, "C \u2014 long tail")
  )), /* @__PURE__ */ React.createElement("div", { style: { width: 1, height: 22, background: C.line } }), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 7 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft } }, "Status"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: account.stage || "customer",
      onChange: (e) => updateAccount(account.id, { stage: e.target.value }),
      style: { ...inputStyle, width: "auto", padding: "5px 8px", fontSize: 12.5, fontWeight: 600, color: stageOf(account).fg }
    },
    /* @__PURE__ */ React.createElement("option", { value: "customer" }, "Customer"),
    /* @__PURE__ */ React.createElement("option", { value: "prospect" }, "Prospect")
  )), /* @__PURE__ */ React.createElement("div", { style: { width: 1, height: 22, background: C.line } }), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 7 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12, color: C.inkSoft, display: "inline-flex", alignItems: "center", gap: 4 } }, /* @__PURE__ */ React.createElement(Icon, { name: "clock", size: 13 }), " Keep in touch"), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: account.cadenceDays || 0,
      onChange: (e) => updateAccount(account.id, { cadenceDays: Number(e.target.value) || null }),
      style: { ...inputStyle, width: "auto", padding: "5px 8px", fontSize: 12.5 }
    },
    CADENCE_OPTS.map(([v, l]) => /* @__PURE__ */ React.createElement("option", { key: v, value: v }, l))
  )), /* @__PURE__ */ React.createElement("div", { style: { flex: 1 } }), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8 } }, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: touch.state === "overdue" ? C.brass : C.inkSoft } }, last ? `last touch ${fmtDate(last)}` : "no contact yet"), /* @__PURE__ */ React.createElement(Btn, { kind: "subtle", small: true, onClick: () => markContacted(account.id), title: "Log that you reached out today" }, "Mark contacted"))), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: account.notes || "",
      onChange: (e) => updateAccount(account.id, { notes: e.target.value }),
      placeholder: "Account notes \u2014 what they buy, terms, preferences, anything worth remembering.",
      rows: 2,
      style: { ...inputStyle, resize: "vertical", marginBottom: 22, fontSize: 13 }
    }
  ), /* @__PURE__ */ React.createElement(
    OrgChart,
    {
      account,
      db,
      addPerson,
      updatePerson,
      deletePerson,
      onCompose: (person) => setCompose(person)
    }
  ), /* @__PURE__ */ React.createElement(SectionLabel, null, "Tasks"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: title,
      onChange: (e) => setTitle(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "New task for this account\u2026",
      style: { ...inputStyle, flex: "1 1 220px" }
    }
  ), /* @__PURE__ */ React.createElement("input", { type: "date", value: due, onChange: (e) => setDue(e.target.value), style: { ...inputStyle, width: 148 } }), /* @__PURE__ */ React.createElement("select", { value: repeatDays, onChange: (e) => setRepeatDays(Number(e.target.value)), style: { ...inputStyle, width: "auto" }, title: "Repeat" }, REPEAT_OPTS.map((o) => /* @__PURE__ */ React.createElement("option", { key: o.v, value: o.v }, o.label))), team.length > 0 && /* @__PURE__ */ React.createElement("select", { value: assigneeId, onChange: (e) => setAssigneeId(e.target.value), style: { ...inputStyle, width: "auto", minWidth: 130 } }, /* @__PURE__ */ React.createElement("option", { value: "" }, "Unassigned"), team.map((m) => /* @__PURE__ */ React.createElement("option", { key: m.id, value: m.id }, m.name))), /* @__PURE__ */ React.createElement(Btn, { icon: "plus", onClick: submit }, "Add")), open.length === 0 && done.length === 0 && /* @__PURE__ */ React.createElement(Empty, { title: "No tasks yet", body: "Add a task above, or include this account in a strategy." }), open.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 12px", marginBottom: 18 } }, open.map((t) => /* @__PURE__ */ React.createElement(
    TaskRow,
    {
      key: t.id,
      task: t,
      strategy: stratById[t.strategyId],
      showAccount: false,
      team,
      canAssign: team.length > 0,
      siblings,
      blockedBy: blockedByTask(t, tById),
      onToggle: () => toggleTask(t.id),
      onUpdate: (p) => updateTask(t.id, p),
      onDelete: () => deleteTask(t.id),
      onSnooze: snoozeTask
    }
  ))), done.length > 0 && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement(SectionLabel, null, "History \xB7 ", done.length), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 12px" } }, done.map((t) => /* @__PURE__ */ React.createElement(
    TaskRow,
    {
      key: t.id,
      task: t,
      strategy: stratById[t.strategyId],
      showAccount: false,
      team,
      canAssign: team.length > 0,
      onToggle: () => toggleTask(t.id),
      onUpdate: (p) => updateTask(t.id, p),
      onDelete: () => deleteTask(t.id)
    }
  )))), compose && /* @__PURE__ */ React.createElement(
    ComposeModal,
    {
      person: compose,
      account,
      templates: db.templates || [],
      onClose: () => setCompose(null),
      onSent: () => markContacted(account.id),
      addTemplate,
      updateTemplate,
      deleteTemplate
    }
  ));
}
function SectionLabel({ children }) {
  return /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontSize: 12, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.06em", color: C.inkSoft, margin: "4px 0 8px" } }, children);
}
function ProjectsView({ db, addProject, onOpen }) {
  const [name, setName] = useState("");
  const projects = db.projects || [];
  const submit = () => {
    if (name.trim()) {
      addProject(name);
      setName("");
    }
  };
  const stats = (id) => {
    const ts = db.tasks.filter((t) => t.projectId === id && !t.done);
    const next = ts.filter((t) => t.due).sort((a, b) => a.due.localeCompare(b.due))[0];
    const overdue = ts.filter((t) => t.due && daysFromToday(t.due) < 0).length;
    return { open: ts.length, next, overdue };
  };
  return /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("h1", { style: { fontFamily: FONT_HEAD, fontSize: 26, fontWeight: 700, margin: "0 0 4px", letterSpacing: "-0.025em" } }, "Projects"), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, margin: "0 0 20px" } }, "Standalone work that isn't tied to a single account \u2014 trade-show prep, internal initiatives, personal to-dos. Their tasks still show up on Today and in your morning digest."), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginBottom: 22, maxWidth: 560 } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: name,
      onChange: (e) => setName(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "Project name (e.g. Spring Trade Show)",
      style: inputStyle
    }
  ), /* @__PURE__ */ React.createElement(Btn, { icon: "plus", onClick: submit }, "Add")), projects.length === 0 && /* @__PURE__ */ React.createElement(Empty, { title: "No projects yet", body: "Add a project to track work that doesn't belong to a specific retailer. Each one holds its own task list." }), /* @__PURE__ */ React.createElement("div", { style: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(268px, 1fr))", gap: 12 } }, [...projects].sort((a, b) => a.name.localeCompare(b.name)).map((p) => {
    const s = stats(p.id);
    return /* @__PURE__ */ React.createElement("button", { key: p.id, className: "sd-card", onClick: () => onOpen(p.id), style: {
      textAlign: "left",
      background: C.surface,
      border: `1px solid ${C.line}`,
      borderRadius: 12,
      padding: 0,
      cursor: "pointer",
      fontFamily: FONT_BODY,
      overflow: "hidden",
      display: "flex"
    } }, /* @__PURE__ */ React.createElement("div", { style: { width: 4, background: C.violet, flex: "none" } }), /* @__PURE__ */ React.createElement("div", { style: { padding: "14px 16px", flex: 1, minWidth: 0 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 7, marginBottom: 6 } }, /* @__PURE__ */ React.createElement("span", { style: { color: C.violet, display: "flex", flex: "none" } }, /* @__PURE__ */ React.createElement(Icon, { name: "project", size: 15 })), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_HEAD, fontWeight: 600, fontSize: 15.5, color: C.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, p.name)), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11.5, color: s.open ? C.green : C.inkSoft } }, s.open ? `${s.open} open` : "no open tasks"), s.overdue > 0 && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11.5, color: C.red } }, s.overdue, " overdue"), s.next && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11.5, color: C.inkSoft } }, "next ", fmtDate(s.next.due)))));
  })));
}
function ProjectDetail({ project, db, stratById, team = [], onBack, updateProject, deleteProject, addTask, toggleTask, updateTask, deleteTask, snoozeTask }) {
  const [title, setTitle] = useState("");
  const [due, setDue] = useState("");
  const [assigneeId, setAssigneeId] = useState("");
  const [repeatDays, setRepeatDays] = useState(0);
  const tasks = db.tasks.filter((t) => t.projectId === project.id);
  const open = tasks.filter((t) => !t.done).sort((a, b) => (a.due || "9999").localeCompare(b.due || "9999"));
  const done = tasks.filter((t) => t.done).sort((a, b) => (b.completedAt || "").localeCompare(a.completedAt || ""));
  const tById = Object.fromEntries(tasks.map((t) => [t.id, t]));
  const siblings = tasks.map((t) => ({ id: t.id, title: t.title, done: t.done }));
  const submit = () => {
    if (!title.trim()) return;
    addTask({ projectId: project.id, title, due: due || null, assigneeId: assigneeId || null, repeatDays: repeatDays || null });
    setTitle("");
    setDue("");
    setRepeatDays(0);
  };
  return /* @__PURE__ */ React.createElement("div", { style: { maxWidth: 740 } }, /* @__PURE__ */ React.createElement("button", { className: "sd-link", onClick: onBack, style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 13, fontWeight: 600, padding: 0, marginBottom: 16, display: "inline-flex", alignItems: "center", gap: 4 } }, "\u2190 All projects"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 10, marginBottom: 4 } }, /* @__PURE__ */ React.createElement("span", { style: { color: C.violet, display: "flex", flex: "none" } }, /* @__PURE__ */ React.createElement(Icon, { name: "project", size: 20 })), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: project.name,
      onChange: (e) => updateProject(project.id, { name: e.target.value }),
      style: { fontFamily: FONT_HEAD, fontSize: 25, fontWeight: 700, border: "none", background: "transparent", color: C.ink, outline: "none", flex: 1, letterSpacing: "-0.025em", padding: 0 }
    }
  ), /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: () => deleteProject(project.id) }, "Delete")), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: project.notes || "",
      onChange: (e) => updateProject(project.id, { notes: e.target.value }),
      placeholder: "Project notes \u2014 goals, context, anything worth keeping here.",
      rows: 2,
      style: { ...inputStyle, resize: "vertical", margin: "12px 0 22px", fontSize: 13 }
    }
  ), /* @__PURE__ */ React.createElement(SectionLabel, null, "Tasks"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginBottom: 16, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: title,
      onChange: (e) => setTitle(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "New task for this project\u2026",
      style: { ...inputStyle, flex: "1 1 220px" }
    }
  ), /* @__PURE__ */ React.createElement("input", { type: "date", value: due, onChange: (e) => setDue(e.target.value), style: { ...inputStyle, width: 148 } }), /* @__PURE__ */ React.createElement("select", { value: repeatDays, onChange: (e) => setRepeatDays(Number(e.target.value)), style: { ...inputStyle, width: "auto" }, title: "Repeat" }, REPEAT_OPTS.map((o) => /* @__PURE__ */ React.createElement("option", { key: o.v, value: o.v }, o.label))), team.length > 0 && /* @__PURE__ */ React.createElement("select", { value: assigneeId, onChange: (e) => setAssigneeId(e.target.value), style: { ...inputStyle, width: "auto", minWidth: 130 } }, /* @__PURE__ */ React.createElement("option", { value: "" }, "Unassigned"), team.map((m) => /* @__PURE__ */ React.createElement("option", { key: m.id, value: m.id }, m.name))), /* @__PURE__ */ React.createElement(Btn, { icon: "plus", onClick: submit }, "Add")), open.length === 0 && done.length === 0 && /* @__PURE__ */ React.createElement(Empty, { title: "No tasks yet", body: "Add the first task for this project above." }), open.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 12px", marginBottom: 18 } }, open.map((t) => /* @__PURE__ */ React.createElement(
    TaskRow,
    {
      key: t.id,
      task: t,
      strategy: stratById[t.strategyId],
      showAccount: false,
      team,
      canAssign: team.length > 0,
      siblings,
      blockedBy: blockedByTask(t, tById),
      onToggle: () => toggleTask(t.id),
      onUpdate: (p) => updateTask(t.id, p),
      onDelete: () => deleteTask(t.id),
      onSnooze: snoozeTask
    }
  ))), done.length > 0 && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement(SectionLabel, null, "History \xB7 ", done.length), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, borderRadius: 12, border: `1px solid ${C.line}`, padding: "2px 12px" } }, done.map((t) => /* @__PURE__ */ React.createElement(
    TaskRow,
    {
      key: t.id,
      task: t,
      strategy: stratById[t.strategyId],
      showAccount: false,
      team,
      canAssign: team.length > 0,
      onToggle: () => toggleTask(t.id),
      onUpdate: (p) => updateTask(t.id, p),
      onDelete: () => deleteTask(t.id)
    }
  )))));
}
function OrgChart({ account, db, addPerson, updatePerson, deletePerson, onCompose }) {
  const mine = (db.people || []).filter((p) => p.accountId === account.id);
  const [open, setOpen] = useState(true);
  const [name, setName] = useState("");
  const [role, setRole] = useState("");
  const [email, setEmail] = useState("");
  const [met, setMet] = useState(false);
  const [managerId, setManagerId] = useState("");
  const nameRef = useRef(null);
  const kids = {};
  mine.forEach((p) => {
    const key = p.managerId && mine.some((m) => m.id === p.managerId) ? p.managerId : "__root";
    (kids[key] = kids[key] || []).push(p);
  });
  const submit = () => {
    if (!name.trim()) return;
    addPerson({ accountId: account.id, name, role, email, met, managerId: managerId || null });
    setName("");
    setRole("");
    setEmail("");
    setMet(false);
    nameRef.current && nameRef.current.focus();
  };
  const quickReport = (id) => {
    setManagerId(id);
    nameRef.current && nameRef.current.focus();
  };
  return /* @__PURE__ */ React.createElement("div", { style: { marginBottom: 22 } }, /* @__PURE__ */ React.createElement("button", { onClick: () => setOpen(!open), style: { background: "none", border: "none", cursor: "pointer", padding: 0, marginBottom: 8, display: "flex", alignItems: "baseline", gap: 8 } }, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_HEAD, fontSize: 12.5, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.04em", color: C.inkSoft } }, open ? "\u25BE" : "\u25B8", " Org chart"), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft } }, mine.length || "")), open && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "14px 16px" } }, /* @__PURE__ */ React.createElement("datalist", { id: "salesdesk-roles" }, (db.roles || []).map((r) => /* @__PURE__ */ React.createElement("option", { key: r, value: r }))), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, flexWrap: "wrap", marginBottom: mine.length ? 14 : 0 } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      ref: nameRef,
      value: name,
      onChange: (e) => setName(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "Person's name",
      style: { ...inputStyle, width: 170, flex: "1 1 150px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: role,
      onChange: (e) => setRole(e.target.value),
      list: "salesdesk-roles",
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "Role \u2014 type new or pick saved",
      style: { ...inputStyle, width: 190, flex: "1 1 170px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: email,
      onChange: (e) => setEmail(e.target.value),
      type: "email",
      onKeyDown: (e) => e.key === "Enter" && submit(),
      placeholder: "Email (optional)",
      style: { ...inputStyle, width: 190, flex: "1 1 170px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: managerId,
      onChange: (e) => setManagerId(e.target.value),
      style: { ...inputStyle, width: "auto", minWidth: 150 }
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "Top level"),
    mine.map((p) => /* @__PURE__ */ React.createElement("option", { key: p.id, value: p.id }, "Reports to ", p.name))
  ), /* @__PURE__ */ React.createElement("label", { style: { display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, color: C.inkSoft, cursor: "pointer", userSelect: "none" } }, /* @__PURE__ */ React.createElement("input", { type: "checkbox", checked: met, onChange: (e) => setMet(e.target.checked), style: { accentColor: C.green, width: 15, height: 15 } }), "Met in person"), /* @__PURE__ */ React.createElement(Btn, { onClick: submit }, "Add person")), mine.length === 0 && /* @__PURE__ */ React.createElement("div", { style: { fontSize: 12.5, color: C.inkSoft, marginTop: 10 } }, "Map who's who at ", account.name, " \u2014 buyers, managers, owners. Roles you add are saved for every account."), (kids.__root || []).map((p) => /* @__PURE__ */ React.createElement(
    PersonNode,
    {
      key: p.id,
      p,
      kids,
      depth: 0,
      mine,
      updatePerson,
      deletePerson,
      onQuickReport: quickReport,
      onCompose
    }
  ))));
}
function PersonNode({ p, kids, depth, mine, updatePerson, deletePerson, onQuickReport, onCompose }) {
  const [edit, setEdit] = useState(false);
  const blocked = useMemo(() => {
    const out = /* @__PURE__ */ new Set([p.id]);
    const walk = (x) => (kids[x] || []).forEach((c) => {
      out.add(c.id);
      walk(c.id);
    });
    walk(p.id);
    return out;
  }, [kids, p.id]);
  return /* @__PURE__ */ React.createElement("div", { style: {
    marginLeft: depth ? 16 : 0,
    borderLeft: depth ? `2px solid ${C.line}` : "none",
    paddingLeft: depth ? 14 : 0,
    marginTop: 8
  } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("span", { style: { width: 7, height: 7, borderRadius: 999, background: depth === 0 ? C.green : "#B9C4BC", minWidth: 7 } }), /* @__PURE__ */ React.createElement("span", { style: { fontSize: 13.5, fontWeight: 600, color: C.ink } }, p.name), p.role && /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11.5, fontWeight: 600, color: C.green, background: C.greenSoft, padding: "2px 8px", borderRadius: 6 } }, p.role), p.email && /* @__PURE__ */ React.createElement(
    "button",
    {
      className: "sd-btn",
      onClick: () => onCompose && onCompose(p),
      title: `Email ${p.name}`,
      style: {
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        fontFamily: FONT_BODY,
        fontSize: 11.5,
        fontWeight: 600,
        color: C.green,
        background: "transparent",
        border: `1px solid ${C.line}`,
        borderRadius: 6,
        padding: "2px 8px",
        cursor: "pointer"
      }
    },
    /* @__PURE__ */ React.createElement(Icon, { name: "mail", size: 13 }),
    " Email"
  ), /* @__PURE__ */ React.createElement(
    "button",
    {
      onClick: () => updatePerson(p.id, { met: !p.met }),
      title: p.met ? "You've met in person \u2014 click to unset" : "Haven't met in person yet \u2014 click to mark as met",
      style: {
        fontFamily: FONT_BODY,
        fontSize: 11,
        fontWeight: 600,
        padding: "2px 8px",
        borderRadius: 999,
        cursor: "pointer",
        border: `1px solid ${p.met ? C.green : C.line}`,
        background: p.met ? C.greenSoft : "transparent",
        color: p.met ? C.green : "#9AA69E"
      }
    },
    p.met ? "\u{1F91D} met" : "not met"
  ), /* @__PURE__ */ React.createElement("div", { style: { flex: 1 } }), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: () => onQuickReport(p.id), title: `Add someone who reports to ${p.name}` }, "+ report"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: () => setEdit(!edit) }, edit ? "Done" : "Edit")), !edit && p.notes && /* @__PURE__ */ React.createElement("div", { style: { marginLeft: 15, marginTop: 4, fontSize: 12, color: C.inkSoft, whiteSpace: "pre-wrap", lineHeight: 1.45 } }, /* @__PURE__ */ React.createElement("span", { style: { color: "#9AA69E" } }, "\u{1F4DD} "), p.notes), edit && /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: 8, marginTop: 8, marginLeft: 15 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: p.name,
      onChange: (e) => updatePerson(p.id, { name: e.target.value }),
      style: { ...inputStyle, width: 150, flex: "1 1 130px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: p.role || "",
      onChange: (e) => updatePerson(p.id, { role: e.target.value }),
      list: "salesdesk-roles",
      placeholder: "Role",
      style: { ...inputStyle, width: 160, flex: "1 1 140px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: p.email || "",
      onChange: (e) => updatePerson(p.id, { email: e.target.value }),
      type: "email",
      placeholder: "Email",
      style: { ...inputStyle, width: 180, flex: "1 1 160px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "select",
    {
      value: p.managerId || "",
      onChange: (e) => updatePerson(p.id, { managerId: e.target.value || null }),
      style: { ...inputStyle, width: "auto", minWidth: 140 }
    },
    /* @__PURE__ */ React.createElement("option", { value: "" }, "Top level"),
    mine.filter((m) => !blocked.has(m.id)).map((m) => /* @__PURE__ */ React.createElement("option", { key: m.id, value: m.id }, "Reports to ", m.name))
  )), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: p.notes || "",
      onChange: (e) => updatePerson(p.id, { notes: e.target.value }),
      placeholder: "Contact notes \u2014 how they like to be reached, what they buy, personal details worth remembering\u2026",
      rows: 2,
      style: { ...inputStyle, resize: "vertical", fontSize: 13 }
    }
  ), /* @__PURE__ */ React.createElement("div", { style: { display: "flex" } }, /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: () => {
    if (window.confirm(`Remove ${p.name}? Their reports move up one level.`)) deletePerson(p.id);
  } }, "Delete contact"))), (kids[p.id] || []).map((c) => /* @__PURE__ */ React.createElement(
    PersonNode,
    {
      key: c.id,
      p: c,
      kids,
      depth: depth + 1,
      mine,
      updatePerson,
      deletePerson,
      onQuickReport,
      onCompose
    }
  )));
}
function StrategiesView({ db, accById, team = [], addStrategy, deleteStrategy, onGoAccount, toggleTask, updateTask }) {
  const [creating, setCreating] = useState(false);
  return /* @__PURE__ */ React.createElement("div", { style: { maxWidth: 760 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 4 } }, /* @__PURE__ */ React.createElement("h1", { style: { fontFamily: FONT_HEAD, fontSize: 24, fontWeight: 700, margin: 0, letterSpacing: "-0.02em" } }, "Strategies"), /* @__PURE__ */ React.createElement(Btn, { onClick: () => setCreating(!creating) }, creating ? "Cancel" : "+ New strategy")), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, margin: "0 0 20px" } }, "Decide a play once \u2014 a task is created for every account you pick, and progress is tracked here."), creating && /* @__PURE__ */ React.createElement(StrategyForm, { accounts: db.accounts, team, onCreate: (s) => {
    addStrategy(s);
    setCreating(false);
  } }), db.strategies.length === 0 && !creating && /* @__PURE__ */ React.createElement(Empty, { title: "No strategies yet", body: 'Example: "Pitch the new spring line" \u2192 one outreach task per retailer, each with its own due date and outcome notes.' }), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: 14 } }, db.strategies.map((s) => /* @__PURE__ */ React.createElement(
    StrategyCard,
    {
      key: s.id,
      strategy: s,
      db,
      accById,
      onDelete: () => deleteStrategy(s.id),
      onGoAccount,
      toggleTask,
      updateTask
    }
  ))));
}
function StrategyForm({ accounts, team = [], onCreate }) {
  const [name, setName] = useState("");
  const [taskText, setTaskText] = useState("");
  const [due, setDue] = useState("");
  const [assigneeId, setAssigneeId] = useState("");
  const [selected, setSelected] = useState(() => new Set(accounts.map((a) => a.id)));
  const toggle = (id) => setSelected((s) => {
    const n = new Set(s);
    n.has(id) ? n.delete(id) : n.add(id);
    return n;
  });
  const allOn = selected.size === accounts.length;
  const go = () => {
    if (!name.trim() || selected.size === 0) return;
    onCreate({ name, taskText: taskText.trim() || name.trim(), due: due || null, accountIds: [...selected], assigneeId: assigneeId || null });
  };
  return /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: 18, marginBottom: 22 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: 10 } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: name,
      onChange: (e) => setName(e.target.value),
      placeholder: 'Strategy name \u2014 e.g. "Pitch the new spring line"',
      style: { ...inputStyle, fontSize: 14 },
      autoFocus: true
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: taskText,
      onChange: (e) => setTaskText(e.target.value),
      placeholder: "Task wording per account (defaults to strategy name)",
      style: inputStyle
    }
  ), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12.5, color: C.inkSoft } }, "Due date for all tasks"), /* @__PURE__ */ React.createElement("input", { type: "date", value: due, onChange: (e) => setDue(e.target.value), style: { ...inputStyle, width: 160 } }), team.length > 0 && /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12.5, color: C.inkSoft, marginLeft: 6 } }, "Assign all to"), /* @__PURE__ */ React.createElement("select", { value: assigneeId, onChange: (e) => setAssigneeId(e.target.value), style: { ...inputStyle, width: "auto", minWidth: 130 } }, /* @__PURE__ */ React.createElement("option", { value: "" }, "Unassigned"), team.map((m) => /* @__PURE__ */ React.createElement("option", { key: m.id, value: m.id }, m.name))))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 } }, /* @__PURE__ */ React.createElement("span", { style: { fontSize: 12.5, fontWeight: 600, color: C.inkSoft } }, "Accounts \xB7 ", selected.size, " selected"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: () => setSelected(allOn ? /* @__PURE__ */ new Set() : new Set(accounts.map((a) => a.id))) }, allOn ? "Deselect all" : "Select all")), accounts.length === 0 ? /* @__PURE__ */ React.createElement("div", { style: { fontSize: 13, color: C.inkSoft } }, "Add accounts first \u2014 the strategy needs someone to fan out to.") : /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexWrap: "wrap", gap: 6 } }, accounts.map((a) => {
    const on = selected.has(a.id);
    return /* @__PURE__ */ React.createElement("button", { key: a.id, onClick: () => toggle(a.id), style: {
      fontFamily: FONT_BODY,
      fontSize: 12.5,
      fontWeight: 600,
      padding: "5px 11px",
      borderRadius: 999,
      cursor: "pointer",
      border: `1px solid ${on ? C.green : C.line}`,
      background: on ? C.greenSoft : "transparent",
      color: on ? C.green : C.inkSoft
    } }, on ? "\u2713 " : "", a.name);
  }))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement(Btn, { onClick: go, style: { opacity: name.trim() && selected.size ? 1 : 0.5 } }, "Create ", selected.size, " task", selected.size === 1 ? "" : "s"))));
}
function StrategyCard({ strategy, db, accById, onDelete, onGoAccount, toggleTask, updateTask }) {
  const [open, setOpen] = useState(false);
  const [editingNote, setEditingNote] = useState(null);
  const tasks = db.tasks.filter((t) => t.strategyId === strategy.id);
  const done = tasks.filter((t) => t.done).length;
  const pct = tasks.length ? Math.round(done / tasks.length * 100) : 0;
  return /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "16px 18px" } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 12 } }, /* @__PURE__ */ React.createElement("div", { style: { flex: 1, minWidth: 0 } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontWeight: 600, fontSize: 15.5 } }, strategy.name), /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft, marginTop: 2 } }, "started ", fmtDate(strategy.createdAt), " \xB7 ", done, "/", tasks.length, " accounts done")), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 13, fontWeight: 500, color: pct === 100 ? C.green : C.ink } }, pct, "%"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: () => setOpen(!open) }, open ? "Hide" : "Detail"), /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: onDelete }, "\u2715")), /* @__PURE__ */ React.createElement("div", { style: { height: 6, background: C.chip, borderRadius: 999, marginTop: 12, overflow: "hidden" } }, /* @__PURE__ */ React.createElement("div", { style: { height: "100%", width: `${pct}%`, background: C.green, borderRadius: 999, transition: "width 0.3s" } })), open && /* @__PURE__ */ React.createElement("div", { style: { marginTop: 12, borderTop: `1px solid ${C.line}` } }, tasks.map((t) => {
    const a = accById[t.accountId];
    return /* @__PURE__ */ React.createElement("div", { key: t.id, style: { borderBottom: `1px solid ${C.line}`, padding: "9px 2px" } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 10 } }, /* @__PURE__ */ React.createElement(Check, { done: t.done, onToggle: () => toggleTask(t.id) }), /* @__PURE__ */ React.createElement(
      "span",
      {
        onClick: () => a && onGoAccount(a.id),
        style: { fontSize: 13.5, fontWeight: 600, color: C.ink, cursor: "pointer", minWidth: 130 }
      },
      a ? a.name : "(deleted account)"
    ), /* @__PURE__ */ React.createElement(
      "span",
      {
        onClick: () => setEditingNote(editingNote === t.id ? null : t.id),
        style: { flex: 1, fontSize: 12.5, color: t.notes ? C.ink : C.inkSoft, cursor: "pointer", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }
      },
      t.notes || "+ add outcome note"
    ), /* @__PURE__ */ React.createElement(DuePill, { due: t.due, done: t.done })), editingNote === t.id && /* @__PURE__ */ React.createElement(
      "textarea",
      {
        value: t.notes || "",
        onChange: (e) => updateTask(t.id, { notes: e.target.value }),
        placeholder: "Outcome \u2014 e.g. left voicemail, wants samples, ordered 2 cases\u2026",
        rows: 2,
        autoFocus: true,
        style: { ...inputStyle, resize: "vertical", marginTop: 8, fontSize: 13 }
      }
    ));
  })));
}
function IdentityPicker({ team, ownerPin, onPick, fontCss }) {
  const [pinFor, setPinFor] = useState(null);
  const [pin, setPin] = useState("");
  const [err, setErr] = useState(false);
  const pick = (m) => {
    if (m.isOwner && ownerPin) {
      setPinFor(m);
      setPin("");
      setErr(false);
      return;
    }
    onPick(m.id);
  };
  const confirmPin = () => {
    if (pin === ownerPin) onPick(pinFor.id);
    else setErr(true);
  };
  return /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_BODY, background: C.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 24 } }, /* @__PURE__ */ React.createElement("style", null, fontCss), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 16, padding: "34px 38px", maxWidth: 440, width: "100%" } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontWeight: 700, fontSize: 24, letterSpacing: "-0.02em", marginBottom: 6, textAlign: "center" } }, "Sales", /* @__PURE__ */ React.createElement("span", { style: { color: C.green } }, "Desk")), !pinFor ? /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 14, textAlign: "center", margin: "0 0 22px" } }, "Who's using SalesDesk on this computer?"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", flexDirection: "column", gap: 8 } }, team.map((m) => /* @__PURE__ */ React.createElement("button", { key: m.id, onClick: () => pick(m), style: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    background: C.chip,
    border: `1px solid ${C.line}`,
    borderRadius: 10,
    padding: "12px 16px",
    cursor: "pointer",
    fontFamily: FONT_BODY,
    fontSize: 15,
    fontWeight: 600,
    color: C.ink,
    textAlign: "left"
  } }, m.name, /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft } }, m.isOwner ? ownerPin ? "owner \u{1F512}" : "owner" : "teammate")))), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 11.5, textAlign: "center", marginTop: 18, marginBottom: 0, lineHeight: 1.5 } }, "This choice is remembered on this device. You can switch anytime from the sidebar.")) : /* @__PURE__ */ React.createElement(React.Fragment, null, /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 14, textAlign: "center", margin: "0 0 16px" } }, "Enter the owner passcode to continue as ", /* @__PURE__ */ React.createElement("b", { style: { color: C.ink } }, pinFor.name), "."), /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "password",
      value: pin,
      autoFocus: true,
      onChange: (e) => {
        setPin(e.target.value);
        setErr(false);
      },
      onKeyDown: (e) => e.key === "Enter" && confirmPin(),
      placeholder: "Owner passcode",
      style: { ...inputStyle, fontSize: 15, padding: "10px 14px", textAlign: "center", borderColor: err ? C.red : C.line }
    }
  ), err && /* @__PURE__ */ React.createElement("p", { style: { color: C.red, fontSize: 12.5, textAlign: "center", margin: "8px 0 0" } }, "That passcode doesn't match."), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 10, justifyContent: "center", marginTop: 16 } }, /* @__PURE__ */ React.createElement(Btn, { onClick: confirmPin }, "Continue"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", onClick: () => setPinFor(null) }, "Back")))));
}
function TeamView({ db, meId, addMember, updateMember, deleteMember, setOwnerPin, chooseIdentity }) {
  const team = db.team || [];
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [pinInput, setPinInput] = useState(db.ownerPin || "");
  const [pinSaved, setPinSaved] = useState(false);
  const hasOwner = team.some((m) => m.isOwner);
  const addTeammate = () => {
    if (!name.trim()) return;
    const firstEver = team.length === 0;
    const id = addMember(name, firstEver || !hasOwner, email);
    if (firstEver) chooseIdentity(id);
    setName("");
    setEmail("");
  };
  const load = (id) => db.tasks.filter((t) => t.assigneeId === id && !t.done).length;
  const missingEmails = team.filter((m) => !(m.email && m.email.trim())).length;
  return /* @__PURE__ */ React.createElement("div", { style: { maxWidth: 660 } }, /* @__PURE__ */ React.createElement("h1", { style: { fontFamily: FONT_HEAD, fontSize: 26, fontWeight: 700, margin: "0 0 4px", letterSpacing: "-0.025em" } }, "Team"), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, margin: "0 0 20px" } }, "Add teammates so you can assign tasks. Each person opens the app and picks their name \u2014 they'll see only the tasks assigned to them. Emails are used for the morning digest."), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: name,
      onChange: (e) => setName(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && addTeammate(),
      placeholder: team.length === 0 ? "Your name (you'll be the owner)" : "Teammate's name",
      style: { ...inputStyle, flex: "1 1 160px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: email,
      onChange: (e) => setEmail(e.target.value),
      onKeyDown: (e) => e.key === "Enter" && addTeammate(),
      type: "email",
      placeholder: "Work email",
      style: { ...inputStyle, flex: "1 1 180px" }
    }
  ), /* @__PURE__ */ React.createElement(Btn, { icon: "plus", onClick: addTeammate }, "Add")), team.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "4px 16px", marginBottom: 24 } }, team.map((m) => /* @__PURE__ */ React.createElement("div", { key: m.id, style: { display: "flex", alignItems: "center", gap: 10, padding: "11px 0", borderBottom: `1px solid ${C.line}`, flexWrap: "wrap" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      value: m.name,
      onChange: (e) => updateMember(m.id, { name: e.target.value }),
      style: { ...inputStyle, border: "none", background: "transparent", padding: 0, fontWeight: 600, fontSize: 14, flex: "1 1 120px" }
    }
  ), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: m.email || "",
      onChange: (e) => updateMember(m.id, { email: e.target.value }),
      type: "email",
      placeholder: "add email\u2026",
      style: { ...inputStyle, border: "none", background: "transparent", padding: 0, fontSize: 12.5, fontFamily: FONT_MONO, color: m.email ? C.inkSoft : C.red, flex: "1 1 150px" }
    }
  ), m.id === meId && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 10.5, color: C.green } }, "you"), /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft } }, load(m.id), " open"), /* @__PURE__ */ React.createElement(
    "button",
    {
      className: "sd-btn",
      onClick: () => updateMember(m.id, { isOwner: !m.isOwner }),
      title: "Owners get full access; teammates see only their own tasks",
      style: {
        fontFamily: FONT_BODY,
        fontSize: 11,
        fontWeight: 600,
        padding: "3px 10px",
        borderRadius: 999,
        cursor: "pointer",
        border: `1px solid ${m.isOwner ? C.green : C.line}`,
        background: m.isOwner ? C.greenSoft : "transparent",
        color: m.isOwner ? C.green : C.inkSoft
      }
    },
    m.isOwner ? "owner" : "teammate"
  ), /* @__PURE__ */ React.createElement(Btn, { kind: "danger", small: true, onClick: () => {
    if (window.confirm(`Remove ${m.name}? Their assigned tasks become unassigned.`)) deleteMember(m.id);
  } }, "\u2715")))), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "16px 18px", marginBottom: 24 } }, /* @__PURE__ */ React.createElement("div", { style: { display: "flex", alignItems: "center", gap: 8, marginBottom: 4 } }, /* @__PURE__ */ React.createElement("span", { style: { width: 26, height: 26, borderRadius: 7, background: C.brassSoft, color: C.brass, display: "flex", alignItems: "center", justifyContent: "center", flex: "none" } }, /* @__PURE__ */ React.createElement(Icon, { name: "bell", size: 15 })), /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontSize: 14, fontWeight: 600 } }, "Morning email digest")), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 12.5, margin: "0 0 10px", lineHeight: 1.55 } }, "A Power Automate flow can email everyone their due & overdue tasks each morning. SalesDesk writes a ready-to-read ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "outbox"), " section into your data file for the flow to use \u2014 no setup needed inside the app. Follow the build guide to create the flow once."), team.length === 0 ? /* @__PURE__ */ React.createElement("div", { style: { fontSize: 12.5, color: C.brass } }, "Add yourself and your teammates above (with emails) to enable digests.") : missingEmails > 0 ? /* @__PURE__ */ React.createElement("div", { style: { fontSize: 12.5, color: C.red } }, "\u26A0 ", missingEmails, " ", missingEmails === 1 ? "person is" : "people are", " missing an email \u2014 they won't receive a digest until you add one.") : /* @__PURE__ */ React.createElement("div", { style: { fontSize: 12.5, color: C.green } }, "\u2713 Everyone has an email. Your data file is ready for the flow.")), team.length > 0 && /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 12, padding: "16px 18px" } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontSize: 13, fontWeight: 600, marginBottom: 4 } }, "Owner passcode ", /* @__PURE__ */ React.createElement("span", { style: { fontWeight: 400, color: C.inkSoft } }, "(optional)")), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 12.5, margin: "0 0 12px", lineHeight: 1.5 } }, "Set a passcode and anyone picking an owner name must enter it \u2014 this keeps teammates from casually opening the full view. It's a light lock, not encryption: the shared file still contains all data."), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center" } }, /* @__PURE__ */ React.createElement(
    "input",
    {
      type: "text",
      value: pinInput,
      onChange: (e) => {
        setPinInput(e.target.value);
        setPinSaved(false);
      },
      placeholder: "No passcode set",
      style: { ...inputStyle, width: 200 }
    }
  ), /* @__PURE__ */ React.createElement(Btn, { onClick: () => {
    setOwnerPin(pinInput.trim());
    setPinSaved(true);
  } }, "Save"), pinInput && /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: () => {
    setPinInput("");
    setOwnerPin("");
    setPinSaved(true);
  } }, "Clear"), pinSaved && /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.green } }, "\u2713 saved"))), /* @__PURE__ */ React.createElement("div", { style: { marginTop: 22, fontSize: 12.5, color: C.inkSoft, lineHeight: 1.6, background: C.chip, borderRadius: 10, padding: "14px 16px" } }, /* @__PURE__ */ React.createElement("b", { style: { color: C.ink } }, "How teammates get in:"), " share the same ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "SalesDesk.html"), " and give them access to the OneDrive folder holding ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "salesdesk-data.json"), ". When they open it and connect that file, they pick their name and see only their tasks."));
}
function fillTemplate(str, person, account) {
  return (str || "").replace(/\{\{\s*name\s*\}\}/gi, (person.name || "").split(" ")[0]).replace(/\{\{\s*fullname\s*\}\}/gi, person.name || "").replace(/\{\{\s*account\s*\}\}/gi, account.name || "").replace(/\{\{\s*role\s*\}\}/gi, person.role || "");
}
function ComposeModal({ person, account, templates, onClose, onSent, addTemplate, updateTemplate, deleteTemplate }) {
  const [subject, setSubject] = useState("");
  const [body, setBody] = useState("");
  const [manageOpen, setManageOpen] = useState(false);
  const [tplName, setTplName] = useState("");
  const applyTemplate = (tpl) => {
    setSubject(fillTemplate(tpl.subject, person, account));
    setBody(fillTemplate(tpl.body, person, account));
  };
  const openInMail = () => {
    const url = `mailto:${encodeURIComponent(person.email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    window.location.href = url;
    if (onSent) onSent();
  };
  const saveTemplate = () => {
    if (!tplName.trim()) return;
    addTemplate({ name: tplName, subject, body });
    setTplName("");
    setManageOpen(false);
  };
  return /* @__PURE__ */ React.createElement("div", { onClick: onClose, style: { position: "fixed", inset: 0, background: "rgba(16,33,25,.38)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, zIndex: 50 } }, /* @__PURE__ */ React.createElement("div", { onClick: (e) => e.stopPropagation(), style: { background: C.surface, borderRadius: 16, width: "100%", maxWidth: 560, maxHeight: "90vh", overflow: "auto", boxShadow: "0 24px 60px rgba(16,33,25,.28)" } }, /* @__PURE__ */ React.createElement("div", { style: { padding: "18px 22px", borderBottom: `1px solid ${C.line}`, display: "flex", alignItems: "center", gap: 10 } }, /* @__PURE__ */ React.createElement("span", { style: { width: 30, height: 30, borderRadius: 8, background: C.greenSoft, color: C.green, display: "flex", alignItems: "center", justifyContent: "center", flex: "none" } }, /* @__PURE__ */ React.createElement(Icon, { name: "mail", size: 17 })), /* @__PURE__ */ React.createElement("div", { style: { flex: 1, minWidth: 0 } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontWeight: 600, fontSize: 15 } }, "Email ", person.name), /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_MONO, fontSize: 11, color: C.inkSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" } }, person.email)), /* @__PURE__ */ React.createElement("button", { className: "sd-link", onClick: onClose, style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 20, lineHeight: 1, padding: 4 } }, "\xD7")), /* @__PURE__ */ React.createElement("div", { style: { padding: "18px 22px", display: "flex", flexDirection: "column", gap: 12 } }, templates.length > 0 && /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("div", { style: { fontSize: 12, color: C.inkSoft, marginBottom: 6 } }, "Start from a template"), /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 6, flexWrap: "wrap" } }, templates.map((tpl) => /* @__PURE__ */ React.createElement("span", { key: tpl.id, style: { display: "inline-flex", alignItems: "center", gap: 4, background: C.chip, borderRadius: 20, padding: "3px 4px 3px 11px" } }, /* @__PURE__ */ React.createElement("button", { className: "sd-link", onClick: () => applyTemplate(tpl), style: { background: "none", border: "none", cursor: "pointer", fontSize: 12, fontWeight: 600, color: C.ink, padding: 0 } }, tpl.name), /* @__PURE__ */ React.createElement("button", { onClick: () => deleteTemplate(tpl.id), title: "Delete template", style: { background: "none", border: "none", cursor: "pointer", color: C.inkSoft, fontSize: 13, lineHeight: 1, padding: "0 3px" } }, "\xD7"))))), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, color: C.inkSoft, display: "block", marginBottom: 4 } }, "Subject"), /* @__PURE__ */ React.createElement("input", { value: subject, onChange: (e) => setSubject(e.target.value), placeholder: "Subject line", style: inputStyle })), /* @__PURE__ */ React.createElement("div", null, /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, color: C.inkSoft, display: "block", marginBottom: 4 } }, "Message"), /* @__PURE__ */ React.createElement(
    "textarea",
    {
      value: body,
      onChange: (e) => setBody(e.target.value),
      rows: 8,
      placeholder: `Hi ${(person.name || "there").split(" ")[0]},

\u2026`,
      style: { ...inputStyle, resize: "vertical", fontSize: 13.5, lineHeight: 1.5 }
    }
  ), /* @__PURE__ */ React.createElement("div", { style: { fontSize: 11, color: C.inkSoft, marginTop: 5 } }, "Placeholders: ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "{{name}}"), ", ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "{{account}}"), ", ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "{{role}}"), " fill in automatically.")), manageOpen ? /* @__PURE__ */ React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center", background: C.chip, borderRadius: 10, padding: "10px 12px" } }, /* @__PURE__ */ React.createElement("input", { value: tplName, onChange: (e) => setTplName(e.target.value), placeholder: "Template name", style: { ...inputStyle, flex: 1 } }), /* @__PURE__ */ React.createElement(Btn, { small: true, onClick: saveTemplate }, "Save template"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", small: true, onClick: () => setManageOpen(false) }, "Cancel")) : /* @__PURE__ */ React.createElement(
    "button",
    {
      className: "sd-link",
      onClick: () => setManageOpen(true),
      style: { background: "none", border: "none", cursor: "pointer", color: C.green, fontSize: 12.5, fontWeight: 600, padding: 0, alignSelf: "flex-start" }
    },
    "+ Save this as a reusable template"
  )), /* @__PURE__ */ React.createElement("div", { style: { padding: "14px 22px", borderTop: `1px solid ${C.line}`, display: "flex", gap: 10, alignItems: "center" } }, /* @__PURE__ */ React.createElement(Btn, { icon: "mail", onClick: openInMail }, "Open in email"), /* @__PURE__ */ React.createElement(Btn, { kind: "ghost", onClick: onClose }, "Cancel"), /* @__PURE__ */ React.createElement("div", { style: { flex: 1 } }), /* @__PURE__ */ React.createElement("span", { style: { fontSize: 11, color: C.inkSoft, textAlign: "right", maxWidth: 220 } }, "Opens a ready-to-send draft in your mail app. You send it \u2014 nothing leaves without you."))));
}
function AIAssistant() {
  return /* @__PURE__ */ React.createElement(
    "a",
    {
      href: "https://chatgpt.com/",
      target: "_blank",
      rel: "noopener noreferrer",
      title: "Open ChatGPT in a new tab",
      className: "sd-btn",
      style: {
        position: "fixed",
        right: 20,
        bottom: 20,
        zIndex: 60,
        display: "inline-flex",
        alignItems: "center",
        gap: 8,
        background: C.green,
        color: "#fff",
        textDecoration: "none",
        border: "none",
        borderRadius: 999,
        padding: "11px 17px",
        fontFamily: FONT_BODY,
        fontWeight: 600,
        fontSize: 13.5,
        cursor: "pointer",
        boxShadow: "0 6px 20px rgba(16,40,30,.24)"
      }
    },
    /* @__PURE__ */ React.createElement(Icon, { name: "spark", size: 16 }),
    " ChatGPT"
  );
}
function ConfigScreen({ cfg, err, fontCss, onSave }) {
  const [clientId, setClientId] = useState(cfg.clientId || "");
  const [path, setPath] = useState(cfg.path || "SalesDesk/salesdesk-data.json");
  const valid = clientId.trim().length > 20 && path.trim().length > 0;
  return /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_BODY, background: C.bg, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", padding: 20 } }, /* @__PURE__ */ React.createElement("style", null, fontCss), /* @__PURE__ */ React.createElement("div", { style: { background: C.surface, border: `1px solid ${C.line}`, borderRadius: 16, padding: "30px 26px", maxWidth: 460, width: "100%" } }, /* @__PURE__ */ React.createElement("div", { style: { fontFamily: FONT_HEAD, fontWeight: 700, fontSize: 25, letterSpacing: "-0.02em", marginBottom: 6, textAlign: "center" } }, "Sales", /* @__PURE__ */ React.createElement("span", { style: { color: C.green } }, "Desk")), /* @__PURE__ */ React.createElement("p", { style: { color: C.inkSoft, fontSize: 13.5, lineHeight: 1.55, margin: "0 0 20px", textAlign: "center" } }, "One-time setup. Paste the Client ID from your Microsoft app registration and confirm where your data file lives in OneDrive."), /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, fontWeight: 600, color: C.ink, display: "block", marginBottom: 5 } }, "Application (client) ID"), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: clientId,
      onChange: (e) => setClientId(e.target.value),
      placeholder: "00000000-0000-0000-0000-000000000000",
      style: { ...inputStyle, fontFamily: FONT_MONO, fontSize: 12.5, marginBottom: 14 }
    }
  ), /* @__PURE__ */ React.createElement("label", { style: { fontSize: 12, fontWeight: 600, color: C.ink, display: "block", marginBottom: 5 } }, "Path to your data file in OneDrive"), /* @__PURE__ */ React.createElement(
    "input",
    {
      value: path,
      onChange: (e) => setPath(e.target.value),
      placeholder: "SalesDesk/salesdesk-data.json",
      style: { ...inputStyle, fontFamily: FONT_MONO, fontSize: 12.5, marginBottom: 6 }
    }
  ), /* @__PURE__ */ React.createElement("p", { style: { fontSize: 11, color: C.inkSoft, margin: "0 0 18px", lineHeight: 1.45 } }, "Relative to the root of your OneDrive. If the file sits in a folder called ", /* @__PURE__ */ React.createElement("span", { style: { fontFamily: FONT_MONO } }, "SalesDesk"), ", the path above is right. The file must already exist \u2014 create it once from the desktop app."), /* @__PURE__ */ React.createElement(
    Btn,
    {
      onClick: () => valid && onSave({ clientId: clientId.trim(), path: path.trim().replace(/^\/+/, "") }),
      style: { width: "100%", justifyContent: "center", fontSize: 14.5, padding: "11px 18px", opacity: valid ? 1 : 0.5 }
    },
    "Save and continue"
  ), err && /* @__PURE__ */ React.createElement("p", { style: { color: C.red, fontSize: 12, marginTop: 12, marginBottom: 0 } }, err), /* @__PURE__ */ React.createElement("p", { style: { fontSize: 11, color: C.inkSoft, marginTop: 14, marginBottom: 0, textAlign: "center" } }, "These settings stay on this device. Follow the setup guide if you haven't registered the app yet.")));
}
ReactDOM.createRoot(document.getElementById("root")).render(/* @__PURE__ */ React.createElement(SalesDesk, null));
