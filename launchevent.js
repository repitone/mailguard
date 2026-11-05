/*
 * MailGuard rule engine — pure functions, no Office.js dependency.
 * The same file is inlined into launchevent.js (classic Outlook does not
 * support imports in event-based runtimes) and loaded by the task pane.
 *
 * Core principle: "1 mail = 1 external counterparty company".
 * Forwarders and neutral parties (banks, inspectors) may join any mail.
 */
var MailGuardEngine = (function () {
  "use strict";

  var PUBLIC_DOMAINS = [
    "gmail.com", "googlemail.com", "outlook.com", "hotmail.com", "live.com",
    "yahoo.com", "yahoo.co.kr", "naver.com", "daum.net", "hanmail.net",
    "nate.com", "icloud.com", "me.com", "qq.com", "163.com", "126.com",
    "aol.com", "proton.me", "protonmail.com"
  ];

  var EMAIL_RE = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;

  // Lines that start the quoted part of a reply/forward.
  var QUOTE_MARKERS = [
    /^-{2,}\s*original message\s*-{2,}/im,
    /^-{2,}\s*원본 메시지\s*-{2,}/im,
    /^_{8,}\s*$/m,
    /^\s*from:\s.+$/im,
    /^\s*보낸 사람:\s.+$/im,
    /^\s*발신:\s.+$/im,
    /^on .{5,120} wrote:\s*$/im
  ];

  var ATTACH_WORDS = /(첨부|attached|attachment|enclosed|please find|pls find|첨부파일|첨부드|첨부합)/i;

  function lower(s) { return String(s || "").trim().toLowerCase(); }

  function domainOf(email) {
    var e = lower(email);
    var i = e.lastIndexOf("@");
    return i >= 0 ? e.slice(i + 1) : "";
  }

  function domainMatches(domain, ruleDomain) {
    var d = lower(domain), r = lower(ruleDomain).replace(/^@/, "");
    return d === r || (d.length > r.length && d.slice(-(r.length + 1)) === "." + r);
  }

  function isInternal(email, rules) {
    var d = domainOf(email);
    var list = (rules && rules.internalDomains) || [];
    for (var i = 0; i < list.length; i++) if (domainMatches(d, list[i])) return true;
    return false;
  }

  // Resolve an address to a company. Exact address first, then domain
  // (public mail domains only resolve by exact address).
  function companyOf(email, rules) {
    var e = lower(email), d = domainOf(e);
    var cs = (rules && rules.companies) || [];
    for (var i = 0; i < cs.length; i++) {
      var addrs = cs[i].addresses || [];
      for (var j = 0; j < addrs.length; j++) if (lower(addrs[j]) === e) return cs[i];
    }
    if (PUBLIC_DOMAINS.indexOf(d) >= 0) return null;
    for (var k = 0; k < cs.length; k++) {
      var ds = cs[k].domains || [];
      for (var m = 0; m < ds.length; m++) if (domainMatches(d, ds[m])) return cs[k];
    }
    return null;
  }

  function splitQuoted(text) {
    var t = String(text || ""), cut = t.length;
    for (var i = 0; i < QUOTE_MARKERS.length; i++) {
      var m = QUOTE_MARKERS[i].exec(t);
      if (m && m.index < cut && m.index > 0) cut = m.index;
    }
    return { fresh: t.slice(0, cut), quoted: t.slice(cut) };
  }

  function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

  // Whole-word-ish match that also works for Korean (no \b support).
  function mentions(text, alias) {
    var a = String(alias || "").trim();
    if (a.length < 3 && !/[가-힣]{2,}/.test(a)) return false;
    var re = new RegExp("(^|[^A-Za-z0-9가-힣])" + escapeRe(a) + "($|[^A-Za-z0-9])", "i");
    return re.test(text);
  }

  function namesOf(c) {
    var out = [];
    if (c.name) out.push(c.name);
    (c.aliases || []).forEach(function (a) { if (a) out.push(a); });
    return out;
  }

  function inAllowedGroup(ids, rules) {
    if (ids.length < 2) return true;
    var groups = (rules && rules.allowedGroups) || [];
    for (var i = 0; i < groups.length; i++) {
      var g = groups[i];
      if (ids.every(function (id) { return g.indexOf(id) >= 0; })) return true;
    }
    return false;
  }

  function uniq(arr) {
    var seen = {}, out = [];
    arr.forEach(function (x) { if (!seen[x]) { seen[x] = 1; out.push(x); } });
    return out;
  }

  /*
   * msg = { recipients:[{email, field:'to'|'cc'|'bcc'}], subject, body (plain text),
   *         attachments:[{name}] }
   * returns { issues:[{code, level:'block'|'warn', title, detail, fix?}], unknownDomains:[], signature }
   */
  function check(msg, rules) {
    rules = rules || {};
    var issues = [];
    var recips = (msg.recipients || []).filter(function (r) { return r && r.email; });
    var external = recips.filter(function (r) { return !isInternal(r.email, rules); });

    if (external.length === 0) {
      return { issues: [], unknownDomains: [], signature: signature(msg, []) };
    }

    var counterparties = {}, forwarders = {}, unknown = [];
    external.forEach(function (r) {
      var c = companyOf(r.email, rules);
      if (!c) { unknown.push(r); return; }
      if (c.type === "forwarder") forwarders[c.id] = c;
      else if (c.type === "neutral") { /* always allowed */ }
      else {
        if (!counterparties[c.id]) counterparties[c.id] = { company: c, emails: [] };
        counterparties[c.id].emails.push(r.email);
      }
    });
    var cpIds = Object.keys(counterparties);
    var fwIds = Object.keys(forwarders);

    // B1 — more than one counterparty company in the same mail
    if (cpIds.length >= 2 && !inAllowedGroup(cpIds, rules)) {
      issues.push({
        code: "B1", level: "block",
        title: "서로 다른 거래처 " + cpIds.length + "곳이 한 메일에 있습니다",
        detail: cpIds.map(function (id) {
          return counterparties[id].company.name + " (" + counterparties[id].emails.join(", ") + ")";
        }).join(" / "),
        fix: { type: "removeRecipients", options: cpIds.map(function (id) {
          return { label: counterparties[id].company.name + " 제외", emails: counterparties[id].emails };
        }) }
      });
    }

    var text = String(msg.body || "");
    var parts = splitQuoted(text);
    var bodyEmails = uniq((text.match(EMAIL_RE) || []).map(lower));
    var recipSet = {};
    recips.forEach(function (r) { recipSet[lower(r.email)] = 1; });

    // B2 — another counterparty's address appears in the body / quoted history
    var leaked = {};
    bodyEmails.forEach(function (e) {
      if (recipSet[e] || isInternal(e, rules)) return;
      var c = companyOf(e, rules);
      if (!c || c.type === "forwarder" || c.type === "neutral") return;
      if (counterparties[c.id]) return;
      if (isGroupMember(c.id, cpIds, rules)) return;
      if (!leaked[c.id]) leaked[c.id] = { company: c, emails: [] };
      leaked[c.id].emails.push(e);
    });
    var leakedIds = Object.keys(leaked);
    if (leakedIds.length) {
      var toCp = cpIds.length > 0;
      var inQuoted = leakedIds.some(function (id) {
        return leaked[id].emails.some(function (e) { return parts.quoted.toLowerCase().indexOf(e) >= 0; });
      });
      issues.push({
        code: "B2", level: toCp ? "block" : "warn",
        title: (inQuoted ? "인용된 이전 메일" : "본문") + "에 다른 거래처 주소가 남아 있습니다",
        detail: leakedIds.map(function (id) {
          return leaked[id].company.name + ": " + leaked[id].emails.join(", ");
        }).join(" / "),
        fix: inQuoted ? { type: "stripQuoted" } : { type: "maskEmails", emails: [].concat.apply([], leakedIds.map(function (id) { return leaked[id].emails; })) }
      });
    }

    // W1 — another counterparty's name/alias is mentioned (only when writing to a counterparty)
    if (cpIds.length) {
      var hay = (msg.subject || "") + "\n" + text;
      var named = [];
      (rules.companies || []).forEach(function (c) {
        if (c.type === "forwarder" || c.type === "neutral") return;
        if (counterparties[c.id] || leaked[c.id]) return;
        if (isGroupMember(c.id, cpIds, rules)) return;
        var hit = namesOf(c).filter(function (n) { return mentions(hay, n); });
        if (hit.length) named.push(c.name + " (" + uniq(hit).join(", ") + ")");
      });
      if (named.length) {
        issues.push({
          code: "W1", level: "warn",
          title: "다른 거래처 이름이 본문·제목에 있습니다",
          detail: named.join(" / ")
        });
      }

      // W3 — attachment file names mention another counterparty
      var badFiles = [];
      (msg.attachments || []).forEach(function (a) {
        (rules.companies || []).forEach(function (c) {
          if (c.type === "forwarder" || c.type === "neutral" || counterparties[c.id]) return;
          if (isGroupMember(c.id, cpIds, rules)) return;
          if (namesOf(c).some(function (n) { return mentions(a.name || "", n); })) badFiles.push(a.name + " → " + c.name);
        });
      });
      if (badFiles.length) {
        issues.push({ code: "W3", level: "warn", title: "첨부 파일명에 다른 거래처 이름이 있습니다", detail: uniq(badFiles).join(" / ") });
      }
    }

    // W2 — forwarder and counterparty together without approval
    if (fwIds.length && cpIds.length) {
      var allIds = cpIds.concat(fwIds);
      if (!inAllowedGroup(allIds, rules)) {
        issues.push({
          code: "W2", level: "warn",
          title: "Forwarder와 거래처가 함께 수신합니다",
          detail: fwIds.map(function (id) { return forwarders[id].name; }).join(", ") + " + " +
            cpIds.map(function (id) { return counterparties[id].company.name; }).join(", ")
        });
      }
    }

    // W4 — mentions an attachment but none attached (fresh text only)
    var realFiles = (msg.attachments || []).filter(function (a) { return !a.isInline; });
    if (ATTACH_WORDS.test(parts.fresh) && !realFiles.length) {
      issues.push({ code: "W4", level: "warn", title: "첨부를 언급했지만 첨부 파일이 없습니다", detail: "" });
    }

    // U — unknown external domains (asked once; caller filters already-seen ones)
    var unknownDomains = uniq(unknown.map(function (r) {
      var d = domainOf(r.email);
      return PUBLIC_DOMAINS.indexOf(d) >= 0 ? lower(r.email) : d;
    })).filter(function (d) { return (rules.seenUnknown || []).indexOf(d) < 0; });
    if (unknownDomains.length) {
      issues.push({
        code: "U1", level: "warn",
        title: "처음 보는 외부 " + (unknownDomains.length > 1 ? "주소 " + unknownDomains.length + "개" : "주소") +
          (cpIds.length ? " (등록된 거래처와 함께 수신)" : ""),
        detail: unknownDomains.join(", "),
        fix: { type: "tagDomains", domains: unknownDomains }
      });
    }

    return { issues: issues, unknownDomains: unknownDomains, signature: signature(msg, issues) };
  }

  function isGroupMember(id, cpIds, rules) {
    if (!cpIds.length) return false;
    return inAllowedGroup(cpIds.concat([id]), rules);
  }

  // Identifies "this message with these problems" so an approved override
  // only applies while recipients and issues stay the same.
  function signature(msg, issues) {
    var r = (msg.recipients || []).map(function (x) { return lower(x.email); }).sort().join(",");
    var c = issues.map(function (i) { return i.code + ":" + i.detail; }).sort().join("|");
    var s = r + "#" + c, h = 5381;
    for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0;
    return h.toString(36);
  }

  // Short Markdown for the Smart Alerts dialog (limit 500 chars).
  function dialogText(result) {
    var lines = [];
    var blocks = result.issues.filter(function (i) { return i.level === "block"; });
    var warns = result.issues.filter(function (i) { return i.level !== "block"; });
    blocks.concat(warns).forEach(function (i) {
      var d = i.detail ? " — " + i.detail : "";
      lines.push("- **" + i.title + "**" + d + "\r");
    });
    var head = blocks.length ? "발송 전에 확인이 필요합니다.\n\n" : "확인해 주세요.\n\n";
    var tail = blocks.length ? "\n\n[수정하기]에서 원클릭 수정 또는 사유 입력 후 발송" : "";
    var body = head + lines.join("");
    var max = 500 - tail.length;
    if (body.length > max) body = body.slice(0, max - 2) + "…\r";
    return body + tail;
  }

  return {
    check: check, dialogText: dialogText, splitQuoted: splitQuoted,
    companyOf: companyOf, domainOf: domainOf, isInternal: isInternal,
    PUBLIC_DOMAINS: PUBLIC_DOMAINS
  };
})();


/*
 * Smart Alerts OnMessageSend handler.
 * Runs automatically every time Send is selected. Clean mail passes with no UI.
 * build.js prepends engine.js to this file to produce dist/launchevent.js.
 */
/* global Office, MailGuardEngine */

var MG_RULES_KEY = "mg_rules";
var MG_OVERRIDE_KEY = "mg_override";

function mgLoadRules() {
  try {
    var raw = Office.context.roamingSettings.get(MG_RULES_KEY);
    if (!raw) return null;
    return typeof raw === "string" ? JSON.parse(raw) : raw;
  } catch (e) {
    return null;
  }
}

function mgAsync(fn) {
  return new Promise(function (resolve) {
    try {
      fn(function (res) {
        resolve(res && res.status === Office.AsyncResultStatus.Succeeded ? res.value : null);
      });
    } catch (e) {
      resolve(null);
    }
  });
}

function mgRecipients(field, name) {
  if (!field) return Promise.resolve([]);
  return mgAsync(function (cb) { field.getAsync(cb); }).then(function (list) {
    return (list || []).map(function (r) { return { email: r.emailAddress, field: name }; });
  });
}

function mgPlain(markdown) {
  return markdown.replace(/\*\*/g, "").replace(/\r/g, "\n").slice(0, 500);
}

function mgMarkSeen(rules, domains) {
  // "Ask once": remember unknown domains so the next mail to them passes.
  return new Promise(function (resolve) {
    try {
      rules.seenUnknown = (rules.seenUnknown || []).concat(domains).slice(-300);
      Office.context.roamingSettings.set(MG_RULES_KEY, JSON.stringify(rules));
      var done = false;
      if (typeof setTimeout === "function") {
        setTimeout(function () { if (!done) { done = true; resolve(); } }, 1500);
      }
      Office.context.roamingSettings.saveAsync(function () { if (!done) { done = true; resolve(); } });
    } catch (e) {
      resolve();
    }
  });
}

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  var rules = mgLoadRules();

  // Not configured yet: never get in the way.
  if (!rules || !rules.internalDomains || !rules.internalDomains.length) {
    event.completed({ allowEvent: true });
    return;
  }

  Promise.all([
    mgRecipients(item.to, "to"),
    mgRecipients(item.cc, "cc"),
    mgRecipients(item.bcc, "bcc"),
    mgAsync(function (cb) { item.subject.getAsync(cb); }),
    mgAsync(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }),
    mgAsync(function (cb) { item.getAttachmentsAsync(cb); }),
    item.sessionData ? mgAsync(function (cb) { item.sessionData.getAsync(MG_OVERRIDE_KEY, cb); }) : Promise.resolve(null)
  ]).then(function (v) {
    var msg = {
      recipients: v[0].concat(v[1], v[2]),
      subject: v[3] || "",
      body: v[4] || "",
      attachments: (v[5] || []).map(function (a) { return { name: a.name, isInline: !!a.isInline }; })
    };
    var result = MailGuardEngine.check(msg, rules);

    if (!result.issues.length) {
      event.completed({ allowEvent: true });
      return;
    }

    // Approved in the task pane ("사유 기록 후 보내기") for exactly this situation.
    if (v[6] && v[6] === result.signature) {
      event.completed({ allowEvent: true });
      return;
    }

    var hasBlock = result.issues.some(function (i) { return i.level === "block"; });
    var md = MailGuardEngine.dialogText(result);
    var options = {
      allowEvent: false,
      errorMessage: mgPlain(md),
      errorMessageMarkdown: md,
      cancelLabel: "수정하기",
      commandId: "mgTaskpaneButton",
      contextData: JSON.stringify({ sig: result.signature })
    };
    // Warnings only: Outlook shows "Send Anyway". Blocks: fix or record a reason first.
    if (!hasBlock && Office.MailboxEnums && Office.MailboxEnums.SendModeOverride) {
      options.sendModeOverride = Office.MailboxEnums.SendModeOverride.PromptUser;
    }

    var pending = result.unknownDomains.length ? mgMarkSeen(rules, result.unknownDomains) : Promise.resolve();
    pending.then(function () { event.completed(options); });
  }).catch(function () {
    // Never trap the user because of our own bug: allow Send Anyway.
    var opt = { allowEvent: false, errorMessage: "MailGuard 검사 중 오류가 발생했습니다. 수신자를 직접 확인한 뒤 보내세요." };
    if (Office.MailboxEnums && Office.MailboxEnums.SendModeOverride) {
      opt.sendModeOverride = Office.MailboxEnums.SendModeOverride.PromptUser;
    }
    event.completed(opt);
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
