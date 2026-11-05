/* global Office, MailGuardEngine */
(function () {
  "use strict";

  var RULES_KEY = "mg_rules", LOG_KEY = "mg_log", OVERRIDE_KEY = "mg_override";
  var LIMIT = 32 * 1024; // roaming settings limit per add-in
  var TYPES = { counterparty: "거래처", forwarder: "Forwarder", neutral: "중립 기관" };

  var rules, log, item, lastResult = null, armed = {};

  // ---------- helpers ----------
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function splitList(s) { return String(s || "").split(/[,\s;]+/).map(function (x) { return x.trim().toLowerCase().replace(/^@/, ""); }).filter(Boolean); }
  function toast(msg) { var t = $("toast"); t.textContent = msg; t.style.display = "block"; clearTimeout(toast.t); toast.t = setTimeout(function () { t.style.display = "none"; }, 2600); }
  function slug(name) { return (String(name).toLowerCase().replace(/[^a-z0-9가-힣]+/g, "-").replace(/^-|-$/g, "") || "c") + "-" + Date.now().toString(36).slice(-4); }
  function call(fn) { return new Promise(function (res, rej) { fn(function (r) { r.status === Office.AsyncResultStatus.Succeeded ? res(r.value) : rej(r.error || new Error("failed")); }); }); }
  function blank() { return { version: 1, internalDomains: [], companies: [], allowedGroups: [], seenUnknown: [] }; }

  function load() {
    var rs = Office.context.roamingSettings;
    try { rules = JSON.parse(rs.get(RULES_KEY) || "null") || blank(); } catch (e) { rules = blank(); }
    try { log = JSON.parse(rs.get(LOG_KEY) || "[]"); } catch (e) { log = []; }
    ["internalDomains", "companies", "allowedGroups", "seenUnknown"].forEach(function (k) { if (!Array.isArray(rules[k])) rules[k] = []; });
  }

  function size() { return JSON.stringify(rules).length * 2 + JSON.stringify(log).length * 2; }

  function save() {
    if (size() > LIMIT) {
      rules.seenUnknown = rules.seenUnknown.slice(-50);
      log = log.slice(-20);
      if (size() > LIMIT) { toast("저장 공간(32KB)을 초과했습니다. 별칭·주소를 줄여 주세요."); return Promise.reject(new Error("size")); }
    }
    var rs = Office.context.roamingSettings;
    rs.set(RULES_KEY, JSON.stringify(rules));
    rs.set(LOG_KEY, JSON.stringify(log));
    return call(function (cb) { rs.saveAsync(cb); });
  }

  function isCompose() { return !!(item && item.to && item.to.setAsync); }

  // ---------- current mail ----------
  function readMessage() {
    function rec(f, name) { return f ? call(function (cb) { f.getAsync(cb); }).then(function (l) { return l.map(function (r) { return { email: r.emailAddress, name: r.displayName, field: name }; }); }) : Promise.resolve([]); }
    return Promise.all([
      rec(item.to, "to"), rec(item.cc, "cc"), rec(item.bcc, "bcc"),
      call(function (cb) { item.subject.getAsync(cb); }),
      call(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); }),
      call(function (cb) { item.getAttachmentsAsync(cb); })
    ]).then(function (v) {
      return { recipients: v[0].concat(v[1], v[2]), subject: v[3], body: v[4],
        attachments: v[5].map(function (a) { return { name: a.name, isInline: !!a.isInline }; }) };
    });
  }

  function runCheck() {
    var out = $("check-out");
    if (!isCompose()) { out.innerHTML = '<p class="muted">메일을 작성하거나 회신할 때 검사 결과가 여기 표시됩니다.</p>'; return; }
    if (!rules.internalDomains.length) {
      out.innerHTML = '<div class="card warn"><div class="title">처음 설정이 필요합니다</div><div class="detail">[설정] 탭에서 내 회사 도메인을 입력하고 거래처 목록을 가져오세요.</div></div>';
      return;
    }
    // Check against rules without the "seen" filter so the pane always shows unknowns.
    var view = JSON.parse(JSON.stringify(rules)); view.seenUnknown = [];
    readMessage().then(function (msg) {
      lastResult = MailGuardEngine.check(msg, view);
      var sigForSend = MailGuardEngine.check(msg, rules).signature;
      renderIssues(lastResult, sigForSend);
    }).catch(function () { out.innerHTML = '<p class="muted">메일 정보를 읽지 못했습니다.</p>'; });
  }

  function renderIssues(res, sig) {
    var out = $("check-out");
    if (!res.issues.length) {
      out.innerHTML = '<div class="card ok"><div class="title">문제가 없습니다</div><div class="detail">그대로 보내셔도 됩니다.</div>' +
        (canSend() ? '<div class="actions"><button class="btn primary" id="send-now">보내기</button></div>' : '') + '</div>';
      if ($("send-now")) $("send-now").onclick = function () { sendNow(); };
      return;
    }
    var html = res.issues.map(function (i, n) {
      return '<div class="card ' + (i.level === "block" ? "block" : "warn") + '"><span class="tag">' + (i.level === "block" ? "차단" : "경고") + '</span>' +
        '<span class="title">' + esc(i.title) + '</span>' + (i.detail ? '<div class="detail">' + esc(i.detail) + '</div>' : "") +
        fixHtml(i, n) + '</div>';
    }).join("");
    html += '<h3>사유 입력 후 그대로 발송</h3><textarea id="ov-reason" placeholder="예: T4가 T3와 직접 거래하는 건으로 사전 합의됨"></textarea>' +
      '<div class="actions"><button class="btn" id="ov-send">사유 기록 후 보내기</button></div>' +
      '<p class="muted small">기록은 [설정] 탭에서 다시 볼 수 있습니다.</p>';
    out.innerHTML = html;

    res.issues.forEach(function (i, n) { bindFix(i, n); });
    $("ov-send").onclick = function () { overrideSend(res, sig); };
  }

  function fixHtml(i, n) {
    if (!i.fix) return "";
    var f = i.fix;
    if (f.type === "removeRecipients") return '<div class="actions">' + f.options.map(function (o, k) { return '<button class="btn" id="fx' + n + '-' + k + '">' + esc(o.label) + '</button>'; }).join("") + '</div>';
    if (f.type === "stripQuoted") return '<div class="actions"><button class="btn danger" id="fx' + n + '">인용된 이전 메일 삭제</button></div>';
    if (f.type === "maskEmails") return '<div class="actions"><button class="btn" id="fx' + n + '">본문에서 주소 지우기</button></div>';
    if (f.type === "tagDomains") return '<div id="fx' + n + '">' + f.domains.map(function (d, k) { return tagRow(d, n, k); }).join("") + '</div>';
    return "";
  }

  function tagRow(d, n, k) {
    var opts = '<option value="">역할 선택…</option><option value="__new">새 거래처</option><option value="__fw">Forwarder</option><option value="__neutral">중립 기관</option><option value="__internal">내 회사</option>';
    if (rules.companies.length) opts += '<optgroup label="기존 회사에 추가">' + rules.companies.map(function (c) { return '<option value="' + esc(c.id) + '">' + esc(c.name) + '</option>'; }).join("") + '</optgroup>';
    return '<label>' + esc(d) + '</label><div class="row"><select id="tg' + n + '-' + k + '">' + opts + '</select>' +
      '<input id="tn' + n + '-' + k + '" placeholder="회사명" hidden><button class="btn" id="tb' + n + '-' + k + '" style="flex:0 0 auto">저장</button></div>';
  }

  function bindFix(i, n) {
    if (!i.fix) return;
    var f = i.fix;
    if (f.type === "removeRecipients") f.options.forEach(function (o, k) { $("fx" + n + "-" + k).onclick = function () { removeEmails(o.emails).then(after); }; });
    if (f.type === "stripQuoted") $("fx" + n).onclick = function () {
      if (!armed.strip) { armed.strip = true; this.textContent = "한 번 더 누르면 삭제됩니다"; return; }
      armed.strip = false; stripQuoted().then(after);
    };
    if (f.type === "maskEmails") $("fx" + n).onclick = function () { maskEmails(f.emails).then(after); };
    if (f.type === "tagDomains") f.domains.forEach(function (d, k) {
      var sel = $("tg" + n + "-" + k), nm = $("tn" + n + "-" + k);
      sel.onchange = function () { nm.hidden = !(sel.value === "__new" || sel.value === "__fw" || sel.value === "__neutral"); nm.value = nm.hidden ? "" : guessName(d); };
      $("tb" + n + "-" + k).onclick = function () { tagDomain(d, sel.value, nm.value).then(after); };
    });
  }

  function after() { toast("반영했습니다"); runCheck(); }

  function guessName(d) { var host = d.indexOf("@") >= 0 ? d.split("@")[0] : d.split(".")[0]; return host.charAt(0).toUpperCase() + host.slice(1); }

  function tagDomain(d, choice, name) {
    if (!choice) { toast("역할을 선택하세요"); return Promise.reject(); }
    var isAddr = d.indexOf("@") >= 0;
    if (choice === "__internal") { if (!isAddr) rules.internalDomains.push(d); }
    else if (choice.indexOf("__") === 0) {
      var type = choice === "__fw" ? "forwarder" : choice === "__neutral" ? "neutral" : "counterparty";
      var nm = (name || guessName(d)).trim();
      rules.companies.push({ id: slug(nm), name: nm, type: type, domains: isAddr ? [] : [d], addresses: isAddr ? [d] : [], aliases: [] });
    } else {
      var c = rules.companies.filter(function (x) { return x.id === choice; })[0];
      if (c) { if (isAddr) (c.addresses = c.addresses || []).push(d); else (c.domains = c.domains || []).push(d); }
    }
    rules.seenUnknown = rules.seenUnknown.filter(function (x) { return x !== d; });
    return save();
  }

  function removeEmails(emails) {
    var drop = {}; emails.forEach(function (e) { drop[e.toLowerCase()] = 1; });
    return Promise.all(["to", "cc", "bcc"].map(function (k) {
      var f = item[k]; if (!f) return null;
      return call(function (cb) { f.getAsync(cb); }).then(function (list) {
        var keep = list.filter(function (r) { return !drop[String(r.emailAddress).toLowerCase()]; });
        if (keep.length === list.length) return null;
        return call(function (cb) { f.setAsync(keep.map(function (r) { return { displayName: r.displayName, emailAddress: r.emailAddress }; }), cb); });
      });
    }));
  }

  // Finds where Outlook/Gmail start the quoted original and cuts the HTML there.
  function stripQuoted() {
    return call(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); }).then(function (html) {
      var markers = ['id="divRplyFwdMsg"', "id='divRplyFwdMsg'", 'id="appendonsend"', 'class="gmail_quote"',
        '-----Original Message-----', '----- 원본 메시지 -----', '<b>From:</b>', '<b>보낸 사람:</b>', '>From:</span>', '>보낸 사람:</span>', '<hr'];
      var bodyStart = Math.max(0, html.toLowerCase().indexOf("<body"));
      var cut = -1;
      markers.forEach(function (m) { var i = html.indexOf(m, bodyStart); if (i > bodyStart && (cut < 0 || i < cut)) cut = i; });
      if (cut < 0) { toast("인용 부분을 찾지 못했습니다. 직접 지워 주세요."); return Promise.reject(); }
      if (html.charAt(cut) !== "<") {
        var tagStart = Math.max(html.lastIndexOf("<div", cut), html.lastIndexOf("<p", cut), html.lastIndexOf("<hr", cut));
        if (tagStart > bodyStart) cut = tagStart;
      }
      return call(function (cb) { item.body.setAsync(html.slice(0, cut) + "</body></html>", { coercionType: Office.CoercionType.Html }, cb); });
    });
  }

  function maskEmails(emails) {
    return call(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); }).then(function (html) {
      emails.forEach(function (e) { html = html.replace(new RegExp(e.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), "[삭제됨]"); });
      return call(function (cb) { item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb); });
    });
  }

  function canSend() { return isCompose() && typeof item.sendAsync === "function"; }

  function sendNow() {
    if (!canSend()) { toast("Outlook의 보내기 버튼을 누르세요"); return; }
    item.sendAsync(function (r) { if (r.status !== Office.AsyncResultStatus.Succeeded) toast("보내기 버튼을 직접 눌러 주세요"); });
  }

  function overrideSend(res, sig) {
    var reason = ($("ov-reason").value || "").trim();
    if (reason.length < 4) { toast("사유를 짧게라도 입력해 주세요"); return; }
    log.push({ t: new Date().toISOString(), reason: reason.slice(0, 200), issues: res.issues.map(function (i) { return i.code + " " + i.title; }).slice(0, 4) });
    log = log.slice(-50);
    save().then(function () {
      return item.sessionData ? call(function (cb) { item.sessionData.setAsync(OVERRIDE_KEY, sig, cb); }) : null;
    }).then(function () {
      if (canSend()) sendNow(); else toast("기록했습니다. 보내기 버튼을 다시 누르세요");
    }).catch(function () { toast("저장에 실패했습니다"); });
  }

  // ---------- companies ----------
  function renderCompanies() {
    var q = ($("cp-filter").value || "").toLowerCase();
    var list = rules.companies.filter(function (c) { return !q || JSON.stringify(c).toLowerCase().indexOf(q) >= 0; })
      .sort(function (a, b) { return a.name.localeCompare(b.name); });
    $("cp-count").textContent = rules.companies.length + "곳";
    $("cp-list").innerHTML = list.map(function (c) {
      return '<div class="item"><div><div class="title">' + esc(c.name) + ' <span class="muted small">' + TYPES[c.type || "counterparty"] + '</span></div>' +
        '<div class="detail small">' + esc((c.domains || []).concat(c.addresses || []).join(", ")) + (c.aliases && c.aliases.length ? " · 별칭: " + esc(c.aliases.join(", ")) : "") + '</div></div>' +
        '<button class="btn" data-edit="' + esc(c.id) + '">편집</button></div>';
    }).join("") || '<p class="muted">등록된 거래처가 없습니다.</p>';
    Array.prototype.forEach.call(document.querySelectorAll("[data-edit]"), function (b) { b.onclick = function () { openForm(b.getAttribute("data-edit")); }; });
    renderGroups();
  }

  function openForm(id) {
    var c = rules.companies.filter(function (x) { return x.id === id; })[0] || { id: "", name: "", type: "counterparty", domains: [], addresses: [], aliases: [] };
    var f = $("cp-form"); f.hidden = false;
    f.innerHTML = '<label>회사명</label><input id="f-name" value="' + esc(c.name) + '">' +
      '<label>구분</label><select id="f-type">' + Object.keys(TYPES).map(function (k) { return '<option value="' + k + '"' + (c.type === k ? " selected" : "") + '>' + TYPES[k] + '</option>'; }).join("") + '</select>' +
      '<label>도메인 (쉼표 구분)</label><input id="f-dom" value="' + esc((c.domains || []).join(", ")) + '">' +
      '<label>개별 주소 (gmail 등 개인 메일)</label><input id="f-addr" value="' + esc((c.addresses || []).join(", ")) + '">' +
      '<label>별칭·약칭·한글명 (본문 언급 검사용)</label><input id="f-alias" value="' + esc((c.aliases || []).join(", ")) + '">' +
      '<div class="actions"><button class="btn primary" id="f-save">저장</button><button class="btn" id="f-cancel">취소</button>' + (c.id ? '<button class="btn danger" id="f-del">삭제</button>' : "") + '</div>';
    $("f-cancel").onclick = function () { f.hidden = true; };
    $("f-save").onclick = function () {
      var name = $("f-name").value.trim(); if (!name) { toast("회사명을 입력하세요"); return; }
      var rec = { id: c.id || slug(name), name: name, type: $("f-type").value, domains: splitList($("f-dom").value), addresses: splitList($("f-addr").value),
        aliases: $("f-alias").value.split(",").map(function (s) { return s.trim(); }).filter(Boolean) };
      rules.companies = rules.companies.filter(function (x) { return x.id !== rec.id; }).concat([rec]);
      save().then(function () { f.hidden = true; renderCompanies(); toast("저장했습니다"); });
    };
    if ($("f-del")) $("f-del").onclick = function () {
      if (!armed.del) { armed.del = true; this.textContent = "한 번 더 누르면 삭제"; return; }
      armed.del = false;
      rules.companies = rules.companies.filter(function (x) { return x.id !== c.id; });
      rules.allowedGroups = rules.allowedGroups.map(function (g) { return g.filter(function (i) { return i !== c.id; }); }).filter(function (g) { return g.length > 1; });
      save().then(function () { f.hidden = true; renderCompanies(); });
    };
  }

  function nameOf(id) { var c = rules.companies.filter(function (x) { return x.id === id; })[0]; return c ? c.name : id; }

  function renderGroups() {
    $("grp-list").innerHTML = rules.allowedGroups.map(function (g, i) {
      return '<div class="item"><div>' + esc(g.map(nameOf).join(" + ")) + '</div><button class="btn danger" data-grp="' + i + '">삭제</button></div>';
    }).join("") || '<p class="muted small">없음</p>';
    Array.prototype.forEach.call(document.querySelectorAll("[data-grp]"), function (b) {
      b.onclick = function () { rules.allowedGroups.splice(+b.getAttribute("data-grp"), 1); save().then(renderGroups); };
    });
  }

  function openGroupForm() {
    var f = $("grp-form"); f.hidden = false;
    f.innerHTML = '<p class="small muted">함께 받아도 되는 회사를 2곳 이상 고르세요.</p>' +
      rules.companies.map(function (c) { return '<label style="display:flex;gap:6px;align-items:center;color:var(--ink)"><input type="checkbox" value="' + esc(c.id) + '" style="width:auto">' + esc(c.name) + ' <span class="muted small">' + TYPES[c.type || "counterparty"] + '</span></label>'; }).join("") +
      '<div class="actions"><button class="btn primary" id="g-save">저장</button><button class="btn" id="g-cancel">취소</button></div>';
    $("g-cancel").onclick = function () { f.hidden = true; };
    $("g-save").onclick = function () {
      var ids = Array.prototype.map.call(f.querySelectorAll("input:checked"), function (x) { return x.value; });
      if (ids.length < 2) { toast("2곳 이상 선택하세요"); return; }
      rules.allowedGroups.push(ids); save().then(function () { f.hidden = true; renderGroups(); });
    };
  }

  // ---------- settings ----------
  function renderSettings() {
    $("set-internal").value = rules.internalDomains.length ? rules.internalDomains.join(", ") : "repitone.com";
    $("set-usage").textContent = "사용 " + Math.round(size() / 1024 * 10) / 10 + "KB / 32KB · 거래처 " + rules.companies.length + "곳 · 처음 본 주소 기록 " + rules.seenUnknown.length + "건";
    $("set-log").innerHTML = log.slice().reverse().map(function (l) {
      return '<div class="item"><div><div>' + esc(l.reason) + '</div><div class="muted">' + esc(l.t.slice(0, 16).replace("T", " ")) + " · " + esc((l.issues || []).join(" / ")) + '</div></div></div>';
    }).join("") || '<p class="muted">기록 없음</p>';
  }

  function normalize(c) {
    return { id: c.id || slug(c.name || "c"), name: String(c.name || c.id || "").trim(), type: TYPES[c.type] ? c.type : "counterparty",
      domains: splitList((c.domains || []).join(",")), addresses: splitList((c.addresses || []).join(",")),
      aliases: (c.aliases || []).map(function (s) { return String(s).trim(); }).filter(Boolean) };
  }

  function importJson(replace) {
    var data;
    try { data = JSON.parse($("set-json").value); } catch (e) { toast("JSON 형식이 올바르지 않습니다"); return; }
    if (Array.isArray(data)) data = { companies: data };
    var incoming = (data.companies || []).map(normalize).filter(function (c) { return c.name; });
    if (replace) {
      rules.companies = incoming;
      rules.allowedGroups = data.allowedGroups || [];
      if (data.internalDomains) rules.internalDomains = splitList(data.internalDomains.join(","));
    } else {
      incoming.forEach(function (c) {
        var ex = rules.companies.filter(function (x) { return x.id === c.id || x.name.toLowerCase() === c.name.toLowerCase(); })[0];
        if (!ex) { rules.companies.push(c); return; }
        ["domains", "addresses", "aliases"].forEach(function (k) { ex[k] = Array.from(new Set((ex[k] || []).concat(c[k]))); });
        ex.type = c.type || ex.type;
      });
      (data.allowedGroups || []).forEach(function (g) { rules.allowedGroups.push(g); });
      if (data.internalDomains) rules.internalDomains = Array.from(new Set(rules.internalDomains.concat(splitList(data.internalDomains.join(",")))));
    }
    save().then(function () { $("set-json").value = ""; renderSettings(); renderCompanies(); toast("가져왔습니다: " + incoming.length + "곳"); });
  }

  // ---------- boot ----------
  function showTab(name) {
    Array.prototype.forEach.call(document.querySelectorAll("header button"), function (b) { b.classList.toggle("on", b.getAttribute("data-tab") === name); });
    ["check", "companies", "settings"].forEach(function (t) { $("tab-" + t).hidden = t !== name; });
    if (name === "check") runCheck();
    if (name === "companies") renderCompanies();
    if (name === "settings") renderSettings();
  }

  Office.onReady(function () {
    item = Office.context.mailbox.item;
    load();
    Array.prototype.forEach.call(document.querySelectorAll("header button"), function (b) { b.onclick = function () { showTab(b.getAttribute("data-tab")); }; });
    $("cp-filter").oninput = renderCompanies;
    $("cp-new").onclick = function () { openForm(""); };
    $("grp-new").onclick = openGroupForm;
    $("set-internal-save").onclick = function () { rules.internalDomains = splitList($("set-internal").value); save().then(function () { toast("저장했습니다"); renderSettings(); }); };
    $("set-import-merge").onclick = function () { importJson(false); };
    $("set-import-replace").onclick = function () { importJson(true); };
    $("set-export").onclick = function () { $("set-json").value = JSON.stringify({ internalDomains: rules.internalDomains, companies: rules.companies, allowedGroups: rules.allowedGroups }, null, 1); };
    $("set-reset-seen").onclick = function () { rules.seenUnknown = []; save().then(renderSettings); };

    if (!rules.internalDomains.length) showTab("settings");
    else showTab(isCompose() ? "check" : "companies");

    if (item && item.addHandlerAsync && Office.EventType.RecipientsChanged) {
      try { item.addHandlerAsync(Office.EventType.RecipientsChanged, function () { if (!$("tab-check").hidden) runCheck(); }); } catch (e) { /* optional */ }
    }
  });
})();
