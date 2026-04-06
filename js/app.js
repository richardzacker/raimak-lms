// Raimak LMS - App Logic v3.0

const State = {
  leads:           [],
  contractors:     [],
  activityLog:     [],
  todaySales:      [],
  currentView:     "dashboard",
  filters:         { status: "all", search: "", assignedTo: "all" },
  editingLeadId:   null,
  loading:         false,
  role:            "agent",
  currentUser:     null,
  salesFeedTimer:  null,
  dripLead:        null,
  selectedLeads:   new Set(),
};

function isAdmin() { return State.role === "admin"; }

function detectRole(user) {
  if (!user) return "agent";
  const email  = (user.email || "").toLowerCase();
  const admins = (Config.roles.admins || []).map(function(a) { return a.toLowerCase(); });
  return admins.includes(email) ? "admin" : "agent";
}

// ── Boot ──────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", async function() {
  try {
    const redirectResult = await Auth.init();
    if (!Auth.isSignedIn()) { showLoginScreen(); return; }
    if (redirectResult) window.history.replaceState({}, document.title, window.location.pathname);
    State.currentUser = Auth.getUser();
    State.role        = detectRole(State.currentUser);
    showAppShell();
    await loadAllData();
    renderDashboard();
  } catch (err) {
    console.error("Boot error:", err);
    showLoginScreen();
  }
});

async function loadAllData() {
  setLoading(true);
  try {
    const [rawLeads, contractors, activityLog, todaySales] = await Promise.all([
      Graph.getLeads(),
      Graph.getContractors(),
      Graph.getActivityLog(),
      Graph.getTodaySales(),
    ]);
    State.contractors = contractors;
    State.leads       = Graph.applyBusinessRules(rawLeads, contractors);
    State.activityLog = activityLog;
    State.todaySales  = todaySales;
  } catch (err) {
    UI.showToast("Failed to load data: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  LOGIN
// ============================================================
function showLoginScreen() {
  document.getElementById("app").innerHTML = `
    <div class="login-screen">
      <div class="login-card">
        <div class="login-logo"><img src="Raimak.png" alt="Raimak"></div>
        <h1>Lead Management</h1>
        <p>Sign in with your Raimak Microsoft account to access the system.</p>
        <button class="btn-primary btn-lg" onclick="Auth.signIn()">
          <svg width="20" height="20" viewBox="0 0 21 21" fill="none" style="margin-right:8px">
            <path d="M10 1H1v9h9V1zM20 1h-9v9h9V1zM10 11H1v9h9v-9zM20 11h-9v9h9v-9z" fill="currentColor"/>
          </svg>
          Sign in with Microsoft
        </button>
        <p class="login-version">v${Config.rules.appVersion} · Raimak Leadship</p>
      </div>
    </div>
  `;
}

// ============================================================
//  APP SHELL
// ============================================================
function showAppShell() {
  const user = State.currentUser;

  const adminNav = isAdmin() ? `
    <a class="nav-item" data-view="drip" onclick="navigate('drip')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 14H9V8h2v8zm4 0h-2V8h2v8z" fill="currentColor"/></svg>
      Drip Feed
    </a>
    <a class="nav-item" data-view="assign" onclick="navigate('assign')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="9" cy="7" r="4" stroke="currentColor" stroke-width="2"/><line x1="19" y1="8" x2="19" y2="14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="22" y1="11" x2="16" y2="11" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      Assign Leads
    </a>
    <a class="nav-item" data-view="report" onclick="navigate('report')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" stroke-width="2"/><polyline points="14,2 14,8 20,8" stroke="currentColor" stroke-width="2"/><line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      Daily Report
    </a>
    <a class="nav-item" data-view="leads" onclick="navigate('leads')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="9" cy="7" r="4" stroke="currentColor" stroke-width="2"/></svg>
      Leads
      <span class="badge" id="badge-leads"></span>
    </a>
    <a class="nav-item" data-view="contractors" onclick="navigate('contractors')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><rect x="2" y="7" width="20" height="14" rx="2" stroke="currentColor" stroke-width="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2" stroke="currentColor" stroke-width="2"/></svg>
      Raimak Team
    </a>
    <a class="nav-item" data-view="activity" onclick="navigate('activity')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><polyline points="22,12 18,12 15,21 9,3 6,12 2,12" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
      Activity Log
    </a>` : "";

  const agentNav = `
    <a class="nav-item" data-view="myleads" onclick="navigate('myleads')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><polyline points="12,6 12,12 16,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      My Leads
    </a>`;

  document.getElementById("app").innerHTML = `
    <div class="app-shell">
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-brand"><img src="Raimak.png" alt="Raimak"></div>
        <nav class="sidebar-nav">
          <a class="nav-item active" data-view="dashboard" onclick="navigate('dashboard')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1" fill="currentColor"/><rect x="14" y="3" width="7" height="7" rx="1" fill="currentColor"/><rect x="3" y="14" width="7" height="7" rx="1" fill="currentColor"/><rect x="14" y="14" width="7" height="7" rx="1" fill="currentColor"/></svg>
            Dashboard
          </a>
          ${agentNav}
          ${adminNav}
        </nav>
        <div class="sidebar-footer">
          <div class="user-info">
            <div class="user-avatar">${((user && user.name) || "U")[0].toUpperCase()}</div>
            <div class="user-meta">
              <span class="user-name">${(user && user.name) || "User"}</span>
              <span class="user-email">${(user && user.email) || ""}</span>
            </div>
          </div>
          <button class="btn-ghost" onclick="Auth.signOut()" title="Sign Out">
            <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="16,17 21,12 16,7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="21" y1="12" x2="9" y2="12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          </button>
        </div>
      </aside>
      <main class="main-content" id="main-content">
        <div class="loading-overlay" id="loading-overlay" style="display:none"><div class="spinner"></div></div>
      </main>
    </div>
    <div id="toast-container"></div>
    <div id="modal-overlay" class="modal-overlay" style="display:none" onclick="closeModal(event)">
      <div class="modal" id="modal"></div>
    </div>
  `;
}

// ============================================================
//  NAVIGATION
// ============================================================
function navigate(view) {
  const adminOnly = ["leads", "drip", "assign", "report", "contractors", "activity"];
  if (!isAdmin() && adminOnly.includes(view)) {
    view = "myleads";
  }
  State.currentView = view;
  document.querySelectorAll(".nav-item").forEach(function(el) { el.classList.remove("active"); });
  const navEl = document.querySelector("[data-view='" + view + "']");
  if (navEl) navEl.classList.add("active");
  switch (view) {
    case "dashboard":   renderDashboard();   break;
    case "leads":       renderLeads();       break;
    case "myleads":     renderMyLeads();     break;
    case "drip":        renderDripFeed();    break;
    case "assign":      renderAssignLeads(); break;
    case "contractors": renderContractors(); break;
    case "activity":    renderActivity();    break;
    case "report":      renderDailyReport(); break;
  }
}

// ============================================================
//  DASHBOARD
// ============================================================
function renderDashboard() {
  const leads      = State.leads;
  const todaySales = State.todaySales;
  const total      = leads.length;
  const active     = leads.filter(function(l) { return !Config.terminalStatuses.includes(l.status); }).length;
  const sold       = leads.filter(function(l) { return l.status === "Sold"; }).length;
  const needRecycle= leads.filter(function(l) { return l.flags && l.flags.includes("needs_recycle"); }).length;
  const coolOff    = leads.filter(function(l) { return l.flags && l.flags.includes("cool_off"); }).length;

  const statusCounts = {};
  Config.leadStatuses.forEach(function(s) {
    statusCounts[s] = leads.filter(function(l) { return l.status === s; }).length;
  });

  const agentSales = {};
  todaySales.forEach(function(l) {
    if (l.assignedTo) agentSales[l.assignedTo] = (agentSales[l.assignedTo] || 0) + 1;
  });
  const top5 = Object.entries(agentSales).sort(function(a,b) { return b[1]-a[1]; }).slice(0,5);
  const recentLeads = leads.slice().sort(function(a,b) { return new Date(b.createdAt)-new Date(a.createdAt); }).slice(0,8);

  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Dashboard</h1>
        <span class="view-subtitle">${isAdmin() ? "// ADMIN VIEW" : "// AGENT VIEW"} · v${Config.rules.appVersion}</span>
      </div>
      ${isAdmin() ? `<button class="btn-primary" onclick="openAddLeadModal()">
        <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="5" y1="12" x2="19" y2="12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg>
        Add Lead
      </button>` : ""}
    </div>

    <div class="kpi-grid">
      <div class="kpi-card kpi-primary">
        <span class="kpi-label">Total Leads</span>
        <span class="kpi-value">${total}</span>
        <span class="kpi-sub">${active} active in pipeline</span>
      </div>
      <div class="kpi-card kpi-success">
        <span class="kpi-label">Sold Today</span>
        <span class="kpi-value">${todaySales.length}</span>
        <span class="kpi-sub">${total ? Math.round((sold/total)*100) : 0}% all-time close rate</span>
      </div>
      ${isAdmin() ? `
      <div class="kpi-card ${needRecycle > 0 ? "kpi-warn" : "kpi-neutral"}" ${needRecycle > 0 ? 'style="cursor:pointer" onclick="document.querySelector(\'[data-recycle-queue]\')?.scrollIntoView({behavior:\'smooth\'})"' : ""}>
        <span class="kpi-label">Needs Recycle</span>
        <span class="kpi-value">${needRecycle}</span>
        <span class="kpi-sub">${needRecycle > 0 ? "↓ See recycle queue below" : "All leads current"}</span>
      </div>` : ""}
      <div class="kpi-card ${coolOff > 0 ? "kpi-info" : "kpi-neutral"}">
        <span class="kpi-label">In Cool-Off</span>
        <span class="kpi-value">${coolOff}</span>
        <span class="kpi-sub">${Config.rules.coolOffDays}-day rule active</span>
      </div>
    </div>

    <div class="two-col">
      <div class="card">
        <div class="card-header"><h2 class="card-title">Pipeline Status</h2></div>
        <div class="status-breakdown">
          ${Config.leadStatuses.map(function(s) {
            const cls = "status-" + s.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"");
            return `
              <div class="status-row">
                <span class="status-badge ${cls}">${s}</span>
                <div class="status-bar-wrap"><div class="status-bar" style="width:${total?(statusCounts[s]/total)*100:0}%"></div></div>
                <span class="status-count">${statusCounts[s]}</span>
              </div>`;
          }).join("")}
        </div>
      </div>

      <div class="card">
        <div class="card-header">
          <h2 class="card-title"><span class="live-dot"></span>Live Sales Feed</h2>
          <span class="card-meta" id="sales-feed-time">Today</span>
        </div>
        <div class="sales-feed" id="sales-feed">
          ${todaySales.length ? todaySales.slice(0,6).map(function(l) { return `
            <div class="sale-entry">
              <div class="sale-icon">&#127881;</div>
              <div class="sale-info">
                <span class="sale-name">${escHtml(l.name)}</span>
                <span class="sale-agent">${escHtml(l.assignedTo || "Unassigned")}</span>
              </div>
              <span class="sale-time">${formatTime(l.modified)}</span>
            </div>`;
          }).join("") : `<p class="empty-state" style="padding:24px">No sales yet today.</p>`}
        </div>
        <div class="card-header" style="border-top:1px solid var(--border);border-bottom:none;margin-top:2px">
          <h2 class="card-title">Top 5 Today</h2>
        </div>
        <div class="top5-list">
          ${top5.length ? top5.map(function(e,i) { return `
            <div class="top5-row">
              <span class="top5-rank rank-${i+1}">${i+1}</span>
              <span class="top5-name">${escHtml(e[0])}</span>
              <span class="top5-count">${e[1]} sale${e[1]!==1?"s":""}</span>
            </div>`;
          }).join("") : `<p class="empty-state" style="padding:16px 20px;font-size:12px">No sales yet today.</p>`}
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-header">
        <h2 class="card-title">Recent Leads</h2>
        ${isAdmin() ? `<button class="btn-ghost-sm" onclick="navigate('leads')">View all</button>` : ""}
      </div>
      ${renderLeadsTable(recentLeads, true)}
    </div>

    ${isAdmin() && needRecycle > 0 ? `
    <div class="card" data-recycle-queue style="border-color:#FFB300;box-shadow:0 0 20px rgba(255,179,0,0.1)">
      <div class="card-header" style="background:#FFF8E1">
        <h2 class="card-title" style="color:#8B6914">
          ⚠️ Recycle Queue — ${needRecycle} lead${needRecycle!==1?"s":""} ready
        </h2>
        <div style="display:flex;gap:8px;align-items:center">
          <span class="card-meta" style="color:#8B6914">3rd Contact · 48hrs+ since last contact</span>
          <button class="btn-primary" style="padding:6px 16px;font-size:12px;background:#FFB300;border-color:#FFB300;color:#1A2640" onclick="recycleAllLeads()">
            ♻️ Recycle All
          </button>
        </div>
      </div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr>
            <th>Name</th><th>Address</th><th>Previously Assigned To</th><th>Last Contacted</th><th>Action</th>
          </tr></thead>
          <tbody>
            ${leads.filter(function(l){return l.flags&&l.flags.includes("needs_recycle");}).map(function(l){
              return `<tr>
                <td><span class="lead-name">${escHtml(l.name)}</span></td>
                <td class="td-mono" style="font-size:11px">${escHtml(l.address||"—")}${l.city?", "+escHtml(l.city):""}</td>
                <td>
                  <div style="display:flex;flex-direction:column;gap:2px">
                    ${l.assignedTo ? `<span style="font-size:13px;font-weight:600;color:#1A2640">${escHtml(l.assignedTo)}</span>` : "—"}
                    ${l.previousAgents ? `<span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8">Previously: ${escHtml(l.previousAgents)}</span>` : ""}
                  </div>
                </td>
                <td class="td-mono">${formatDate(l.lastContacted)||"—"}</td>
                <td>
                  <button class="btn-primary" style="padding:6px 14px;font-size:12px" onclick="recycleLeadAction('${l.id}','${escHtml(l.assignedTo||"")}','${escHtml(l.name)}')">
                    Recycle
                  </button>
                </td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>
    </div>` : ""}
  `;
  updateBadges();
  startSalesFeedPolling();
}

async function recycleLeadAction(leadId, currentAgent, leadName) {
  if (!confirm("Recycle \"" + leadName + "\"?\n\nThis will:\n• Reset status to New\n• Unassign from " + (currentAgent||"current agent") + "\n• Record previous assignment history\n\nThe lead can then be reassigned to a different agent.")) return;
  setLoading(true);
  try {
    await Graph.recycleLead(leadId, currentAgent);
    await Graph.logActivity({
      LeadID:     leadId,
      Title:      leadName,
      ActionType: "Recycled",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
      Notes:      "Recycled by admin — previous agent: " + (currentAgent || "unknown"),
    });
    UI.showToast(leadName + " recycled and ready to reassign!", "success");
    await loadAllData();
    renderDashboard();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

async function recycleAllLeads() {
  const recycleLeads = State.leads.filter(function(l) { return l.flags && l.flags.includes("needs_recycle"); });
  if (!recycleLeads.length) { UI.showToast("No leads to recycle.", "info"); return; }
  if (!confirm("Recycle ALL " + recycleLeads.length + " lead" + (recycleLeads.length !== 1 ? "s" : "") + " in the queue?\n\nThis will:\n• Reset all their statuses to New\n• Unassign them from their current agents\n• Record previous assignment history\n\nThis cannot be undone.")) return;
  setLoading(true);
  try {
    for (var i = 0; i < recycleLeads.length; i++) {
      const lead = recycleLeads[i];
      await Graph.recycleLead(lead.id, lead.assignedTo || "");
      await Graph.logActivity({
        LeadID:     lead.id,
        Title:      lead.name,
        ActionType: "Recycled",
        AgentEmail: (State.currentUser && State.currentUser.email) || "",
        Notes:      "Bulk recycled by admin — previous agent: " + (lead.assignedTo || "unknown"),
      });
    }
    UI.showToast("Recycled " + recycleLeads.length + " lead" + (recycleLeads.length !== 1 ? "s" : "") + " successfully!", "success");
    await loadAllData();
    renderDashboard();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

function startSalesFeedPolling() {
  if (State.salesFeedTimer) clearInterval(State.salesFeedTimer);

  // Track by IDs not count — prevents false positives on navigation
  let knownSaleIds = new Set((State.todaySales || []).map(function(l) { return l.id; }));

  State.salesFeedTimer = setInterval(async function() {
    try {
      const newSales = await Graph.getTodaySales();

      // Only trigger banner for genuinely new sales not seen before
      const newOnes = newSales.filter(function(l) { return !knownSaleIds.has(l.id); });
      if (newOnes.length) {
        const latest   = newOnes[newOnes.length - 1];
        const soldBy   = latest && latest.soldBy;
        const assignee = latest && latest.assignedTo;
        UI.showSaleBanner(
          (latest && latest.name) || "a customer",
          soldBy || assignee || "An agent",
          soldBy && assignee && soldBy !== assignee ? assignee : null
        );
        UI.showConfetti();
        newOnes.forEach(function(l) { knownSaleIds.add(l.id); });
      }

      State.todaySales = newSales;

      if (State.currentView === "dashboard") {
        const feed = document.getElementById("sales-feed");
        const time = document.getElementById("sales-feed-time");
        if (!feed) return;
        if (time) time.textContent = "Updated " + formatTime(new Date().toISOString());
        if (!newSales.length) return;
        feed.innerHTML = newSales.slice(0,6).map(function(l) { return `
          <div class="sale-entry">
            <div class="sale-icon">&#127881;</div>
            <div class="sale-info">
              <span class="sale-name">${escHtml(l.name)}</span>
              <span class="sale-agent">${escHtml(l.assignedTo||"Unassigned")}</span>
            </div>
            <span class="sale-time">${formatTime(l.modified)}</span>
          </div>`;
        }).join("");
      }
    } catch(e) { /* silent */ }
  }, Config.salesFeedInterval);
}

// ============================================================
//  ADMIN — DRIP FEED
// ============================================================
function renderDripFeed() {
  if (!isAdmin()) { navigate("myleads"); return; }

  const unassigned = State.leads.filter(function(l) {
    return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
  });

  if (!State.dripLead && unassigned.length) {
    State.dripLead = unassigned[0];
  }

  const lead      = State.dripLead;
  const remaining = unassigned.length;

  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Drip Feed</h1>
        <span class="view-subtitle">// ASSIGN ONE LEAD AT A TIME · ${remaining} unassigned</span>
      </div>
      <button class="btn-ghost" onclick="skipDripLead()">
        Skip This Lead
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><polyline points="9,18 15,12 9,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      </button>
    </div>

    ${!lead ? `
      <div class="drip-card">
        <div style="text-align:center;padding:40px 0">
          <div style="font-size:52px;margin-bottom:16px">&#10003;</div>
          <h3 style="font-family:var(--font-head);font-size:24px;text-transform:uppercase;letter-spacing:1px;color:var(--green)">All Leads Assigned!</h3>
          <p style="color:var(--text-2);margin-top:8px">No unassigned leads remaining in the pipeline.</p>
        </div>
      </div>
    ` : `
      <div class="drip-card">
        <div class="drip-header">
          <span class="drip-title">Next Lead</span>
          <div style="display:flex;gap:8px;align-items:center">
            ${lead.leadType ? `<span class="lead-type-badge lead-type-${(lead.leadType||"").toLowerCase()}">${escHtml(lead.leadType)}</span>` : ""}
            <span class="status-badge status-${lead.status.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}">${lead.status}</span>
          </div>
        </div>
        <div class="drip-lead-name">${escHtml(lead.name)}</div>
        <div class="drip-meta">
          ${lead.phone ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 12 19.79 19.79 0 0 1 1.61 3.38 2 2 0 0 1 3.6 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L7.91 8.96a16 16 0 0 0 6 6l.92-.92a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 21.73 16.92z" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.phone)}</span>` : ""}
          ${lead.email ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" stroke="currentColor" stroke-width="2"/><polyline points="22,6 12,13 2,6" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.email)}</span>` : ""}
          ${lead.currentMRC ? `<span class="feed-meta">MRC: $${escHtml(lead.currentMRC)}/mo</span>` : ""}
          ${lead.currentProducts ? `<span class="feed-meta">Has: ${escHtml(lead.currentProducts)}</span>` : ""}
        </div>
        ${lead.notes ? `<div class="feed-notes">${escHtml(lead.notes)}</div>` : ""}
        <div style="margin-top:8px">
          <span class="feed-label">Assign To Agent</span>
          <div style="display:flex;gap:10px;align-items:center;margin-top:8px;flex-wrap:wrap">
            <select id="drip-agent-select" class="filter-select" style="min-width:220px">
              <option value="">Select an agent...</option>
              ${State.contractors.map(function(c) {
                const count = State.leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
                const full  = count >= Config.rules.maxLeadsPerAgent;
                return `<option value="${escHtml(c.name)}" ${full?"disabled":""}>${escHtml(c.name)} — ${count}/${Config.rules.maxLeadsPerAgent}${full?" (FULL)":""}</option>`;
              }).join("")}
            </select>
            <button class="btn-primary" onclick="confirmDripAssign('${lead.id}')">
              <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><polyline points="20,6 9,17 4,12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
              Assign &amp; Next
            </button>
            <button class="btn-ghost" onclick="skipDripLead()">Skip</button>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-header"><h2 class="card-title">Remaining Unassigned (${remaining})</h2></div>
        ${renderLeadsTable(unassigned.slice(0,10), true)}
      </div>
    `}
  `;
}

async function confirmDripAssign(leadId) {
  const select = document.getElementById("drip-agent-select");
  const agent  = select && select.value;
  if (!agent) { UI.showToast("Please select an agent first.", "error"); return; }
  if (!Graph.canAgentTakeLead(agent, State.leads)) {
    UI.showToast(agent + " is at the " + Config.rules.maxLeadsPerAgent + "-lead limit.", "error");
    return;
  }
  const lead = State.leads.find(function(l) { return l.id === leadId; });
  setLoading(true);
  try {
    await Graph.assignAgent(leadId, agent);
    await Graph.logActivity({
      LeadID:     leadId,
      Title:      lead ? lead.name : "",
      ActionType: "Drip Assigned",
      AgentEmail: agent,
      Notes:      "Drip-assigned by " + ((State.currentUser && State.currentUser.name) || "Admin"),
    });
    UI.showToast(lead.name + " assigned to " + agent, "success");
    await loadAllData();
    const remaining = State.leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); });
    State.dripLead = remaining.length ? remaining[0] : null;
    renderDripFeed();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

function skipDripLead() {
  const unassigned = State.leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); });
  const currentIdx = unassigned.findIndex(function(l) { return State.dripLead && l.id === State.dripLead.id; });
  const nextIdx    = (currentIdx + 1) % unassigned.length;
  State.dripLead   = unassigned[nextIdx] || null;
  renderDripFeed();
}

// ============================================================
//  AGENT — MY LEADS
// ============================================================
function getStatusColor(status) {
  const colors = Config.statusColors || {};
  if ((colors.red    || []).includes(status)) return "#FF4444";
  if ((colors.yellow || []).includes(status)) return "#FFD700";
  if ((colors.green  || []).includes(status)) return "#00FF88";
  if ((colors.blue   || []).includes(status)) return "#4D79FF";
  if ((colors.cyan   || []).includes(status)) return "#00E5FF";
  if ((colors.white  || []).includes(status)) return "#FFFFFF";
  return "#7A98C8";
}

function getStatusDot(status) {
  const color = getStatusColor(status);
  return `<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};box-shadow:0 0 6px ${color};flex-shrink:0;margin-right:6px"></span>`;
}

let _leadSaved = false;
let _currentFeedIndex = 0;

function renderMyLeads() {
  const user      = State.currentUser;
  const userName  = ((user && user.name)  || "").toLowerCase().trim();
  const userEmail = ((user && user.email) || "").toLowerCase().trim();

  const contractor = State.contractors.find(function(c) {
    return (c.email || "").toLowerCase().trim() === userEmail ||
           (c.name  || "").toLowerCase().trim() === userName;
  });
  const agentName = contractor ? contractor.name.toLowerCase().trim() : userName;

  const myLeads = State.leads.filter(function(l) {
    const assigned = (l.assignedTo || "").toLowerCase().replace(/\s+/g, " ").trim();
    return assigned && (
      assigned === agentName.replace(/\s+/g, " ") ||
      assigned === userName.replace(/\s+/g, " ")  ||
      assigned === userEmail.replace(/\s+/g, " ")
    ) && !Config.terminalStatuses.includes(l.status);
  });

  const allMyLeads = State.leads.filter(function(l) {
    const assigned = (l.assignedTo || "").toLowerCase().replace(/\s+/g, " ").trim();
    return assigned && (
      assigned === agentName.replace(/\s+/g, " ") ||
      assigned === userName.replace(/\s+/g, " ")  ||
      assigned === userEmail.replace(/\s+/g, " ")
    ) && !Config.terminalStatuses.includes(l.status);
  });
  window._myLeads   = allMyLeads;
  window._agentName = agentName;
  _leadSaved        = false;

  if (_currentFeedIndex >= myLeads.length) _currentFeedIndex = 0;

  const contactsToday = Graph.agentContactsToday((user && user.email) || "", State.activityLog);
  const atLimit       = contactsToday >= Config.rules.maxContactsPerDay;

  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">My Leads</h1>
        <span class="view-subtitle">// ${myLeads.length} remaining · lead ${Math.min(_currentFeedIndex+1, myLeads.length)} of ${myLeads.length}</span>
      </div>
      <div class="contacts-today-badge ${atLimit ? "badge-full" : ""}">
        <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><polyline points="12,6 12,12 16,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        ${contactsToday}/${Config.rules.maxContactsPerDay} contacts today
      </div>
    </div>

    ${atLimit ? `<div class="alert alert-info">
      <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><line x1="12" y1="8" x2="12" y2="12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="16" x2="12.01" y2="16" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      Daily limit reached — ${Config.rules.maxContactsPerDay} contacts today. Great work!
    </div>` : ""}

    <div id="lead-feed-wrap">${renderLeadFeedCard(myLeads, contactsToday)}</div>

    <div id="lead-search-section" style="display:${isAdmin() ? 'block' : 'none'};margin-top:20px">
      <div class="card">
        <div class="card-header">
          <h2 class="card-title">Find a Lead</h2>
          <span class="card-meta">Search by name, phone or address</span>
        </div>
        <div style="padding:16px 20px">
          <div class="search-wrap">
            <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8" stroke="currentColor" stroke-width="2"/><line x1="21" y1="21" x2="16.65" y2="16.65" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
            <input type="text" class="search-input" placeholder="Search name, phone, address..." oninput="searchMyLeads(this.value)" id="my-leads-search">
          </div>
        </div>
        <div id="my-leads-table"></div>
      </div>
    </div>
  `;
}

function searchMyLeads(q) {
  const leads    = window._myLeads || [];
  const filtered = !q.trim() ? [] : leads.filter(function(l) {
    return l.name.toLowerCase().includes(q.toLowerCase()) ||
           (l.phone   || "").includes(q) ||
           (l.btn     || "").includes(q) ||
           (l.cbr     || "").includes(q) ||
           (l.address || "").toLowerCase().includes(q.toLowerCase());
  });
  const wrap = document.getElementById("my-leads-table");
  if (!wrap) return;
  if (!q.trim()) { wrap.innerHTML = ""; return; }
  wrap.innerHTML = filtered.length ? renderLeadsTable(filtered, false, true) : `<div class="empty-state">No leads found for "${escHtml(q)}"</div>`;
  if (filtered.length === 1) {
    const feedWrap = document.getElementById("lead-feed-wrap");
    if (feedWrap) {
      _leadSaved = false;
      feedWrap.innerHTML = renderLeadFeedCard(filtered, 0, true);
    }
  }
}

let _stagedStatus = null;

function renderLeadFeedCard(myLeads, contactsToday, forceFirst) {
  const lead    = forceFirst ? myLeads[0] : myLeads.find(function(l) { return !Graph.isInCoolOff(l); });
  const atLimit = contactsToday >= Config.rules.maxContactsPerDay;
  _stagedStatus = null;

  if (!lead) return `
    <div class="feed-card feed-card-empty">
      <div class="feed-empty-icon">&#10003;</div>
      <h3>All Caught Up!</h3>
      <p>${myLeads.length > 0 ? "Remaining leads are in the cool-off period." : "No leads assigned yet — ask your manager."}</p>
    </div>`;

  return `
    <div class="feed-card">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
        <span class="feed-label">Next Lead</span>
        <div style="display:flex;gap:6px">
          ${lead.leadType ? `<span class="lead-type-badge lead-type-${(lead.leadType||"").toLowerCase()}">${escHtml(lead.leadType)}</span>` : ""}
          <span class="status-badge status-${lead.status.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}" id="feed-current-status">${lead.status}</span>
        </div>
      </div>
      <div class="feed-name">${escHtml(lead.name)}</div>
      ${Graph.isInCoolOff(lead) ? `<div style="background:#FFF8E1;border:1px solid #FFD700;border-radius:6px;padding:8px 14px;margin-bottom:12px;font-family:var(--font-mono);font-size:11px;color:#8B6914">⏱ This lead is in the ${Config.rules.coolOffDays}-day cool-off period — you can still update it if the customer reached out.</div>` : ""}
      <div class="feed-meta-row">
        ${lead.phone    ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 12 19.79 19.79 0 0 1 1.61 3.38 2 2 0 0 1 3.6 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L7.91 8.96a16 16 0 0 0 6 6l.92-.92a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 21.73 16.92z" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.phone)}</span>` : ""}
        ${lead.email    ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" stroke="currentColor" stroke-width="2"/><polyline points="22,6 12,13 2,6" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.email)}</span>` : ""}
        ${lead.address  ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z" stroke="currentColor" stroke-width="2"/><circle cx="12" cy="10" r="3" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.address)}${lead.city ? ", " + escHtml(lead.city) : ""}${lead.state ? " " + escHtml(lead.state) : ""}${lead.zip ? " " + escHtml(lead.zip) : ""}</span>` : ""}
      </div>

      ${lead.notes ? `<div class="feed-notes">${escHtml(lead.notes)}</div>` : ""}

      <div class="feed-customer-info">
        <div class="form-group">
          <label>BTN</label>
          <input type="text" id="feed-btn" class="form-input" placeholder="Enter BTN" value="${escHtml(lead.btn||"")}">
        </div>
        <div class="form-group">
          <label>Package / Current Products</label>
          <select id="feed-products" class="form-input">
            <option value="">Select products...</option>
            ${Config.currentProducts.map(function(p) { return `<option value="${p}" ${lead.currentProducts===p?"selected":""}>${p}</option>`; }).join("")}
          </select>
        </div>
        <div class="form-group">
          <label>Price (MRC)</label>
          <input type="text" id="feed-mrc" class="form-input" placeholder="e.g. $104.49" value="${escHtml(lead.currentMRC||"")}">
        </div>
        <div class="form-group">
          <label>CBR</label>
          <input type="text" id="feed-cbr" class="form-input" placeholder="Enter CBR" value="${escHtml(lead.cbr||"")}">
        </div>
      </div>

      <!-- Sold By — always visible, required before saving -->
      <div style="margin-bottom:16px">
        <div style="font-family:var(--font-mono);font-size:10px;color:#6B85B0;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">
          Sold By <span style="color:var(--red)">* Required</span>
        </div>
        <select id="feed-sold-by" class="form-input" style="max-width:300px">
          <option value="">Select agent who made the sale...</option>
          ${State.contractors.map(function(c) {
            return `<option value="${escHtml(c.name)}">${escHtml(c.name)}</option>`;
          }).join("")}
        </select>
      </div>

      <!-- AutoPay -->
      <div style="margin-bottom:16px">
        <div style="font-family:var(--font-mono);font-size:10px;color:#6B85B0;margin-bottom:8px;text-transform:uppercase;letter-spacing:1px">
          AutoPay <span style="color:var(--red)">* Required</span>
        </div>
        <div style="display:flex;gap:8px;flex-wrap:wrap">
          ${["ACH - Debit Card","ACH - Credit Card","No Auto Pay"].map(function(opt) {
            return `<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;color:#1A2640;background:#F4F7FD;border:1px solid #D0DCF0;padding:8px 14px;border-radius:6px;transition:all 0.15s">
              <input type="radio" name="feed-autopay" value="${opt}" ${lead.autoPay===opt?"checked":""} style="accent-color:#2563B0"> ${opt}
            </label>`;
          }).join("")}
        </div>
      </div>

      <div class="feed-status-row">
        <span class="feed-label">Select Status — click Save to confirm</span>
        <div class="feed-status-buttons" id="feed-status-buttons">
          ${Config.leadStatuses.filter(function(s) { return s !== "New"; }).map(function(s) {
            const cls      = "status-btn-" + s.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"");
            const isTDM    = s === "TDM";
            const disabled = atLimit && !Config.terminalStatuses.includes(s);
            return `<button class="status-btn ${cls}" id="sbtn-${s.replace(/\s+/g,"-")}"
              onclick="stageStatus('${lead.id}','${s}')"
              ${disabled ? "disabled title='Daily limit reached'" : ""}
              >${s}${isTDM ? " ↩" : ""}</button>`;
          }).join("")}
        </div>
      </div>

      <div id="feed-staged-notice" style="display:none;margin-top:6px;font-family:var(--font-mono);font-size:11px;color:var(--amber)"></div>

      <div style="margin-top:16px">
        <span class="feed-label">Notes</span>
        ${lead.notes ? `<div style="background:#F4F7FD;border:1px solid #D0DCF0;border-radius:6px;padding:10px 14px;margin-top:6px;margin-bottom:8px;max-height:140px;overflow-y:auto">
          ${(lead.notes||"").split("\n").filter(function(l){return l.trim();}).map(function(line) {
            const match = line.match(/^\[(\d{2}\/\d{2}(?:\/\d{2})?)(.*?)\]\s*(.*)/);
            if (match) {
              const date  = match[1];
              const agent = match[2] ? match[2].replace(/^\s*-\s*/,"") : "";
              const text  = match[3];
              return `<div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
                <div style="display:flex;gap:8px;align-items:center;margin-bottom:3px">
                  <span style="font-family:var(--font-mono);font-size:10px;color:#2563B0;font-weight:700;background:#E8F0FF;padding:1px 6px;border-radius:3px">${date}</span>
                  ${agent ? `<span style="font-family:var(--font-mono);font-size:10px;color:#6B85B0">${escHtml(agent)}</span>` : ""}
                </div>
                <span style="font-size:13px;color:#1A2640">${escHtml(text)}</span>
              </div>`;
            }
            return `<div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8"><div style="margin-bottom:3px"><span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8;background:#F4F7FD;padding:1px 6px;border-radius:3px">Legacy note — author unknown</span></div><span style="font-size:13px;color:#4A6080">${escHtml(line)}</span></div>`;
          }).join("")}
        </div>` : ""}
        <div style="font-family:var(--font-mono);font-size:10px;color:#6B85B0;margin-top:8px;margin-bottom:4px;text-transform:uppercase;letter-spacing:1px">
          ${new Date().toLocaleDateString("en-US",{month:"2-digit",day:"2-digit",year:"2-digit"})} — Today's Note
        </div>
        <div class="feed-note-row">
          <textarea id="feed-notes" class="form-input form-textarea" placeholder="Add a note for today..."></textarea>
          <button class="btn-primary" id="feed-save-btn" onclick="agentSaveAll('${lead.id}')">Save</button>
        </div>
      </div>

      <div id="feed-next-row" style="display:none;margin-top:12px">
        <button class="btn-cyan" style="width:100%;justify-content:center;font-size:16px;padding:14px" onclick="advanceToNextLead()">
          Next Lead →
        </button>
      </div>
    </div>`;
}

function stageStatus(leadId, newStatus) {
  const lead = State.leads.find(function(l) { return l.id === leadId; });
  if (!lead) return;

  if (Graph.isInCoolOff(lead) && !Config.terminalStatuses.includes(newStatus)) {
    UI.showToast("Note: this lead is in the " + Config.rules.coolOffDays + "-day cool-off period.", "info");
  }

  _stagedStatus = newStatus;

  document.querySelectorAll(".status-btn").forEach(function(btn) {
    btn.style.borderColor = "";
    btn.style.color       = "";
    btn.style.background  = "";
    btn.style.boxShadow   = "";
  });
  const selectedBtn = document.getElementById("sbtn-" + newStatus.replace(/\s+/g, "-"));
  if (selectedBtn) {
    selectedBtn.style.borderColor = "var(--cyan)";
    selectedBtn.style.color       = "var(--cyan)";
    selectedBtn.style.background  = "var(--cyan-dim)";
    selectedBtn.style.boxShadow   = "0 0 12px var(--cyan-glow)";
  }

  const notice = document.getElementById("feed-staged-notice");
  if (notice) {
    notice.style.display = "block";
    notice.textContent   = "⚡ \"" + newStatus + "\" staged — click Save to confirm";
  }

  const badge = document.getElementById("feed-current-status");
  if (badge) {
    badge.textContent   = newStatus + " (staged)";
    badge.style.opacity = "0.7";
  }
}

async function agentSaveAll(leadId) {
  const user = State.currentUser;
  const lead = State.leads.find(function(l) { return l.id === leadId; });
  if (!lead) return;

  const newStatus  = _stagedStatus || lead.status;
  const mrc        = (document.getElementById("feed-mrc")      || {}).value || "";
  const products   = (document.getElementById("feed-products") || {}).value || "";
  const newNote    = (document.getElementById("feed-notes")    || {}).value || "";
  const cbr        = (document.getElementById("feed-cbr")      || {}).value || "";
  const btn        = (document.getElementById("feed-btn")      || {}).value || "";
  const autoPayEl  = document.querySelector('input[name="feed-autopay"]:checked');
  const autoPay    = autoPayEl ? autoPayEl.value : "";
  const soldByEl   = document.getElementById("feed-sold-by");
  const soldByName = soldByEl ? soldByEl.value : "";

  // Require AutoPay
  if (!autoPay) { UI.showToast("Please select an AutoPay option before saving.", "error"); return; }

  // Require Sold By if status is Sold
  if (newStatus === Config.soldStatus && !soldByName) {
    UI.showToast("Please select who made this sale in the Sold By field.", "error");
    return;
  }

  // Resolve sold by agent's email for activity log
  const soldByContractor = soldByName
    ? State.contractors.find(function(c) { return c.name === soldByName; })
    : null;
  const soldByEmail = soldByContractor
    ? (soldByContractor.email || soldByName)
    : (user && user.email) || "";

  // For non-sold statuses, use current user's email
  const activityEmail = newStatus === Config.soldStatus ? soldByEmail : (user && user.email) || "";

  let notes = lead.notes || "";
  if (newNote.trim()) {
    const today     = new Date();
    const dateStamp = (today.getMonth()+1).toString().padStart(2,"0") + "/" + today.getDate().toString().padStart(2,"0") + "/" + String(today.getFullYear()).slice(-2);
    const agentTag  = (user && user.name) ? " - " + user.name : "";
    const stamped   = "[" + dateStamp + agentTag + "] " + newNote.trim();
    notes = notes ? stamped + "\n" + notes : stamped;
  }

  setLoading(true);
  try {
    const today      = new Date().toISOString().split("T")[0];
    const saveFields = { Status: newStatus, LastTouchedOn: today };
    if (mrc)      saveFields["MonthlyRecurringCharge_x0028_MRC"] = mrc;
    if (products) saveFields["CurrentProducts"] = products;
    if (cbr)      saveFields["CBR"] = cbr;
    if (btn)      saveFields["BTN"] = btn;
    if (notes)    saveFields["Notes"] = notes;
    if (autoPay)  saveFields["AutoPay"] = autoPay;
    await Graph.updateLead(leadId, saveFields);

    if (newStatus === "TDM") {
      await Graph.assignAgent(leadId, "");
      UI.showToast("TDM — lead returned to admin queue.", "info");
    }

    // Log activity — credit goes to soldByEmail for Sold, current user for everything else
    await Graph.logActivity({
      LeadID:     leadId,
      Title:      lead.name,
      ActionType: "Status: " + newStatus,
      AgentEmail: activityEmail,
      Notes:      notes + (newStatus === Config.soldStatus && soldByName && soldByName !== (user && user.name)
        ? " [Sold by " + soldByName + ", recorded by " + ((user && user.name) || "admin") + "]"
        : ""),
    });

    if (newStatus === Config.soldStatus) {
      UI.showConfetti();
      const savingAgentName = (user && user.name) || "";
      // If saving for someone else, show "Noah just made a sale for Jon"
      // If saving own lead, show "Jon just closed a sale!"
      if (soldByName && savingAgentName && soldByName !== savingAgentName) {
        UI.showSaleBanner(lead.name, soldByName, lead.assignedTo !== soldByName ? lead.assignedTo : null);
      } else {
        UI.showSaleBanner(lead.name, soldByName || savingAgentName, null);
      }
    } else if (newStatus !== "TDM") {
      UI.showToast("Saved!", "success");
    }

    _stagedStatus = null;
    _leadSaved    = true;
    await loadAllData();

    const nextRow   = document.getElementById("feed-next-row");
    const searchSec = document.getElementById("lead-search-section");
    const saveBtn   = document.getElementById("feed-save-btn");
    if (nextRow)   { nextRow.style.display   = "block"; }
    if (searchSec) { searchSec.style.display = "block"; }
    if (saveBtn)   { saveBtn.textContent = "Saved ✓"; saveBtn.disabled = true; saveBtn.style.background = "var(--green)"; }
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

function advanceToNextLead() {
  _currentFeedIndex++;
  _leadSaved = false;
  renderMyLeads();
}

// ============================================================
//  ASSIGN LEADS (Admin only)
// ============================================================
function renderAssignLeads() {
  if (!isAdmin()) { navigate("myleads"); return; }

  const { leads, contractors } = State;
  const unassigned = leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); });
  const max        = Config.rules.maxLeadsPerAgent;

  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Assign Leads</h1>
        <span class="view-subtitle">// ${unassigned.length} unassigned</span>
      </div>
      <div style="display:flex;gap:8px">
        <button class="btn-cyan" onclick="navigate('drip')">
          <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><polyline points="12,8 12,12 14,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Drip Feed Mode
        </button>
        <button class="btn-primary" onclick="autoAssignLeads()">Auto-Assign Evenly</button>
      </div>
    </div>

    <div class="card" style="margin-bottom:20px;border-color:#2563B0">
      <div class="card-header" style="background:#EEF4FB">
        <h2 class="card-title" style="color:#0D1B3E">Assign by Quantity</h2>
        <span class="card-meta">Set how many leads each agent should receive</span>
      </div>
      <div style="padding:16px 20px">
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(260px,1fr));gap:12px;margin-bottom:16px">
          ${contractors.map(function(c) {
            const current = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
            return `
              <div style="display:flex;align-items:center;gap:10px;padding:12px;background:#F4F7FD;border:1px solid #D0DCF0;border-radius:8px">
                <div class="contractor-avatar" style="width:36px;height:36px;font-size:16px">${c.name[0].toUpperCase()}</div>
                <div style="flex:1;min-width:0">
                  <div style="font-family:var(--font-head);font-size:13px;font-weight:700;color:#0D1B3E;text-transform:uppercase">${escHtml(c.name)}</div>
                  <div style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8;margin-top:2px">${current} currently assigned</div>
                </div>
                <div style="display:flex;align-items:center;gap:6px">
                  <input type="number" id="qty-${escHtml(c.name)}" class="form-input" min="0" max="${unassigned.length}" placeholder="0"
                    style="width:70px;text-align:center;padding:6px 8px;font-size:14px;font-weight:600">
                  <span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8">leads</span>
                </div>
              </div>`;
          }).join("")}
        </div>
        <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">
          <button class="btn-primary" onclick="bulkAssignByQuantity()">
            <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><polyline points="20,6 9,17 4,12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
            Assign by Quantity
          </button>
          <span id="qty-preview" style="font-family:var(--font-mono);font-size:12px;color:#6B85B0"></span>
        </div>
      </div>
    </div>

    <div class="assign-agent-grid">
      ${contractors.map(function(c) {
        const count = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
        const pct   = Math.min(100, Math.round((count/max)*100));
        return `
          <div class="assign-agent-card ${count >= max ? "agent-full" : ""}">
            <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
            <div class="assign-agent-info">
              <span class="contractor-name">${escHtml(c.name)}</span>
              <div class="load-bar-wrap"><div class="load-bar ${pct>=100?"load-full":pct>=80?"load-high":""}" style="width:${pct}%"></div></div>
              <span class="assign-count ${count>=max?"text-danger":""}">${count} assigned</span>
            </div>
          </div>`;
      }).join("")}
    </div>

    <div class="card">
      <div class="card-header"><h2 class="card-title">Unassigned Leads (${unassigned.length})</h2></div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Name</th><th>Type</th><th>Phone</th><th>Status</th><th>Assign To</th></tr></thead>
          <tbody>
            ${unassigned.length ? unassigned.map(function(lead) { return `
              <tr>
                <td><span class="lead-name">${escHtml(lead.name)}</span></td>
                <td>${lead.leadType ? `<span class="lead-type-badge lead-type-${(lead.leadType||"").toLowerCase()}">${escHtml(lead.leadType)}</span>` : "—"}</td>
                <td class="td-mono">${escHtml(lead.phone)}</td>
                <td><span class="status-badge status-${lead.status.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}">${lead.status}</span></td>
                <td>
                  <div class="assign-select-row">
                    <select class="filter-select assign-select" id="assign-${lead.id}">
                      <option value="">Select agent</option>
                      ${contractors.map(function(c) {
                        const cnt = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
                        return `<option value="${escHtml(c.name)}">${escHtml(c.name)} (${cnt} assigned)</option>`;
                      }).join("")}
                    </select>
                    <button class="btn-primary" style="padding:6px 14px;font-size:12px" onclick="assignLead('${lead.id}')">Assign</button>
                  </div>
                </td>
              </tr>`;
            }).join("") : `<tr><td colspan="5" class="empty-state">All leads are assigned!</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  State.contractors.forEach(function(c) {
    const input = document.getElementById("qty-" + c.name);
    if (input) {
      input.addEventListener("input", function() {
        const total = State.contractors.reduce(function(sum, agent) {
          const val = parseInt((document.getElementById("qty-" + agent.name)||{}).value||"0") || 0;
          return sum + val;
        }, 0);
        const preview         = document.getElementById("qty-preview");
        const unassignedCount = State.leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); }).length;
        if (preview) preview.textContent = total + " of " + unassignedCount + " unassigned leads allocated";
        if (preview) preview.style.color = total > unassignedCount ? "#FF4444" : "#2563B0";
      });
    }
  });
}

async function assignLead(leadId) {
  const select = document.getElementById("assign-" + leadId);
  const agent  = select && select.value;
  if (!agent) { UI.showToast("Please select an agent.", "error"); return; }
  if (!Graph.canAgentTakeLead(agent, State.leads)) { UI.showToast(agent + " is at the lead limit.", "error"); return; }
  const lead = State.leads.find(function(l) { return l.id === leadId; });
  setLoading(true);
  try {
    await Graph.assignAgent(leadId, agent);
    await Graph.logActivity({ LeadID: leadId, Title: lead ? lead.name : "", ActionType: "Assigned", AgentEmail: (State.currentUser && State.currentUser.email) || "", Notes: "Assigned by " + ((State.currentUser && State.currentUser.name) || "Admin") });
    UI.showToast("Assigned to " + agent, "success");
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

async function bulkAssignByQuantity() {
  const { leads, contractors } = State;
  const unassigned = leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); });

  const plan = [];
  contractors.forEach(function(c) {
    const qty = parseInt((document.getElementById("qty-" + c.name)||{}).value||"0") || 0;
    if (qty > 0) plan.push({ agent: c.name, qty: qty });
  });

  if (!plan.length) { UI.showToast("Please enter a quantity for at least one agent.", "error"); return; }

  const totalRequested = plan.reduce(function(s,p){return s+p.qty;},0);
  if (totalRequested > unassigned.length) {
    UI.showToast("Total (" + totalRequested + ") exceeds unassigned leads (" + unassigned.length + "). Reduce quantities.", "error");
    return;
  }

  const summary = plan.map(function(p){return p.qty + " → " + p.agent;}).join("\n");
  if (!confirm("Assign leads by quantity?\n\n" + summary + "\n\nTotal: " + totalRequested + " leads")) return;

  setLoading(true);
  try {
    let idx = 0;
    for (var p = 0; p < plan.length; p++) {
      for (var q = 0; q < plan[p].qty; q++) {
        if (idx >= unassigned.length) break;
        await Graph.assignAgent(unassigned[idx].id, plan[p].agent);
        idx++;
      }
    }
    UI.showToast("Assigned " + totalRequested + " leads successfully!", "success");
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

// ============================================================
//  LEADS VIEW (Admin only)
// ============================================================
function renderLeads() {
  if (!isAdmin()) { navigate("myleads"); return; }

  State.selectedLeads.clear();
  const contractors = State.contractors.map(function(c) { return c.name; });
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">All Leads</h1>
        <span class="view-subtitle" id="leads-subtitle">// ${State.leads.length} total</span>
      </div>
      <div class="header-actions">
        <button class="btn-ghost" onclick="refreshData()">
          <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><polyline points="23,4 23,10 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="1,20 1,14 7,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Refresh
        </button>
        <button class="btn-ghost" onclick="exportCSV()">
          <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="7,10 12,15 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Export CSV
        </button>
        <button class="btn-ghost btn-danger-ghost" onclick="confirmClearAll()">
          <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Clear All
        </button>
        <button class="btn-primary" onclick="openAddLeadModal()">+ Add Lead</button>
      </div>
    </div>

    <div class="bulk-bar" id="bulk-bar" style="display:none">
      <span class="bulk-count" id="bulk-count">0 selected</span>
      <div class="bulk-actions">
        <select class="filter-select bulk-assign-select" id="bulk-assign-select" style="min-width:180px;font-size:12px;padding:6px 10px">
          <option value="">Assign to agent...</option>
          ${contractors.map(function(c) { return `<option value="${escHtml(c)}">${escHtml(c)}</option>`; }).join("")}
        </select>
        <button class="btn-cyan bulk-btn" onclick="bulkAssign()">Assign Agent</button>
        <select class="filter-select bulk-assign-select" id="bulk-type-select" style="min-width:140px;font-size:12px;padding:6px 10px">
          <option value="">Assign type...</option>
          ${Config.leadTypes.map(function(t) { return `<option value="${escHtml(t)}">${escHtml(t)}</option>`; }).join("")}
        </select>
        <button class="btn-cyan bulk-btn" onclick="bulkAssignType()">Assign Type</button>
        <button class="btn-ghost bulk-btn" onclick="bulkExportSelected()">Export Selected</button>
        <button class="bulk-btn bulk-delete-btn" onclick="bulkDelete()">
          <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Delete Selected
        </button>
        <button class="btn-ghost bulk-btn" onclick="clearSelection()">Clear Selection</button>
      </div>
    </div>

    <div class="filters-bar">
      <div class="search-wrap">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8" stroke="currentColor" stroke-width="2"/><line x1="21" y1="21" x2="16.65" y2="16.65" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        <input type="text" id="search-input" class="search-input" placeholder="Search name, email..." value="${State.filters.search}" oninput="applyFilters()">
      </div>
      <select class="filter-select" id="filter-status" onchange="applyFilters()">
        <option value="all">All Statuses</option>
        ${Config.leadStatuses.map(function(s) { return `<option value="${s}" ${State.filters.status===s?"selected":""}>${s}</option>`; }).join("")}
      </select>
      <select class="filter-select" id="filter-agent" onchange="applyFilters()">
        <option value="all">All Agents</option>
        ${contractors.map(function(c) { return `<option value="${c}" ${State.filters.assignedTo===c?"selected":""}>${c}</option>`; }).join("")}
      </select>
    </div>
    <div class="card" id="leads-table-wrap">${renderLeadsTable(getFilteredLeads())}</div>
  `;
}

function toggleLeadSelect(id, checked) {
  if (checked) { State.selectedLeads.add(id); }
  else         { State.selectedLeads.delete(id); }
  updateBulkBar();
}

function toggleSelectAll(checked) {
  const checkboxes = document.querySelectorAll(".lead-checkbox");
  checkboxes.forEach(function(cb) {
    cb.checked = checked;
    if (checked) { State.selectedLeads.add(cb.dataset.id); }
    else         { State.selectedLeads.delete(cb.dataset.id); }
  });
  updateBulkBar();
}

function updateBulkBar() {
  const bar   = document.getElementById("bulk-bar");
  const count = document.getElementById("bulk-count");
  const n     = State.selectedLeads.size;
  if (!bar) return;
  bar.style.display = n > 0 ? "flex" : "none";
  if (count) count.textContent = n + " lead" + (n !== 1 ? "s" : "") + " selected";
  const allCbs = document.querySelectorAll(".lead-checkbox");
  const selAll = document.getElementById("select-all-cb");
  if (selAll && allCbs.length) {
    selAll.indeterminate = n > 0 && n < allCbs.length;
    selAll.checked       = n === allCbs.length;
  }
}

function clearSelection() {
  State.selectedLeads.clear();
  document.querySelectorAll(".lead-checkbox").forEach(function(cb) { cb.checked = false; });
  const selAll = document.getElementById("select-all-cb");
  if (selAll) { selAll.checked = false; selAll.indeterminate = false; }
  updateBulkBar();
}

async function bulkDelete() {
  const ids = Array.from(State.selectedLeads);
  if (!ids.length) return;
  if (!confirm("Permanently delete " + ids.length + " lead" + (ids.length !== 1 ? "s" : "") + "? This cannot be undone.")) return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) { await Graph.deleteLead(ids[i]); }
    UI.showToast("Deleted " + ids.length + " lead" + (ids.length !== 1 ? "s" : ""), "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

async function bulkAssign() {
  const ids   = Array.from(State.selectedLeads);
  const agent = (document.getElementById("bulk-assign-select") || {}).value;
  if (!ids.length) return;
  if (!agent) { UI.showToast("Please select an agent first.", "error"); return; }
  if (!confirm("Assign " + ids.length + " lead" + (ids.length !== 1 ? "s" : "") + " to " + agent + "?")) return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) { await Graph.assignAgent(ids[i], agent); }
    UI.showToast("Assigned " + ids.length + " leads to " + agent, "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

async function bulkAssignType() {
  const ids  = Array.from(State.selectedLeads);
  const type = (document.getElementById("bulk-type-select") || {}).value;
  if (!ids.length) return;
  if (!type) { UI.showToast("Please select a lead type first.", "error"); return; }
  if (!confirm("Set type to \"" + type + "\" for " + ids.length + " lead" + (ids.length !== 1 ? "s" : "") + "?")) return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) { await Graph.updateLead(ids[i], { Lead_x0020_Type: type }); }
    UI.showToast("Set " + ids.length + " leads to type: " + type, "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

function bulkExportSelected() {
  const ids   = Array.from(State.selectedLeads);
  const leads = State.leads.filter(function(l) { return ids.includes(l.id); });
  if (!leads.length) return;
  const today = new Date().toISOString().slice(0,10);
  const csv   = ["Name,Type,Email,Phone,Status,Source,Assigned To,MRC,Current Products,Last Contacted,Notes"]
    .concat(leads.map(function(l) {
      return [l.name,l.leadType,l.email,l.phone,l.status,l.source,l.assignedTo,l.currentMRC,l.currentProducts,l.lastContacted,l.notes]
        .map(function(v){ return '"'+String(v||"").replace(/"/g,'""')+'"'; }).join(",");
    })).join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
  a.download = "raimak-leads-selected-" + today + ".csv";
  a.click();
  UI.showToast("Exported " + leads.length + " leads!", "success");
}

async function confirmClearAll() {
  const total = State.leads.length;
  if (!total) { UI.showToast("No leads to clear.", "info"); return; }
  const input = prompt("This will permanently delete ALL " + total + " leads.\n\nType DELETE to confirm:");
  if (input !== "DELETE") { UI.showToast("Clear all cancelled.", "info"); return; }
  setLoading(true);
  try {
    for (var i = 0; i < State.leads.length; i++) { await Graph.deleteLead(State.leads[i].id); }
    UI.showToast("All " + total + " leads deleted. Clean slate!", "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

function getFilteredLeads() {
  let leads = State.leads.slice();
  const { status, search, assignedTo } = State.filters;
  if (status !== "all")     leads = leads.filter(function(l) { return l.status === status; });
  if (assignedTo !== "all") leads = leads.filter(function(l) { return l.assignedTo === assignedTo; });
  if (search.trim()) {
    const q = search.toLowerCase();
    leads = leads.filter(function(l) { return l.name.toLowerCase().includes(q) || l.email.toLowerCase().includes(q); });
  }
  return leads;
}

function applyFilters() {
  State.filters.search     = (document.getElementById("search-input")  || {}).value || "";
  State.filters.status     = (document.getElementById("filter-status") || {}).value || "all";
  State.filters.assignedTo = (document.getElementById("filter-agent")  || {}).value || "all";
  const wrap = document.getElementById("leads-table-wrap");
  if (wrap) wrap.innerHTML = renderLeadsTable(getFilteredLeads());
}

function renderLeadsTable(leads, compact, agentView) {
  if (!leads.length) return `<div class="empty-state"><p>No leads found.</p></div>`;
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr>
          ${compact ? "" : `<th style="width:36px"><input type="checkbox" id="select-all-cb" class="lead-cb" onchange="toggleSelectAll(this.checked)" title="Select all"></th>`}
          <th>Name</th><th>Type</th><th>Status</th><th>Phone</th><th>Assigned To</th><th>Address</th><th>Last Contacted</th>
          ${compact ? "" : "<th>CBR</th><th>BTN</th><th>Flags</th><th></th>"}
        </tr></thead>
        <tbody>
          ${leads.map(function(lead) {
            const statusCls = "status-" + lead.status.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"");
            const typeCls   = lead.leadType ? "lead-type-" + lead.leadType.toLowerCase() : "";
            const isChecked = State.selectedLeads.has(lead.id);
            const rowClick  = agentView
              ? "loadLeadInFeed('" + lead.id + "')"
              : (isAdmin() ? "openEditLeadModal('" + lead.id + "')" : "");
            return `
              <tr class="lead-row ${lead.flags && lead.flags.includes("needs_recycle") ? "row-warn" : ""} ${isChecked ? "row-selected" : ""}"
                  onclick="${rowClick}" style="${agentView ? "cursor:pointer" : ""}">
                ${compact ? "" : `<td onclick="event.stopPropagation()" style="width:36px"><input type="checkbox" class="lead-checkbox lead-cb" data-id="${lead.id}" ${isChecked?"checked":""} onchange="toggleLeadSelect('${lead.id}',this.checked)"></td>`}
                <td><span class="lead-name">${escHtml(lead.name)}</span></td>
                <td>${lead.leadType ? `<span class="lead-type-badge ${typeCls}">${escHtml(lead.leadType)}</span>` : "—"}</td>
                <td><span class="status-badge ${statusCls}">${lead.status}</span></td>
                <td class="td-mono">${escHtml(lead.phone || "—")}</td>
                <td>${escHtml(lead.assignedTo || "—")}</td>
                <td class="td-mono" style="font-size:11px">${lead.address ? escHtml(lead.address) + (lead.city ? ", " + escHtml(lead.city) : "") + (lead.state ? " " + escHtml(lead.state) : "") : "—"}</td>
                <td class="td-mono">${formatDate(lead.lastContacted) || "—"}</td>
                ${compact ? "" : `
                <td class="td-mono">${escHtml(lead.cbr || "—")}</td>
                <td class="td-mono">${escHtml(lead.btn || "—")}</td>
                <td class="td-flags">${(lead.flags||[]).map(function(f) { return `<span class="flag flag-${f}">${flagLabel(f)}</span>`; }).join("")}</td>
                <td class="td-actions">
                  ${isAdmin() ? `
                    <button class="btn-icon" onclick="event.stopPropagation();openEditLeadModal('${lead.id}')" title="Edit">
                      <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
                    </button>
                    <button class="btn-icon btn-danger" onclick="event.stopPropagation();deleteLead('${lead.id}')" title="Delete">
                      <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
                    </button>` : ""}
                </td>`}
              </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>`;
}

function loadLeadInFeed(leadId) {
  const lead = (window._myLeads || []).find(function(l) { return l.id === leadId; });
  if (!lead) return;
  const feedWrap = document.getElementById("lead-feed-wrap");
  if (feedWrap) {
    _leadSaved = false;
    feedWrap.innerHTML = renderLeadFeedCard([lead], 0, true);
    feedWrap.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

// ============================================================
//  DAILY REPORT (Admin only)
// ============================================================
async function renderDailyReport() {
  if (!isAdmin()) { navigate("myleads"); return; }
  document.getElementById("main-content").innerHTML = `
    <div class="view-header"><h1 class="view-title">Daily Report</h1></div>
    <div class="card"><div class="empty-state" style="padding:40px">Loading report...</div></div>`;
  try {
    const stats  = await Graph.getDailyStats();
    const today  = new Date().toLocaleDateString("en-GB", { weekday:"long", year:"numeric", month:"long", day:"numeric" });
    document.getElementById("main-content").innerHTML = `
      <div class="view-header">
        <div>
          <h1 class="view-title">Daily Report</h1>
          <span class="view-subtitle">// ${today}</span>
        </div>
        <button class="btn-ghost" onclick="exportReportCSV()">Export CSV</button>
      </div>
      <div class="kpi-grid">
        <div class="kpi-card kpi-primary"><span class="kpi-label">Total Contacts Today</span><span class="kpi-value">${stats.reduce(function(s,a){return s+a.contacts;},0)}</span></div>
        <div class="kpi-card kpi-success"><span class="kpi-label">Total Sales Today</span><span class="kpi-value">${State.todaySales.length}</span></div>
        <div class="kpi-card kpi-info"><span class="kpi-label">Active Agents</span><span class="kpi-value">${stats.length}</span></div>
        <div class="kpi-card kpi-neutral"><span class="kpi-label">Avg Contacts/Agent</span><span class="kpi-value">${stats.length?Math.round(stats.reduce(function(s,a){return s+a.contacts;},0)/stats.length):0}</span></div>
      </div>
      <div class="card">
        <div class="card-header"><h2 class="card-title">Agent Breakdown</h2></div>
        <div class="table-wrap">
          <table class="data-table">
            <thead><tr><th>Agent</th><th>Contacts Today</th><th>Sales Today</th><th>Limit</th><th>Last Action</th></tr></thead>
            <tbody>
              ${stats.length ? stats.map(function(a) {
                const pct  = Math.round((a.contacts/Config.rules.maxContactsPerDay)*100);
                const last = a.actions.length ? a.actions[0] : null;
                return `<tr>
                  <td><span class="lead-name">${escHtml(a.agent)}</span></td>
                  <td>
                    <div style="display:flex;align-items:center;gap:10px">
                      <span class="td-mono">${a.contacts}</span>
                      <div class="load-bar-wrap" style="flex:1;max-width:80px"><div class="load-bar ${pct>=100?"load-full":pct>=80?"load-high":""}" style="width:${Math.min(100,pct)}%"></div></div>
                    </div>
                  </td>
                  <td><span class="status-badge status-sold">${a.sold}</span></td>
                  <td class="td-mono">${a.contacts}/${Config.rules.maxContactsPerDay}</td>
                  <td class="td-mono" style="color:var(--text-3)">${last?formatDateTime(last.timestamp):"—"}</td>
                </tr>`;
              }).join("") : `<tr><td colspan="5" class="empty-state">No activity today yet.</td></tr>`}
            </tbody>
          </table>
        </div>
      </div>`;
    window._reportStats = stats;
  } catch (err) {
    UI.showToast("Failed to load report: " + err.message, "error");
  }
}

function exportReportCSV() {
  const stats = window._reportStats || [];
  const today = new Date().toISOString().split("T")[0];
  const csv   = ["Agent,Contacts Today,Sales Today,Date"]
    .concat(stats.map(function(a) { return [a.agent,a.contacts,a.sold,today].map(function(v){return '"'+String(v||"").replace(/"/g,'""')+'"';}).join(","); }))
    .join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
  a.download = "raimak-report-" + today + ".csv";
  a.click();
  UI.showToast("Report exported!", "success");
}

// ============================================================
//  RAIMAK TEAM (Admin only)
// ============================================================
function renderContractors() {
  if (!isAdmin()) { navigate("myleads"); return; }
  const { contractors, leads } = State;
  const max = Config.rules.maxLeadsPerAgent;
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Raimak Team</h1>
      <span class="view-subtitle">// ${contractors.length} agents</span>
    </div>
    <div class="contractor-grid">
      ${contractors.map(function(c) {
        const count    = leads.filter(function(l){return l.assignedTo===c.name&&!Config.terminalStatuses.includes(l.status);}).length;
        const pct      = Math.min(100,Math.round((count/max)*100));
        const contacts = Graph.agentContactsToday(c.email || c.name, State.activityLog);
        return `
          <div class="contractor-card">
            <div class="contractor-header">
              <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
              <div><div class="contractor-name">${escHtml(c.name)}</div><div class="contractor-role">${escHtml(c.role)}</div></div>
              <span class="status-dot ${c.active?"dot-active":"dot-inactive"}"></span>
            </div>
            <div class="contractor-email">${escHtml(c.email||"No email")}</div>
            <div class="load-label"><span>Lead Load</span><span class="${count>=max?"text-danger":""}">${count}/${max}</span></div>
            <div class="load-bar-wrap"><div class="load-bar ${pct>=100?"load-full":pct>=80?"load-high":""}" style="width:${pct}%"></div></div>
            <div class="load-label" style="margin-top:10px"><span>Contacts Today</span><span>${contacts}/${Config.rules.maxContactsPerDay}</span></div>
            <div class="load-bar-wrap"><div class="load-bar ${contacts>=Config.rules.maxContactsPerDay?"load-full":""}" style="width:${Math.min(100,Math.round((contacts/Config.rules.maxContactsPerDay)*100))}%"></div></div>
          </div>`;
      }).join("")}
    </div>`;
}

// ============================================================
//  ACTIVITY LOG (Admin only)
// ============================================================
function renderActivity() {
  if (!isAdmin()) { navigate("myleads"); return; }
  const { activityLog } = State;
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Activity Log</h1>
      <span class="view-subtitle">// ${activityLog.length} entries</span>
    </div>
    <div class="card">
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Time</th><th>Lead</th><th>Action</th><th>Agent</th><th>Notes</th></tr></thead>
          <tbody>
            ${activityLog.length ? activityLog.map(function(e) { return `
              <tr>
                <td class="td-mono">${formatDateTime(e.timestamp)}</td>
                <td>${escHtml(e.leadName||e.leadId||"—")}</td>
                <td><span class="action-badge">${escHtml(e.action||"—")}</span></td>
                <td>${escHtml(e.agent||"—")}</td>
                <td class="td-notes">${escHtml(e.notes||"")}</td>
              </tr>`;
            }).join("") : `<tr><td colspan="5" class="empty-state">No activity yet.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>`;
}

// ============================================================
//  LEAD MODAL (Admin — Add/Edit)
// ============================================================
function openAddLeadModal() { if (!isAdmin()) return; State.editingLeadId = null; renderLeadModal(null); }

function openEditLeadModal(id) {
  if (!isAdmin()) return;
  const lead = State.leads.find(function(l) { return l.id === id; });
  if (!lead) return;
  State.editingLeadId = id;
  renderLeadModal(lead);
}

function renderLeadModal(lead) {
  const isEdit      = !!lead;
  const contractors = State.contractors.map(function(c) { return c.name; });

  document.getElementById("modal").innerHTML = `
    <div class="modal-header">
      <h2>${isEdit ? "Edit Lead" : "New Lead"}</h2>
      <button class="btn-icon" onclick="closeModal()">
        <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="6" y1="6" x2="18" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      </button>
    </div>
    <div class="modal-form">
      <div class="form-row">
        <div class="form-group"><label>First Name *</label><input type="text" id="f-firstname" class="form-input" value="${escHtml((lead&&lead.firstName)||"")}"></div>
        <div class="form-group"><label>Last Name *</label><input type="text" id="f-lastname" class="form-input" value="${escHtml((lead&&lead.lastName)||"")}"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Email</label><input type="email" id="f-email" class="form-input" value="${escHtml((lead&&lead.email)||"")}"></div>
        <div class="form-group"><label>Phone</label><input type="tel" id="f-phone" class="form-input" value="${escHtml((lead&&lead.phone)||"")}"></div>
      </div>
      <div class="form-section-title">Address</div>
      <div class="form-group form-group-full" style="margin-bottom:12px">
        <label>Street Address</label>
        <input type="text" id="f-address" class="form-input" placeholder="e.g. 125 Brown Rd" value="${escHtml((lead&&lead.address)||"")}">
      </div>
      <div class="form-row">
        <div class="form-group"><label>City</label><input type="text" id="f-city" class="form-input" value="${escHtml((lead&&lead.city)||"")}"></div>
        <div class="form-group"><label>State</label><input type="text" id="f-state" class="form-input" value="${escHtml((lead&&lead.state)||"")}"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Zip</label><input type="text" id="f-zip" class="form-input" value="${escHtml((lead&&lead.zip)||"")}"></div>
        <div class="form-group">
          <label>Lead Type</label>
          <select id="f-leadtype" class="form-input">
            <option value="">Select type...</option>
            ${Config.leadTypes.map(function(t) { return `<option value="${t}" ${lead&&lead.leadType===t?"selected":""}>${t}</option>`; }).join("")}
          </select>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Status</label>
          <select id="f-status" class="form-input">
            ${Config.leadStatuses.map(function(s) { return `<option value="${s}" ${((lead&&lead.status)||"New")===s?"selected":""}>${s}</option>`; }).join("")}
          </select>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Assigned To</label>
          <select id="f-assigned" class="form-input">
            <option value="">Unassigned</option>
            ${contractors.map(function(c) { return `<option value="${c}" ${lead&&lead.assignedTo===c?"selected":""}>${c}</option>`; }).join("")}
          </select>
        </div>
        <div class="form-group"><label>Last Contacted</label><input type="date" id="f-lastcontacted" class="form-input" value="${lead&&lead.lastContacted?lead.lastContacted.split("T")[0]:""}"></div>
      </div>
      <div class="form-section-title">Customer Info</div>
      <div class="form-row">
        <div class="form-group">
          <label>Monthly Recurring Charge (MRC)</label>
          <input type="text" id="f-mrc" class="form-input" placeholder="e.g. $104.49" value="${escHtml((lead&&lead.currentMRC)||"")}">
        </div>
        <div class="form-group">
          <label>Current Products</label>
          <select id="f-products" class="form-input">
            <option value="">Select products...</option>
            ${Config.currentProducts.map(function(p) { return `<option value="${p}" ${lead&&lead.currentProducts===p?"selected":""}>${p}</option>`; }).join("")}
          </select>
        </div>
      </div>
      <div class="form-group form-group-full" style="margin-bottom:16px">
        <label>AutoPay</label>
        <div style="display:flex;gap:16px;flex-wrap:wrap;margin-top:6px">
          ${["ACH - Debit Card","ACH - Credit Card","No Auto Pay"].map(function(opt) {
            const checked = lead&&lead.autoPay===opt?"checked":"";
            return `<label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer">
              <input type="radio" name="f-autopay" value="${opt}" ${checked} style="accent-color:#2563B0;width:14px;height:14px"> ${opt}
            </label>`;
          }).join("")}
        </div>
      </div>
      <div class="form-group form-group-full">
        <label>Notes History</label>
        ${lead && lead.notes ? `
        <div style="background:#F4F7FD;border:1px solid #D0DCF0;border-radius:6px;padding:12px 14px;margin-bottom:10px;max-height:180px;overflow-y:auto">
          ${(lead.notes||"").split("\n").filter(function(l){return l.trim();}).map(function(line) {
            const match = line.match(/^\[(\d{2}\/\d{2}(?:\/\d{2})?)(.*?)\]\s*(.*)/);
            if (match) {
              const date  = match[1];
              const agent = match[2] ? match[2].replace(/^\s*-\s*/,"") : "";
              const text  = match[3];
              return `<div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
                <div style="display:flex;gap:8px;align-items:center;margin-bottom:3px">
                  <span style="font-family:var(--font-mono);font-size:10px;color:#2563B0;font-weight:700;background:#E8F0FF;padding:1px 6px;border-radius:3px">${date}</span>
                  ${agent ? `<span style="font-family:var(--font-mono);font-size:10px;color:#6B85B0">${escHtml(agent)}</span>` : ""}
                </div>
                <span style="font-size:13px;color:#1A2640">${escHtml(text)}</span>
              </div>`;
            }
            return `<div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8"><div style="margin-bottom:3px"><span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8;background:#F4F7FD;padding:1px 6px;border-radius:3px">Legacy note — author unknown</span></div><span style="font-size:13px;color:#4A6080">${escHtml(line)}</span></div>`;
          }).join("")}
        </div>` : `<div style="font-size:12px;color:#8EA5C8;margin-bottom:10px;font-family:var(--font-mono)">No notes yet.</div>`}
        <label style="margin-top:4px">Add Note</label>
        <textarea id="f-notes" class="form-input form-textarea" placeholder="Add a note — will be date-stamped automatically on save..."></textarea>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" onclick="${isEdit ? 'submitEditLead()' : 'submitAddLead()'}">${isEdit ? 'Save Changes' : 'Add Lead'}</button>
    </div>
  `;
  document.getElementById("modal-overlay").style.display = "flex";
}

async function submitAddLead() {
  const fields    = collectLeadForm();
  if (!fields) return;
  const agentName = fields._agentName;
  delete fields._agentName;
  setLoading(true);
  try {
    const newLead = await Graph.addLead(fields);
    if (agentName) await Graph.assignAgent(newLead.id, agentName);
    await Graph.logActivity({ LeadID: newLead.id, Title: fields.Title, ActionType: "Lead Created", AgentEmail: (State.currentUser&&State.currentUser.email)||"" });
    await refreshData();
    closeModal();
    UI.showToast("Lead added!", "success");
  } catch (err) { UI.showToast("Failed: " + err.message, "error"); }
  finally { setLoading(false); }
}

async function submitEditLead() {
  const fields    = collectLeadForm();
  if (!fields) return;
  const agentName = fields._agentName;
  delete fields._agentName;
  setLoading(true);
  try {
    await Graph.updateLead(State.editingLeadId, fields);
    if (agentName) await Graph.assignAgent(State.editingLeadId, agentName);
    await Graph.logActivity({ LeadID: State.editingLeadId, Title: fields.Title, ActionType: "Lead Updated", AgentEmail: (State.currentUser&&State.currentUser.email)||"" });
    await refreshData();
    closeModal();
    UI.showToast("Lead updated!", "success");
  } catch (err) { UI.showToast("Failed: " + err.message, "error"); }
  finally { setLoading(false); }
}

function collectLeadForm() {
  const firstName = ((document.getElementById("f-firstname")||{}).value||"").trim();
  const lastName  = ((document.getElementById("f-lastname") ||{}).value||"").trim();
  const agentName = (document.getElementById("f-assigned")||{}).value || "";
  const nameEl    = document.getElementById("f-name");
  const fullName  = nameEl ? ((nameEl.value||"").trim()) : (firstName + " " + lastName).trim();

  if (!firstName && !lastName && !fullName) { UI.showToast("Name is required.", "error"); return null; }

  const fields = { _agentName: agentName };
  if (firstName) fields["FirstName"] = firstName;
  if (lastName)  fields["LastName"]  = lastName;
  if (fullName && !firstName) fields["Title"] = fullName;

  const add = function(key, elId, trim) {
    const el  = document.getElementById(elId);
    const val = el ? (trim ? (el.value||"").trim() : (el.value||"")) : "";
    if (val) fields[key] = val;
  };

  add("Lead_x0020_Type",                 "f-leadtype");
  add("Email",                            "f-email",        true);
  add("Phone",                            "f-phone",        true);
  add("Status",                           "f-status");
  add("LastTouchedOn",                    "f-lastcontacted");
  add("MonthlyRecurringCharge_x0028_MRC", "f-mrc",          true);
  add("CurrentProducts",                  "f-products");
  add("CBR",                              "f-cbr",          true);
  add("BTN",                              "f-btn",          true);
  add("WorkAddress",                      "f-address",      true);
  add("WorkCity",                         "f-city",         true);
  add("State",                            "f-state",        true);
  add("Zip",                              "f-zip",          true);
  add("AutoPay",                          "f-autopay");

  const notesEl = document.getElementById("f-notes");
  if (notesEl && notesEl.value.trim()) {
    const today     = new Date();
    const dateStamp = (today.getMonth()+1).toString().padStart(2,"0") + "/" + today.getDate().toString().padStart(2,"0") + "/" + String(today.getFullYear()).slice(-2);
    const adminName = (State.currentUser && State.currentUser.name) ? " - " + State.currentUser.name : "";
    const lead      = State.editingLeadId ? State.leads.find(function(l){return l.id===State.editingLeadId;}) : null;
    const existing  = (lead && lead.notes) || "";
    const stamped   = "[" + dateStamp + adminName + "] " + notesEl.value.trim();
    fields["Notes"] = existing ? stamped + "\n" + existing : stamped;
  }

  if (!fields.Status) fields.Status = "New";
  return fields;
}

async function deleteLead(id) {
  const lead = State.leads.find(function(l) { return l.id === id; });
  if (!confirm("Delete \"" + (lead&&lead.name) + "\"? This cannot be undone.")) return;
  setLoading(true);
  try {
    await Graph.deleteLead(id);
    await refreshData();
    UI.showToast("Lead deleted.", "success");
  } catch (err) { UI.showToast("Failed: " + err.message, "error"); }
  finally { setLoading(false); }
}

function closeModal(event) {
  if (event && event.target !== document.getElementById("modal-overlay")) return;
  document.getElementById("modal-overlay").style.display = "none";
}

// ============================================================
//  UTILITIES
// ============================================================
async function refreshData() { await loadAllData(); navigate(State.currentView); }

function setLoading(on) {
  State.loading = on;
  const o = document.getElementById("loading-overlay");
  if (o) o.style.display = on ? "flex" : "none";
}

function updateBadges() {
  const n = State.leads.filter(function(l) { return l.flags && (l.flags.includes("needs_recycle")||l.flags.includes("agent_overloaded")); }).length;
  const b = document.getElementById("badge-leads");
  if (b) { b.textContent = n > 0 ? n : ""; b.style.display = n > 0 ? "inline-flex" : "none"; }
}

function exportCSV() {
  const leads = getFilteredLeads();
  const today = new Date().toISOString().slice(0,10);
  const csv   = ["Name,Type,Email,Phone,Status,Source,Assigned To,MRC,Current Products,Last Contacted,Notes"]
    .concat(leads.map(function(l) {
      return [l.name,l.leadType,l.email,l.phone,l.status,l.source,l.assignedTo,l.currentMRC,l.currentProducts,l.lastContacted,l.notes]
        .map(function(v){return '"'+String(v||"").replace(/"/g,'""')+'"';}).join(",");
    })).join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv],{type:"text/csv"}));
  a.download = "raimak-leads-" + today + ".csv";
  a.click();
  UI.showToast("Exported!", "success");
}

function flagLabel(f) { return {cool_off:"Cool-off",needs_recycle:"Recycle",agent_overloaded:"Overloaded"}[f]||f; }
function formatDate(d) { if (!d) return ""; return new Date(d).toLocaleDateString("en-GB",{day:"2-digit",month:"short",year:"numeric"}); }
function formatTime(d) { if (!d) return ""; return new Date(d).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}); }
function formatDateTime(d) { if (!d) return ""; return new Date(d).toLocaleString("en-GB",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"}); }
function escHtml(str) { return String(str||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

const UI = {
  showToast: function(msg, type) {
    type = type || "info";
    const c = document.getElementById("toast-container");
    if (!c) return;
    const t = document.createElement("div");
    t.className = "toast toast-" + type;
    t.textContent = msg;
    c.appendChild(t);
    setTimeout(function(){t.classList.add("show");},10);
    setTimeout(function(){t.classList.remove("show");setTimeout(function(){t.remove();},300);},4000);
  },
  showConfetti: function() {
    const el = document.createElement("div");
    el.className = "confetti-burst";
    el.innerHTML = "&#127881; SOLD! &#127881;";
    document.body.appendChild(el);
    setTimeout(function(){el.remove();},2600);
  },
  // soldBy = agent who made the sale
  // forAgent = lead's assigned agent (if different from soldBy)
  showSaleBanner: function(leadName, soldBy, forAgent) {
    const existing = document.getElementById("sale-banner");
    if (existing) existing.remove();
    const banner = document.createElement("div");
    banner.id = "sale-banner";

    // Build message:
    // "Noah just made a sale for Jon — Customer Name"
    // "Stephanie just closed a sale! — Customer Name"
    const message = forAgent
      ? `<strong>${escHtml(soldBy)}</strong> just made a sale for <strong>${escHtml(forAgent)}</strong> — <strong>${escHtml(leadName)}</strong>`
      : `<strong>${escHtml(soldBy || "Someone")}</strong> just closed a sale! — <strong>${escHtml(leadName)}</strong>`;

    banner.innerHTML = `
      <div style="display:flex;align-items:center;gap:16px;justify-content:center;flex:1">
        <span style="font-size:20px">&#127881;</span>
        <span>${message}</span>
        <span style="font-size:20px">&#127881;</span>
      </div>
      <button onclick="document.getElementById('sale-banner').remove()"
        style="background:transparent;border:1px solid rgba(255,255,255,0.3);color:white;padding:4px 12px;border-radius:4px;cursor:pointer;font-family:'Rajdhani',sans-serif;font-size:13px;flex-shrink:0">
        Dismiss
      </button>`;
    banner.style.cssText = `
      position:fixed;top:32px;left:0;right:0;z-index:9997;
      background:linear-gradient(90deg,#0D1B3E,#1B4F8A,#0D1B3E);
      color:#FFFFFF;padding:12px 24px;display:flex;align-items:center;gap:16px;
      font-family:'Rajdhani',sans-serif;font-size:18px;font-weight:600;letter-spacing:1px;
      border-bottom:3px solid #00FF88;box-shadow:0 4px 24px rgba(0,255,136,0.3);
      animation:bannerSlide 0.4s ease both;
    `;
    document.body.appendChild(banner);
  },
};
