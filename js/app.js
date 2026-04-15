// Raimak LMS - App Logic v3.0

const State = {
  leads: [],
  contractors: [],
  activityLog: [],
  todaySales: [],
  currentView: "dashboard",
  filters: { status: "all", search: "", assignedTo: "all" },
  editingLeadId: null,
  loading: false,
  role: "agent",
  currentUser: null,
  salesFeedTimer: null,
  dripLead: null,
  selectedLeads: new Set(),
};

const stateTimezones = {
  AL: "America/Chicago",
  AK: "America/Anchorage",
  AZ: "America/Phoenix",
  AR: "America/Chicago",
  CA: "America/Los_Angeles",
  CO: "America/Denver",
  CT: "America/New_York",
  DE: "America/New_York",
  FL: "America/New_York",
  GA: "America/New_York",
  HI: "America/Honolulu",
  ID: "America/Boise",
  IL: "America/Chicago",
  IN: "America/Indianapolis",
  IA: "America/Chicago",
  KS: "America/Chicago",
  KY: "America/New_York",
  LA: "America/Chicago",
  ME: "America/New_York",
  MD: "America/New_York",
  MA: "America/New_York",
  MI: "America/Detroit",
  MN: "America/Chicago",
  MS: "America/Chicago",
  MO: "America/Chicago",
  MT: "America/Denver",
  NE: "America/Chicago",
  NV: "America/Los_Angeles",
  NH: "America/New_York",
  NJ: "America/New_York",
  NM: "America/Denver",
  NY: "America/New_York",
  NC: "America/New_York",
  ND: "America/Chicago",
  OH: "America/New_York",
  OK: "America/Chicago",
  OR: "America/Los_Angeles",
  PA: "America/New_York",
  RI: "America/New_York",
  SC: "America/New_York",
  SD: "America/Chicago",
  TN: "America/Chicago",
  TX: "America/Chicago",
  UT: "America/Denver",
  VT: "America/New_York",
  VA: "America/New_York",
  WA: "America/Los_Angeles",
  WV: "America/New_York",
  WI: "America/Chicago",
  WY: "America/Denver",
};

function isAdmin() {
  return State.role === "admin";
}

function detectRole(user) {
  if (!user) return "agent";
  const email = (user.email || "").toLowerCase();
  const admins = (Config.roles.admins || []).map(function (a) {
    return a.toLowerCase();
  });
  return admins.includes(email) ? "admin" : "agent";
}

// ── Boot ──────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", async function () {
  //separated login page html and js
  const loginBtn = document.getElementById("ms-login-btn");
  if (loginBtn) {
    loginBtn.addEventListener("click", function () {
      Auth.signIn();
    });
  }

  try {
    const redirectResult = await Auth.init();

    if (!Auth.isSignedIn()) {
      showLoginScreen();
      return;
    }

    if (redirectResult) {
      window.history.replaceState({}, document.title, window.location.pathname);
    }

    State.currentUser = Auth.getUser();
    State.role = detectRole(State.currentUser);
    showAppShell();
    await loadAllData();
    renderDashboard();
    Ticker.update();
  } catch (err) {
    console.error("Boot error:", err);
    showLoginScreen();
  }
});

async function loadAllData() {
  setLoading(true);
  try {
    // 1. Fetch the base data (Notice we only call getActivityLogForToday)
    const [rawLeads, contractors, todayLogs] = await Promise.all([
      Graph.getLeads(),
      Graph.getContractors(),
      Graph.getActivityLogForToday(),
    ]);

    State.contractors = contractors;
    State.leads = Graph.applyBusinessRules(rawLeads, contractors);

    // 2. Feed todayLogs directly in — NO double fetching!
    State.todaySales = await Graph.getTodaySales(todayLogs);

    // 3. Conditionally load the massive historical log ONLY for admins
    if (isAdmin()) {
      State.activityLog = await Graph.getActivityLog();
    } else {
      State.activityLog = todayLogs; // Standard agents just keep the lightweight log in state
    }
  } catch (err) {
    UI.showToast("Failed to load data: " + err.message, "error");
    console.error("Data Load Error:", err);
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  LOGIN
// ============================================================
function showLoginScreen() {
  // 1. Hide the main app wrapper
  document.getElementById("app-shell").style.display = "none";

  // 2. Show the login view
  document.getElementById("login-view").style.display = "flex";

  // 3. Safely inject the version number
  if (typeof Config !== "undefined" && Config.rules) {
    document.getElementById("app-version-text").textContent =
      Config.rules.appVersion;
  }
}

// ============================================================
//  APP SHELL
// ============================================================
function showAppShell() {
  const user = State.currentUser;

  // 1. Hide Login, Show App Shell
  document.getElementById("login-view").style.display = "none";
  document.getElementById("app-shell").style.display = "flex";

  // 2. Populate User Data dynamically into the HTML we just made
  if (user) {
    document.getElementById("ui-user-name").textContent = user.name || "User";
    document.getElementById("ui-user-email").textContent = user.email || "";
    document.getElementById("ui-user-initial").textContent = (user.name ||
      "U")[0].toUpperCase();
  }

  // 3. The "Cardboard" Security Guard (Frontend Role Check)
  // Hide all elements with the 'admin-only' class if they aren't an admin.
  if (!isAdmin()) {
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.style.display = "none";
    });
  } else {
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.style.display = "flex"; // Or 'block', depending on your CSS
    });
  }
}

// ============================================================
//  NAVIGATION
// ============================================================
function navigate(view) {
  if (window._clockTimer) clearInterval(window._clockTimer);
  const adminOnly = [
    "leads",
    "drip",
    "assign",
    "report",
    "contractors",
    "activity",
  ];
  if (!isAdmin() && adminOnly.includes(view)) {
    view = "myleads";
  }
  State.currentView = view;
  document.querySelectorAll(".nav-item").forEach(function (el) {
    el.classList.remove("active");
  });
  const navEl = document.querySelector("[data-view='" + view + "']");
  if (navEl) navEl.classList.add("active");
  switch (view) {
    case "dashboard":
      renderDashboard();
      break;
    case "leads":
      renderLeads();
      break;
    case "myleads":
      renderMyLeads();
      break;
    case "drip":
      renderDripFeed();
      break;
    case "assign":
      renderAssignLeads();
      break;
    case "contractors":
      renderContractors();
      break;
    case "activity":
      renderActivity();
      break;
    case "report":
      renderDailyReport();
      break;
  }
}

// ============================================================
//  DASHBOARD
// ============================================================
function renderDashboard() {
  const leads = State.leads;
  const todaySales = State.todaySales;
  const total = leads.length;

  // -- Keep all his math/counting logic exactly the same --
  const active = leads.filter(
    (l) => !Config.terminalStatuses.includes(l.status),
  ).length;
  const sold = leads.filter((l) => l.status === "Sold").length;
  const needRecycle = leads.filter(
    (l) => l.flags && l.flags.includes("needs_recycle"),
  ).length;
  const coolOff = leads.filter(
    (l) => l.flags && l.flags.includes("cool_off"),
  ).length;

  const statusCounts = {};
  Config.leadStatuses.forEach((s) => {
    statusCounts[s] = leads.filter((l) => l.status === s).length;
  });

  const agentSales = {};
  todaySales.forEach((l) => {
    if (l.assignedTo)
      agentSales[l.assignedTo] = (agentSales[l.assignedTo] || 0) + 1;
  });
  const top5 = Object.entries(agentSales)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  const recentLeads = leads
    .slice()
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
    .slice(0, 8);

  // ==========================================
  //  THE NEW RENDER LOGIC
  // ==========================================
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = ""; // Clear existing screen

  // 1. Clone the HTML blueprint
  const template = document.getElementById("tmpl-dashboard");
  const clone = template.content.cloneNode(true);

  // 2. Handle Admin Security (Rip out admin elements if they are an agent)
  if (!isAdmin()) {
    clone.querySelectorAll(".admin-only").forEach((el) => el.remove());
  }

  // 3. Populate Header & KPIs
  clone.getElementById("dash-subtitle").textContent =
    `${isAdmin() ? "// ADMIN VIEW" : "// AGENT VIEW"} · v${Config.rules.appVersion}`;
  clone.getElementById("kpi-total").textContent = total;
  clone.getElementById("kpi-active-sub").textContent =
    `${active} active in pipeline`;
  clone.getElementById("kpi-sold-today").textContent = todaySales.length;
  clone.getElementById("kpi-close-rate").textContent =
    `${total ? Math.round((sold / total) * 100) : 0}% all-time close rate`;
  clone.getElementById("kpi-cooloff").textContent = coolOff;
  clone.getElementById("kpi-cooloff-sub").textContent =
    `${Config.rules.coolOffDays}-day rule active`;

  // Apply conditional styling to KPI cards
  const coolOffCard = clone.getElementById("kpi-cooloff-card");
  coolOffCard.className = `kpi-card ${coolOff > 0 ? "kpi-info" : "kpi-neutral"}`;

  if (isAdmin()) {
    const recycleCard = clone.getElementById("kpi-recycle-card");
    clone.getElementById("kpi-recycle-count").textContent = needRecycle;
    clone.getElementById("kpi-recycle-sub").textContent =
      needRecycle > 0 ? "↓ See recycle queue below" : "All leads current";
    recycleCard.className = `kpi-card admin-only ${needRecycle > 0 ? "kpi-warn" : "kpi-neutral"}`;
    if (needRecycle > 0) {
      recycleCard.style.cursor = "pointer";
      recycleCard.onclick = () =>
        document
          .getElementById("dash-recycle-section")
          .scrollIntoView({ behavior: "smooth" });
    }
  }

  // 4. Inject Dynamic Lists (Much smaller innerHTML blocks now!)
  if (isAdmin()) {
    clone.getElementById("dash-pipeline-status").innerHTML = Config.leadStatuses
      .map((s) => {
        const cls =
          "status-" +
          s
            .toLowerCase()
            .replace(/\s+/g, "-")
            .replace(/[^a-z0-9-]/g, "");
        return `<div class="status-row">
                <span class="status-badge ${cls}">${s}</span>
                <div class="status-bar-wrap"><div class="status-bar" style="width:${total ? (statusCounts[s] / total) * 100 : 0}%"></div></div>
                <span class="status-count">${statusCounts[s]}</span>
              </div>`;
      })
      .join("");
  }

  clone.getElementById("dash-top5").innerHTML = top5.length
    ? top5
        .map(
          (e, i) => `
      <div class="top5-row">
        <span class="top5-rank rank-${i + 1}">${i + 1}</span>
        <span class="top5-name">${escHtml(e[0])}</span>
        <span class="top5-count">${e[1]} sale${e[1] !== 1 ? "s" : ""}</span>
      </div>`,
        )
        .join("")
    : `<p class="empty-state">No sales yet today.</p>`;

  // 5. Inject the Heavy Tables
  clone
    .getElementById("dash-recent-table")
    .replaceChildren(renderLeadsTable(recentLeads, true));

  if (isAdmin() && needRecycle > 0) {
    clone.getElementById("dash-recycle-title").textContent =
      `⚠️ Recycle Queue — ${needRecycle} lead${needRecycle !== 1 ? "s" : ""} ready`;

    // 1. Filter out only the leads that need recycling
    const recycleQueue = leads.filter(
      (l) => l.flags && l.flags.includes("needs_recycle"),
    );

    // 2. Build the table HTML
    clone.getElementById("dash-recycle-table").innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            <th>Name</th><th>Address</th><th>Previously Assigned To</th><th>Last Contacted</th><th>Action</th>
          </tr>
        </thead>
        <tbody>
          ${recycleQueue
            .map(
              (l) => `
            <tr>
              <td><span class="lead-name">${escHtml(l.name)}</span></td>
              <td class="td-mono" style="font-size:11px">${escHtml(l.address || "—")}${l.city ? ", " + escHtml(l.city) : ""}</td>
              <td>
                <div style="display:flex;flex-direction:column;gap:2px">
                  ${l.assignedTo ? `<span style="font-size:13px;font-weight:600;color:#1A2640">${escHtml(l.assignedTo)}</span>` : "—"}
                  ${l.previousAgents ? `<span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8">Previously: ${escHtml(l.previousAgents)}</span>` : ""}
                </div>
              </td>
              <td class="td-mono">${formatDate(l.lastContacted) || "—"}</td>
              <td>
                <button class="btn-primary" style="padding:6px 14px;font-size:12px" onclick="recycleLeadAction('${l.id}','${escHtml(l.assignedTo || "")}','${escHtml(l.name)}')">
                  Recycle
                </button>
              </td>
            </tr>
          `,
            )
            .join("")}
        </tbody>
      </table>
    `;
  } else if (isAdmin()) {
    // Hide the whole section if no leads to recycle
    clone.getElementById("dash-recycle-section").style.display = "none";
  }
  // 6. Mount it to the screen!
  mainContent.appendChild(clone);

  updateBadges();
  startSalesFeedPolling();
}

async function recycleLeadAction(leadId, currentAgent, leadName) {
  if (
    !confirm(
      'Recycle "' +
        leadName +
        '"?\n\nThis will:\n• Reset status to New\n• Unassign from ' +
        (currentAgent || "current agent") +
        "\n• Record previous assignment history\n\nThe lead can then be reassigned to a different agent.",
    )
  )
    return;
  setLoading(true);
  try {
    await Graph.recycleLead(leadId, currentAgent);
    await Graph.logActivity({
      LeadID: leadId,
      Title: leadName,
      ActionType: "Recycled",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
      Notes:
        "Recycled by admin — previous agent: " + (currentAgent || "unknown"),
    });
    UI.showToast(leadName + " recycled and ready to reassign!", "success");
    await loadAllData();
    renderDashboard();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function recycleAllLeads() {
  const recycleLeads = State.leads.filter(function (l) {
    return l.flags && l.flags.includes("needs_recycle");
  });
  if (!recycleLeads.length) {
    UI.showToast("No leads to recycle.", "info");
    return;
  }
  if (
    !confirm(
      "Recycle ALL " +
        recycleLeads.length +
        " lead" +
        (recycleLeads.length !== 1 ? "s" : "") +
        " in the queue?\n\nThis will:\n• Reset all their statuses to New\n• Unassign them from their current agents\n• Record previous assignment history\n\nThis cannot be undone.",
    )
  )
    return;
  setLoading(true);
  try {
    for (var i = 0; i < recycleLeads.length; i++) {
      const lead = recycleLeads[i];
      await Graph.recycleLead(lead.id, lead.assignedTo || "");
      await Graph.logActivity({
        LeadID: lead.id,
        Title: lead.name,
        ActionType: "Recycled",
        AgentEmail: (State.currentUser && State.currentUser.email) || "",
        Notes:
          "Bulk recycled by admin — previous agent: " +
          (lead.assignedTo || "unknown"),
      });
    }
    UI.showToast(
      "Recycled " +
        recycleLeads.length +
        " lead" +
        (recycleLeads.length !== 1 ? "s" : "") +
        " successfully!",
      "success",
    );
    await loadAllData();
    renderDashboard();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function startSalesFeedPolling() {
  if (State.salesFeedTimer) clearInterval(State.salesFeedTimer);

  let knownSaleIds = new Set(
    (State.todaySales || []).map(function (l) {
      return l.id;
    }),
  );

  // 1. Wrap the entire fetch and render logic in a named inner function
  async function pollSalesData() {
    try {
      const newSales = await Graph.getTodaySales();

      const newOnes = newSales.filter(function (l) {
        return !knownSaleIds.has(l.id);
      });

      if (newOnes.length) {
        Ticker.update();
        UI.showConfetti();
        newOnes.forEach(function (l) {
          knownSaleIds.add(l.id);
        });
      }

      State.todaySales = newSales;

      if (State.currentView === "dashboard") {
        const feed = document.getElementById("dash-sales-feed");
        const time = document.getElementById("sales-feed-time");

        if (!feed) return;

        if (time)
          time.textContent = "Updated " + formatTime(new Date().toISOString());

        if (!newSales || !newSales.length) {
          feed.innerHTML = `<p class="empty-state" style="padding:24px; text-align:center;">No sales yet today.</p>`;
          return;
        }

        const nameLookup = {};
        (State.contractors || []).forEach((c) => {
          if (c.email) nameLookup[c.email.toLowerCase().trim()] = c.name;
        });

        function formatAgentName(rawString) {
          if (!rawString) return "Unassigned";
          const lower = rawString.toLowerCase().trim();
          return nameLookup[lower] || rawString;
        }

        feed.innerHTML = [...newSales]
          .sort(function (a, b) {
            return new Date(b.modified) - new Date(a.modified);
          })
          .slice(0, 6)
          .map(function (l) {
            const displayAgent = formatAgentName(l.soldBy || l.assignedTo);
            return `
          <div class="sale-entry">
            <div class="sale-icon">&#127881;</div>
            <div class="sale-info">
              <span class="sale-name">${escHtml(l.name)}</span>
              <span class="sale-agent">${escHtml(displayAgent)}</span>
            </div>
            <span class="sale-time">${formatTime(l.modified)}</span>
          </div>`;
          })
          .join("");
      }
    } catch (e) {
      console.error("Sales feed polling error:", e);
    }
  }

  // 2. THE FIX: Run it immediately right now, THEN start the interval timer
  pollSalesData();
  State.salesFeedTimer = setInterval(pollSalesData, Config.salesFeedInterval);
}

// ============================================================
//  ADMIN — DRIP FEED
// ============================================================
function renderDripFeed() {
  // 1. Security & Data Prep (Kept exactly the same)
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  const unassigned = State.leads.filter(
    (l) => !l.assignedTo && !Config.terminalStatuses.includes(l.status),
  );

  if (!State.dripLead && unassigned.length) {
    State.dripLead = unassigned[0];
  }

  const lead = State.dripLead;
  const remaining = unassigned.length;

  // 2. Setup Template
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  const template = document.getElementById("tmpl-drip-feed");
  const clone = template.content.cloneNode(true);

  clone.getElementById("drip-subtitle").textContent =
    `// ASSIGN ONE LEAD AT A TIME · ${remaining} unassigned`;
  if (!lead) {
    clone.getElementById("drip-empty-state").style.display = "block";
    clone.getElementById("drip-header-skip").style.display = "none";
  } else {
    clone.getElementById("drip-active-state").style.display = "block";

    const typeBadge = clone.getElementById("drip-lead-type");
    if (lead.leadType) {
      typeBadge.textContent = lead.leadType;
      typeBadge.className = `lead-type-badge lead-type-${lead.leadType.toLowerCase()}`;
    } else {
      typeBadge.style.display = "none";
    }

    const statusBadge = clone.getElementById("drip-lead-status");
    statusBadge.textContent = lead.status;
    statusBadge.className = `status-badge status-${lead.status
      .toLowerCase()
      .replace(/\s+/g, "-")
      .replace(/[^a-z0-9-]/g, "")}`;

    // Build Text Fields
    clone.getElementById("drip-lead-name").textContent = lead.name;

    const notesEl = clone.getElementById("drip-notes");
    if (lead.notes) notesEl.textContent = lead.notes;
    else notesEl.style.display = "none";

    // Build Meta Icons (Phone, Email, etc.)
    let metaHtml = "";
    if (lead.phone)
      metaHtml += `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 12 19.79 19.79 0 0 1 1.61 3.38 2 2 0 0 1 3.6 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L7.91 8.96a16 16 0 0 0 6 6l.92-.92a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 21.73 16.92z" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.phone)}</span>`;
    if (lead.email)
      metaHtml += `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" stroke="currentColor" stroke-width="2"/><polyline points="22,6 12,13 2,6" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.email)}</span>`;
    if (lead.currentMRC)
      metaHtml += `<span class="feed-meta">MRC: $${escHtml(lead.currentMRC)}/mo</span>`;
    if (lead.currentProducts)
      metaHtml += `<span class="feed-meta">Has: ${escHtml(lead.currentProducts)}</span>`;
    clone.getElementById("drip-meta-container").innerHTML = metaHtml;

    // Build Agent Dropdown
    const selectEl = clone.getElementById("drip-agent-select");
    let optionsHtml = `<option value="">Select an agent...</option>`;
    State.contractors.forEach((c) => {
      const count = State.leads.filter(
        (l) =>
          l.assignedTo === c.name &&
          !Config.terminalStatuses.includes(l.status),
      ).length;
      const full = count >= Config.rules.maxLeadsPerAgent;
      optionsHtml += `<option value="${escHtml(c.name)}" ${full ? "disabled" : ""}>${escHtml(c.name)} — ${count}/${Config.rules.maxLeadsPerAgent}${full ? " (FULL)" : ""}</option>`;
    });
    selectEl.innerHTML = optionsHtml;

    // Wire up the dynamic ID to the assign button
    clone.getElementById("drip-assign-btn").onclick = () =>
      confirmDripAssign(lead.id);

    // Build Remaining Table
    clone.getElementById("drip-remaining-title").textContent =
      `Remaining Unassigned (${remaining})`;
    clone
      .getElementById("drip-remaining-table")
      .replaceChildren(renderLeadsTable(unassigned.slice(0, 10), true));
  }

  // 4. Mount
  mainContent.appendChild(clone);
}

async function confirmDripAssign(leadId) {
  const select = document.getElementById("drip-agent-select");
  const agent = select && select.value;
  if (!agent) {
    UI.showToast("Please select an agent first.", "error");
    return;
  }
  if (!Graph.canAgentTakeLead(agent, State.leads)) {
    UI.showToast(
      agent + " is at the " + Config.rules.maxLeadsPerAgent + "-lead limit.",
      "error",
    );
    return;
  }
  const lead = State.leads.find(function (l) {
    return l.id === leadId;
  });
  setLoading(true);
  try {
    await Graph.assignAgent(leadId, agent);
    await Graph.logActivity({
      LeadID: leadId,
      Title: lead ? lead.name : "",
      ActionType: "Drip Assigned",
      AgentEmail: agent,
      Notes:
        "Drip-assigned by " +
        ((State.currentUser && State.currentUser.name) || "Admin"),
    });
    UI.showToast(lead.name + " assigned to " + agent, "success");
    await loadAllData();
    const remaining = State.leads.filter(function (l) {
      return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
    });
    State.dripLead = remaining.length ? remaining[0] : null;
    renderDripFeed();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function skipDripLead() {
  const unassigned = State.leads.filter(function (l) {
    return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
  });
  const currentIdx = unassigned.findIndex(function (l) {
    return State.dripLead && l.id === State.dripLead.id;
  });
  const nextIdx = (currentIdx + 1) % unassigned.length;
  State.dripLead = unassigned[nextIdx] || null;
  renderDripFeed();
}

// ============================================================
//  AGENT — MY LEADS
// ============================================================
function getStatusColor(status) {
  const colors = Config.statusColors || {};
  if ((colors.red || []).includes(status)) return "#FF4444";
  if ((colors.yellow || []).includes(status)) return "#FFD700";
  if ((colors.green || []).includes(status)) return "#00FF88";
  if ((colors.blue || []).includes(status)) return "#4D79FF";
  if ((colors.cyan || []).includes(status)) return "#00E5FF";
  if ((colors.white || []).includes(status)) return "#FFFFFF";
  return "#7A98C8";
}

function getStatusDot(status) {
  const color = getStatusColor(status);
  return `<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};box-shadow:0 0 6px ${color};flex-shrink:0;margin-right:6px"></span>`;
}

let _leadSaved = false;
let _currentFeedIndex = 0;

function renderMyLeads() {
  const user = State.currentUser;
  const userName = ((user && user.name) || "").toLowerCase().trim();
  const userEmail = ((user && user.email) || "").toLowerCase().trim();

  const contractor = State.contractors.find((c) => {
    return (
      (c.email || "").toLowerCase().trim() === userEmail ||
      (c.name || "").toLowerCase().trim() === userName
    );
  });
  const agentName = contractor
    ? contractor.name.toLowerCase().trim()
    : userName;

  // We only run this filter loop ONCE now!
  const myLeads = State.leads.filter((l) => {
    const assigned = (l.assignedTo || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
    return (
      assigned &&
      (assigned === agentName.replace(/\s+/g, " ") ||
        assigned === userName.replace(/\s+/g, " ") ||
        assigned === userEmail.replace(/\s+/g, " ")) &&
      !Config.terminalStatuses.includes(l.status)
    );
  });

  // Keep his global variables intact so we don't break the rest of his code
  window._myLeads = myLeads;
  window._agentName = agentName;
  _leadSaved = false;

  if (_currentFeedIndex >= myLeads.length) _currentFeedIndex = 0;

  if (!window._forceShowLead) {
    while (
      _currentFeedIndex < myLeads.length &&
      Graph.isInCoolOff(myLeads[_currentFeedIndex])
    ) {
      _currentFeedIndex++;
    }
  }

  // Instantly reset the flag so the "Next Lead" button goes back to skipping cool-offs normally
  window._forceShowLead = false;

  const contactsToday = Graph.agentContactsToday(
    (user && user.email) || "",
    State.activityLog,
  );
  const atLimit = contactsToday >= Config.rules.maxContactsPerDay;

  // ==========================================
  //  THE NEW RENDER LOGIC
  // ==========================================
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  const template = document.getElementById("tmpl-my-leads");
  const clone = template.content.cloneNode(true);

  // 1. Handle Admin Security (Removes the search box if they are an agent)
  if (!isAdmin()) {
    clone.querySelectorAll(".admin-only").forEach((el) => el.remove());
  }

  // 2. Populate Header
  clone.getElementById("myleads-subtitle").textContent =
    `// ${myLeads.length} remaining · lead ${Math.min(_currentFeedIndex + 1, myLeads.length || 1)} of ${myLeads.length}`;
  clone.getElementById("myleads-contact-text").textContent =
    `${contactsToday}/${Config.rules.maxContactsPerDay} contacts today`;

  if (atLimit) {
    clone.getElementById("myleads-contact-badge").classList.add("badge-full");
    clone.getElementById("myleads-limit-alert").style.display = "flex";
    clone.getElementById("myleads-limit-text").textContent =
      `Daily limit reached — ${Config.rules.maxContactsPerDay} contacts today. Great work!`;
  }

  // 3. Inject the active lead card
  // (Assuming renderLeadFeedCard still returns an HTML string for now)
  clone.getElementById("lead-feed-wrap").innerHTML = ""; // Clear it first
  clone
    .getElementById("lead-feed-wrap")
    .appendChild(renderLeadFeedCard(myLeads, contactsToday));

  // 4. Mount
  mainContent.appendChild(clone);

  // ==========================================
  //  5. LIVE CLOCK LOGIC (Smart Timezones)
  // ==========================================
  const clockEl = document.getElementById("myleads-clock");

  // 1. Figure out which lead they are currently looking at
  const activeLead = myLeads[_currentFeedIndex];
  console.log("Here is the full Lead Object:", activeLead);
  // 2. Extract the state (if it exists)
  const leadState =
    activeLead && activeLead.state
      ? activeLead.state.toUpperCase().trim()
      : null;
  console.log(leadState);

  const updateClock = () => {
    if (!clockEl) return;

    // Default to the Agent's local computer time
    let tz = Intl.DateTimeFormat().resolvedOptions().timeZone;

    // If the lead has a state, and we have it in our dictionary, override the timezone!
    if (leadState && stateTimezones[leadState]) {
      tz = stateTimezones[leadState];
    }

    try {
      // Now including seconds, and passing our dynamic timezone!
      clockEl.textContent = new Date().toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit", // Added seconds
        timeZone: tz,
        timeZoneName: "short",
      });
    } catch (e) {
      // Safe fallback if something goes weird
      clockEl.textContent = new Date().toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
    }
  };
  // Run it once immediately so it doesn't say "--:--" for the first second
  updateClock();

  // Clear any existing timer from previous visits to this page
  if (window._clockTimer) clearInterval(window._clockTimer);

  // Start a new timer that ticks every 1 second (1000ms)
  window._clockTimer = setInterval(updateClock, 1000);
}

function searchMyLeads(q) {
  const leads = window._myLeads || [];
  const wrap = document.getElementById("my-leads-table");
  if (!wrap) return;

  if (!q.trim()) {
    wrap.innerHTML = "";
    return;
  }

  const filtered = leads.filter((l) => {
    return (
      l.name.toLowerCase().includes(q.toLowerCase()) ||
      (l.phone || "").includes(q) ||
      (l.btn || "").includes(q) ||
      (l.cbr || "").includes(q) ||
      (l.address || "").toLowerCase().includes(q.toLowerCase())
    );
  });

  if (filtered.length) {
    wrap.replaceChildren(renderLeadsTable(filtered, false, true));
  } else {
    wrap.innerHTML = `<div class="empty-state">No leads found for "${escHtml(q)}"</div>`;
  }
}

let _stagedStatus = null;

function renderLeadFeedCard(myLeads, contactsToday) {
  // 1. FIXED LOGIC: Grab the exact lead we are supposed to be looking at!
  let lead = myLeads[_currentFeedIndex];

  // If that exact lead happens to be in cool-off, we should probably warn the logic,
  // but we still want to show them the correct lead!
  const isCoolOff = lead ? Graph.isInCoolOff(lead) : false;

  const atLimit = contactsToday >= Config.rules.maxContactsPerDay;
  _stagedStatus = null;

  // 2. Setup Template
  const template = document.getElementById("tmpl-lead-feed-card");
  const clone = template.content.cloneNode(true);

  const emptyState = clone.getElementById("feed-card-empty");
  const activeState = clone.getElementById("feed-card-active");

  // 3. Handle Empty State
  if (!lead) {
    emptyState.style.display = "flex"; // Show empty state
    clone.getElementById("feed-empty-text").textContent =
      myLeads.length > 0
        ? "Remaining leads are in the cool-off period."
        : "No leads assigned yet — ask your manager.";

    // We create a wrapper div to hold the clone and return it
    const wrapper = document.createElement("div");
    wrapper.appendChild(clone);
    return wrapper;
  }

  // 4. Handle Active Lead State
  activeState.style.display = "block"; // Show active form

  // Badges & Name
  const typeBadge = clone.getElementById("feed-lead-type");
  if (lead.leadType) {
    typeBadge.textContent = lead.leadType;
    typeBadge.className = `lead-type-badge lead-type-${lead.leadType.toLowerCase()}`;
  } else {
    typeBadge.style.display = "none";
  }

  const statusBadge = clone.getElementById("feed-current-status");
  statusBadge.textContent = lead.status;
  statusBadge.className = `status-badge status-${lead.status
    .toLowerCase()
    .replace(/\s+/g, "-")
    .replace(/[^a-z0-9-]/g, "")}`;

  clone.getElementById("feed-lead-name").textContent = lead.name;

  // Cooloff Alert
  if (isCoolOff) {
    const alert = clone.getElementById("feed-cooloff-alert");
    alert.style.display = "block";
    alert.textContent = `⏱ This lead is in the ${Config.rules.coolOffDays}-day cool-off period — you can still update it if the customer reached out.`;
  }

  // Meta Row (Icons)
  let metaHtml = "";
  if (lead.phone)
    metaHtml += `<span class="feed-meta">📞 ${escHtml(lead.phone)}</span>`;
  if (lead.email)
    metaHtml += `<span class="feed-meta">✉️ ${escHtml(lead.email)}</span>`;
  if (lead.address)
    metaHtml += `<span class="feed-meta">📍 ${escHtml(lead.address)}${lead.city ? ", " + escHtml(lead.city) : ""}${lead.state ? " " + escHtml(lead.state) : ""}${lead.zip ? " " + escHtml(lead.zip) : ""}</span>`;
  clone.getElementById("feed-meta-container").innerHTML = metaHtml;

  // Form Inputs
  clone.getElementById("feed-btn").value = lead.btn || "";
  clone.getElementById("feed-mrc").value = lead.currentMRC || "";
  clone.getElementById("feed-cbr").value = lead.cbr || "";

  // Products Dropdown
  const productsSelect = clone.getElementById("feed-products");
  Config.currentProducts.forEach((p) => {
    const option = document.createElement("option");
    option.value = p;
    option.textContent = p;
    if (lead.currentProducts === p) option.selected = true;
    productsSelect.appendChild(option);
  });

  // Sold By Dropdown
  const soldBySelect = clone.getElementById("feed-sold-by");
  State.contractors.forEach((c) => {
    const option = document.createElement("option");
    option.value = c.name;
    option.textContent = c.name;
    // You can add logic here to auto-select if needed!
    soldBySelect.appendChild(option);
  });

  // AutoPay Radios
  const autoPayContainer = clone.getElementById("feed-autopay-container");
  ["ACH - Debit Card", "ACH - Credit Card", "No Auto Pay"].forEach((opt) => {
    const isChecked = lead.autoPay === opt ? "checked" : "";
    autoPayContainer.innerHTML += `
      <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;color:#1A2640;background:#F4F7FD;border:1px solid #D0DCF0;padding:8px 14px;border-radius:6px;">
        <input type="radio" name="feed-autopay" value="${opt}" ${isChecked} style="accent-color:#2563B0"> ${opt}
      </label>`;
  });

  // Status Buttons
  const statusContainer = clone.getElementById("feed-status-buttons");
  Config.leadStatuses
    .filter((s) => s !== "New")
    .forEach((s) => {
      const disabled =
        atLimit && !Config.terminalStatuses.includes(s)
          ? "disabled title='Daily limit reached'"
          : "";
      const isTDM = s === "TDM" ? " ↩" : "";
      const cls =
        "status-btn-" +
        s
          .toLowerCase()
          .replace(/\s+/g, "-")
          .replace(/[^a-z0-9-]/g, "");

      // We use innerHTML here purely for brevity of building buttons, it's safe enough!
      statusContainer.innerHTML += `<button class="status-btn ${cls}" id="sbtn-${s.replace(/\s+/g, "-")}" onclick="stageStatus('${lead.id}','${s}')" ${disabled}>${s}${isTDM}</button>`;
    });

  clone.getElementById("feed-today-date").textContent =
    new Date().toLocaleDateString("en-US", {
      month: "2-digit",
      day: "2-digit",
      year: "2-digit",
    });

  // THE RESTORED LEGACY NOTES PARSER
  const pastNotesContainer = clone.getElementById("feed-past-notes-container");
  if (lead.notes && lead.notes.trim()) {
    pastNotesContainer.style.display = "block"; // Unhide the box

    const notesHtml = lead.notes
      .split("\n")
      .filter((l) => l.trim()) // Remove empty lines
      .map((line) => {
        // Look for the [MM/DD/YY - Agent Name] pattern
        const match = line.match(/^\[(\d{2}\/\d{2}(?:\/\d{2})?)(.*?)\]\s*(.*)/);

        if (match) {
          const date = match[1];
          const agent = match[2] ? match[2].replace(/^\s*-\s*/, "") : "";
          const text = match[3];

          return `
            <div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
              <div style="display:flex;gap:8px;align-items:center;margin-bottom:3px">
                <span style="font-family:var(--font-mono);font-size:10px;color:#2563B0;font-weight:700;background:#E8F0FF;padding:1px 6px;border-radius:3px">${date}</span>
                ${agent ? `<span style="font-family:var(--font-mono);font-size:10px;color:#6B85B0">${escHtml(agent)}</span>` : ""}
              </div>
              <span style="font-size:13px;color:#1A2640">${escHtml(text)}</span>
            </div>`;
        }

        // Fallback if the note doesn't match the standard format
        return `
          <div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
            <div style="margin-bottom:3px">
              <span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8;background:#F4F7FD;padding:1px 6px;border-radius:3px">Legacy note — author unknown</span>
            </div>
            <span style="font-size:13px;color:#4A6080">${escHtml(line)}</span>
          </div>`;
      })
      .join("");

    pastNotesContainer.innerHTML = notesHtml;
  }

  // Save Button Action
  clone.getElementById("feed-save-btn").onclick = () => agentSaveAll(lead.id);

  // Wrap and Return
  const wrapper = document.createElement("div");
  wrapper.appendChild(clone);
  return wrapper;
}

function stageStatus(leadId, newStatus) {
  const lead = State.leads.find(function (l) {
    return l.id === leadId;
  });
  if (!lead) return;

  if (Graph.isInCoolOff(lead) && !Config.terminalStatuses.includes(newStatus)) {
    UI.showToast(
      "Note: this lead is in the " +
        Config.rules.coolOffDays +
        "-day cool-off period.",
      "info",
    );
  }

  _stagedStatus = newStatus;

  document.querySelectorAll(".status-btn").forEach(function (btn) {
    btn.style.borderColor = "";
    btn.style.color = "";
    btn.style.background = "";
    btn.style.boxShadow = "";
  });
  const selectedBtn = document.getElementById(
    "sbtn-" + newStatus.replace(/\s+/g, "-"),
  );
  if (selectedBtn) {
    selectedBtn.style.borderColor = "var(--cyan)";
    selectedBtn.style.color = "var(--cyan)";
    selectedBtn.style.background = "var(--cyan-dim)";
    selectedBtn.style.boxShadow = "0 0 12px var(--cyan-glow)";
  }

  const notice = document.getElementById("feed-staged-notice");
  if (notice) {
    notice.style.display = "block";
    notice.textContent =
      '⚡ "' + newStatus + '" staged — click Save to confirm';
  }

  const badge = document.getElementById("feed-current-status");
  if (badge) {
    badge.textContent = newStatus + " (staged)";
    badge.style.opacity = "0.7";
  }
}

async function agentSaveAll(leadId) {
  const user = State.currentUser;
  const lead = State.leads.find((l) => l.id === leadId);
  if (!lead) return;

  const newStatus = _stagedStatus || lead.status;
  const mrc = (document.getElementById("feed-mrc") || {}).value || "";
  const products = (document.getElementById("feed-products") || {}).value || "";
  const newNote = (document.getElementById("feed-notes") || {}).value || "";
  const cbr = (document.getElementById("feed-cbr") || {}).value || "";
  const btn = (document.getElementById("feed-btn") || {}).value || "";
  const autoPayEl = document.querySelector(
    'input[name="feed-autopay"]:checked',
  );
  const autoPay = autoPayEl ? autoPayEl.value : "";
  const soldByEl = document.getElementById("feed-sold-by");
  const soldByName = soldByEl ? soldByEl.value : "";

  // Validation
  if (!autoPay) {
    UI.showToast("Please select an AutoPay option before saving.", "error");
    return;
  }
  if (newStatus === Config.soldStatus && !soldByName) {
    UI.showToast(
      "Please select who made this sale in the Sold By field.",
      "error",
    );
    return;
  }

  // Activity Log Email Resolution
  const soldByContractor = soldByName
    ? State.contractors.find((c) => c.name === soldByName)
    : null;
  const soldByEmail = soldByContractor
    ? soldByContractor.email || soldByName
    : (user && user.email) || "";
  const activityEmail =
    newStatus === Config.soldStatus ? soldByEmail : (user && user.email) || "";

  // Note Formatting
  let notes = lead.notes || "";
  if (newNote.trim()) {
    const today = new Date();
    const dateStamp =
      (today.getMonth() + 1).toString().padStart(2, "0") +
      "/" +
      today.getDate().toString().padStart(2, "0") +
      "/" +
      String(today.getFullYear()).slice(-2);
    const agentTag = user && user.name ? " - " + user.name : "";
    const stamped = "[" + dateStamp + agentTag + "] " + newNote.trim();
    notes = notes ? stamped + "\n" + notes : stamped;
  }

  // Setup Payload
  const todayDate = new Date().toISOString().split("T")[0];
  const saveFields = { Status: newStatus, LastTouchedOn: todayDate };
  if (mrc) saveFields["MonthlyRecurringCharge_x0028_MRC"] = mrc;
  if (products) saveFields["CurrentProducts"] = products;
  if (cbr) saveFields["CBR"] = cbr;
  if (btn) saveFields["BTN"] = btn;
  if (notes) saveFields["Notes"] = notes;
  if (autoPay) saveFields["AutoPay"] = autoPay;

  setLoading(true);
  try {
    // 1. THE TURBO BOOST: Fire all network requests concurrently
    const networkTasks = [
      Graph.updateLead(leadId, saveFields),
      Graph.logActivity({
        LeadID: leadId,
        Title: lead.name,
        ActionType: "Status: " + newStatus,
        AgentEmail: activityEmail,
        Notes:
          notes +
          (newStatus === Config.soldStatus &&
          soldByName &&
          soldByName !== (user && user.name)
            ? " [Sold by " +
              soldByName +
              ", recorded by " +
              ((user && user.name) || "admin") +
              "]"
            : ""),
      }),
    ];

    if (newStatus === "TDM") {
      networkTasks.push(Graph.assignAgent(leadId, ""));
    }

    // Wait for all of them to finish at the exact same time
    await Promise.all(networkTasks);

    // 2. IN-MEMORY UPDATE: Update the local state so we don't need loadAllData()
    lead.status = newStatus;
    lead.notes = notes;
    if (mrc) lead.mrc = mrc; // Assuming your state keys match your variables
    if (products) lead.products = products;
    if (autoPay) lead.autoPay = autoPay;
    if (newStatus === "TDM") lead.assignedTo = "";

    // 3. INSTANT UI UPDATES
    if (newStatus === "TDM") {
      UI.showToast("TDM — lead returned to admin queue.", "info");
    } else {
      UI.showToast("Saved!", "success");
    }

    Ticker.update();
    _stagedStatus = null;
    _leadSaved = true;

    const nextRow = document.getElementById("feed-next-row");
    const searchSec = document.getElementById("lead-search-section");
    const saveBtn = document.getElementById("feed-save-btn");

    if (nextRow) nextRow.style.display = "block";
    if (searchSec) searchSec.style.display = "block";
    if (saveBtn) {
      saveBtn.textContent = "Saved ✓";
      saveBtn.disabled = true;
      saveBtn.style.background = "var(--green)";
      setTimeout(() => {
        saveBtn.textContent = "Save";
        saveBtn.disabled = false;
        saveBtn.style.background = "";
      }, 2000);
    }
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
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
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  const { leads, contractors } = State;
  const unassigned = leads.filter(function (l) {
    const isValidLead = l && l.id && (l.name || l.phone);
    const isAvailable =
      !l.assignedTo && !Config.terminalStatuses.includes(l.status);
    return isValidLead && isAvailable;
  });

  const agentCounts = {};
  contractors.forEach((c) => (agentCounts[c.name] = 0));
  leads.forEach((l) => {
    if (
      l.assignedTo &&
      !Config.terminalStatuses.includes(l.status) &&
      agentCounts[l.assignedTo] !== undefined
    ) {
      agentCounts[l.assignedTo]++;
    }
  });

  const uniqueStates = [
    ...new Set(
      unassigned
        .map((l) => (l.state || "").trim().toUpperCase())
        .filter(Boolean),
    ),
  ].sort();

  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Assign Leads</h1>
        <span class="view-subtitle">// ${unassigned.length} total unassigned</span>
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
        <h2 class="card-title" style="color:#0D1B3E">Bulk Assign to Agent</h2>
        <span class="card-meta">Select an agent, lead type, state, and quantity</span>
      </div>
      <div style="padding:16px 20px; display:flex; gap:12px; align-items:center; flex-wrap:wrap;">
        
        <select id="bulk-agent-select" class="form-input" style="min-width: 180px; flex: 1;">
          <option value="">Select Agent...</option>
          ${contractors.map((c) => `<option value="${escHtml(c.name)}">${escHtml(c.name)} (${agentCounts[c.name]} assigned)</option>`).join("")}
        </select>

        <div style="display:flex; align-items:center; gap:8px;">
          <select id="bulk-type-select" class="form-input" style="width: 120px; padding-right: 24px;">
            <option value="all">Any Type</option>
            ${(Config.leadTypes || []).map((t) => `<option value="${escHtml(t)}">${escHtml(t)}</option>`).join("")}
          </select>

          <select id="bulk-state-select" class="form-input" style="width: 120px; padding-right: 24px;">
            <option value="all">Any State</option>
            ${uniqueStates.map((s) => `<option value="${escHtml(s)}">${escHtml(s)}</option>`).join("")}
          </select>

          <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#0D1B3E; cursor:pointer; margin-left:4px; white-space:nowrap;">
            <input type="checkbox" id="bulk-unworked-check" style="cursor:pointer; width:15px; height:15px;">
            Unworked Only
          </label>
        </div>

        <div style="display:flex; align-items:center; gap:8px;">
          <input type="number" id="bulk-agent-qty" class="form-input" min="1" max="${unassigned.length}" placeholder="Qty" style="width: 75px;">
          <span id="bulk-type-count" style="font-size: 13px; color: #6B85B0; white-space: nowrap;">of ${unassigned.length} available</span>
        </div>

        <button class="btn-primary" onclick="bulkAssignToSelectedAgent()" style="white-space: nowrap;">
          Assign
        </button>
      </div>
    </div>

    <div class="card">
      <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
        <div>
          <h2 class="card-title">Unassigned Leads Preview</h2>
          <span class="card-meta" id="table-meta-count">Loading...</span>
        </div>
        
        <div style="display:flex; gap:8px; align-items:center;">
          <button id="btn-prev-page" class="btn-secondary" style="padding: 4px 10px;">&larr; Prev</button>
          <span id="page-indicator" style="font-family:var(--font-mono); font-size: 13px; font-weight: 600; color: #0D1B3E;">Page 1</span>
          <button id="btn-next-page" class="btn-secondary" style="padding: 4px 10px;">Next &rarr;</button>
        </div>
      </div>
      
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Name</th><th>Type</th><th>Phone</th><th>Status</th><th style="text-align: right;">Assign To</th></tr></thead>
          <tbody id="assign-tbody">
            </tbody>
        </table>
      </div>
    </div>
  `;

  // 3. Internal State & DOM Pointers
  let currentPage = 1;
  const itemsPerPage = 25; // Change this if you ever want more/less per page!
  const unworkedCheck = document.getElementById("bulk-unworked-check");
  const typeSelect = document.getElementById("bulk-type-select");
  const stateSelect = document.getElementById("bulk-state-select");
  const qtyInput = document.getElementById("bulk-agent-qty");
  const countDisplay = document.getElementById("bulk-type-count");

  const tbody = document.getElementById("assign-tbody");
  const prevBtn = document.getElementById("btn-prev-page");
  const nextBtn = document.getElementById("btn-next-page");
  const pageIndicator = document.getElementById("page-indicator");
  const tableMetaCount = document.getElementById("table-meta-count");

  // 4. The Smart Table Renderer
  function updateTableAndMath() {
    const selectedType = typeSelect ? typeSelect.value : "all";
    const selectedState = stateSelect ? stateSelect.value : "all";
    const requireUnworked = unworkedCheck ? unworkedCheck.checked : false;

    // Filter master pool based on dropdowns AND checkbox
    const filteredLeads = unassigned.filter(function (l) {
      const typeMatch =
        selectedType === "all" ||
        (l.leadType && l.leadType.toLowerCase() === selectedType.toLowerCase());
      const stateMatch =
        selectedState === "all" ||
        (l.state && l.state.toUpperCase() === selectedState.toUpperCase());

      // 3. THE NEW RULE: If checked, lead must have NO previous agents
      const unworkedMatch = !requireUnworked || !l.previousAgents;

      return typeMatch && stateMatch && unworkedMatch;
    });

    // Update Bulk Assign max math
    const total = filteredLeads.length;
    if (countDisplay) countDisplay.textContent = `of ${total} available`;
    if (qtyInput) {
      qtyInput.max = total;
      if (parseInt(qtyInput.value, 10) > total) qtyInput.value = total;
    }

    // Calculate Pagination Math
    const totalPages = Math.max(1, Math.ceil(total / itemsPerPage));
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    // Update UI text and button disabled states
    pageIndicator.textContent = `Page ${currentPage} of ${totalPages}`;
    tableMetaCount.textContent = `Showing ${total} leads matching filters`;

    prevBtn.disabled = currentPage === 1;
    prevBtn.style.opacity = currentPage === 1 ? "0.4" : "1";

    nextBtn.disabled = currentPage === totalPages;
    nextBtn.style.opacity = currentPage === totalPages ? "0.4" : "1";

    // Slice for the current page and Draw HTML
    const startIndex = (currentPage - 1) * itemsPerPage;
    const displayLeads = filteredLeads.slice(
      startIndex,
      startIndex + itemsPerPage,
    );

    if (displayLeads.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-state">No unassigned leads match these filters!</td></tr>`;
      return;
    }

    tbody.innerHTML = displayLeads
      .map(function (lead) {
        return `
        <tr>
          <td><span class="lead-name">${escHtml(lead.name)}</span></td>
          
          <td>${lead.leadType ? `<span class="lead-type-badge lead-type-${(lead.leadType || "").toLowerCase()}">${escHtml(lead.leadType)}</span>` : "—"}</td>
          <td class="td-mono">${escHtml(lead.phone)}</td>
          <td><span class="status-badge status-${lead.status
            .toLowerCase()
            .replace(/\s+/g, "-")
            .replace(/[^a-z0-9-]/g, "")}">${lead.status}</span></td>
          
          <td>
            <div class="assign-select-row" style="display:flex; gap:6px; align-items:center; justify-content: flex-end;">
              <select class="filter-select assign-select" id="assign-${lead.id}">
                <option value="">Select agent</option>
                ${contractors.map((c) => `<option value="${escHtml(c.name)}">${escHtml(c.name)} (${agentCounts[c.name]} assigned)</option>`).join("")}
              </select>
              
              <button class="btn-primary" style="padding:6px 14px;font-size:12px" onclick="assignLead('${lead.id}')">Assign</button>
              
              <button class="btn-secondary" style="padding:6px 14px;font-size:12px" 
              onclick="renderLeadModal(State.leads.find(l => l.id === '${lead.id}'))">
                View
              </button>
            </div>
          </td>
        </tr>`;
      })
      .join("");
  }

  // 5. Attach Event Listeners
  if (unworkedCheck)
    unworkedCheck.addEventListener("change", () => {
      currentPage = 1;
      updateTableAndMath();
    });
  if (typeSelect)
    typeSelect.addEventListener("change", () => {
      currentPage = 1;
      updateTableAndMath();
    });
  if (stateSelect)
    stateSelect.addEventListener("change", () => {
      currentPage = 1;
      updateTableAndMath();
    });

  if (prevBtn)
    prevBtn.addEventListener("click", () => {
      if (currentPage > 1) {
        currentPage--;
        updateTableAndMath();
      }
    });
  if (nextBtn)
    nextBtn.addEventListener("click", () => {
      currentPage++;
      updateTableAndMath();
    });

  // 6. Initialize the math & table on first load
  updateTableAndMath();
}

async function assignLead(leadId) {
  const select = document.getElementById("assign-" + leadId);
  const agent = select && select.value;
  if (!agent) {
    UI.showToast("Please select an agent.", "error");
    return;
  }
  if (!Graph.canAgentTakeLead(agent, State.leads)) {
    UI.showToast(agent + " is at the lead limit.", "error");
    return;
  }
  const lead = State.leads.find(function (l) {
    return l.id === leadId;
  });
  setLoading(true);
  try {
    await Graph.assignAgent(leadId, agent);
    await Graph.logActivity({
      LeadID: leadId,
      Title: lead ? lead.name : "",
      ActionType: "Assigned",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
      Notes:
        "Assigned by " +
        ((State.currentUser && State.currentUser.name) || "Admin"),
    });
    UI.showToast("Assigned to " + agent, "success");
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssignToSelectedAgent() {
  const agentName = document.getElementById("bulk-agent-select").value;
  const qty = parseInt(document.getElementById("bulk-agent-qty").value, 10);
  const selectedType = document.getElementById("bulk-type-select").value;
  const selectedState = document.getElementById("bulk-state-select").value;

  if (!agentName) {
    UI.showToast("Please select an agent first.", "warning");
    return;
  }
  if (!qty || qty <= 0) {
    UI.showToast("Please enter a valid number of leads.", "warning");
    return;
  }

  // 1. Filter by unassigned, terminal status, type, AND state
  const unassigned = State.leads.filter(function (l) {
    const isAvailable =
      !l.assignedTo && !Config.terminalStatuses.includes(l.status);

    const matchesType =
      selectedType === "all" ||
      (l.leadType && l.leadType.toLowerCase() === selectedType.toLowerCase());

    const matchesState =
      selectedState === "all" ||
      (l.state && l.state.toUpperCase() === selectedState.toUpperCase());

    return isAvailable && matchesType && matchesState;
  });

  // 2. Filter out leads the agent has previously worked
  const validLeads = unassigned.filter(function (l) {
    const prevAgents = (l.previousAgents || "").toLowerCase();
    return !prevAgents.includes(agentName.toLowerCase());
  });

  // Dynamic labels for the Toast notifications
  const stateLabel = selectedState === "all" ? "" : `${selectedState} `;
  const typeLabel = selectedType === "all" ? "fresh" : selectedType;
  const combinedLabel = `${stateLabel}${typeLabel}`.trim();

  // 3. Validation Warnings
  if (validLeads.length === 0) {
    UI.showToast(
      `${agentName} has no available ${combinedLabel} leads left to work!`,
      "warning",
    );
    return;
  }

  if (qty > validLeads.length) {
    UI.showToast(
      `Only ${validLeads.length} ${combinedLabel} leads available for ${agentName}.`,
      "warning",
    );
    return;
  }

  const leadsToAssign = validLeads.slice(0, qty);

  setLoading(true);
  try {
    await Promise.all(
      leadsToAssign.map(async (lead) => {
        await Graph.updateLead(lead.id, {
          Agent_x0020_Assigned: agentName,
        });
        lead.assignedTo = agentName;
      }),
    );

    UI.showToast(
      `Successfully assigned ${qty} ${combinedLabel} leads to ${agentName}!`,
      "success",
    );
    renderAssignLeads();
  } catch (err) {
    console.error("Bulk Assign Error:", err);
    UI.showToast("Failed to assign leads: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssignByQuantity() {
  const { leads, contractors } = State;
  const unassigned = leads.filter(function (l) {
    return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
  });

  const plan = [];
  contractors.forEach(function (c) {
    const qty =
      parseInt((document.getElementById("qty-" + c.name) || {}).value || "0") ||
      0;
    if (qty > 0) plan.push({ agent: c.name, qty: qty });
  });

  if (!plan.length) {
    UI.showToast("Please enter a quantity for at least one agent.", "error");
    return;
  }

  const totalRequested = plan.reduce(function (s, p) {
    return s + p.qty;
  }, 0);
  if (totalRequested > unassigned.length) {
    UI.showToast(
      "Total (" +
        totalRequested +
        ") exceeds unassigned leads (" +
        unassigned.length +
        "). Reduce quantities.",
      "error",
    );
    return;
  }

  const summary = plan
    .map(function (p) {
      return p.qty + " → " + p.agent;
    })
    .join("\n");
  if (
    !confirm(
      "Assign leads by quantity?\n\n" +
        summary +
        "\n\nTotal: " +
        totalRequested +
        " leads",
    )
  )
    return;

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
    UI.showToast(
      "Assigned " + totalRequested + " leads successfully!",
      "success",
    );
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  LEADS VIEW (Admin only)
// ============================================================
function renderLeads() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  // 1. Security & Data Prep
  State.selectedLeads.clear();
  const contractors = State.contractors.map((c) => c.name);

  // 2. Setup Template
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  const template = document.getElementById("tmpl-all-leads");
  const clone = template.content.cloneNode(true);

  // 3. Header
  clone.getElementById("leads-subtitle").textContent =
    `// ${State.leads.length} total`;

  // 4. Populate Bulk Bar Dropdowns
  const bulkAssignSelect = clone.getElementById("bulk-assign-select");
  contractors.forEach((c) => {
    const option = document.createElement("option");
    option.value = c;
    option.textContent = c;
    bulkAssignSelect.appendChild(option);
  });

  const bulkTypeSelect = clone.getElementById("bulk-type-select");
  Config.leadTypes.forEach((t) => {
    const option = document.createElement("option");
    option.value = t;
    option.textContent = t;
    bulkTypeSelect.appendChild(option);
  });

  // 5. Populate Filters (and restore any previous search state)
  clone.getElementById("search-input").value = State.filters.search || "";

  const statusFilter = clone.getElementById("filter-status");
  Config.leadStatuses.forEach((s) => {
    const option = document.createElement("option");
    option.value = s;
    option.textContent = s;
    if (State.filters.status === s) option.selected = true; // Restore state
    statusFilter.appendChild(option);
  });

  const agentFilter = clone.getElementById("filter-agent");
  contractors.forEach((c) => {
    const option = document.createElement("option");
    option.value = c;
    option.textContent = c;
    if (State.filters.assignedTo === c) option.selected = true; // Restore state
    agentFilter.appendChild(option);
  });

  // 6. Draw the Table
  clone
    .getElementById("leads-table-wrap")
    .replaceChildren(renderLeadsTable(getFilteredLeads()));

  // 7. Mount!
  mainContent.appendChild(clone);
}

function toggleLeadSelect(id, checked) {
  if (checked) {
    State.selectedLeads.add(id);
  } else {
    State.selectedLeads.delete(id);
  }
  updateBulkBar();
}

function toggleSelectAll(checked) {
  const checkboxes = document.querySelectorAll(".lead-checkbox");
  checkboxes.forEach(function (cb) {
    cb.checked = checked;
    if (checked) {
      State.selectedLeads.add(cb.dataset.id);
    } else {
      State.selectedLeads.delete(cb.dataset.id);
    }
  });
  updateBulkBar();
}

function updateBulkBar() {
  const bar = document.getElementById("bulk-bar");
  const count = document.getElementById("bulk-count");
  const n = State.selectedLeads.size;
  if (!bar) return;
  bar.style.display = n > 0 ? "flex" : "none";
  if (count)
    count.textContent = n + " lead" + (n !== 1 ? "s" : "") + " selected";
  const allCbs = document.querySelectorAll(".lead-checkbox");
  const selAll = document.getElementById("select-all-cb");
  if (selAll && allCbs.length) {
    selAll.indeterminate = n > 0 && n < allCbs.length;
    selAll.checked = n === allCbs.length;
  }
}

function clearSelection() {
  State.selectedLeads.clear();
  document.querySelectorAll(".lead-checkbox").forEach(function (cb) {
    cb.checked = false;
  });
  const selAll = document.getElementById("select-all-cb");
  if (selAll) {
    selAll.checked = false;
    selAll.indeterminate = false;
  }
  updateBulkBar();
}

async function bulkDelete() {
  const ids = Array.from(State.selectedLeads);
  if (!ids.length) return;
  if (
    !confirm(
      "Permanently delete " +
        ids.length +
        " lead" +
        (ids.length !== 1 ? "s" : "") +
        "? This cannot be undone.",
    )
  )
    return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) {
      await Graph.deleteLead(ids[i]);
    }
    UI.showToast(
      "Deleted " + ids.length + " lead" + (ids.length !== 1 ? "s" : ""),
      "success",
    );
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssign() {
  const ids = Array.from(State.selectedLeads);
  const agent = (document.getElementById("bulk-assign-select") || {}).value;
  if (!ids.length) return;
  if (!agent) {
    UI.showToast("Please select an agent first.", "error");
    return;
  }
  if (
    !confirm(
      "Assign " +
        ids.length +
        " lead" +
        (ids.length !== 1 ? "s" : "") +
        " to " +
        agent +
        "?",
    )
  )
    return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) {
      await Graph.assignAgent(ids[i], agent);
    }
    UI.showToast("Assigned " + ids.length + " leads to " + agent, "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssignType() {
  const ids = Array.from(State.selectedLeads);
  const type = (document.getElementById("bulk-type-select") || {}).value;
  if (!ids.length) return;
  if (!type) {
    UI.showToast("Please select a lead type first.", "error");
    return;
  }
  if (
    !confirm(
      'Set type to "' +
        type +
        '" for ' +
        ids.length +
        " lead" +
        (ids.length !== 1 ? "s" : "") +
        "?",
    )
  )
    return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) {
      await Graph.updateLead(ids[i], { Lead_x0020_Type: type });
    }
    UI.showToast("Set " + ids.length + " leads to type: " + type, "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function bulkExportSelected() {
  const ids = Array.from(State.selectedLeads);
  const leads = State.leads.filter(function (l) {
    return ids.includes(l.id);
  });
  if (!leads.length) return;
  const today = new Date().toISOString().slice(0, 10);
  const csv = [
    "Name,Type,Email,Phone,Status,Source,Assigned To,MRC,Current Products,Last Contacted,Notes",
  ]
    .concat(
      leads.map(function (l) {
        return [
          l.name,
          l.leadType,
          l.email,
          l.phone,
          l.status,
          l.source,
          l.assignedTo,
          l.currentMRC,
          l.currentProducts,
          l.lastContacted,
          l.notes,
        ]
          .map(function (v) {
            return '"' + String(v || "").replace(/"/g, '""') + '"';
          })
          .join(",");
      }),
    )
    .join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = "raimak-leads-selected-" + today + ".csv";
  a.click();
  UI.showToast("Exported " + leads.length + " leads!", "success");
}

async function confirmClearAll() {
  const total = State.leads.length;
  if (!total) {
    UI.showToast("No leads to clear.", "info");
    return;
  }
  const input = prompt(
    "This will permanently delete ALL " +
      total +
      " leads.\n\nType DELETE to confirm:",
  );
  if (input !== "DELETE") {
    UI.showToast("Clear all cancelled.", "info");
    return;
  }
  setLoading(true);
  try {
    for (var i = 0; i < State.leads.length; i++) {
      await Graph.deleteLead(State.leads[i].id);
    }
    UI.showToast("All " + total + " leads deleted. Clean slate!", "success");
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function getFilteredLeads() {
  let leads = State.leads.slice();
  const { status, search, assignedTo } = State.filters;
  if (status !== "all")
    leads = leads.filter(function (l) {
      return l.status === status;
    });
  if (assignedTo !== "all")
    leads = leads.filter(function (l) {
      return l.assignedTo === assignedTo;
    });
  if (search.trim()) {
    const q = search.toLowerCase();
    leads = leads.filter(function (l) {
      return (
        l.name.toLowerCase().includes(q) || l.email.toLowerCase().includes(q)
      );
    });
  }
  return leads;
}

function applyFilters() {
  State.filters.search =
    (document.getElementById("search-input") || {}).value || "";
  State.filters.status =
    (document.getElementById("filter-status") || {}).value || "all";
  State.filters.assignedTo =
    (document.getElementById("filter-agent") || {}).value || "all";

  const wrap = document.getElementById("leads-table-wrap");

  if (wrap) {
    wrap.replaceChildren(renderLeadsTable(getFilteredLeads()));
  }
}

function renderLeadsTable(leads, compact = false, agentView = false) {
  // 1. Handle Empty State
  if (!leads.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.innerHTML = "<p>No leads found.</p>";
    return empty;
  }

  // 2. Setup Template
  const template = document.getElementById("tmpl-leads-table");
  const clone = template.content.cloneNode(true);

  // We wrap it so we can return the outer element properly
  const wrapper = document.createElement("div");
  wrapper.appendChild(clone);

  const thead = wrapper.querySelector("#table-header-row");
  const tbody = wrapper.querySelector("#table-body");

  // 3. Build Headers dynamically based on 'compact' mode
  let headers = "<tr>";
  if (!compact)
    headers += `<th style="width:36px"><input type="checkbox" id="select-all-cb" class="lead-cb" onchange="toggleSelectAll(this.checked)" title="Select all"></th>`;
  headers += `<th>Name</th><th>Type</th><th>Status</th><th>Phone</th><th>Assigned To</th><th>Address</th><th>Last Contacted</th>`;
  if (!compact) headers += `<th>CBR</th><th>BTN</th><th>Flags</th><th></th>`;
  headers += "</tr>";
  thead.innerHTML = headers;

  // 4. Build Rows instantly using our helper
  const admin = isAdmin();
  tbody.innerHTML = leads
    .map((lead) => buildLeadRowHtml(lead, compact, agentView, admin))
    .join("");

  return wrapper.firstElementChild; // Return the living DOM element!
}

function buildLeadRowHtml(lead, compact, agentView, admin) {
  const statusCls =
    "status-" +
    lead.status
      .toLowerCase()
      .replace(/\s+/g, "-")
      .replace(/[^a-z0-9-]/g, "");
  const typeCls = lead.leadType
    ? "lead-type-" + lead.leadType.toLowerCase()
    : "";
  const isChecked = State.selectedLeads.has(lead.id);

  const rowClick = agentView
    ? `loadLeadInFeed('${lead.id}')`
    : admin
      ? `openEditLeadModal('${lead.id}')`
      : "";
  const rowStyle = agentView ? "cursor:pointer" : "";
  const warnCls =
    lead.flags && lead.flags.includes("needs_recycle") ? "row-warn" : "";
  const selCls = isChecked ? "row-selected" : "";

  let html = `<tr class="lead-row ${warnCls} ${selCls}" onclick="${rowClick}" style="${rowStyle}">`;

  if (!compact) {
    html += `<td onclick="event.stopPropagation()" style="width:36px">
              <input type="checkbox" class="lead-checkbox lead-cb" data-id="${lead.id}" ${isChecked ? "checked" : ""} onchange="toggleLeadSelect('${lead.id}',this.checked)">
             </td>`;
  }

  html += `<td><span class="lead-name">${escHtml(lead.name)}</span></td>`;
  html += `<td>${lead.leadType ? `<span class="lead-type-badge ${typeCls}">${escHtml(lead.leadType)}</span>` : "—"}</td>`;
  html += `<td><span class="status-badge ${statusCls}">${lead.status}</span></td>`;
  html += `<td class="td-mono">${escHtml(lead.phone || "—")}</td>`;
  html += `<td>${escHtml(lead.assignedTo || "—")}</td>`;
  html += `<td class="td-mono" style="font-size:11px">${lead.address ? escHtml(lead.address) : "—"}${lead.city ? ", " + escHtml(lead.city) : ""}${lead.state ? " " + escHtml(lead.state) : ""}</td>`;
  html += `<td class="td-mono">${formatDate(lead.lastContacted) || "—"}</td>`;

  if (!compact) {
    html += `<td class="td-mono">${escHtml(lead.cbr || "—")}</td>`;
    html += `<td class="td-mono">${escHtml(lead.btn || "—")}</td>`;

    const flagsHtml = (lead.flags || [])
      .map((f) => `<span class="flag flag-${f}">${flagLabel(f)}</span>`)
      .join("");
    html += `<td class="td-flags">${flagsHtml}</td>`;

    html += `<td class="td-actions">`;
    if (admin) {
      html += `
        <button class="btn-icon" onclick="event.stopPropagation();openEditLeadModal('${lead.id}')" title="Edit">
          <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        </button>
        <button class="btn-icon btn-danger" onclick="event.stopPropagation();deleteLead('${lead.id}')" title="Delete">
          <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        </button>`;
    }
    html += `</td>`;
  }

  html += `</tr>`;
  return html;
}

function loadLeadInFeed(leadId) {
  const realIndex = (window._myLeads || []).findIndex((l) => l.id === leadId);
  if (realIndex !== -1) {
    _leadSaved = false;
    _currentFeedIndex = realIndex;
    window._forceShowLead = true;
    renderMyLeads();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }
}

// ============================================================
//  DAILY REPORT (Admin only)
// ============================================================
async function renderDailyReport() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }
  document.getElementById("main-content").innerHTML = `
    <div class="view-header"><h1 class="view-title">Daily Report</h1></div>
    <div class="card"><div class="empty-state" style="padding:40px">Loading report...</div></div>`;
  try {
    const stats = await Graph.getDailyStats();
    const today = new Date().toLocaleDateString("en-GB", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    document.getElementById("main-content").innerHTML = `
      <div class="view-header">
        <div>
          <h1 class="view-title">Daily Report</h1>
          <span class="view-subtitle">// ${today}</span>
        </div>
        <button class="btn-ghost" onclick="exportReportCSV()">Export CSV</button>
      </div>
      <div class="kpi-grid">
        <div class="kpi-card kpi-primary"><span class="kpi-label">Total Contacts Today</span><span class="kpi-value">${stats.reduce(
          function (s, a) {
            return s + a.contacts;
          },
          0,
        )}</span></div>
        <div class="kpi-card kpi-success"><span class="kpi-label">Total Sales Today</span><span class="kpi-value">${State.todaySales.length}</span></div>
        <div class="kpi-card kpi-info"><span class="kpi-label">Active Agents</span><span class="kpi-value">${stats.length}</span></div>
        <div class="kpi-card kpi-neutral"><span class="kpi-label">Avg Contacts/Agent</span><span class="kpi-value">${
          stats.length
            ? Math.round(
                stats.reduce(function (s, a) {
                  return s + a.contacts;
                }, 0) / stats.length,
              )
            : 0
        }</span></div>
      </div>
      <div class="card">
        <div class="card-header"><h2 class="card-title">Agent Breakdown</h2></div>
        <div class="table-wrap">
          <table class="data-table">
            <thead><tr><th>Agent</th><th>Contacts Today</th><th>Sales Today</th><th>Limit</th><th>Last Action</th></tr></thead>
            <tbody>
              ${
                stats.length
                  ? stats
                      .map(function (a) {
                        const pct = Math.round(
                          (a.contacts / Config.rules.maxContactsPerDay) * 100,
                        );
                        const last = a.actions.length ? a.actions[0] : null;
                        return `<tr>
                  <td><span class="lead-name">${escHtml(a.agent)}</span></td>
                  <td>
                    <div style="display:flex;align-items:center;gap:10px">
                      <span class="td-mono">${a.contacts}</span>
                      <div class="load-bar-wrap" style="flex:1;max-width:80px"><div class="load-bar ${pct >= 100 ? "load-full" : pct >= 80 ? "load-high" : ""}" style="width:${Math.min(100, pct)}%"></div></div>
                    </div>
                  </td>
                  <td><span class="status-badge status-sold">${a.sold}</span></td>
                  <td class="td-mono">${a.contacts}/${Config.rules.maxContactsPerDay}</td>
                  <td class="td-mono" style="color:var(--text-3)">${last ? formatDateTime(last.timestamp) : "—"}</td>
                </tr>`;
                      })
                      .join("")
                  : `<tr><td colspan="5" class="empty-state">No activity today yet.</td></tr>`
              }
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
  const csv = ["Agent,Contacts Today,Sales Today,Date"]
    .concat(
      stats.map(function (a) {
        return [a.agent, a.contacts, a.sold, today]
          .map(function (v) {
            return '"' + String(v || "").replace(/"/g, '""') + '"';
          })
          .join(",");
      }),
    )
    .join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = "raimak-report-" + today + ".csv";
  a.click();
  UI.showToast("Report exported!", "success");
}

// ============================================================
//  RAIMAK TEAM (Admin only)
// ============================================================
function renderContractors() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }
  const { contractors, leads } = State;
  const max = Config.rules.maxLeadsPerAgent;
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Raimak Team</h1>
      <span class="view-subtitle">// ${contractors.length} agents</span>
    </div>
    <div class="contractor-grid">
      ${contractors
        .map(function (c) {
          const count = leads.filter(function (l) {
            return (
              l.assignedTo === c.name &&
              !Config.terminalStatuses.includes(l.status)
            );
          }).length;
          const pct = Math.min(100, Math.round((count / max) * 100));
          const contacts = Graph.agentContactsToday(
            c.email || c.name,
            State.activityLog,
          );
          return `
          <div class="contractor-card">
            <div class="contractor-header">
              <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
              <div><div class="contractor-name">${escHtml(c.name)}</div><div class="contractor-role">${escHtml(c.role)}</div></div>
              <span class="status-dot ${c.active ? "dot-active" : "dot-inactive"}"></span>
            </div>
            <div class="contractor-email">${escHtml(c.email || "No email")}</div>
            <div class="load-label"><span>Lead Load</span><span class="${count >= max ? "text-danger" : ""}">${count}/${max}</span></div>
            <div class="load-bar-wrap"><div class="load-bar ${pct >= 100 ? "load-full" : pct >= 80 ? "load-high" : ""}" style="width:${pct}%"></div></div>
            <div class="load-label" style="margin-top:10px"><span>Contacts Today</span><span>${contacts}/${Config.rules.maxContactsPerDay}</span></div>
            <div class="load-bar-wrap"><div class="load-bar ${contacts >= Config.rules.maxContactsPerDay ? "load-full" : ""}" style="width:${Math.min(100, Math.round((contacts / Config.rules.maxContactsPerDay) * 100))}%"></div></div>
          </div>`;
        })
        .join("")}
    </div>`;
}

// ============================================================
//  ACTIVITY LOG (Admin only)
// ============================================================
function renderActivity() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }
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
            ${
              activityLog.length
                ? activityLog
                    .map(function (e) {
                      return `
              <tr>
                <td class="td-mono">${formatDateTime(e.timestamp)}</td>
                <td>${escHtml(e.leadName || e.leadId || "—")}</td>
                <td><span class="action-badge">${escHtml(e.action || "—")}</span></td>
                <td>${escHtml(e.agent || "—")}</td>
                <td class="td-notes">${escHtml(e.notes || "")}</td>
              </tr>`;
                    })
                    .join("")
                : `<tr><td colspan="5" class="empty-state">No activity yet.</td></tr>`
            }
          </tbody>
        </table>
      </div>
    </div>`;
}

// ============================================================
//  LEAD MODAL (Admin — Add/Edit)
// ============================================================
function openAddLeadModal() {
  if (!isAdmin()) return;
  State.editingLeadId = null;
  renderLeadModal(null);
}

function openEditLeadModal(id) {
  if (!isAdmin()) return;
  const lead = State.leads.find(function (l) {
    return l.id === id;
  });
  if (!lead) return;
  State.editingLeadId = id;
  renderLeadModal(lead);
}

function renderLeadModal(lead) {
  const isEdit = !!lead;
  const contractors = State.contractors.map((c) => c.name);
  const modalContainer = document.getElementById("modal");
  modalContainer.innerHTML = ""; // Clear existing
  const template = document.getElementById("tmpl-lead-modal");
  const clone = template.content.cloneNode(true);

  // 2. Header & Button Logic
  clone.getElementById("modal-title").textContent = isEdit
    ? "Edit Lead"
    : "New Lead";
  const submitBtn = clone.getElementById("modal-submit-btn");
  submitBtn.textContent = isEdit ? "Save Changes" : "Add Lead";
  submitBtn.onclick = () => (isEdit ? submitEditLead() : submitAddLead());

  // 3. Populate Standard Text Inputs (Safely falls back to empty string if creating new lead)
  const safeVal = (val) => val || "";

  clone.getElementById("f-firstname").value = safeVal(lead?.firstName);
  clone.getElementById("f-lastname").value = safeVal(lead?.lastName);
  clone.getElementById("f-email").value = safeVal(lead?.email);
  clone.getElementById("f-phone").value = safeVal(lead?.phone);
  clone.getElementById("f-address").value = safeVal(lead?.address);
  clone.getElementById("f-city").value = safeVal(lead?.city);
  clone.getElementById("f-state").value = safeVal(lead?.state);
  clone.getElementById("f-zip").value = safeVal(lead?.zip);
  clone.getElementById("f-mrc").value = safeVal(lead?.currentMRC);

  if (lead && lead.lastContacted) {
    clone.getElementById("f-lastcontacted").value =
      lead.lastContacted.split("T")[0];
  }

  // 4. Populate Dropdowns dynamically
  const leadTypeSelect = clone.getElementById("f-leadtype");
  Config.leadTypes.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t;
    opt.textContent = t;
    if (lead && lead.leadType === t) opt.selected = true;
    leadTypeSelect.appendChild(opt);
  });

  const statusSelect = clone.getElementById("f-status");
  Config.leadStatuses.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    if ((lead?.status || "New") === s) opt.selected = true;
    statusSelect.appendChild(opt);
  });

  const assignedSelect = clone.getElementById("f-assigned");
  contractors.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    if (lead && lead.assignedTo === c) opt.selected = true;
    assignedSelect.appendChild(opt);
  });

  const productsSelect = clone.getElementById("f-products");
  Config.currentProducts.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    if (lead && lead.currentProducts === p) opt.selected = true;
    productsSelect.appendChild(opt);
  });

  // 5. Build AutoPay Radios
  const autopayContainer = clone.getElementById("f-autopay-container");
  ["ACH - Debit Card", "ACH - Credit Card", "No Auto Pay"].forEach((opt) => {
    const isChecked = lead && lead.autoPay === opt ? "checked" : "";
    autopayContainer.innerHTML += `
      <label style="display:flex;align-items:center;gap:8px;font-size:13px;cursor:pointer">
        <input type="radio" name="f-autopay" value="${opt}" ${isChecked} style="accent-color:#2563B0;width:14px;height:14px"> ${opt}
      </label>`;
  });

  // 6. Notes History Parser
  const notesHistory = clone.getElementById("modal-notes-history");
  if (lead && lead.notes && lead.notes.trim()) {
    const notesHtml = lead.notes
      .split("\n")
      .filter((l) => l.trim())
      .map((line) => {
        const match = line.match(/^\[(\d{2}\/\d{2}(?:\/\d{2})?)(.*?)\]\s*(.*)/);
        if (match) {
          const date = match[1];
          const agent = match[2] ? match[2].replace(/^\s*-\s*/, "") : "";
          const text = match[3];
          return `
            <div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
              <div style="display:flex;gap:8px;align-items:center;margin-bottom:3px">
                <span style="font-family:var(--font-mono);font-size:10px;color:#2563B0;font-weight:700;background:#E8F0FF;padding:1px 6px;border-radius:3px">${date}</span>
                ${agent ? `<span style="font-family:var(--font-mono);font-size:10px;color:#6B85B0">${escHtml(agent)}</span>` : ""}
              </div>
              <span style="font-size:13px;color:#1A2640">${escHtml(text)}</span>
            </div>`;
        }
        return `<div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8"><div style="margin-bottom:3px"><span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8;background:#F4F7FD;padding:1px 6px;border-radius:3px">Legacy note — author unknown</span></div><span style="font-size:13px;color:#4A6080">${escHtml(line)}</span></div>`;
      })
      .join("");

    notesHistory.innerHTML = `<div style="background:#F4F7FD;border:1px solid #D0DCF0;border-radius:6px;padding:12px 14px;margin-bottom:10px;max-height:180px;overflow-y:auto">${notesHtml}</div>`;
  } else {
    notesHistory.innerHTML = `<div style="font-size:12px;color:#8EA5C8;margin-bottom:10px;font-family:var(--font-mono)">No notes yet.</div>`;
  }

  // 7. Mount & Display
  modalContainer.appendChild(clone);
  document.getElementById("modal-overlay").style.display = "flex";
}

async function submitAddLead() {
  const fields = collectLeadForm();
  if (!fields) return;
  const agentName = fields._agentName;
  delete fields._agentName;
  setLoading(true);
  try {
    const newLead = await Graph.addLead(fields);
    if (agentName) await Graph.assignAgent(newLead.id, agentName);
    await Graph.logActivity({
      LeadID: newLead.id,
      Title: fields.Title,
      ActionType: "Lead Created",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
    });
    await refreshData();
    closeModal();
    UI.showToast("Lead added!", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function submitEditLead() {
  const fields = collectLeadForm();
  if (!fields) return;
  const agentName = fields._agentName;
  delete fields._agentName;
  setLoading(true);
  try {
    await Graph.updateLead(State.editingLeadId, fields);
    if (agentName) await Graph.assignAgent(State.editingLeadId, agentName);
    await Graph.logActivity({
      LeadID: State.editingLeadId,
      Title: fields.Title,
      ActionType: "Lead Updated",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
    });
    await refreshData();
    closeModal();
    UI.showToast("Lead updated!", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function collectLeadForm() {
  const firstName = (
    (document.getElementById("f-firstname") || {}).value || ""
  ).trim();
  const lastName = (
    (document.getElementById("f-lastname") || {}).value || ""
  ).trim();
  const agentName = (document.getElementById("f-assigned") || {}).value || "";
  const nameEl = document.getElementById("f-name");
  const fullName = nameEl
    ? (nameEl.value || "").trim()
    : (firstName + " " + lastName).trim();

  if (!firstName && !lastName && !fullName) {
    UI.showToast("Name is required.", "error");
    return null;
  }

  const fields = { _agentName: agentName };
  if (firstName) fields["FirstName"] = firstName;
  if (lastName) fields["LastName"] = lastName;
  if (fullName && !firstName) fields["Title"] = fullName;

  const add = function (key, elId, trim) {
    const el = document.getElementById(elId);
    const val = el ? (trim ? (el.value || "").trim() : el.value || "") : "";
    if (val) fields[key] = val;
  };

  add("Lead_x0020_Type", "f-leadtype");
  add("Email", "f-email", true);
  add("Phone", "f-phone", true);
  add("Status", "f-status");
  add("LastTouchedOn", "f-lastcontacted");
  add("MonthlyRecurringCharge_x0028_MRC", "f-mrc", true);
  add("CurrentProducts", "f-products");
  add("CBR", "f-cbr", true);
  add("BTN", "f-btn", true);
  add("WorkAddress", "f-address", true);
  add("WorkCity", "f-city", true);
  add("State", "f-state", true);
  add("Zip", "f-zip", true);
  add("AutoPay", "f-autopay");

  const notesEl = document.getElementById("f-notes");
  if (notesEl && notesEl.value.trim()) {
    const today = new Date();
    const dateStamp =
      (today.getMonth() + 1).toString().padStart(2, "0") +
      "/" +
      today.getDate().toString().padStart(2, "0") +
      "/" +
      String(today.getFullYear()).slice(-2);
    const adminName =
      State.currentUser && State.currentUser.name
        ? " - " + State.currentUser.name
        : "";
    const lead = State.editingLeadId
      ? State.leads.find(function (l) {
          return l.id === State.editingLeadId;
        })
      : null;
    const existing = (lead && lead.notes) || "";
    const stamped = "[" + dateStamp + adminName + "] " + notesEl.value.trim();
    fields["Notes"] = existing ? stamped + "\n" + existing : stamped;
  }

  if (!fields.Status) fields.Status = "New";
  return fields;
}

async function deleteLead(id) {
  const lead = State.leads.find(function (l) {
    return l.id === id;
  });
  if (!confirm('Delete "' + (lead && lead.name) + '"? This cannot be undone.'))
    return;
  setLoading(true);
  try {
    await Graph.deleteLead(id);
    await refreshData();
    UI.showToast("Lead deleted.", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function closeModal(event) {
  if (event && event.target !== document.getElementById("modal-overlay"))
    return;
  document.getElementById("modal-overlay").style.display = "none";
}

// ============================================================
//  UTILITIES
// ============================================================
async function refreshData() {
  await loadAllData();
  navigate(State.currentView);
}

function setLoading(on) {
  State.loading = on;
  const o = document.getElementById("loading-overlay");
  if (o) o.style.display = on ? "flex" : "none";
}

function updateBadges() {
  const n = State.leads.filter(function (l) {
    return (
      l.flags &&
      (l.flags.includes("needs_recycle") ||
        l.flags.includes("agent_overloaded"))
    );
  }).length;
  const b = document.getElementById("badge-leads");
  if (b) {
    b.textContent = n > 0 ? n : "";
    b.style.display = n > 0 ? "inline-flex" : "none";
  }
}

function exportCSV() {
  const leads = getFilteredLeads();
  const today = new Date().toISOString().slice(0, 10);
  const csv = [
    "Name,Type,Email,Phone,Status,Source,Assigned To,MRC,Current Products,Last Contacted,Notes",
  ]
    .concat(
      leads.map(function (l) {
        return [
          l.name,
          l.leadType,
          l.email,
          l.phone,
          l.status,
          l.source,
          l.assignedTo,
          l.currentMRC,
          l.currentProducts,
          l.lastContacted,
          l.notes,
        ]
          .map(function (v) {
            return '"' + String(v || "").replace(/"/g, '""') + '"';
          })
          .join(",");
      }),
    )
    .join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = "raimak-leads-" + today + ".csv";
  a.click();
  UI.showToast("Exported!", "success");
}

function flagLabel(f) {
  return (
    {
      cool_off: "Cool-off",
      needs_recycle: "Recycle",
      agent_overloaded: "Overloaded",
    }[f] || f
  );
}
function formatDate(d) {
  if (!d) return "";
  return new Date(d).toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}
function formatTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  });
}
function formatDateTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}
function escHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

const UI = {
  showToast: function (msg, type) {
    type = type || "info";
    const c = document.getElementById("toast-container");
    if (!c) return;
    const t = document.createElement("div");
    t.className = "toast toast-" + type;
    t.textContent = msg;
    c.appendChild(t);
    setTimeout(function () {
      t.classList.add("show");
    }, 10);
    setTimeout(function () {
      t.classList.remove("show");
      setTimeout(function () {
        t.remove();
      }, 300);
    }, 4000);
  },
  showConfetti: function () {
    const el = document.createElement("div");
    el.className = "confetti-burst";
    el.innerHTML = "&#127881; SOLD! &#127881;";
    document.body.appendChild(el);
    setTimeout(function () {
      el.remove();
    }, 2600);
  },
};

const Ticker = {
  update: function () {
    const tickerEl = document.getElementById("sales-ticker-content");
    if (!tickerEl) return;

    // Grab today's date string (e.g., "Fri Apr 10 2026")
    const todayStr = new Date().toDateString();

    // 1. Find all leads sold TODAY
    const soldToday = State.leads.filter((l) => {
      const isSold = l.status && l.status.toLowerCase().includes("sold");
      // Check if the lead was last updated today
      const isFromToday =
        l.lastContacted &&
        new Date(l.lastContacted).toDateString() === todayStr;

      return isSold && isFromToday;
    });

    // 2. Calculate Top 5 Agents (Today Only)
    const agentSales = {};
    soldToday.forEach((l) => {
      const seller = l.soldBy || l.assignedTo;
      if (seller) {
        agentSales[seller] = (agentSales[seller] || 0) + 1;
      }
    });

    const topAgents = Object.entries(agentSales)
      .sort((a, b) => b[1] - a[1]) // Sort highest to lowest
      .slice(0, 5) // Grab top 5
      .map(
        (entry, i) =>
          `<strong>#${i + 1} ${escHtml(entry[0])}</strong> (${entry[1]})`,
      );

    // 3. Grab the 5 most recent sales (Today Only)
    const recentSales = soldToday
      .slice(-5)
      .reverse()
      .map((l) => {
        const soldBy = l.soldBy || l.assignedTo || "Someone";
        const forAgent = l.assignedTo;

        if (forAgent && soldBy !== forAgent) {
          return `🎉 <strong>${escHtml(soldBy)}</strong> just made a sale for <strong>${escHtml(forAgent)}</strong> — ${escHtml(l.name)}`;
        } else {
          return `🎉 <strong>${escHtml(soldBy)}</strong> just closed a sale! — ${escHtml(l.name)}`;
        }
      });

    // 4. Build the string with new "TODAY" labels
    let textParts = [];
    if (recentSales.length > 0) {
      textParts.push(`🔥 TODAY'S RECENT: ${recentSales.join("  •  ")}`);
    }
    if (topAgents.length > 0) {
      textParts.push(`🏆 TODAY'S LEADERS: ${topAgents.join("  •  ")}`);
    }

    // 5. Inject it
    tickerEl.innerHTML =
      textParts.length > 0
        ? textParts.join("  |  ")
        : "🚀 Let's make some sales today!";
  },
};
