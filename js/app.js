// Raimak LMS - App Logic v3.0
window._isWorkingCallback = false;
const cachedSkips = sessionStorage.getItem("_skippedSessionLeads");
window._skippedSessionLeads = cachedSkips ? JSON.parse(cachedSkips) : [];
const savedSyncDate = localStorage.getItem("RaimakActivityLastSyncDate");

const State = {
  leads: [],
  contractors: [],
  activityLog: [],
  todaySales: [],
  drafts: {},
  agentScores: [],
  currentView: "dashboard",
  filters: { status: "all", search: "", assignedTo: "all" },
  editingLeadId: null,
  loading: false,
  role: "agent",
  currentUser: null,
  salesFeedTimer: null,
  dripLead: null,
  selectedLeads: new Set(),

  // 🚀 THE NEW TIME-BASED TRACKER
  lastSyncDate: savedSyncDate || null,
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
    await LocalDB.init();
    State.currentUser = Auth.getUser();
    State.role = detectRole(State.currentUser);
    showAppShell();
    Points.initHUDAutoHider();
    await loadAllData();
    Points.updateHUD();
    renderDashboard();
    Ticker.update();
  } catch (err) {
    console.error("Boot error:", err);
    showLoginScreen();
  }
});

async function loadAllData() {
  setLoading(true);
  UI.showToast("Syncing floor data...", "info");

  try {
    // 1. Resolve IDs once
    await Graph.resolveSiteIds();

    // 🚀 THE FIX Part 1: We pull the Activity Log OUT of the concurrent race.
    // Leads are heavy, but Contractors and Points are tiny, so they can safely race together!
    const [rawLeads, contractors, pointsData] = await Promise.all([
      Graph.getLeads().then((data) => {
        UI.showToast("✅ Leads synced!", "success");
        return data;
      }),
      Graph.getContractors().then((data) => {
        UI.showToast("✅ Contractors synced!", "success");
        return data;
      }),
      Points.fetchBalances().then((data) => {
        UI.showToast("✅ Points balances synced!", "success");
        return data;
      }),
    ]);

    State.contractors = contractors;
    State.leads = Graph.applyBusinessRules(rawLeads, contractors);

    // 🚀 THE FIX Part 2: The Smart Activity Fetch
    // Now that the massive Leads download is finished, we safely ask for the Logs
    let todayLogs = [];

    if (isAdmin()) {
      UI.showToast("Syncing historical admin logs...", "info");

      // Admins use the hyper-fast Delta Sync to get the entire database
      // 🚀 UPDATED: Swapped highestActivityId for lastSyncDate
      const logData = await Graph.getActivityLog(
        State.lastSyncDate,
        State.activityLog,
      );

      State.activityLog = logData.updatedLogs;
      // 🚀 UPDATED: Swapped newHighestId for newSyncDate
      State.lastSyncDate = logData.newSyncDate;

      // Extract today's logs purely from RAM so we don't have to fetch them again!
      const todayStr = new Date().toDateString();
      todayLogs = State.activityLog.filter(
        (log) =>
          log.timestamp && new Date(log.timestamp).toDateString() === todayStr,
      );

      UI.showToast("✅ Admin logs synced!", "success");
    } else {
      UI.showToast("Syncing today's activity...", "info");

      // Standard agents just get the fast daily log
      todayLogs = await Graph.getActivityLogForToday();
      State.activityLog = todayLogs;

      UI.showToast("✅ Activity synced!", "success");
    }

    // 3. Instant, synchronous math!
    State.todaySales = Graph.getTodaySales(todayLogs);
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
    case "callbacks":
      renderCallBacks();
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

  const todayStr = new Date().toLocaleDateString();
  const agentUniqueLeads = {};

  const aliasMap = {
    "j.torres@raimak.com": "JULIAN TORRES",
    // ANY EMAILS THAT ARE SHOWING UP ON THE LEADERBOARD, DO THIS FOR THEM.
  };
  // 1. Loop through all logs and count unique leads touched per agent today
  (State.activityLog || []).forEach((log) => {
    let isToday = false;
    if (log.timestamp) {
      isToday = new Date(log.timestamp).toLocaleDateString() === todayStr;
    }

    const actionStr = log.action || log.ActionType || "";
    const isContact =
      actionStr.startsWith("Status:") ||
      actionStr === "1st Contact" ||
      actionStr === "2nd Contact" ||
      actionStr === "3rd Contact";

    if (isToday && isContact) {
      let rawAgent = (log.agent || log.AgentEmail || "Unknown")
        .toLowerCase()
        .trim();

      // THE FIX: Check the alias map first, then fall back to the contractor list
      let displayName = aliasMap[rawAgent];

      if (!displayName) {
        const contractor = State.contractors.find(
          (c) =>
            (c.email || "").toLowerCase().trim() === rawAgent ||
            (c.name || "").toLowerCase().trim() === rawAgent,
        );
        displayName = contractor
          ? contractor.name
          : log.agent || log.AgentEmail || "Unknown";
      }

      const leadId = log.leadId || log.LeadID;

      if (!agentUniqueLeads[displayName])
        agentUniqueLeads[displayName] = new Set();
      if (leadId) agentUniqueLeads[displayName].add(leadId);
    }
  });

  // 2. Convert sets to numbers, sort highest to lowest, and grab the top 5
  const top5Contacts = Object.entries(agentUniqueLeads)
    .map(([name, leadSet]) => [name, leadSet.size])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  // 3. Inject exactly like the sales leaderboard, but with blue text for contacts!
  const dashContactsEl = clone.getElementById("dash-top5-contacts");
  if (dashContactsEl) {
    dashContactsEl.innerHTML = top5Contacts.length
      ? top5Contacts
          .map(
            (e, i) => `
        <div class="top5-row">
          <span class="top5-rank rank-${i + 1}">${i + 1}</span>
          <span class="top5-name">${escHtml(e[0])}</span>
          <span class="top5-count" style="color: var(--blue, #3b82f6);">${e[1]} contact${e[1] !== 1 ? "s" : ""}</span>
        </div>`,
          )
          .join("")
      : `<p class="empty-state">No contacts logged yet today.</p>`;
  }

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
      // 🚀 1. The Lightweight Delta Sync (Runs silently!)
      // UPDATED: Swapped highestActivityId for lastSyncDate
      const logData = await Graph.getActivityLog(
        State.lastSyncDate,
        State.activityLog,
      );

      // Update global state with the merged logs and the new timestamp
      State.activityLog = logData.updatedLogs;
      // UPDATED: Swapped newHighestId for newSyncDate
      State.lastSyncDate = logData.newSyncDate;

      // 🚀 2. In-Memory Math (Zero network requests)
      const newSales = Graph.getTodaySales(State.activityLog);

      // Update State FIRST so if Ticker.update() relies on it, the data is ready
      State.todaySales = newSales;

      // 3. The Confetti Trigger
      const newOnes = newSales.filter(function (l) {
        return !knownSaleIds.has(l.id);
      });

      if (newOnes.length > 0) {
        // Scrubbed the window. prefix!
        if (Ticker && Ticker.update) Ticker.update();
        if (UI && UI.showConfetti) UI.showConfetti();

        newOnes.forEach(function (l) {
          knownSaleIds.add(l.id);
        });
      }

      // 4. DOM Updates
      if (State.currentView === "dashboard") {
        const feed = document.getElementById("dash-sales-feed");
        const time = document.getElementById("sales-feed-time");

        if (!feed) return;

        if (time) {
          time.textContent = "Updated " + formatTime(new Date().toISOString());
        }

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

        // 🚀 Translated for the Activity Log object structure
        feed.innerHTML = [...newSales]
          .sort(function (a, b) {
            return new Date(b.saleTime) - new Date(a.saleTime);
          })
          .slice(0, 6)
          .map(function (l) {
            // 1. soldBy already contains the translated name from getTodaySales
            const displayAgent = l.soldBy || "Unassigned";

            // 2. 'name' is what getTodaySales uses instead of 'leadName'
            const displayName = l.name || "Unknown Lead";

            return `
      <div class="sale-entry">
        <div class="sale-icon">🎉</div>
        <div class="sale-info">
          <span class="sale-name">${escHtml(displayName)}</span>
          <span class="sale-agent">${escHtml(displayAgent)}</span>
        </div>
        <span class="sale-time">${formatTime(l.saleTime)}</span>
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

  // ==========================================
  //  THE STRICT BOUNCER
  // ==========================================
  let myLeads;
  if (window._forceShowLead && window._myLeads && window._myLeads.length > 0) {
    myLeads = window._myLeads;
  } else {
    // Otherwise, run the normal filter to build the agent's daily queue
    myLeads = State.leads.filter((l) => {
      const assigned = (l.assignedTo || "")
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim();
      const matchesAgent =
        assigned &&
        (assigned === agentName.replace(/\s+/g, " ") ||
          assigned === userName.replace(/\s+/g, " ") ||
          assigned === userEmail.replace(/\s+/g, " "));

      // Bouncer 1: Is it a terminal status?
      const isTerminal =
        Config.terminalStatuses.includes(l.status) ||
        l.status === "3rd Contact";

      // Bouncer 2: Is it in cool-off?
      const inCoolOff = Graph.isInCoolOff(l);

      // Bouncer 3: The Callback & Install Check
      let waitingForDate = false;
      let isDueCallback = false;

      if (l.callbackAt) {
        // Strip time to compare pure calendar days for BOTH rules
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const scheduledDate = new Date(l.callbackAt);
        scheduledDate.setHours(0, 0, 0, 0);

        if (l.status === "Pending Order") {
          // INSTALL RULE: Keep hidden on the day of the install.
          if (today <= scheduledDate) waitingForDate = true;
        } else {
          // CALLBACK RULE: Calendar Day Checking!
          // Keep it hidden ONLY if the scheduled date is tomorrow or later.
          // This ensures a 4:30 PM callback is visible immediately at 8:00 AM.
          if (today < scheduledDate) waitingForDate = true;
        }

        // If it's today (or overdue), they get the VIP pass to bypass cool-off!
        if (!waitingForDate) {
          isDueCallback = true;
        }
      }

      // THE ULTIMATE DECISION:
      // If it's a due callback, it gets a VIP Pass to bypass the cool-off restriction!
      const passedCoolOff = isDueCallback ? true : !inCoolOff;
      const isDismissed =
        window._skippedSessionLeads &&
        window._skippedSessionLeads.includes(l.id);
      // ONLY keep actionable, unworked leads that have passed their waiting period!
      return (
        matchesAgent &&
        !isTerminal &&
        passedCoolOff &&
        !waitingForDate &&
        !isDismissed
      );
    });

    // ==========================================
    //  THE SORT: FORCE CALLBACKS TO THE TOP
    // ==========================================
    myLeads.sort((a, b) => {
      const aHasCallback = !!a.callbackAt;
      const bHasCallback = !!b.callbackAt;

      // Rule 1: Callbacks always rise above standard leads
      if (aHasCallback && !bHasCallback) return -1;
      if (!aHasCallback && bHasCallback) return 1;

      // Rule 2: If BOTH are callbacks, sort by exact time (earliest calls first)
      if (aHasCallback && bHasCallback) {
        return new Date(a.callbackAt) - new Date(b.callbackAt);
      }

      // Rule 3: Leave standard leads alone
      return 0;
    });
  }

  // Keep global variables intact
  window._myLeads = myLeads;
  window._agentName = agentName;
  _leadSaved = false;

  if (_currentFeedIndex >= myLeads.length) _currentFeedIndex = 0;

  // We completely deleted the old `while` loop here because the Bouncer
  // physically removed the cool-off leads from the array!

  window._forceShowLead = false;

  // ==========================================
  //  THE RENDER LOGIC
  // ==========================================
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";
  const contactsToday = getMyContactsToday();
  const template = document.getElementById("tmpl-my-leads");
  const clone = template.content.cloneNode(true);
  const textEl = clone.getElementById("myleads-contact-text");
  if (textEl) {
    textEl.textContent = contactsToday;
  }
  clone.getElementById("myleads-subtitle").textContent =
    `// ${myLeads.length} remaining · lead ${Math.min(_currentFeedIndex + 1, myLeads.length || 1)} of ${myLeads.length}`;

  clone.getElementById("lead-feed-wrap").innerHTML = "";
  clone
    .getElementById("lead-feed-wrap")
    .appendChild(renderLeadFeedCard(myLeads));

  mainContent.appendChild(clone);

  // ==========================================
  //  LIVE CLOCK LOGIC (Smart Timezones)
  // ==========================================
  const clockEl = document.getElementById("myleads-clock");

  const activeLead = myLeads[_currentFeedIndex];
  const leadState =
    activeLead && activeLead.state
      ? activeLead.state.toUpperCase().trim()
      : null;

  const updateClock = () => {
    if (!clockEl) return;
    let tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    if (leadState && stateTimezones[leadState]) {
      tz = stateTimezones[leadState];
    }
    try {
      clockEl.textContent = new Date().toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        timeZone: tz,
        timeZoneName: "short",
      });
    } catch (e) {
      clockEl.textContent = new Date().toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
    }
  };

  updateClock();
  if (window._clockTimer) clearInterval(window._clockTimer);
  window._clockTimer = setInterval(updateClock, 1000);
}

function searchMyLeads(q) {
  const wrap = document.getElementById("my-leads-table");
  if (!wrap) return;

  if (!q.trim()) {
    wrap.innerHTML = "";
    return;
  }

  const queryLower = q.trim().toLowerCase();
  const toggleEl = document.getElementById("toggle-search-all");
  const searchAll = toggleEl ? toggleEl.checked : true;

  // Grab the current agent's identity (matching the logic from your render function)
  const agentName = (window._agentName || "")
    .toLowerCase()
    .replace(/\s+/g, " ");
  const userName = ((State.currentUser && State.currentUser.name) || "")
    .toLowerCase()
    .replace(/\s+/g, " ");
  const userEmail = ((State.currentUser && State.currentUser.email) || "")
    .toLowerCase()
    .replace(/\s+/g, " ");

  // THE FIX: Always search the master database so we catch cool-off and terminal leads!
  const filtered = (State.leads || []).filter((l) => {
    const assigned = (l.assignedTo || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();

    let passesAssignmentFilter = false;

    if (searchAll) {
      // Searching everything: just verify it's assigned to someone
      passesAssignmentFilter = assigned !== "";
    } else {
      // Searching "My Leads": verify it belongs to THIS agent specifically
      passesAssignmentFilter =
        assigned &&
        (assigned === agentName ||
          assigned === userName ||
          assigned === userEmail);
    }

    const matchesSearch =
      (l.name && l.name.toLowerCase().includes(queryLower)) ||
      (l.phone && l.phone.includes(queryLower)) ||
      (l.btn && l.btn.includes(queryLower)) ||
      (l.cbr && l.cbr.includes(queryLower)) ||
      (l.address && l.address.toLowerCase().includes(queryLower));

    return passesAssignmentFilter && matchesSearch;
  });

  if (filtered.length) {
    // THE QOL UPGRADE: Slice down to 25 max for instant rendering
    const displayLeads = filtered.slice(0, 25);

    // Draw the table
    wrap.replaceChildren(renderLeadsTable(displayLeads, false, true));

    // Add a helpful hint if we truncated the list
    if (filtered.length > 25) {
      const hint = document.createElement("div");
      hint.style.cssText =
        "text-align:center; padding:12px; font-size:12px; color:#64748b; font-style:italic; border-top:1px solid #e2e8f0;";
      hint.textContent = `+ ${filtered.length - 25} more matches. Keep typing to narrow it down.`;
      wrap.appendChild(hint);
    }
  } else {
    wrap.innerHTML = `<div class="empty-state">No leads found for "${escHtml(q)}"</div>`;
  }
}

function getMyContactsToday() {
  const logs = State.activityLog || [];
  const user = State.currentUser || {};

  const myEmail = (user.email || "").toLowerCase().trim();
  let myName = (user.name || "").toLowerCase().trim();

  const contractor = State.contractors.find(
    (c) =>
      (c.email || "").toLowerCase().trim() === myEmail ||
      (c.name || "").toLowerCase().trim() === myName,
  );
  if (contractor) myName = contractor.name.toLowerCase().trim();

  // 1. Get today's local date (e.g., "4/18/2026")
  const todayString = new Date().toLocaleDateString();

  const uniqueLeads = new Set();

  logs.forEach((log) => {
    const entryAgent = (log.agent || log.AgentEmail || "").toLowerCase().trim();
    const actionStr = log.action || log.ActionType || "";
    const leadId = log.leadId || log.LeadID;

    // 2. Convert the log's timestamp to a local date string and compare
    let isToday = false;
    if (log.timestamp) {
      isToday = new Date(log.timestamp).toLocaleDateString() === todayString;
    }

    const isMyLog = entryAgent === myEmail || entryAgent === myName;
    const isContact =
      actionStr.startsWith("Status:") ||
      actionStr === "1st Contact" ||
      actionStr === "2nd Contact" ||
      actionStr === "3rd Contact";

    // 3. The new requirement: It MUST happen today
    if (isMyLog && isContact && isToday && leadId) {
      uniqueLeads.add(leadId);
    }
  });

  return uniqueLeads.size;
}
let _stagedStatus = null;

function renderLeadFeedCard(myLeads) {
  // 1. FIXED LOGIC: Grab the exact lead we are supposed to be looking at!
  let lead = myLeads[_currentFeedIndex];

  // If that exact lead happens to be in cool-off, we should probably warn the logic,
  // but we still want to show them the correct lead!
  const isCoolOff = lead ? Graph.isInCoolOff(lead) : false;

  _stagedStatus = null;
  window._forceShowLead = false;
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

  // ==========================================
  // NEW: 1. THE CALLBACK / INSTALL ALERT BADGE
  // ==========================================
  const callbackAlert = clone.getElementById("feed-callback-alert");
  if (callbackAlert && lead.callbackAt) {
    const targetDate = new Date(lead.callbackAt);
    const today = new Date();

    // Create midnight versions for pure day comparison
    const todayMidnight = new Date(today);
    todayMidnight.setHours(0, 0, 0, 0);
    const targetMidnight = new Date(targetDate);
    targetMidnight.setHours(0, 0, 0, 0);

    if (todayMidnight >= targetMidnight) {
      // Format time like "2:30 PM"
      const timeString = targetDate.toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
      });

      if (lead.status === "Pending Order") {
        callbackAlert.innerHTML =
          "📅 INSTALL DUE: Check if fiber is active today!";
      } else {
        callbackAlert.innerHTML = `📅 ACTION REQUIRED: Scheduled follow-up today at ${timeString}`;
      }
      callbackAlert.style.display = "block";
    }
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

  // ==========================================
  // NEW: 2. PULLING THE SAVED CALLBACK DATE UI
  // ==========================================
  const callbackInput = clone.getElementById("f-callback-date");
  const callbackWrap = clone.getElementById("callback-wrapper");
  const callbackBtn = clone.getElementById("btn-toggle-callback");
  const callbackLabel = clone.getElementById("callback-label");

  if (callbackInput && lead.callbackAt) {
    // SharePoint gives us ISO strings (UTC). HTML datetime-local needs YYYY-MM-DDThh:mm in local time.
    const localDate = new Date(lead.callbackAt);
    const tzOffset = localDate.getTimezoneOffset() * 60000;
    const localISOTime = new Date(localDate - tzOffset)
      .toISOString()
      .slice(0, 16);

    callbackInput.value = localISOTime;

    // Slide the menu open so the agent sees the date is set
    if (callbackWrap) {
      callbackWrap.style.width = "200px";
      callbackWrap.style.opacity = "1";
      callbackWrap.style.overflow = "visible";
      callbackWrap.dataset.manuallyOpened = "false";
    }

    // If it's a pending order, apply the visual lockdown
    if (lead.status === "Pending Order" && callbackBtn) {
      if (callbackLabel)
        callbackLabel.innerHTML =
          'Scheduled Install <span style="color: var(--red)">*</span>';
      callbackBtn.disabled = true;
      callbackBtn.style.opacity = "0.4";
      callbackBtn.style.cursor = "not-allowed";
      callbackInput.required = true;
    }
  }

  if (callbackInput) {
    callbackInput.addEventListener("change", (e) => {
      const selectedVal = e.target.value;

      if (callbackWrap) {
        callbackWrap.dataset.manuallyOpened = "true";
      }

      if (selectedVal) {
        // 1. The Green Flash on the input box
        callbackInput.style.transition =
          "background-color 0.3s, border-color 0.3s";
        callbackInput.style.backgroundColor = "var(--green-dim, #e6f8f3)";
        callbackInput.style.borderColor = "var(--green, #10b981)";

        setTimeout(() => {
          callbackInput.style.backgroundColor = ""; // Fade out background, leave the border
        }, 600);

        // 2. Only apply the cooldown and label change if it's NOT a Pending Order
        if (lead.status !== "Pending Order" && callbackBtn) {
          // Update the label to the "Aha!" state
          if (callbackLabel) {
            const d = new Date(selectedVal);
            const formattedStr =
              d.toLocaleDateString([], { month: "short", day: "numeric" }) +
              " @ " +
              d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            callbackLabel.innerHTML = `✅ Set: ${formattedStr}`;
            callbackLabel.style.color = "var(--green, #10b981)";
          }

          // Gray it out to prevent the instant double-click
          callbackBtn.disabled = true;
          callbackBtn.style.opacity = "0.4";
          callbackBtn.style.cursor = "not-allowed";

          // Bring it back to life after 2 seconds
          setTimeout(() => {
            // Double check it wasn't miraculously changed to a pending order in the last 2 seconds
            if (lead.status !== "Pending Order") {
              callbackBtn.disabled = false;
              callbackBtn.style.opacity = "1";
              callbackBtn.style.cursor = "pointer";
            }
          }, 2000);
        }
      } else {
        // 3. THE MISSING ELSE BLOCK (If they manually clear the input box)
        callbackInput.style.borderColor = "";
        if (callbackLabel) {
          callbackLabel.innerHTML = "Callback date and time";
          callbackLabel.style.color = "#6b85b0";
        }
      }
    });
  }
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
    soldBySelect.appendChild(option);
  });

  // THE DRAFT PEEK
  const draft = State.drafts[lead.id] || {};
  const activeAutoPay =
    draft.autoPay !== undefined ? draft.autoPay : lead.autoPay;

  // AutoPay Radios
  const autoPayContainer = clone.getElementById("feed-autopay-container");
  ["ACH - Debit Card", "ACH - Credit Card", "No Auto Pay"].forEach((opt) => {
    const isChecked = activeAutoPay === opt ? "checked" : "";
    autoPayContainer.innerHTML += `
      <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;color:#1A2640;background:#F4F7FD;border:1px solid #D0DCF0;padding:8px 14px;border-radius:6px;">
        <input type="radio" name="feed-autopay" value="${opt}" ${isChecked} style="accent-color:#2563B0"> ${opt}
      </label>`;
  });

  // Status Buttons
  const statusContainer = clone.getElementById("feed-status-buttons");
  const hiddenStatuses = ["New", "TD Non-Reg", "D2D Lead"];

  Config.leadStatuses
    .filter((s) => !hiddenStatuses.includes(s))
    .forEach((s) => {
      const isTDM = s === "TDM" ? " ↩" : "";
      const cls =
        "status-btn-" +
        s
          .toLowerCase()
          .replace(/\s+/g, "-")
          .replace(/[^a-z0-9-]/g, "");
      statusContainer.innerHTML += `<button class="status-btn ${cls}" id="sbtn-${s.replace(/\s+/g, "-")}" onclick="stageStatus('${lead.id}','${s}')">${s}${isTDM}</button>`;
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

  // THE DRAFT MEMORY ENABLER
  const inputsToDraft = [
    { id: "feed-btn", key: "btn" },
    { id: "feed-mrc", key: "mrc" },
    { id: "feed-cbr", key: "cbr" },
    { id: "feed-notes", key: "notes" },
    { id: "feed-products", key: "products" },
    { id: "feed-sold-by", key: "soldBy" },
  ];

  inputsToDraft.forEach((item) => {
    const el = clone.getElementById(item.id);
    if (el) {
      if (draft[item.key] !== undefined) {
        el.value = draft[item.key];
      }
      const eventType = el.tagName === "SELECT" ? "change" : "input";
      el.addEventListener(eventType, (e) =>
        updateLeadDraft(lead.id, item.key, e.target.value),
      );
    }
  });

  const allAutoPayRadios = clone.querySelectorAll('input[name="feed-autopay"]');
  if (allAutoPayRadios.length > 0) {
    allAutoPayRadios.forEach((r) => {
      r.addEventListener("change", (e) =>
        updateLeadDraft(lead.id, "autoPay", e.target.value),
      );
    });
  }

  // Wrap and Return
  const wrapper = document.createElement("div");
  wrapper.appendChild(clone);
  return wrapper;
}

function toggleCallbackDate() {
  const wrap = document.getElementById("callback-wrapper");
  const btn = document.getElementById("btn-toggle-callback");
  const input = document.getElementById("f-callback-date");
  const label = document.getElementById("callback-label");

  // Prevent toggling if it's locked (like for Pending Orders or during our 2-second cooldown!)
  if (btn.disabled) return;

  if (wrap.style.width === "0px" || wrap.style.width === "") {
    // Slide & Fade IN
    wrap.style.width = "200px";
    wrap.style.opacity = "1";

    // THE MEMORY: Remember that the agent specifically asked for this to be open
    wrap.dataset.manuallyOpened = "true";

    setTimeout(() => {
      wrap.style.overflow = "visible";
    }, 300);
  } else {
    // Slide & Fade OUT
    wrap.style.width = "0px";
    wrap.style.opacity = "0";
    wrap.style.overflow = "hidden";

    // THE MEMORY: Remember that the agent closed it
    wrap.dataset.manuallyOpened = "false";

    // Scrub the data and reset the UI *after* the menu finishes sliding shut
    setTimeout(() => {
      if (input) {
        input.value = "";
        input.style.borderColor = "";
        input.style.backgroundColor = "";
      }

      // Reset the label ONLY if it's our dynamic "✅ Set:" text
      if (label && label.innerHTML.includes("✅ Set:")) {
        label.innerHTML = "Callback date and time"; // <-- Updated to match template!
        label.style.color = "#6b85b0";
      }
    }, 300);
  }
}

function updateCallbackUIForStatus(status) {
  const wrap = document.getElementById("callback-wrapper");
  const label = document.getElementById("callback-label");
  const btn = document.getElementById("btn-toggle-callback");
  const dateInput = document.getElementById("f-callback-date");

  if (!wrap || !label || !btn || !dateInput) return;

  if (status === "Pending Order") {
    // Force the menu open instantly and lock the button
    wrap.style.width = "200px";
    wrap.style.opacity = "1";
    wrap.style.overflow = "visible";

    label.innerHTML =
      'Scheduled Install <span style="color: var(--red)">*</span>';
    btn.disabled = true;
    btn.style.opacity = "0.4";
    btn.style.cursor = "not-allowed";
    dateInput.required = true;
  } else {
    // Return everything to normal callback mode
    label.textContent = "Callback date and time";
    btn.disabled = false;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
    dateInput.required = false;

    // THE FIX: Check the memory! If they didn't manually open it before, slide it shut.
    if (wrap.dataset.manuallyOpened !== "true") {
      wrap.style.width = "0px";
      wrap.style.opacity = "0";
      wrap.style.overflow = "hidden";

      // Wait for the slide animation to finish before clearing the data
      setTimeout(() => {
        if (wrap.style.width === "0px") dateInput.value = "";
      }, 300);
    }
  }
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
  updateCallbackUIForStatus(newStatus);
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

  // Your original edit modal IDs - untouched!
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

  // NEW: Grab the universal callback date from the sliding toggle UI
  let rawCallbackDate =
    (document.getElementById("f-callback-date") || {}).value || "";

  // ==========================================
  // THE INTERCEPT: Prevent Phantom Callbacks
  // ==========================================
  if (window._isWorkingCallback) {
    const wrap = document.getElementById("callback-wrapper");
    // If they didn't actively schedule a NEW time, scrub the old time out of the variable!
    if (!wrap || wrap.dataset.manuallyOpened !== "true") {
      rawCallbackDate = "";
    }
  }

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

  if (!mrc || mrc.trim() === "") {
    UI.showToast("Please enter an MRC amount.", "error");
    return;
  }

  const cleanBtn = btn.replace(/\D/g, "");
  if (cleanBtn.length !== 10) {
    UI.showToast("Please enter a valid 10-digit BTN.", "error");
    return;
  }

  const cleanCbr = cbr.replace(/\D/g, "");
  if (cleanCbr.length !== 10) {
    UI.showToast("Please enter a valid 10-digit CBR.", "error");
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

  // THE MAGIC: Format the callback date for SharePoint
  if (rawCallbackDate) {
    // Converts the HTML datetime picker into the exact ISO string SharePoint demands
    saveFields["CallbackDateTime"] = new Date(rawCallbackDate).toISOString();
  } else if (newStatus !== "Pending Order") {
    // Because "Callback" isn't a status, if there is no date, WIPE IT!
    // (We only spare it if it's a Pending Order, to protect the install date)
    saveFields["CallbackDateTime"] = null;
  }

  setLoading(true);
  try {
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

    // Wait for all network calls to finish simultaneously
    await Promise.all(networkTasks);

    State.activityLog.push({
      leadId: leadId,
      agent: activityEmail,
      action: "Status: " + newStatus,
      timestamp: new Date().toISOString(),
    });
    delete State.drafts[leadId];

    // --- THE OPTIMISTIC UI UPDATE ---
    // Instantly inject new data into local memory so the Bouncer & Search act immediately
    lead.status = newStatus;
    lead.notes = notes;
    if (mrc) lead.currentMRC = mrc;
    if (products) lead.currentProducts = products;
    if (cbr) lead.cbr = cbr;
    if (btn) lead.btn = btn;
    if (autoPay) lead.autoPay = autoPay;
    if (newStatus === "TDM") lead.assignedTo = "";

    // Crucial for the Bouncer: update the local callback string!
    lead.callbackAt = rawCallbackDate || null;
    // --------------------------------
    Points.awardPoints(newStatus, leadId);
    if (newStatus === "TDM") {
      if (window.UI && UI.showToast)
        UI.showToast("TDM — lead returned to admin queue.", "info");
    } else {
      if (window.UI && UI.showToast) UI.showToast("Saved!", "success");
    }

    // This updates the UI and immediately hides the lead if it's on cool-off
    Ticker.update();
    _stagedStatus = null;
    _leadSaved = true;

    const nextRow = document.getElementById("feed-next-row");
    const searchSec = document.getElementById("lead-search-section");
    const saveBtn = document.getElementById("feed-save-btn");

    if (nextRow) {
      nextRow.style.display = "block";
      const nextBtn = nextRow.querySelector("button");

      if (nextBtn && window._isWorkingCallback) {
        window._isWorkingCallback = false;
        window._forceShowLead = false;
        nextBtn.innerHTML = "Complete Callback ✓";
        nextBtn.classList.remove("btn-cyan");
        nextBtn.classList.add("btn-green");

        // 2. Point the click action to our new completion function
        nextBtn.onclick = () => completeCallbackLead(lead.id);
      }
    }
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
    if (window.UI && UI.showToast)
      UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function advanceToNextLead() {
  const currentLead = window._myLeads[_currentFeedIndex];

  if (currentLead) {
    let isDismissedCallback = false;

    if (currentLead.callbackAt) {
      // 1. Initialize the memory bank just in case
      if (!window._skippedSessionLeads) window._skippedSessionLeads = [];

      // 2. Prevent duplicates and memorize the dismiss
      if (!window._skippedSessionLeads.includes(currentLead.id)) {
        window._skippedSessionLeads.push(currentLead.id);
        sessionStorage.setItem(
          "_skippedSessionLeads",
          JSON.stringify(window._skippedSessionLeads),
        );
      }
      isDismissedCallback = true;
    }

    const isTerminal = Config.terminalStatuses.includes(currentLead.status);
    const isExhausted = currentLead.status === "3rd Contact";
    const inCoolOff = Graph.isInCoolOff(currentLead);

    // 3. Only increment the index if the lead is staying in the active queue!
    // If it's terminal, exhausted, cool-off, OR a dismissed callback,
    // the Bouncer will remove it during renderMyLeads(), naturally sliding the next lead into this slot.
    if (!isTerminal && !isExhausted && !inCoolOff && !isDismissedCallback) {
      _currentFeedIndex++;
    }
  }

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

      // THE UPGRADE: Must have NO previous agents AND NO MRC value
      const unworkedMatch =
        !requireUnworked || (!l.previousAgents && !l.currentMRC);

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
  const unworkedElement = document.getElementById("bulk-unworked-check");
  const requireUnworked = unworkedElement ? unworkedElement.checked : false;
  if (!agentName) {
    UI.showToast("Please select an agent first.", "warning");
    return;
  }
  if (!qty || qty <= 0) {
    UI.showToast("Please enter a valid number of leads.", "warning");
    return;
  }

  const unassigned = State.leads.filter(function (l) {
    // THE GHOST RECORD SHIELD
    const isValidLead = l && l.id && (l.name || l.phone);

    const isAvailable =
      !l.assignedTo && !Config.terminalStatuses.includes(l.status);

    const matchesType =
      selectedType === "all" ||
      (l.leadType && l.leadType.toLowerCase() === selectedType.toLowerCase());

    const matchesState =
      selectedState === "all" ||
      (l.state && l.state.toUpperCase() === selectedState.toUpperCase());

    // THE FIX: Use the actual Status to determine if it's unworked,
    // and safely check previousAgents so it doesn't crash on undefined strings!
    const isFresh =
      l.status === "New" &&
      (!l.previousAgents || l.previousAgents.trim() === "");
    const matchesUnworked = !requireUnworked || isFresh;

    return (
      isValidLead &&
      isAvailable &&
      matchesType &&
      matchesState &&
      matchesUnworked
    );
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
// CALLBACKS
// ============================================================
function renderCallBacks() {
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = ""; // Clear existing screen

  // 1. Setup Template
  const template = document.getElementById("tmpl-callbacks-page");
  const clone = template.content.cloneNode(true);
  const wrap = clone.getElementById("callbacks-list-wrap");

  // 2. Identify the current agent
  const userName = ((State.currentUser && State.currentUser.name) || "")
    .toLowerCase()
    .trim();
  const userEmail = ((State.currentUser && State.currentUser.email) || "")
    .toLowerCase()
    .trim();

  const contractor = (State.contractors || []).find((c) => {
    return (
      (c.email || "").toLowerCase().trim() === userEmail ||
      (c.name || "").toLowerCase().trim() === userName
    );
  });

  const agentName = contractor
    ? contractor.name.toLowerCase().trim()
    : userName;

  // 3. Filter the master database
  const callbacks = (State.leads || []).filter((l) => {
    const assigned = (l.assignedTo || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();

    const isAssignedToMe =
      assigned &&
      (assigned === agentName.replace(/\s+/g, " ") ||
        assigned === userName.replace(/\s+/g, " ") ||
        assigned === userEmail.replace(/\s+/g, " "));

    const isTerminal =
      l.status === "TDM" || l.status === (Config.soldStatus || "Sold");

    return isAssignedToMe && !isTerminal && l.callbackAt;
  });

  // 4. Sort chronologically
  callbacks.sort((a, b) => {
    const dateA = new Date(a.callbackAt);
    if (a.status === "Pending Order") dateA.setDate(dateA.getDate() + 1);

    const dateB = new Date(b.callbackAt);
    if (b.status === "Pending Order") dateB.setDate(dateB.getDate() + 1);

    return dateA - dateB;
  });

  // 5. Handle the Empty State
  if (callbacks.length === 0) {
    // NEW: Added the animation class to the empty state so it fades in too!
    wrap.innerHTML = `
      <div class="animate-fade-up" style="padding: 60px 20px; text-align: center; color: #64748b;">
        <div style="font-size: 40px; margin-bottom: 16px;">📅</div>
        <h3 style="margin: 0 0 8px 0; color: #0a1a3f; font-size: 18px;">Pipeline Clear</h3>
        <p style="margin: 0; font-size: 14px;">You have no upcoming callbacks or installations scheduled.</p>
      </div>
    `;
  } else {
    // 6. Build the Table
    let html = `
      <table class="animate-fade-up" style="width: 100%; border-collapse: collapse; text-align: left;">
        <thead>
          <tr style="background: #f8fafc; border-bottom: 1px solid #e2e8f0; font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 1px;">
            <th style="padding: 14px 16px; font-weight: 600;">Scheduled For</th>
            <th style="padding: 14px 16px; font-weight: 600;">Customer Info</th>
            <th style="padding: 14px 16px; font-weight: 600;">Status</th>
            <th style="padding: 14px 16px; font-weight: 600; text-align: right;">Action</th>
          </tr>
        </thead>
        <tbody>
    `;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // NEW: Added the 'index' parameter to the loop so we can multiply it for the staggered delay
    callbacks.forEach((l, index) => {
      // 1. Calculate the actual Action Date
      const actionDate = new Date(l.callbackAt);
      const isInstall = l.status === "Pending Order";

      if (isInstall) {
        // Push the required action to the day AFTER the install
        actionDate.setDate(actionDate.getDate() + 1);
      }

      const actionMidnight = new Date(actionDate);
      actionMidnight.setHours(0, 0, 0, 0);

      let dateColor = "#1e293b";
      let dateWeight = "normal";
      let badge = "";

      // 2. Check the new action date against Today
      if (actionMidnight < today) {
        dateColor = "var(--red, #ef4444)";
        dateWeight = "bold";
        badge = `<span style="background: #fee2e2; color: #991b1b; padding: 2px 6px; border-radius: 4px; font-size: 10px; margin-left: 8px; font-weight: bold;">OVERDUE</span>`;
      } else if (actionMidnight.getTime() === today.getTime()) {
        dateColor = "var(--blue, #2563B0)";
        dateWeight = "bold";
        badge = `<span style="background: #e0f2fe; color: #0369a1; padding: 2px 6px; border-radius: 4px; font-size: 10px; margin-left: 8px; font-weight: bold;">TODAY</span>`;
      }

      // 3. Format the strings for the UI
      const dateStr = actionDate.toLocaleDateString([], {
        weekday: "short",
        month: "short",
        day: "numeric",
      });
      const timeStr = actionDate.toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
      });

      // ---> THE UI TOUCH SNIPPET <---
      // Give the agent clear context on the original install date
      const typeStr = isInstall
        ? `Install Check (Installed ${new Date(l.callbackAt).toLocaleDateString([], { month: "short", day: "numeric" })})`
        : "Scheduled Call";

      const statusCls = `status-${(l.status || "")
        .toLowerCase()
        .replace(/\s+/g, "-")
        .replace(/[^a-z0-9-]/g, "")}`;

      // NEW: Added the animate-fade-up class and the dynamic staggered delay using the index
      html += `
        <tr class="animate-row-fade" style="border-bottom: 1px solid #f1f5f9; transition: background 0.2s; animation-delay: ${index * 0.05 + 0.1}s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='white'">
          <td style="padding: 16px; font-size: 13px; color: ${dateColor}; font-weight: ${dateWeight};">
            <div style="display: flex; align-items: center; font-size: 14px;">${dateStr} @ ${timeStr} ${badge}</div>
            <div style="font-size: 11px; color: #64748b; margin-top: 4px; font-weight: normal; text-transform: uppercase;">${typeStr}</div>
          </td>
          <td style="padding: 16px;">
            <div style="font-size: 14px; font-weight: 600; color: #0a1a3f;">${escHtml(l.name || "Unknown")}</div>
            <div style="font-size: 12px; color: #64748b; margin-top: 2px;">📞 ${escHtml(l.cbr || "No Phone Number")}</div>
          </td>
          <td style="padding: 16px;">
            <span class="status-badge ${statusCls}">${l.status}</span>
          </td>
          <td style="padding: 16px; text-align: right; white-space: nowrap;">
            <button class="btn-secondary" style="font-size: 12px; padding: 6px 12px; margin-right: 8px;" onclick="viewCallbackLead('${l.id}')">
              View Callback
            </button>
            <button class="btn-primary" style="font-size: 12px; padding: 6px 12px;" onclick="workCallbackLead('${l.id}')">
              Work Callback
            </button>
          </td>
          </td>
        </tr>
      `;
    });

    html += `</tbody></table>`;
    wrap.innerHTML = html;
  }

  // 7. Return the fully built DOM element for your router to mount!
  mainContent.appendChild(clone);
}

function viewCallbackLead(leadId) {
  const targetLead = State.leads.find((l) => l.id === leadId);
  if (!targetLead) return;

  const currentQueue = window._myLeads || [];
  const filteredQueue = currentQueue.filter((l) => l.id !== leadId);
  window._myLeads = [targetLead, ...filteredQueue];

  window._forceShowLead = true;
  window._currentFeedIndex = 0;

  // NEW: Make sure the app knows we are just viewing
  window._isWorkingCallback = false;

  navigate("myleads");
}

function workCallbackLead(leadId) {
  const targetLead = State.leads.find((l) => l.id === leadId);
  if (!targetLead) return;
  const currentQueue = window._myLeads || [];
  const filteredQueue = currentQueue.filter((l) => l.id !== leadId);
  window._myLeads = [targetLead, ...filteredQueue];

  window._forceShowLead = true;
  window._currentFeedIndex = 0;

  // NEW: Tell the app to intercept the "Next Lead" button
  window._isWorkingCallback = true;

  navigate("myleads");
}

function completeCallbackLead(leadId) {
  const targetLead = State.leads.find((l) => l.id === leadId);

  if (targetLead) {
    // 1. NOW we wipe the date since the interaction is over
    targetLead.callbackAt = null;

    // 2. Secretly update SharePoint in the background
    if (window.Graph && Graph.updateLead) {
      Graph.updateLead(leadId, { CallbackDateTime: null }).catch((err) =>
        console.error("Failed to clear callback", err),
      );
    }
  }

  // 3. Reset all our bypass flags so the Bouncer wakes back up
  window._isWorkingCallback = false;
  window._forceShowLead = false;

  // 4. Send them back to their pipeline
  navigate("callbacks");
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

  // ==========================================
  // 🚀 4. CSV IMPORTER WIRING
  // ==========================================
  const importBtn = clone.getElementById("importLeadsBtn");
  const fileInput = clone.getElementById("leadFileInput");

  if (importBtn && fileInput) {
    importBtn.onclick = () => {
      // 1. Build a temporary overlay and modal
      const overlay = document.createElement("div");
      overlay.style.cssText =
        "position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(13, 27, 62, 0.6); z-index:9999; display:flex; align-items:center; justify-content:center; backdrop-filter: blur(3px);";

      const modal = document.createElement("div");
      // Mimicking your clean CRM styling
      modal.style.cssText =
        "background:#fff; padding:24px; border-radius:12px; width:320px; box-shadow:0 10px 25px rgba(0,0,0,0.2); display:flex; flex-direction:column; gap:16px;";

      modal.innerHTML = `
        <div>
          <h3 style="margin:0 0 4px 0; font-size:18px; color:#0D1B3E;">Upload Leads</h3>
          <p style="margin:0; font-size:13px; color:#666;">What type of leads are in this file?</p>
        </div>
        <select id="tempLeadType" class="filter-select" style="width:100%; padding:10px; border-radius:6px;">
          <option value="OFS">OFS Leads</option>
          <option value="MLR">MLR Leads</option>
          <option value="Forced">Forced Leads</option>
        </select>
        <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:8px;">
          <button id="cancelTypeBtn" class="btn-ghost" style="padding:8px 16px;">Cancel</button>
          <button id="confirmTypeBtn" class="btn-primary" style="padding:8px 16px;">Choose File</button>
        </div>
      `;

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      // 2. Handle Modal Clicks
      document.getElementById("cancelTypeBtn").onclick = () => overlay.remove();

      document.getElementById("confirmTypeBtn").onclick = () => {
        // Grab the choice and securely attach it to the file input's dataset
        const selectedType = document.getElementById("tempLeadType").value;
        fileInput.dataset.leadType = selectedType;

        // Destroy the modal and open the file browser
        overlay.remove();
        fileInput.click();
      };
    };

    // Attach our parser function to the hidden input
    fileInput.addEventListener("change", handleFileSelect, false);
  }

  // 5. Populate Bulk Bar Dropdowns
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

  // 6. Populate Filters (and restore any previous search state)
  const searchInput = clone.getElementById("search-input");
  if (searchInput) searchInput.value = State.filters.search || "";

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

  // 7. THE SMART PAGINATION RENDERER
  let currentPage = 1;
  const itemsPerPage = 50;

  // Build the pagination controls dynamically
  const paginationWrap = document.createElement("div");
  paginationWrap.style.cssText =
    "display:flex; gap:6px; align-items:center; justify-content:flex-end; padding: 12px 0; margin-top: 8px;";
  paginationWrap.innerHTML = `
    <button id="leads-prev-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">&larr; Prev</button>
    <span id="leads-page-indicator" style="font-family:var(--font-mono); font-size:13px; font-weight:600; color:#0D1B3E; min-width: 80px; text-align: center;">Pg 1</span>
    <button id="leads-next-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">Next &rarr;</button>
  `;

  // Insert the controls right below the table wrapper inside the clone
  const tableWrap = clone.getElementById("leads-table-wrap");
  if (tableWrap) {
    tableWrap.parentNode.insertBefore(paginationWrap, tableWrap.nextSibling);
  }

  const prevBtn = paginationWrap.querySelector("#leads-prev-page");
  const nextBtn = paginationWrap.querySelector("#leads-next-page");
  const pageIndicator = paginationWrap.querySelector("#leads-page-indicator");

  function updateTable() {
    // A. Grab the master list using your existing logic
    const filtered = getFilteredLeads();

    // B. Run Pagination Math
    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / itemsPerPage));

    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    // C. Update UI Text & Buttons
    if (pageIndicator)
      pageIndicator.textContent = `Pg ${currentPage} / ${totalPages}`;

    if (prevBtn) {
      prevBtn.disabled = currentPage === 1;
      prevBtn.style.opacity = currentPage === 1 ? "0.4" : "1";
    }
    if (nextBtn) {
      nextBtn.disabled = currentPage === totalPages;
      nextBtn.style.opacity = currentPage === totalPages ? "0.4" : "1";
    }

    // D. Slice down to 50 items and draw
    const startIndex = (currentPage - 1) * itemsPerPage;
    const displayLeads = filtered.slice(startIndex, startIndex + itemsPerPage);

    // This updates the live DOM perfectly because \`tableWrap\` maintains its
    // pointer to the element even after the clone is mounted to the screen.
    if (tableWrap) {
      tableWrap.replaceChildren(renderLeadsTable(displayLeads));
    }
  }

  // Attach button listeners
  if (prevBtn)
    prevBtn.addEventListener("click", () => {
      if (currentPage > 1) {
        currentPage--;
        updateTable();
      }
    });
  if (nextBtn)
    nextBtn.addEventListener("click", () => {
      currentPage++;
      updateTable();
    });

  // Initial draw before mounting
  updateTable();

  // 8. Mount!
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

function getFilteredLeads() {
  const { status, search, assignedTo } = State.filters;
  const q = (search || "").trim().toLowerCase();

  return State.leads.filter(function (l) {
    // 1. Status Match
    const matchStatus = status === "all" || l.status === status;

    // 2. Search Match (with safety fallbacks to prevent crashes)
    const matchSearch =
      !q ||
      (l.name && l.name.toLowerCase().includes(q)) ||
      (l.email && l.email.toLowerCase().includes(q)) ||
      (l.phone && l.phone.includes(q));

    // 3. STRICT Agent Match (No override)
    const matchAgent = assignedTo === "all" || l.assignedTo === assignedTo;

    // The lead must pass all three tests to show up on the screen
    return matchStatus && matchSearch && matchAgent;
  });
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
  let realIndex = (window._myLeads || []).findIndex((l) => l.id === leadId);
  if (realIndex === -1) {
    const masterList = State.leads || State.allLeads || [];
    const globalLead = masterList.find((l) => l.id === leadId);

    if (globalLead) {
      // Temporarily inject it at the very beginning of their personal feed
      window._myLeads = window._myLeads || [];
      window._myLeads.unshift(globalLead);
      realIndex = 0;
    } else {
      if (window.UI && UI.showToast)
        UI.showToast("Lead not found in database.", "error");
      return;
    }
  }

  // 3. Render the card (now guaranteed to work)
  _leadSaved = false;
  _currentFeedIndex = realIndex;
  window._forceShowLead = true;
  renderMyLeads();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function handleFileSelect(event) {
  const files = event.target.files;
  if (!files || files.length === 0) return;

  // 🎯 Grab the Lead Type we saved from the modal
  const selectedLeadType = event.target.dataset.leadType || "OFS";
  let combinedCSVData = [];

  UI.showToast(`📄 Reading ${files.length} file(s)...`, "info");

  // Helper function to turn the FileReader into an awaitable Promise
  const readSingleFile = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        // Use your existing parseCSV function to turn the text into objects
        const parsed = parseCSV(e.target.result);
        resolve(parsed);
      };
      reader.readAsText(file);
    });
  };

  // Loop through every file they selected and parse it
  for (let i = 0; i < files.length; i++) {
    const parsedData = await readSingleFile(files[i]);
    // Merge the new data into our master list
    combinedCSVData = combinedCSVData.concat(parsedData);
  }

  console.log(
    `📂 Merged ${files.length} files. Total raw leads: ${combinedCSVData.length}`,
  );

  // Send the massive combined list to your uploader
  // Your "Bouncer" inside this function will automatically filter duplicates across all the files!
  await uploadLeadsToSharePoint(combinedCSVData, selectedLeadType);

  // Reset input so they can upload again later
  event.target.value = "";
}

// The CSV Parser (Handles quotes and commas flawlessly)
function parseCSV(text) {
  const lines = text.split("\n").filter((line) => line.trim() !== "");
  const headers = lines[0].split(",").map((h) => h.trim().replace(/"/g, ""));
  const data = [];

  for (let i = 1; i < lines.length; i++) {
    const values = lines[i]
      .split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/)
      .map((v) => v.trim().replace(/"/g, ""));
    let row = {};

    headers.forEach((h, index) => {
      row[h] = values[index] || "";
    });
    data.push(row);
  }
  return data;
}

async function uploadLeadsToSharePoint(csvData, leadType) {
  // ==========================================
  // 🛡️ 1. THE IN-MEMORY BOUNCER (Deduplication)
  // ==========================================
  const generateKey = (first, last, address) => {
    // The exact same sponge used for the combo multiplier!
    const clean = (str) =>
      (str || "")
        .replace(/[^\w\s]/gi, "")
        .toLowerCase()
        .trim();
    return `${clean(first)}|${clean(last)}|${clean(address)}`;
  };

  const existingKeys = new Set();

  // A. Log all existing leads into the Bouncer's ledger
  (State.leads || []).forEach((lead) => {
    // Make sure these match the properties returned by your normalizeLeadItem function
    const key = generateKey(lead.firstName, lead.lastName, lead.address);
    existingKeys.add(key);
  });

  const validLeads = [];
  let duplicateCount = 0;

  // B. Scan the incoming CSV against the ledger
  csvData.forEach((row) => {
    const key = generateKey(
      row["FirstName"],
      row["LastName"],
      row["StreetAddress"],
    );

    if (existingKeys.has(key)) {
      duplicateCount++; // Blocked!
    } else {
      validLeads.push(row); // Allowed!
      existingKeys.add(key); // Add to ledger to prevent duplicates WITHIN the CSV itself
    }
  });

  const totalLeads = validLeads.length;

  // C. Safety Check: Did the Bouncer block everything?
  if (totalLeads === 0) {
    UI.showToast(
      `❌ Upload aborted: All ${csvData.length} leads in the file are already in the system.`,
      "error",
    );
    return;
  }

  console.log(
    `🚀 BATCH UPLOAD: ${totalLeads} valid leads... (${duplicateCount} duplicates skipped)`,
  );
  UI.showToast(
    `🚀 Starting batch upload of ${totalLeads} leads... (Skipped ${duplicateCount} duplicates)`,
    "info",
  );

  // ==========================================
  // 🛑 2. HIJACK THE UI
  // ==========================================
  const importBtn = document.getElementById("importLeadsBtn");
  const originalBtnHTML = importBtn ? importBtn.innerHTML : "Import CSV";
  if (importBtn) {
    importBtn.disabled = true;
    importBtn.style.cursor = "not-allowed";
  }

  // ==========================================
  // 🔀 3. URL SETUP FOR BATCHING
  // ==========================================
  const host = Config.sharePoint.hostname;
  const sitePath = Config.sharePoint.sites.team;
  const listId = Config.sharePoint.lists.leadsList;
  const batchUrl = `${Config.sharePoint.graphBase}/$batch`;
  const relativeUploadUrl = `/sites/${host}:/${sitePath}:/lists/${listId}/items`;

  let successCount = 0;
  let failCount = 0;
  const batchSize = 20;

  // ==========================================
  // 📦 4. THE BATCHING ENGINE
  // ==========================================
  for (let i = 0; i < totalLeads; i += batchSize) {
    // Chunking the cleaned validLeads array
    const chunk = validLeads.slice(i, i + batchSize);

    const currentBatchNum = Math.ceil(i / batchSize) + 1;
    const totalBatches = Math.ceil(totalLeads / batchSize);
    if (importBtn) {
      importBtn.innerHTML = `⏳ Uploading Batch ${currentBatchNum} of ${totalBatches}... (${i} leads saved)`;
    }

    const batchRequests = chunk.map((row, index) => {
      return {
        id: String(index + 1),
        method: "POST",
        url: relativeUploadUrl,
        headers: { "Content-Type": "application/json" },
        body: {
          fields: {
            FirstName: row["FirstName"],
            LastName: row["LastName"],
            WorkAddress: row["StreetAddress"],
            WorkCity: row["City"],
            State: row["State"],
            Zip: row["Zip"],
            Lead_x0020_Type: leadType,
            Agent_x0020_Assigned: "",
            Status: "New",
          },
        },
      };
    });

    const payload = { requests: batchRequests };

    try {
      const batchResponse = await Graph.apiFetch(batchUrl, "POST", payload);

      if (batchResponse && batchResponse.responses) {
        batchResponse.responses.forEach((res) => {
          if (res.status >= 200 && res.status < 300) {
            successCount++;
          } else {
            console.error(`❌ Batch item ${res.id} failed:`, res.body);
            failCount++;
          }
        });
      }
    } catch (error) {
      console.error("❌ Entire batch failed due to network error:", error);
      failCount += chunk.length;
      UI.showToast(
        `⚠️ Network error on batch ${currentBatchNum}. Pausing...`,
        "error",
      );
    }
  }

  // ==========================================
  // ✅ 5. RESTORE THE UI & REPORT
  // ==========================================
  if (importBtn) {
    importBtn.disabled = false;
    importBtn.style.cursor = "pointer";
    importBtn.innerHTML = originalBtnHTML;
  }

  if (failCount === 0) {
    UI.showToast(
      `✅ Upload complete! ${successCount} leads added. (Skipped ${duplicateCount} duplicates)`,
      "success",
    );
  } else if (successCount > 0 && failCount > 0) {
    UI.showToast(
      `⚠️ Upload finished: ${successCount} added, ${failCount} failed. (Skipped ${duplicateCount} duplicates)`,
      "warning",
    );
  } else {
    UI.showToast(`❌ Upload failed. 0 leads were added.`, "error");
  }

  // ==========================================
  // 🔄 6. RELOAD DATA
  // ==========================================
  await loadAllData();
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

  const { activityLog, contractors } = State;

  // 1. THE PREDICTABLE IDENTITY MAP
  const identityMap = {};

  function getStandardEmail(name) {
    if (!name) return "";
    const parts = name.trim().toLowerCase().split(/\s+/);
    if (parts.length >= 2) {
      return `${parts[0].charAt(0)}.${parts[parts.length - 1]}@raimak.com`;
    }
    return "";
  }

  function registerPerson(name, knownEmail) {
    if (!name) return;
    const officialName = name.trim();
    const lowerName = officialName.toLowerCase();

    identityMap[lowerName] = officialName;

    const generatedEmail = getStandardEmail(officialName);
    if (generatedEmail) identityMap[generatedEmail] = officialName;

    if (knownEmail) identityMap[knownEmail.trim().toLowerCase()] = officialName;
  }

  (contractors || []).forEach((c) => registerPerson(c.name, c.email));

  const user = State.currentUser;
  if (user) registerPerson(user.name, user.email);

  // 2. THE BULLETPROOF NORMALIZER
  function getCleanAgentName(rawString) {
    if (!rawString) return "";
    const cleanString = rawString.trim().toLowerCase();

    if (identityMap[cleanString]) {
      return identityMap[cleanString];
    }

    return rawString.trim().replace(/\w\S*/g, function (txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
  }

  // 3. EXTRACT UNIQUE DATA FOR DROPDOWNS
  const uniqueAgents = [
    ...new Set(
      activityLog.map((e) => getCleanAgentName(e.agent)).filter(Boolean),
    ),
  ].sort();

  const uniqueActions = [
    ...new Set(activityLog.map((e) => (e.action || "").trim()).filter(Boolean)),
  ].sort();

  // Draw the layout skeleton
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Activity Log</h1>
        <span class="view-subtitle">// ${activityLog.length} total entries</span>
      </div>
    </div>

    <div class="card">
      <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
        
        <div>
          <h2 class="card-title">Recent Activity</h2>
          <span class="card-meta" id="activity-meta-count">Loading...</span>
        </div>
        
        <div style="display:flex; gap:16px; align-items:center; flex-wrap:wrap; justify-content:flex-end; flex:1;">
          
          <div style="display:flex; gap:8px; align-items:center;">
            
            <div id="date-inputs-container" style="display:flex; align-items:center; gap:6px; max-width:0px; opacity:0; overflow:hidden; transition:all 0.3s ease-out; white-space:nowrap; pointer-events:none;">
              <input type="date" id="filter-start-date" class="form-input" style="padding:6px 10px; font-size:13px;">
              <span style="font-size:13px; color:#666; font-weight:600;">to</span>
              <input type="date" id="filter-end-date" class="form-input" style="padding:6px 10px; font-size:13px;">
            </div>

            <label style="display:flex; align-items:center; gap:4px; font-size:13px; font-weight:600; cursor:pointer; color:#0D1B3E; margin:0; white-space:nowrap;">
              <input type="checkbox" id="toggle-date-filter" style="cursor:pointer; margin:0; width:14px; height:14px;">
              Date
            </label>
          </div>

          <select id="filter-agent" class="form-input" style="padding:6px 24px 6px 10px; font-size:13px; min-width:130px; max-width:160px;">
            <option value="all">All Agents</option>
            ${uniqueAgents.map((a) => `<option value="${escHtml(a)}">${escHtml(a)}</option>`).join("")}
          </select>

          <select id="filter-action" class="form-input" style="padding:6px 24px 6px 10px; font-size:13px; min-width:130px; max-width:160px;">
            <option value="all">All Actions</option>
            ${uniqueActions.map((a) => `<option value="${escHtml(a)}">${escHtml(a)}</option>`).join("")}
          </select>

          <select id="sort-date" class="form-input" style="padding:6px 24px 6px 10px; font-size:13px; min-width:130px; max-width:160px;">
            <option value="desc">Newest First</option>
            <option value="asc">Oldest First</option>
          </select>

          <div style="display:flex; gap:6px; align-items:center; white-space:nowrap; border-left: 1px solid #e2e8f0; padding-left: 12px; margin-left: 4px;">
            <button id="btn-prev-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">&larr; Prev</button>
            <span id="page-indicator" style="font-family:var(--font-mono); font-size:13px; font-weight:600; color:#0D1B3E; min-width: 50px; text-align: center;">Pg 1</span>
            <button id="btn-next-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">Next &rarr;</button>
          </div>

        </div>
      </div>
      
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Time</th><th>Lead</th><th>Action</th><th>Agent</th><th>Notes</th></tr></thead>
          <tbody id="activity-tbody">
            </tbody>
        </table>
      </div>
    </div>
  `;

  // Internal State & DOM Pointers
  let currentPage = 1;
  const itemsPerPage = 50;

  const tbody = document.getElementById("activity-tbody");
  const prevBtn = document.getElementById("btn-prev-page");
  const nextBtn = document.getElementById("btn-next-page");
  const pageIndicator = document.getElementById("page-indicator");

  const agentFilter = document.getElementById("filter-agent");
  const actionFilter = document.getElementById("filter-action");
  const dateSort = document.getElementById("sort-date");
  const metaCount = document.getElementById("activity-meta-count");

  // Date filter pointers
  const dateToggle = document.getElementById("toggle-date-filter");
  const dateContainer = document.getElementById("date-inputs-container");
  const startDateFilter = document.getElementById("filter-start-date");
  const endDateFilter = document.getElementById("filter-end-date");

  // The Smart Table Renderer
  function updateTable() {
    const selectedAgent = agentFilter ? agentFilter.value : "all";
    const selectedAction = actionFilter ? actionFilter.value : "all";
    const sortOrder = dateSort ? dateSort.value : "desc";

    const isDateActive = dateToggle ? dateToggle.checked : false;
    const startDateStr = startDateFilter ? startDateFilter.value : "";
    const endDateStr = endDateFilter ? endDateFilter.value : "";

    const startTimestamp =
      isDateActive && startDateStr
        ? new Date(startDateStr + "T00:00:00").getTime()
        : 0;
    const endTimestamp =
      isDateActive && endDateStr
        ? new Date(endDateStr + "T23:59:59").getTime()
        : Infinity;

    // Step A: Filter by Agent, Action, AND Date
    let processedLog = activityLog.filter(function (e) {
      const matchAgent =
        selectedAgent === "all" || getCleanAgentName(e.agent) === selectedAgent;
      const matchAction =
        selectedAction === "all" || (e.action || "").trim() === selectedAction;

      const logTime = new Date(e.timestamp || 0).getTime();
      const matchDate = logTime >= startTimestamp && logTime <= endTimestamp;

      return matchAgent && matchAction && matchDate;
    });

    // Step B: Sort by Date
    processedLog.sort(function (a, b) {
      const timeA = new Date(a.timestamp || 0).getTime();
      const timeB = new Date(b.timestamp || 0).getTime();
      return sortOrder === "desc" ? timeB - timeA : timeA - timeB;
    });

    // Step C: Pagination Math
    const total = processedLog.length;
    const totalPages = Math.max(1, Math.ceil(total / itemsPerPage));

    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    // Update UI Text
    pageIndicator.textContent = `Pg ${currentPage} / ${totalPages}`;
    if (metaCount) metaCount.textContent = `Showing ${total} entries`;

    prevBtn.disabled = currentPage === 1;
    prevBtn.style.opacity = currentPage === 1 ? "0.4" : "1";

    nextBtn.disabled = currentPage === totalPages;
    nextBtn.style.opacity = currentPage === totalPages ? "0.4" : "1";

    // Step D: Slice and Draw HTML
    const startIndex = (currentPage - 1) * itemsPerPage;
    const displayLog = processedLog.slice(
      startIndex,
      startIndex + itemsPerPage,
    );

    if (displayLog.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-state">No activity matches these filters.</td></tr>`;
      return;
    }

    tbody.innerHTML = displayLog
      .map(function (e) {
        return `
        <tr>
          <td class="td-mono">${formatDateTime(e.timestamp)}</td>
          <td>${escHtml(e.leadName || e.leadId || "—")}</td>
          <td><span class="action-badge">${escHtml(e.action || "—")}</span></td>
          <td>${escHtml(getCleanAgentName(e.agent) || "—")}</td>
          <td class="td-notes">${escHtml(e.notes || "")}</td>
        </tr>`;
      })
      .join("");
  }

  // Attach Event Listeners
  if (prevBtn)
    prevBtn.addEventListener("click", () => {
      if (currentPage > 1) {
        currentPage--;
        updateTable();
      }
    });
  if (nextBtn)
    nextBtn.addEventListener("click", () => {
      currentPage++;
      updateTable();
    });

  if (agentFilter)
    agentFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });
  if (actionFilter)
    actionFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });
  if (dateSort)
    dateSort.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });

  if (startDateFilter)
    startDateFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });
  if (endDateFilter)
    endDateFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });

  // THE SMOOTH ANIMATION TOGGLE
  if (dateToggle) {
    dateToggle.addEventListener("change", (e) => {
      if (e.target.checked) {
        // Slide open to the left
        dateContainer.style.maxWidth = "350px";
        dateContainer.style.opacity = "1";
        dateContainer.style.pointerEvents = "auto";
      } else {
        // Slide closed to the right
        dateContainer.style.maxWidth = "0px";
        dateContainer.style.opacity = "0";
        dateContainer.style.pointerEvents = "none";

        // Clear values and reset table
        if (startDateFilter) startDateFilter.value = "";
        if (endDateFilter) endDateFilter.value = "";
        currentPage = 1;
        updateTable();
      }
    });
  }

  // Initialize on first load
  updateTable();
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
  clone.getElementById("f-phone").value = safeVal(lead?.cbr);
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

async function exportD2DLeads() {
  // Set up our diagnostic counters
  let workedBy1 = 0;
  let workedBy2 = 0;
  let workedBy3Plus = 0;

  // 1. Filter the master list
  const d2dLeads = (State.leads || []).filter((l) => {
    let count = 0;

    if (Array.isArray(l.previousAgents)) {
      count = l.previousAgents.length;
    } else if (typeof l.previousAgents === "string") {
      count = l.previousAgents
        .split(",")
        .map((a) => a.trim())
        .filter((a) => a !== "").length;
    } else {
      count = parseInt(l.previousAgents) || 0;
    }

    if (count === 1) workedBy1++;
    else if (count === 2) workedBy2++;
    else if (count >= 3) workedBy3Plus++;

    // 🛡️ THE NEW SAFEGUARD: Check if the lead is currently unassigned
    // (Note: If your database uses "agent" or "assignedAgent" instead of "assignedTo", just swap the name below!)
    const isUnassigned = !l.assignedTo || l.assignedTo.trim() === "";

    // Return true ONLY if they have 2+ touches AND are currently unassigned
    return count >= 3 && isUnassigned;
  });

  console.log("--- D2D AGENT TOUCH COUNTS ---");
  console.log(`Leads worked by exactly 1 agent: ${workedBy1}`);
  console.log(`Leads worked by exactly 2 agents: ${workedBy2}`);
  console.log(`Leads worked by 3+ agents: ${workedBy3Plus}`);
  console.log("------------------------------");

  if (d2dLeads.length === 0) {
    UI.showToast("No leads meet the criteria for D2D export.", "warning");
    return;
  }

  // 2. Set up the exact requested headers
  const headers = [
    "First Name",
    "Last Name",
    "Address",
    "City",
    "State",
    "BTN",
    "CBR",
    "currentMRC",
    "currentProducts",
  ];

  // 3. Map the data
  const rows = d2dLeads.map((l) => {
    const firstName = l.firstName || "";
    const lastName = l.lastName || "";
    const address = l.address || "";
    const city = l.city || "";
    const state = l.state || "";
    const btn = l.BTN || l.btn || l.phone || "";
    const cbr = l.CBR || l.cbr || l.altPhone || "";
    const mrc = l.currentMRC || l.mrc || "";
    const products = l.currentProducts || l.products || "";

    return `"${firstName}","${lastName}","${address}","${city}","${state}","${btn}","${cbr}","${mrc}","${products}"`;
  });

  // 4. Build & Trigger CSV Download
  const csvContent = headers.join(",") + "\n" + rows.join("\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute(
    "download",
    `D2D_Export_${new Date().toLocaleDateString().replace(/\//g, "-")}.csv`,
  );

  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  // 🔔 Notify the user that the file downloaded and the database is updating
  UI.showToast(
    `📁 File downloaded! Moving ${d2dLeads.length} leads to D2D status...`,
    "info",
  );

  // ==========================================
  // 🚀 5. THE GRAPH API BATCH UPDATER
  // ==========================================
  const host = Config.sharePoint.hostname;
  const sitePath = Config.sharePoint.sites.team;
  const listId = Config.sharePoint.lists.leadsList;
  const batchUrl = `${Config.sharePoint.graphBase}/$batch`;

  const batchSize = 20;
  let updateCount = 0;

  for (let i = 0; i < d2dLeads.length; i += batchSize) {
    const chunk = d2dLeads.slice(i, i + batchSize);

    const batchRequests = chunk.map((lead, index) => {
      return {
        id: String(index + 1),
        method: "PATCH", // 🎯 PATCH updates an existing item without overwriting other columns
        url: `/sites/${host}:/${sitePath}:/lists/${listId}/items/${lead.id}`, // Notice we target the specific lead.id here
        headers: { "Content-Type": "application/json" },
        body: {
          fields: {
            Status: "D2D Lead", // 🎯 The terminal status update
          },
        },
      };
    });

    try {
      await Graph.apiFetch(batchUrl, "POST", { requests: batchRequests });
      updateCount += chunk.length;
    } catch (error) {
      console.error("❌ Batch update failed:", error);
    }
  }

  // 6. Final success message and UI refresh!
  UI.showToast(
    `✅ Successfully locked ${updateCount} leads as D2D!`,
    "success",
  );

  // Reloading the data applies your new filter, instantly wiping them off the screen
  await loadAllData();
}

function updateLeadDraft(leadId, fieldName, value) {
  // If this lead doesn't have a draft object yet, create one
  if (!State.drafts[leadId]) {
    State.drafts[leadId] = {};
  }
  // Save the keystroke
  State.drafts[leadId][fieldName] = value;
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

    let c = document.getElementById("toast-container");

    if (!c) {
      c = document.createElement("div");
      c.id = "toast-container";

      // THE FIX: Nuke the z-index to a billion, and ensure absolute highest priority
      c.style.cssText =
        "position: fixed !important; bottom: 20px !important; right: 20px !important; z-index: 2147483647 !important; display: flex !important; flex-direction: column !important; gap: 10px !important; pointer-events: none !important;";

      // THE FIX: Attach it directly to the HTML document element, bypassing the body entirely
      document.documentElement.appendChild(c);
    }

    const t = document.createElement("div");
    t.className = "toast toast-" + type;
    t.textContent = msg;
    t.style.pointerEvents = "auto";

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
