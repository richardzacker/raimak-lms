// Raimak LMS - Graph API / SharePoint Data Layer v3.0

const Graph = (() => {
  const base = Config.sharePoint.graphBase;
  const host = Config.sharePoint.hostname;
  const lists = Config.sharePoint.lists;
  let siteIds = { leadship: null, team: null };
  let agentCache = null;

  // ── Generic Fetch ──────────────────────────────────────────
  async function apiFetch(url, options = {}) {
    // 1. Auth Check
    const token = await Auth.getToken();
    if (!token) {
      console.warn("No auth token available — redirecting to sign in.");
      Auth.signIn();
      return null;
    }

    // 2. Setup Default Headers & Merge Custom Ones
    const headers = {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
      // 🚀 THE PREFERENCE HEADER: Tells SharePoint to try harder on large lists
      // Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      ...options.headers, // This allows updateLead to inject "If-Match": "*"
    };

    // 3. Prepare Fetch Options
    const method = options.method || "GET";
    const fetchOpts = {
      method: method,
      headers: headers,
    };

    // If a body was passed, ensure it is stringified
    if (options.body) {
      fetchOpts.body =
        typeof options.body === "string"
          ? options.body
          : JSON.stringify(options.body);
    }

    const maxRetries = options.maxRetries || 3;

    // 4. Execution Loop with Retry Logic
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      const res = await fetch(url, fetchOpts);

      if (res.ok) {
        if (res.status === 204) return null;
        return res.json();
      }

      // Intercept 429 Throttling
      if (res.status === 429) {
        const retryAfterStr = res.headers.get("Retry-After");
        const waitMs = retryAfterStr
          ? parseInt(retryAfterStr) * 1000
          : Math.pow(2, attempt) * 1000;

        console.warn(
          `Graph API Throttled! Pausing for ${waitMs}ms (Attempt ${attempt} of ${maxRetries})`,
        );
        await new Promise((resolve) => setTimeout(resolve, waitMs));
        continue;
      }

      // Handle other errors
      const err = await res.json().catch(() => ({}));
      throw new Error((err.error && err.error.message) || "HTTP " + res.status);
    }

    throw new Error("Microsoft Graph request failed after max retries.");
  }

  // ── Resolve Site IDs ───────────────────────────────────────
  async function resolveSiteIds() {
    if (siteIds.leadship && siteIds.team) return;

    try {
      const [s1, s2] = await Promise.all([
        // 🚀 Pattern Shift: Passing an empty object or { method: "GET" }
        // ensures it hits the upgraded apiFetch signature correctly.
        apiFetch(
          base + "/sites/" + host + ":/" + Config.sharePoint.sites.leadship,
          { method: "GET" },
        ),
        apiFetch(
          base + "/sites/" + host + ":/" + Config.sharePoint.sites.team,
          { method: "GET" },
        ),
      ]);

      siteIds.leadship = s1.id;
      siteIds.team = s2.id;
    } catch (err) {
      console.error(
        "Critical Error: Could not resolve SharePoint Site IDs.",
        err,
      );
      throw err;
    }
  }

  // ── Build Agent ID Cache ───────────────────────────────────
  async function resolveAgentCache() {
    if (agentCache) return agentCache;
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.contractorList +
      "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    agentCache = {};
    raw.forEach(function (item) {
      const name = (
        item.fields &&
        (item.fields.Title || item.fields.ContractorName || "")
      )
        .toLowerCase()
        .trim();
      if (name) agentCache[name] = parseInt(item.id, 10);
    });
    return agentCache;
  }

  async function resolveAgentId(agentName) {
    if (!agentName) return null;
    const cache = await resolveAgentCache();
    return cache[agentName.toLowerCase().trim()] || null;
  }

  async function assignAgent(itemId, agentName) {
    await updateLead(itemId, { Agent_x0020_Assigned: agentName });
  }

  // ── Paginate ───────────────────────────────────────────────
  async function getAllItems(url) {
    const items = [];
    let next = url;

    try {
      while (next) {
        // 🚀 Pattern Shift: Explicitly passing the method in an options object
        const data = await apiFetch(next, { method: "GET" });

        if (data && data.value) {
          items.push(...data.value);
        }

        next = data["@odata.nextLink"] || null;
      }
    } catch (error) {
      console.error("❌ Pagination failed on URL:", next, error);
      if (window.UI && UI.showToast) {
        UI.showToast("Network interrupted. Partial data loaded.", "warning");
      }
    }

    return items;
  }

  // ============================================================
  //  LEADS
  // ============================================================

  async function getLeads(lastSyncDate = null, existingLeads = []) {
    await resolveSiteIds();

    // 🚀 STEP 1: If RAM is empty (F5 refresh), load from IndexedDB instantly
    if (!existingLeads || existingLeads.length === 0) {
      existingLeads = await LocalDB.getAllItems("leads");

      // Apply your business rules to the cached data immediately
      existingLeads = existingLeads.filter(
        (lead) => lead.status !== "D2D Lead" && lead.status !== "TDM Non-Reg",
      );
    }

    const expandQuery = "expand=fields";
    let url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      `/items?${expandQuery}&$select=id,lastModifiedDateTime,createdDateTime&$top=5000`;

    // 🚀 STEP 2: The Delta Filter (Modified vs Created)
    if (lastSyncDate && typeof lastSyncDate === "string") {
      const safeDate = lastSyncDate.split(".")[0] + "Z";
      url += `&$filter=fields/Modified gt '${safeDate}'`;
    }

    const raw = await getAllItems(url);

    // If no new/modified leads, just return what we already have
    if (raw.length === 0) {
      return existingLeads;
    }

    // 🚀 STEP 3: Normalize the new/updated leads
    const updatedBatch = raw.map(normalizeLeadItem);

    // 🚀 STEP 4: Save to IndexedDB (Upsert)
    // IndexedDB will automatically overwrite old versions of these leads
    // because they share the same 'id' primary key.
    await LocalDB.saveItems("leads", updatedBatch);

    // 🚀 STEP 5: Smart Merge for the UI
    // We turn the existing list into a Map for O(1) lookups
    const leadMap = new Map();
    existingLeads.forEach((l) => leadMap.set(l.id, l));

    // Add or Replace with the updated data
    updatedBatch.forEach((l) => leadMap.set(l.id, l));

    // Convert back to array and re-apply filters (in case a status changed to D2D)
    const finalizedLeads = Array.from(leadMap.values()).filter(
      (lead) => lead.status !== "D2D Lead" && lead.status !== "TDM Non-Reg",
    );

    // 🚀 STEP 6: Update the Leads Sync Date
    const validTimestamps = updatedBatch
      .map((l) => new Date(l.modified || l.lastModifiedDateTime).getTime())
      .filter((t) => !isNaN(t));

    if (validTimestamps.length > 0) {
      const maxTime = Math.max(...validTimestamps);
      const newSyncDate = new Date(maxTime).toISOString();
      localStorage.setItem("RaimakLeadsLastSyncDate", newSyncDate);
    }

    return finalizedLeads;
  }

  async function getNextLeadForAgent(agentEmail) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    const leads = raw.map(normalizeLeadItem);
    return (
      leads.find(
        (l) => l.status === "New" && !l.assignedTo && !isInCoolOff(l),
      ) || null
    );
  }

  async function addLead(fields) {
    await resolveSiteIds();
    const url =
      base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items";

    const res = await apiFetch(url, {
      method: "POST",
      body: { fields },
    });

    return normalizeLeadItem(res);
  }

  async function updateLead(itemId, fields) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      "/items/" +
      itemId +
      "/fields";

    // 🚀 THE FIX: We package the request into an options object
    // This allows us to pass the specific header needed to kill 409 errors.
    const options = {
      method: "PATCH",
      body: JSON.stringify(fields),
      headers: {
        "Content-Type": "application/json",
        // 🛡️ THE NUCLEAR OPTION:
        // Tells SharePoint "Overwrite this no matter what version is on the server."
        "If-Match": "*",
      },
    };

    return await apiFetch(url, options);
  }

  async function deleteLead(itemId) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      "/items/" +
      itemId;

    // 🚀 Pattern Shift: Using the options object with If-Match to prevent version conflicts
    await apiFetch(url, {
      method: "DELETE",
      headers: {
        "If-Match": "*",
      },
    });
  }

  function normalizeLeadItem(item) {
    const f = item.fields || {};
    const first = f.FirstName || f.First_x0020_Name || "";
    const last = f.LastName || f.Last_x0020_Name || "";
    const name = (first + " " + last).trim() || f.Title || f.LeadName || "";
    return {
      id: item.id,
      name: name,
      firstName: first,
      lastName: last,
      email: f.Email || f.EmailAddress || "",
      phone: f.Phone || f.PhoneNumber || "",
      status: f.Status || "New",
      source: f.Campaign || f.LeadSource || f.Source || "",
      assignedTo:
        f.Agent_x0020_Assigned ||
        f.AgentAssigned ||
        f.AssignedTo ||
        f.Agent ||
        "",
      notes: f.Notes || "",
      address: f.WorkAddress || f.Address || "",
      city: f.WorkCity || f.City || "",
      state: f.State || "",
      zip: f.Zip || f.ZipCode || "",
      cbr: f.CBR || "",
      btn: f.BTN || "",
      lockFlag: f.LockFlag || false,
      callbackAt: f.CallbackDateTime || null,
      lastContacted: f.LastTouchedOn || f.LastContacted || null,
      createdAt: item.createdDateTime || f.Created || null,
      modified: item.lastModifiedDateTime || null,
      leadType:
        f.Lead_x0020_Type || f.Type || f.Item_x0020_Type || f.LeadType || "",
      currentMRC:
        f.MonthlyRecurringCharge_x0028_MRC || f.CurrentMRC || f.MRC || "",
      currentProducts: f.CurrentProducts || "",
      autoPay: f.AutoPay || "",
      previousAgents: f.PreviousAgents || "",
    };
  }

  // ============================================================
  //  CONTRACTORS / AGENTS
  // ============================================================

  async function getContractors() {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.contractorList +
      "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    return raw.map((item) => {
      const f = item.fields || {};
      return {
        id: item.id,
        name: f.Title || f.ContractorName || "",
        email: f.Email || "",
        phone: f.Phone || "",
        role: f.Role || "Agent",
        active: f.Active !== undefined ? f.Active : true,
      };
    });
  }

  async function getAgentScores() {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores + // <-- Now using the dynamic config ID
      "/items?expand=fields($select=AgentEmail,AgentName,CurrentPoints,LifetimePoints)&$top=500";

    const raw = await getAllItems(url);

    return raw.map((item) => {
      const f = item.fields || {};
      return {
        id: item.id,
        AgentEmail: f.AgentEmail || "",
        AgentName: f.AgentName || "",
        CurrentPoints:
          typeof f.CurrentPoints === "number" ? f.CurrentPoints : 0,
        LifetimePoints:
          typeof f.LifetimePoints === "number" ? f.LifetimePoints : 0,
      };
    });
  }

  async function createAgentScore(email, name) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores +
      "/items";

    const payload = {
      fields: {
        AgentEmail: email,
        AgentName: name,
        CurrentPoints: 0,
        LifetimePoints: 0,
      },
    };

    // 🚀 Pattern Shift: Using the options object for the POST request
    const res = await apiFetch(url, {
      method: "POST",
      body: payload,
    });

    return {
      id: res.id,
      AgentEmail: email,
      AgentName: name,
      CurrentPoints: 0,
      LifetimePoints: 0,
    };
  }

  async function checkLedgerForDuplicate(leadId, actionType) {
    if (!leadId || !actionType) return false;

    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScoresLedger +
      `/items?$expand=fields&$filter=fields/LeadID eq '${leadId}' and fields/ActionType eq '${actionType}'`;

    try {
      // 🚀 Pattern Shift: Using the options object for the GET request
      const res = await apiFetch(url, { method: "GET" });

      if (res && res.value && res.value.length > 0) {
        return true;
      }
      return false;
    } catch (err) {
      console.error("Ledger Check Error:", err);
      // Safe-fail: Assume duplicate if check fails to prevent double-point awarding
      return true;
    }
  }

  // ── THE RECEIPT: Write the transaction to the Ledger ──
  async function writeLedgerTransaction(
    agentEmail,
    actionType,
    pointValue,
    leadId = "",
  ) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScoresLedger +
      "/items";

    // Generate a unique transaction ID
    const transactionId = `${agentEmail}_${actionType}_${Date.now()}`;

    const payload = {
      fields: {
        Title: transactionId,
        AgentEmail: agentEmail,
        ActionType: actionType,
        PointValue: pointValue,
        LeadID: leadId,
      },
    };

    // 🚀 Pattern Shift: Using the options object for the POST request
    await apiFetch(url, {
      method: "POST",
      body: payload,
    });
  }

  async function updateAgentScore(itemId, currentPoints, lifetimePoints) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores +
      "/items/" +
      itemId +
      "/fields";

    const payload = {
      CurrentPoints: currentPoints,
      LifetimePoints: lifetimePoints,
    };

    // 🚀 Pattern Shift: Using the options object with If-Match to prevent score conflicts
    await apiFetch(url, {
      method: "PATCH",
      body: payload,
      headers: {
        "If-Match": "*",
      },
    });
  }

  // ============================================================
  //  ACTIVITY LOG
  // ============================================================
  async function getActivityLog(
    lastSyncDate = null,
    existingLogs = [],
    isDeltaRefresh = false,
  ) {
    await resolveSiteIds();

    // 🚀 STEP 1: If RAM is empty (F5 refresh), try to load from the Local Database first
    if (!existingLogs || existingLogs.length === 0) {
      existingLogs = await LocalDB.getAllItems("activity_logs");

      // Sort them newest first so the UI stays consistent
      existingLogs.sort(
        (a, b) => new Date(b.timestamp) - new Date(a.timestamp),
      );

      if (existingLogs.length > 0) {
        UI.showToast(
          `🚀 Loaded ${existingLogs.length} logs from local cache.`,
          "success",
        );
      }
    }

    // 🚀 STEP 2: Decide if we actually need to download anything
    // If RAM was empty and the Database was empty, we do a full sync (null)
    // Otherwise, we use the saved lastSyncDate to get the delta
    let effectiveSyncDate = existingLogs.length === 0 ? null : lastSyncDate;

    const selectedFields =
      "LeadID,LeadId,Title,LeadName,ActionType,Action,Activity,AgentEmail,Agent,Notes,Created";
    let url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      `/items?expand=fields($select=${selectedFields})&$select=id,createdDateTime&$top=5000`;

    if (effectiveSyncDate) {
      // The millisecond-scrubbing fix for Microsoft Graph
      const safeDate = effectiveSyncDate.split(".")[0] + "Z";
      url += `&$filter=fields/Created gt '${safeDate}'`;
    }

    const raw = await getAllItems(url);

    // If no new items exist, we're done! Return what we have.
    if (raw.length === 0) {
      if (isDeltaRefresh) UI.showToast("✅ Logs are up to date.", "success");
      return { updatedLogs: existingLogs, newSyncDate: lastSyncDate };
    }

    // 🚀 STEP 3: Map the new items from Microsoft
    const newLogs = raw.map((item) => {
      const f = item.fields || {};
      return {
        id: item.id, // Primary Key for IndexedDB
        leadId: String(f.LeadID || f.LeadId || ""),
        leadName: f.Title || f.LeadName || "",
        action: f.ActionType || f.Action || f.Activity || "",
        agent: f.AgentEmail || f.Agent || "",
        agentEmail: f.AgentEmail || "",
        notes: f.Notes || "",
        timestamp: item.createdDateTime || f.Created || null,
      };
    });

    // 🚀 STEP 4: Save the new items to IndexedDB permanently!
    // This happens in the background so the UI doesn't lag.
    await LocalDB.saveItems("activity_logs", newLogs);

    // 🚀 STEP 5: Merge and Sort
    // New items go to the front, followed by the existing local logs
    const finalizedLogs = [...newLogs.reverse(), ...existingLogs];

    // 🚀 STEP 6: Update the Sync Date (High-Water Mark)
    const validTimestamps = newLogs
      .map((log) => new Date(log.timestamp).getTime())
      .filter((time) => !isNaN(time));

    let newLastSyncDate = lastSyncDate;
    if (validTimestamps.length > 0) {
      const maxTime = Math.max(...validTimestamps);
      newLastSyncDate = new Date(maxTime).toISOString();
      localStorage.setItem("RaimakActivityLastSyncDate", newLastSyncDate);
    }

    if (effectiveSyncDate && isDeltaRefresh) {
      UI.showToast(`✅ Synced ${newLogs.length} new logs.`, "success");
    }

    return {
      updatedLogs: finalizedLogs,
      newSyncDate: newLastSyncDate,
    };
  }

  async function getActivityLogForToday() {
    await resolveSiteIds();

    const selectedFields =
      "LeadID,LeadId,Title,LeadName,ActionType,Action,Activity,AgentEmail,Agent,Notes,Created";

    const url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      `/items?expand=fields($select=${selectedFields})&$select=id,createdDateTime&$top=5000`;

    let next = url;
    const todayStr = new Date().toDateString();
    const todayLogs = [];

    try {
      while (next) {
        // 🚀 Pattern Shift: Method and Headers wrapped into the options object
        const response = await apiFetch(next, {
          method: "GET",
          headers: {
            Prefer: "allow-throttleable-queries",
          },
        });

        const items = response.value || [];

        for (const item of items) {
          const f = item.fields || {};
          const timestamp = item.createdDateTime || f.Created;

          if (!timestamp) continue;

          if (new Date(timestamp).toDateString() === todayStr) {
            todayLogs.push({
              id: item.id,
              leadId: String(f.LeadID || f.LeadId || ""),
              leadName: f.Title || f.LeadName || "",
              action: f.ActionType || f.Action || f.Activity || "",
              agent: f.AgentEmail || f.Agent || "",
              agentEmail: f.AgentEmail || "",
              notes: f.Notes || "",
              timestamp: timestamp,
            });
          }
        }

        next = response["@odata.nextLink"] || null;
      }
    } catch (err) {
      console.error("Failed to fetch today's logs:", err);
    }

    // Newest logs at index 0 for the UI
    return todayLogs.reverse();
  }

  async function logActivity(entry) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.leadship + // 🚀 Using your 'leadship' site ID
      "/lists/" +
      lists.activityLog +
      "/items";

    // 🚀 THE FIX: Wrap the method and body into a single options object
    // to match your upgraded apiFetch signature.
    await apiFetch(url, {
      method: "POST",
      body: {
        fields: entry,
      },
    });
  }

  // Get today's sold leads based on activity log entries.
  // Fetches leads directly to avoid timing issues with State.leads.
  function getTodaySales(todayLogs = []) {
    // 1. THE MAPPER: Build the translation dictionary from contractors
    const nameLookup = {};
    (State.contractors || []).forEach((c) => {
      if (c.email) nameLookup[c.email.toLowerCase().trim()] = c.name;
    });

    const sales = [];
    const seenLeadIds = new Set(); // The Bouncer
    const todayStr = new Date().toDateString(); // Get today's date for filtering

    todayLogs.forEach(function (e) {
      // 🛑 CRITICAL: Make sure the log actually happened today!
      const isToday =
        e.timestamp && new Date(e.timestamp).toDateString() === todayStr;

      if (isToday && e.action === "Status: " + Config.soldStatus && e.leadId) {
        // Only count the sale if we haven't seen this exact lead ID yet today
        if (!seenLeadIds.has(e.leadId)) {
          seenLeadIds.add(e.leadId);

          // 2. THE INTERCEPT: Translate the raw agent ID before saving it
          const rawAgent = e.agent || "Unknown";
          const displayName =
            nameLookup[rawAgent.toLowerCase().trim()] || rawAgent;

          sales.push({
            id: e.leadId,
            name: e.leadName || "Unknown Lead",
            soldBy: displayName, // Saves "John Doe" instead of "jdoe@..."
            assignedTo: displayName, // Saves "John Doe" instead of "jdoe@..."
            modified: e.timestamp,
            saleTime: e.timestamp,
          });
        }
      }
    });

    return sales;
  }

  // Get daily activity stats per agent for the report.
  // Maps agent emails back to display names via the contractors list.
  async function getDailyStats() {
    const log = State.activityLog || [];
    const today = new Date().toDateString();

    const todayEntries = log.filter(
      (e) => e.timestamp && new Date(e.timestamp).toDateString() === today,
    );

    const emailToName = {};
    (State.contractors || []).forEach((c) => {
      if (c.email) emailToName[c.email.toLowerCase().trim()] = c.name;
    });

    const stats = {};
    for (const entry of todayEntries) {
      const agentEmail = (entry.agent || "").toLowerCase().trim();
      const agent = emailToName[agentEmail] || entry.agent || "Unknown";

      if (!stats[agent]) {
        stats[agent] = {
          agent,
          actions: [],
          uniqueLeads: new Set(),
          uniqueSales: new Set(),
        };
      }

      const isContact =
        entry.action &&
        (entry.action.indexOf("Status:") === 0 ||
          entry.action.includes("Contact"));
      if (isContact && entry.leadId) stats[agent].uniqueLeads.add(entry.leadId);
      if (entry.action === "Status: " + Config.soldStatus && entry.leadId)
        stats[agent].uniqueSales.add(entry.leadId);

      stats[agent].actions.push(entry);
    }

    return Object.values(stats)
      .map(function (s) {
        // 🚀 CALCULATE AVERAGE CADENCE
        // Sort actions by time (oldest to newest)
        const sorted = s.actions.sort(
          (a, b) => new Date(a.timestamp) - new Date(b.timestamp),
        );
        let totalDiff = 0;
        let intervalCount = 0;

        for (let i = 1; i < sorted.length; i++) {
          const diff =
            new Date(sorted[i].timestamp) - new Date(sorted[i - 1].timestamp);

          // Only count intervals under 30 minutes (exclude lunch/long breaks)
          if (diff > 0 && diff < 30 * 60 * 1000) {
            totalDiff += diff;
            intervalCount++;
          }
        }

        const avgMs = intervalCount > 0 ? totalDiff / intervalCount : 0;
        const mins = Math.floor(avgMs / 60000);
        const secs = Math.floor((avgMs % 60000) / 1000);

        s.avgTime = avgMs > 0 ? `${mins}m ${secs}s` : "—";
        s.contacts = s.uniqueLeads.size;
        s.sold = s.uniqueSales.size;

        delete s.uniqueLeads;
        delete s.uniqueSales;
        return s;
      })
      .sort((a, b) => b.contacts - a.contacts);
  }

  // ============================================================
  //  BUSINESS RULES
  // ============================================================

  function applyBusinessRules(leads, contractors) {
    const now = new Date();
    const { coolOffDays, maxLeadsPerAgent, recycleAfterDays } = Config.rules;

    const agentCounts = {};
    for (const lead of leads) {
      if (lead.assignedTo && !Config.terminalStatuses.includes(lead.status)) {
        agentCounts[lead.assignedTo] = (agentCounts[lead.assignedTo] || 0) + 1;
      }
    }

    return leads.map(function (lead) {
      const flags = [];

      if (lead.lastContacted) {
        const daysSince = (now - new Date(lead.lastContacted)) / 86400000;
        if (
          daysSince < coolOffDays &&
          !Config.terminalStatuses.includes(lead.status)
        ) {
          flags.push("cool_off");
        }
        if (lead.status === "3rd Contact" && daysSince >= coolOffDays) {
          flags.push("needs_recycle");
        }
      }

      const ref = lead.lastContacted || lead.createdAt;
      if (
        ref &&
        !Config.terminalStatuses.includes(lead.status) &&
        lead.status !== "3rd Contact"
      ) {
        const daysSince = (now - new Date(ref)) / 86400000;
        if (daysSince > recycleAfterDays) flags.push("needs_recycle");
      }

      if (lead.assignedTo && agentCounts[lead.assignedTo] > maxLeadsPerAgent) {
        flags.push("agent_overloaded");
      }

      return Object.assign({}, lead, {
        flags: flags,
        agentLeadCount: agentCounts[lead.assignedTo] || 0,
      });
    });
  }

  function canAgentTakeLead(agentName, leads) {
    const count = leads.filter(function (l) {
      return (
        l.assignedTo === agentName &&
        !Config.terminalStatuses.includes(l.status)
      );
    }).length;
    return count < Config.rules.maxLeadsPerAgent;
  }

  // Recycle a lead — record previous agent, unassign, reset to New
  async function recycleLead(leadId, currentAgent) {
    await resolveSiteIds();
    const lead = State
      ? State.leads.find(function (l) {
          return l.id === leadId;
        })
      : null;
    const prev = lead ? lead.previousAgents || "" : "";
    const newPrev = prev ? prev + ", " + currentAgent : currentAgent;
    await updateLead(leadId, {
      Status: "New",
      Agent_x0020_Assigned: null,
      PreviousAgents: newPrev,
      LastTouchedOn: null,
    });
  }

  function isInCoolOff(lead) {
    if (!lead.lastContacted) return false;
    const daysSince = (new Date() - new Date(lead.lastContacted)) / 86400000;
    return daysSince < Config.rules.coolOffDays;
  }

  // Count unique leads an agent contacted today.
  // Accepts email directly for reliable matching against activity log.
  function agentContactsToday(agentEmail, activityLog) {
    const emailLower = (agentEmail || "").toLowerCase().trim();
    const uniqueLeads = new Set();

    activityLog.forEach(function (e) {
      // Safety check: Handle both your mapped frontend names AND raw Graph API names
      const entryAgent = (e.agent || e.AgentEmail || "").toLowerCase().trim();
      const actionStr = e.action || e.ActionType || "";
      const leadId = e.leadId || e.LeadID;

      // Is it an actual lead touch?
      const isContact =
        actionStr.startsWith("Status:") ||
        actionStr === "1st Contact" ||
        actionStr === "2nd Contact" ||
        actionStr === "3rd Contact";

      // If it's a contact, by this agent, add the lead ID to the unique Set
      if (isContact && entryAgent === emailLower && leadId) {
        uniqueLeads.add(leadId);
      }
    });

    // A Set automatically prevents duplicates, so its size is the exact number of unique leads worked!
    return uniqueLeads.size;
  }

  return {
    apiFetch,
    getLeads,
    addLead,
    updateLead,
    deleteLead,
    assignAgent,
    recycleLead,
    getNextLeadForAgent,
    getContractors,
    getActivityLog,
    getActivityLogForToday,
    logActivity,
    getTodaySales,
    getDailyStats,
    resolveSiteIds,
    applyBusinessRules,
    canAgentTakeLead,
    isInCoolOff,
    agentContactsToday,
    getAgentScores,
    createAgentScore,
    checkLedgerForDuplicate,
    writeLedgerTransaction,
    updateAgentScore,
  };
})();
