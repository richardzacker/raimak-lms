// Raimak LMS - Graph API / SharePoint Data Layer v3.0

const Graph = (() => {
  const base = Config.sharePoint.graphBase;
  const host = Config.sharePoint.hostname;
  const lists = Config.sharePoint.lists;
  let siteIds = { leadship: null, team: null };
  let agentCache = null;

  // ── Generic Fetch ──────────────────────────────────────────
  async function apiFetch(url, method = "GET", body = null, maxRetries = 3) {
    const token = await Auth.getToken();
    if (!token) {
      console.warn("No auth token available — redirecting to sign in.");
      Auth.signIn();
      return null;
    }

    const opts = {
      method,
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
    };
    if (body) opts.body = JSON.stringify(body);

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      const res = await fetch(url, opts);

      if (res.ok) {
        if (res.status === 204) return null;
        return res.json();
      }

      // Intercept 429 Throttling
      if (res.status === 429) {
        // Microsoft provides the wait time in seconds.
        // If missing, we use exponential backoff (2s, 4s, 8s...)
        const retryAfterStr = res.headers.get("Retry-After");
        const waitMs = retryAfterStr
          ? parseInt(retryAfterStr) * 1000
          : 2 ** attempt * 1000;

        console.warn(
          `Graph API Throttled! Pausing for ${waitMs}ms (Attempt ${attempt} of ${maxRetries})`,
        );

        // Pause execution without freezing the browser
        await new Promise((resolve) => setTimeout(resolve, waitMs));
        continue; // Loop restarts and tries the fetch again
      }

      // If it's a different error (401, 404, etc.), throw it normally
      const err = await res.json().catch(() => ({}));
      throw new Error((err.error && err.error.message) || "HTTP " + res.status);
    }

    throw new Error("Microsoft Graph request failed after max retries.");
  }

  // ── Resolve Site IDs ───────────────────────────────────────
  async function resolveSiteIds() {
    if (siteIds.leadship && siteIds.team) return;
    const [s1, s2] = await Promise.all([
      apiFetch(
        base + "/sites/" + host + ":/" + Config.sharePoint.sites.leadship,
      ),
      apiFetch(base + "/sites/" + host + ":/" + Config.sharePoint.sites.team),
    ]);
    siteIds.leadship = s1.id;
    siteIds.team = s2.id;
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
    let items = [],
      next = url;
    while (next) {
      const data = await apiFetch(next);
      items = items.concat(data.value || []);
      next = data["@odata.nextLink"] || null;
    }
    return items;
  }

  // ============================================================
  //  LEADS
  // ============================================================

  async function getLeads() {
    await resolveSiteIds();

    // THE GOLDILOCKS QUERY:
    // expand=fields (with no select) grabs ALL your custom CRM columns.
    // $select=id... at the root blocks all the heavy Microsoft metadata.
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.leadsList +
      `/items?expand=fields&$select=id,lastModifiedDateTime,createdDateTime&$top=2000`;

    const raw = await getAllItems(url);
    return raw.map(normalizeLeadItem);
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
    const res = await apiFetch(url, "POST", { fields });
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
    await apiFetch(url, "PATCH", fields);
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
    await apiFetch(url, "DELETE");
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

    const res = await apiFetch(url, "POST", payload);

    // Return a clean object formatted exactly like getAgentScores does
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

    // We use the OData $filter parameter to ask SharePoint to only return exact matches.
    // Because you indexed these columns, this query will execute in milliseconds.
    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScoresLedger + // <-- Using your new config variable!
      `/items?$expand=fields&$filter=fields/LeadID eq '${leadId}' and fields/ActionType eq '${actionType}'`;

    try {
      const res = await apiFetch(url);

      // If the array has anything in it, a receipt already exists. Return true (Duplicate found!)
      if (res && res.value && res.value.length > 0) {
        return true;
      }
      return false; // Array is empty, the action is fresh!
    } catch (err) {
      console.error("Ledger Check Error:", err);
      // Fail safely: if the check crashes, assume it's a duplicate so we don't accidentally give out infinite money.
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

    // Generate a unique transaction ID (e.g., "jdoe@email.com_SoldLead_1714241234567")
    const transactionId = `${agentEmail}_${actionType}_${Date.now()}`;

    const payload = {
      fields: {
        Title: transactionId, // Using the default Title column we repurposed
        AgentEmail: agentEmail,
        ActionType: actionType,
        PointValue: pointValue,
        LeadID: leadId,
      },
    };

    await apiFetch(url, "POST", payload);
  }

  async function updateAgentScore(itemId, currentPoints, lifetimePoints) {
    await resolveSiteIds();

    const url =
      base +
      "/sites/" +
      siteIds.team +
      "/lists/" +
      lists.agentScores + // The main bank list
      "/items/" +
      itemId +
      "/fields";

    const payload = {
      CurrentPoints: currentPoints,
      LifetimePoints: lifetimePoints,
    };

    await apiFetch(url, "PATCH", payload);
  }

  // ============================================================
  //  ACTIVITY LOG
  // ============================================================
  async function getActivityLog() {
    await resolveSiteIds();
    const selectedFields =
      "LeadID,LeadId,Title,LeadName,ActionType,Action,Activity,AgentEmail,Agent,Notes,Created";
    const url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      `/items?expand=fields($select=${selectedFields})&$select=id,createdDateTime&$top=2000`;

    const raw = await getAllItems(url);

    return raw
      .map((item) => {
        const f = item.fields || {};
        return {
          id: item.id,
          leadId: String(f.LeadID || f.LeadId || ""),
          leadName: f.Title || f.LeadName || "",
          action: f.ActionType || f.Action || f.Activity || "",
          agent: f.AgentEmail || f.Agent || "",
          agentEmail: f.AgentEmail || "",
          notes: f.Notes || "",
          timestamp: item.createdDateTime || f.Created || null,
        };
      })
      .reverse(); // Put newest items at the top
  }

  async function getActivityLogForToday() {
    await resolveSiteIds();

    // THE DIET QUERY: We pull 2000 items per page, but ONLY the specific fields we need.
    // This drops the payload size by over 90%, bypassing the 429 throttling limits.
    const selectedFields =
      "LeadID,LeadId,Title,LeadName,ActionType,Action,Activity,AgentEmail,Agent,Notes,Created";
    const url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      `/items?expand=fields($select=${selectedFields})&$select=id,createdDateTime&$top=2000`;

    // getAllItems automatically pages through the database safely
    const raw = await getAllItems(url);

    const todayStr = new Date().toDateString();
    const todayLogs = [];

    // Brute-force through the massive pile to find today's data
    for (const item of raw) {
      const f = item.fields || {};
      const timestamp = item.createdDateTime || f.Created;

      if (timestamp && new Date(timestamp).toDateString() === todayStr) {
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

    // Graph gave them to us oldest-first. We reverse the array so the live feed
    // gets the newest sales at the very top (index 0).
    return todayLogs.reverse();
  }

  async function logActivity(entry) {
    await resolveSiteIds();
    const url =
      base +
      "/sites/" +
      siteIds.leadship +
      "/lists/" +
      lists.activityLog +
      "/items";
    await apiFetch(url, "POST", { fields: entry });
  }

  // Get today's sold leads based on activity log entries.
  // Fetches leads directly to avoid timing issues with State.leads.
  async function getTodaySales(todayLogs) {
    await resolveSiteIds();

    // If no logs were passed in, fetch them (used by the 30-sec polling timer)
    if (!todayLogs) {
      todayLogs = await getActivityLogForToday();
    }

    // 1. THE MAPPER: Build the translation dictionary from contractors
    const nameLookup = {};
    (State.contractors || []).forEach((c) => {
      if (c.email) nameLookup[c.email.toLowerCase().trim()] = c.name;
    });

    const sales = [];
    const seenLeadIds = new Set(); // The Bouncer

    todayLogs.forEach(function (e) {
      if (e.action === "Status: " + Config.soldStatus && e.leadId) {
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
    const log = await getActivityLog(2000);
    const today = new Date().toDateString();
    const todayEntries = log.filter(function (e) {
      return e.timestamp && new Date(e.timestamp).toDateString() === today;
    });

    // Build email → display name map from contractors
    const emailToName = {};
    (State.contractors || []).forEach(function (c) {
      if (c.email) emailToName[c.email.toLowerCase().trim()] = c.name;
    });

    const stats = {};
    for (const entry of todayEntries) {
      const agentEmail = (entry.agent || "").toLowerCase().trim();
      const agent = emailToName[agentEmail] || entry.agent || "Unknown";

      const isContact =
        entry.action &&
        (entry.action.indexOf("Status:") === 0 ||
          entry.action === "1st Contact" ||
          entry.action === "2nd Contact" ||
          entry.action === "3rd Contact");
      if (!stats[agent])
        stats[agent] = {
          agent,
          contacts: 0,
          sold: 0,
          actions: [],
          uniqueLeads: new Set(),
        };
      if (isContact && entry.leadId) stats[agent].uniqueLeads.add(entry.leadId);
      if (entry.action === "Status: " + Config.soldStatus) stats[agent].sold++;
      stats[agent].actions.push(entry);
    }

    Object.values(stats).forEach(function (s) {
      s.contacts = s.uniqueLeads.size;
      delete s.uniqueLeads;
    });
    return Object.values(stats).sort(function (a, b) {
      return b.contacts - a.contacts;
    });
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
