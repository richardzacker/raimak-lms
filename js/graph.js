// Raimak LMS - Graph API / SharePoint Data Layer v3.0

const Graph = (() => {

  const base  = Config.sharePoint.graphBase;
  const host  = Config.sharePoint.hostname;
  const lists = Config.sharePoint.lists;
  let siteIds    = { leadship: null, team: null };
  let agentCache = null; // { name (lowercase) -> numeric sharepoint id }

  // ── Generic Fetch ──────────────────────────────────────────
  async function apiFetch(url, method = "GET", body = null) {
    const token = await Auth.getToken();
    if (!token) throw new Error("Not authenticated");
    const opts = {
      method,
      headers: {
        Authorization: "Bearer " + token,
        "Content-Type": "application/json",
      },
    };
    if (body) opts.body = JSON.stringify(body);
    const res = await fetch(url, opts);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err.error && err.error.message) || ("HTTP " + res.status));
    }
    if (res.status === 204) return null;
    return res.json();
  }

  // ── Resolve Site IDs ───────────────────────────────────────
  async function resolveSiteIds() {
    if (siteIds.leadship && siteIds.team) return;
    const [s1, s2] = await Promise.all([
      apiFetch(base + "/sites/" + host + ":/" + Config.sharePoint.sites.leadship),
      apiFetch(base + "/sites/" + host + ":/" + Config.sharePoint.sites.team),
    ]);
    siteIds.leadship = s1.id;
    siteIds.team     = s2.id;
  }

  // ── Build Agent ID Cache from Contractor & Employee List ───
  async function resolveAgentCache() {
    if (agentCache) return agentCache;
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.contractorList + "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    agentCache = {};
    raw.forEach(function(item) {
      const name = (item.fields && (item.fields.Title || item.fields.ContractorName || "")).toLowerCase().trim();
      if (name) agentCache[name] = parseInt(item.id, 10);
    });
    return agentCache;
  }

  // Resolve an agent name to its numeric SharePoint Lookup ID
  async function resolveAgentId(agentName) {
    if (!agentName) return null;
    const cache = await resolveAgentCache();
    return cache[agentName.toLowerCase().trim()] || null;
  }

  // ── Assign agent on a lead (handles Lookup field correctly) ─
  async function assignAgent(itemId, agentName) {
    const agentId = await resolveAgentId(agentName);
    console.log("assignAgent:", agentName, "→ ID:", agentId);
    if (!agentId) throw new Error("Agent \"" + agentName + "\" not found in Contractor & Employee List.");
    // Try both common SharePoint Lookup ID field name formats
    try {
      await updateLead(itemId, { Agent_x0020_AssignedId: agentId });
    } catch(e) {
      await updateLead(itemId, { Agent_x0020_AssignedLookupId: agentId });
    }
  }

  // ── Paginate ───────────────────────────────────────────────
  async function getAllItems(url) {
    let items = [], next = url;
    while (next) {
      const data = await apiFetch(next);
      items = items.concat(data.value || []);
      next  = data["@odata.nextLink"] || null;
    }
    return items;
  }

  // ============================================================
  //  LEADS
  // ============================================================

  async function getLeads() {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    return raw.map(normalizeLeadItem);
  }

  // Get a single unassigned lead for the agent feed
  async function getNextLeadForAgent(agentEmail) {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    const leads = raw.map(normalizeLeadItem);
    // Find unassigned New lead not in cool-off
    return leads.find(l =>
      l.status === "New" &&
      !l.assignedTo &&
      !isInCoolOff(l)
    ) || null;
  }

  async function addLead(fields) {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items";
    const res = await apiFetch(url, "POST", { fields });
    return normalizeLeadItem(res);
  }

  async function updateLead(itemId, fields) {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items/" + itemId + "/fields";
    await apiFetch(url, "PATCH", fields);
  }

  async function deleteLead(itemId) {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items/" + itemId;
    await apiFetch(url, "DELETE");
  }

  function normalizeLeadItem(item) {
    const f    = item.fields || {};
    if (f.Agent_x0020_Assigned || f.AgentAssigned) console.log("AGENT FIELDS:", JSON.stringify(Object.entries(f).filter(([k]) => k.toLowerCase().includes('agent'))));
    const first = f.FirstName || f.First_x0020_Name || "";
    const last  = f.LastName  || f.Last_x0020_Name  || "";
    const name  = (first + " " + last).trim() || f.Title || f.LeadName || "";
    return {
      id:              item.id,
      name:            name,
      firstName:       first,
      lastName:        last,
      email:           f.Email         || f.EmailAddress || "",
      phone:           f.Phone         || f.PhoneNumber  || "",
      status:          f.Status        || "New",
      source:          f.Campaign      || f.LeadSource   || f.Source || "",
      assignedTo:      f.Agent_x0020_Assigned || f.AgentAssigned || f.AssignedTo || f.Agent || "",
      notes:           f.Notes         || "",
      address:         f.WorkAddress   || f.Address      || "",
      city:            f.WorkCity      || f.City         || "",
      state:           f.State         || "",
      zip:             f.Zip           || f.ZipCode      || "",
      cbr:             f.CBR           || "",
      btn:             f.BTN           || "",
      lockFlag:        f.LockFlag      || false,
      callbackAt:      f.CallbackDateTime || null,
      lastContacted:   f.LastTouchedOn || f.LastContacted || null,
      createdAt:       item.createdDateTime || f.Created || null,
      modified:        item.lastModifiedDateTime || null,
      leadType:        f.Lead_x0020_Type || f.Type || f.Item_x0020_Type || f.LeadType || "",
      currentMRC:      f.CurrentMRC    || "",
      currentProducts: f.CurrentProducts || "",
    };
  }

  // ============================================================
  //  CONTRACTORS / AGENTS
  // ============================================================

  async function getContractors() {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.contractorList + "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    return raw.map(item => {
      const f = item.fields || {};
      return {
        id:     item.id,
        name:   f.Title || f.ContractorName || "",
        email:  f.Email || "",
        phone:  f.Phone || "",
        role:   f.Role  || "Agent",
        active: f.Active !== undefined ? f.Active : true,
      };
    });
  }

  // ============================================================
  //  ACTIVITY LOG
  // ============================================================

  async function getActivityLog(limit) {
    limit = limit || 200;
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.leadship + "/lists/" + lists.activityLog + "/items?expand=fields&$orderby=createdDateTime desc&$top=" + limit;
    const raw = await getAllItems(url);
    return raw.map(item => {
      const f = item.fields || {};
      return {
        id:        item.id,
        leadId:    f.LeadId    || f.LeadID    || "",
        leadName:  f.LeadName  || f.Title     || "",
        action:    f.Action    || f.Activity  || "",
        agent:     f.Agent     || f.AssignedTo || "",
        agentEmail:f.AgentEmail || "",
        notes:     f.Notes     || "",
        timestamp: item.createdDateTime || f.Created || null,
      };
    });
  }

  async function logActivity(entry) {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.leadship + "/lists/" + lists.activityLog + "/items";
    await apiFetch(url, "POST", { fields: entry });
  }

  // Get today's sold leads for live feed
  async function getTodaySales() {
    await resolveSiteIds();
    const url = base + "/sites/" + siteIds.team + "/lists/" + lists.leadsList + "/items?expand=fields&$top=500";
    const raw = await getAllItems(url);
    const today = new Date().toDateString();
    return raw.map(normalizeLeadItem).filter(l => {
      if (l.status !== Config.soldStatus) return false;
      const mod = l.modified ? new Date(l.modified).toDateString() : null;
      return mod === today;
    });
  }

  // Get daily activity stats per agent for the report
  async function getDailyStats() {
    const log = await getActivityLog(500);
    const today = new Date().toDateString();
    const todayEntries = log.filter(e => e.timestamp && new Date(e.timestamp).toDateString() === today);

    // Count contacts per agent today
    const stats = {};
    for (const entry of todayEntries) {
      const agent = entry.agent || "Unknown";
      if (!stats[agent]) stats[agent] = { agent, contacts: 0, sold: 0, actions: [] };
      stats[agent].contacts++;
      if (entry.action === "Status: " + Config.soldStatus) stats[agent].sold++;
      stats[agent].actions.push(entry);
    }
    return Object.values(stats).sort((a, b) => b.contacts - a.contacts);
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

    return leads.map(lead => {
      const flags = [];

      if (lead.lastContacted) {
        const daysSince = (now - new Date(lead.lastContacted)) / 86400000;
        if (daysSince < coolOffDays && !Config.terminalStatuses.includes(lead.status)) {
          flags.push("cool_off");
        }
      }

      const ref = lead.lastContacted || lead.createdAt;
      if (ref && !Config.terminalStatuses.includes(lead.status)) {
        const daysSince = (now - new Date(ref)) / 86400000;
        if (daysSince > recycleAfterDays) flags.push("needs_recycle");
      }

      if (lead.assignedTo && agentCounts[lead.assignedTo] > maxLeadsPerAgent) {
        flags.push("agent_overloaded");
      }

      return Object.assign({}, lead, { flags: flags, agentLeadCount: agentCounts[lead.assignedTo] || 0 });
    });
  }

  function canAgentTakeLead(agentName, leads) {
    const count = leads.filter(function(l) {
      return l.assignedTo === agentName && !Config.terminalStatuses.includes(l.status);
    }).length;
    return count < Config.rules.maxLeadsPerAgent;
  }

  function isInCoolOff(lead) {
    if (!lead.lastContacted) return false;
    const daysSince = (new Date() - new Date(lead.lastContacted)) / 86400000;
    return daysSince < Config.rules.coolOffDays;
  }

  // Count how many leads an agent contacted today
  function agentContactsToday(agentName, activityLog) {
    const today = new Date().toDateString();
    return activityLog.filter(function(e) {
      return e.agent === agentName &&
             e.timestamp &&
             new Date(e.timestamp).toDateString() === today;
    }).length;
  }

  return {
    getLeads, addLead, updateLead, deleteLead, assignAgent,
    getNextLeadForAgent,
    getContractors,
    getActivityLog, logActivity,
    getTodaySales, getDailyStats,
    applyBusinessRules, canAgentTakeLead, isInCoolOff, agentContactsToday,
  };
})();
