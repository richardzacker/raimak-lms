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
    await updateLead(itemId, { Agent_x0020_Assigned: agentName });
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
      currentMRC:      f.MonthlyRecurringCharge_x0028_MRC || f.CurrentMRC || f.MRC || "",
      currentProducts: f.CurrentProducts || "",
      autoPay:         f.AutoPay         || "",
      previousAgents:  f.PreviousAgents  || "",
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
        leadId:    f.LeadID    || f.LeadId    || "",
        leadName:  f.Title     || f.LeadName  || "",
        action:    f.ActionType || f.Action   || f.Activity || "",
        agent:     f.AgentEmail || f.Agent    || "",
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
    const log   = await getActivityLog(500);
    const today = new Date().toDateString();
    const todayEntries = log.filter(function(e) { return e.timestamp && new Date(e.timestamp).toDateString() === today; });

    const stats = {};
    for (const entry of todayEntries) {
      const agent     = entry.agent || "Unknown";
      const isContact = entry.action && (
        entry.action.indexOf("Status:") === 0 ||
        entry.action === "1st Contact" ||
        entry.action === "2nd Contact" ||
        entry.action === "3rd Contact"
      );
      if (!stats[agent]) stats[agent] = { agent, contacts: 0, sold: 0, actions: [], uniqueLeads: new Set() };
      if (isContact && entry.leadId) stats[agent].uniqueLeads.add(entry.leadId);
      if (entry.action === "Status: " + Config.soldStatus) stats[agent].sold++;
      stats[agent].actions.push(entry);
    }
    // Convert Set size to contacts count
    Object.values(stats).forEach(function(s) {
      s.contacts = s.uniqueLeads.size;
      delete s.uniqueLeads;
    });
    return Object.values(stats).sort(function(a, b) { return b.contacts - a.contacts; });
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

    return leads.map(function(lead) {
      const flags = [];

      if (lead.lastContacted) {
        const daysSince = (now - new Date(lead.lastContacted)) / 86400000;
        if (daysSince < coolOffDays && !Config.terminalStatuses.includes(lead.status)) {
          flags.push("cool_off");
        }
        // 3rd Contact leads that have cooled off 48hrs+ need recycling
        if (lead.status === "3rd Contact" && daysSince >= coolOffDays) {
          flags.push("needs_recycle");
        }
      }

      // Also flag any lead inactive longer than recycleAfterDays
      const ref = lead.lastContacted || lead.createdAt;
      if (ref && !Config.terminalStatuses.includes(lead.status) && lead.status !== "3rd Contact") {
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

  // Recycle a lead — record previous agent, unassign, reset to New
  async function recycleLead(leadId, currentAgent) {
  await resolveSiteIds();
  const lead = State ? State.leads.find(function(l){ return l.id === leadId; }) : null;
  const prev = lead ? (lead.previousAgents || "") : "";
  const newPrev = prev ? prev + ", " + currentAgent : currentAgent;
  await updateLead(leadId, {
    Status:               "New",
    Agent_x0020_Assigned: null,
    PreviousAgents:       newPrev,
    LastTouchedOn:        null,
  });
}

  function isInCoolOff(lead) {
    if (!lead.lastContacted) return false;
    const daysSince = (new Date() - new Date(lead.lastContacted)) / 86400000;
    return daysSince < Config.rules.coolOffDays;
  }

  // Count unique leads an agent contacted today (status changes only)
  function agentContactsToday(agentName, activityLog) {
    const today      = new Date().toDateString();
    const agentLower = (agentName || "").toLowerCase().trim();
    const uniqueLeads = new Set();
    activityLog.forEach(function(e) {
      const entryAgent  = (e.agent || "").toLowerCase().trim();
      const isContact   = e.action && (
        e.action.indexOf("Status:") === 0 ||
        e.action === "1st Contact" ||
        e.action === "2nd Contact" ||
        e.action === "3rd Contact"
      );
      if (isContact &&
          entryAgent === agentLower &&
          e.timestamp &&
          new Date(e.timestamp).toDateString() === today &&
          e.leadId) {
        uniqueLeads.add(e.leadId);
      }
    });
    return uniqueLeads.size;
  }

  return {
    getLeads, addLead, updateLead, deleteLead, assignAgent, recycleLead,
    getNextLeadForAgent,
    getContractors,
    getActivityLog, logActivity,
    getTodaySales, getDailyStats,
    applyBusinessRules, canAgentTakeLead, isInCoolOff, agentContactsToday,
  };
})();
