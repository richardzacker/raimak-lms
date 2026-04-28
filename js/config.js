// Raimak LMS - App Configuration v3.0

const Config = {
  azure: {
    clientId: "ad1b153f-8b6a-4f1c-8ab2-58fcf03cf5c2",
    tenantId: "39e14190-0b23-4ecd-99f9-606ad1215881",
    redirectUri: window.location.origin + window.location.pathname,
  },

  sharePoint: {
    hostname: "raimak.sharepoint.com",
    sites: {
      leadship: "sites/RaimakLeadship",
      team: "TeamSite",
    },
    lists: {
      activityLog: "2adb1260-e635-45cd-bb3b-87dd57a2d022",
      contractorList: "bd5df38a-9cb6-411d-87e8-3e79934213d3",
      leadsList: "5a01419d-e2c9-4aad-8484-6ed97233f305",
      ordersAndInstalls: "9d5c9b0b-10d1-4b15-988c-051ef8117d40",
      productPerformance: "870acf73-e3fd-44b4-98cd-a49e9497fff1",
      agentPerformance: "93b94795-ae98-4134-9cbc-c92618856012",
      statePerformance: "9c202207-343c-41af-ae36-bd3797c3f372",
      operationsHealth: "534d595e-521a-4f02-af02-23e5ee427a74",
      agentScores: "242d56aa-dfcd-4399-919b-0ce608ec932c",
      agentScoresLedger: "7c6f1604-bd68-4ceb-9f0c-9271faf5e0b8",
    },
    graphBase: "https://graph.microsoft.com/v1.0",
  },

  rules: {
    coolOffDays: 2,
    maxLeadsPerAgent: 9999,
    maxContactsPerDay: 9999,
    recycleAfterDays: 30,
    appVersion: "3.0",
  },

  // Pipeline statuses
  leadStatuses: [
    "New",
    "1st Contact",
    "2nd Contact",
    "3rd Contact",
    "Do Not Call",
    "Sold",
    "Pending Order",
    "FNQ",
    "Already has Fiber",
    "TDM",
  ],

  // Terminal statuses — removed from agent queue, admin only
  terminalStatuses: ["Do Not Call", "Sold", "FNQ", "Already has Fiber", "TDM"],

  // TDM is kicked back to admin (D2D only)
  adminOnlyStatuses: ["TDM"],

  soldStatus: "Sold",

  // Lead types
  leadTypes: ["OFS", "MLR", "Forced"],

  // Current products options — alphabetical
  currentProducts: [
    "Home Phone + Internet + VAS",
    "Internet",
    "Internet + Phone",
    "Internet + TV",
    "Internet + TV + Phone",
    "Internet + VAS",
    "Other",
    "Phone",
    "TV",
    "TV + Phone",
  ],

  leadSources: [
    "Web Form",
    "Referral",
    "Cold Call",
    "Email Campaign",
    "Social Media",
    "Trade Show",
    "Other",
  ],

  roles: {
    admins: [
      "RichardZacker@raimak.com",
      "B.Hinesley@raimak.com",
      "S.Balleste@raimak.com",
      "m.stevens@raimak.com",
      "N.Caldwell@raimak.com",
      "C.Scarrett@raimak.com",
      "m.mcalpine@raimak.com",
      "J.Scroggins@raimak.com",
      "antoinette.bickel@raimak.com",
    ],
  },

  scopes: ["Sites.ReadWrite.All", "User.Read"],

  salesFeedInterval: 30000,
};
