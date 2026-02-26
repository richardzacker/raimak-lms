// Raimak LMS - App Configuration v2.0

const Config = {

  azure: {
    clientId:    "YOUR_AZURE_APP_CLIENT_ID",
    tenantId:    "YOUR_TENANT_ID",
    redirectUri: window.location.origin + window.location.pathname,
  },

  sharePoint: {
    hostname: "raimak.sharepoint.com",
    sites: {
      leadship: "sites/RaimakLeadship",
      team:     "TeamSite",
    },
    lists: {
      activityLog:    "2adb1260-e635-45cd-bb3b-87dd57a2d022",
      contractorList: "bd5df38a-9cb6-411d-87e8-3e79934213d3",
      leadsList:      "5a01419d-e2c9-4aad-8484-6ed97233f305",
    },
    graphBase: "https://graph.microsoft.com/v1.0",
  },

  // Business rules
  rules: {
    coolOffDays:         2,
    maxLeadsPerAgent:    15,
    maxContactsPerDay:   5,
    recycleAfterDays:    30,
    appVersion:          "2.0",
  },

  // Updated pipeline statuses
  leadStatuses: [
    "New",
    "1st Contact",
    "2nd Contact",
    "3rd Contact",
    "Do Not Call",
    "Sold",
    "Pending Order",
    "FNQ"
  ],

  // Terminal statuses - no further action expected
  terminalStatuses: ["Do Not Call", "Sold", "FNQ"],

  // Sold status - triggers live sales feed
  soldStatus: "Sold",

  leadSources: [
    "Web Form",
    "Referral",
    "Cold Call",
    "Email Campaign",
    "Social Media",
    "Trade Show",
    "Other"
  ],

  // Role-based access
  // Add email addresses of leadership/admins who can assign leads
  roles: {
    admins: [
      // Add admin/leadership emails here e.g:
      // "antoinette.bickel@raimak.com",
      // "S.Balleste@raimak.com",
      // "B.Hinesley@raimak.com",
      // "RichardZacker@raimak.com",
    ],
    // Anyone NOT in the admins list is treated as a standard Agent
  },

  scopes: [
    "Sites.ReadWrite.All",
    "User.Read",
  ],

  // Live sales feed poll interval (ms)
  salesFeedInterval: 30000,
};
