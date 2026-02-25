const Config = {
  azure: {
    clientId: "ad1b153f-8b6a-4f1c-8ab2-58fcf03cf5c2",
    tenantId: "39e14190-0b23-4ecd-99f9-606ad1215881",
    redirectUri: window.location.origin + window.location.pathname
  },
  sharePoint: {
    hostname: "raimak.sharepoint.com",
    sites: {
      leadship: "sites/RaimakLeadship",
      team: "TeamSite"
    },
    lists: {
      activityLog: "2adb1260-e635-45cd-bb3b-87dd57a2d022",
      contractorList: "bd5df38a-9cb6-411d-87e8-3e79934213d3",
      leadsList: "5a01419d-e2c9-4aad-8484-6ed97233f305"
    },
    graphBase: "https://graph.microsoft.com/v1.0"
  },
  rules: {
    coolOffDays: 2,
    maxLeadsPerAgent: 15,
    recycleAfterDays: 30,
    appVersion: "1.0"
  },
  leadStatuses: ["New","Contacted","Qualified","Proposal Sent","Negotiating","Won","Lost","Recycled"],
  leadSources: ["Web Form","Referral","Cold Call","Email Campaign","Social Media","Trade Show","Other"],
  scopes: ["Sites.ReadWrite.All","User.Read"]
};