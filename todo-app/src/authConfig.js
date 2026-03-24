// Azure AD App Registration configuration
// To get these values:
// 1. Go to https://portal.azure.com
// 2. Navigate to Azure Active Directory > App registrations > New registration
// 3. Set redirect URI to http://localhost:5173 (Single-page application)
// 4. Under API permissions, add Microsoft Graph > Delegated > Tasks.ReadWrite, User.Read
// 5. Copy the Application (client) ID and Directory (tenant) ID below

export const msalConfig = {
  auth: {
    clientId: "YOUR_CLIENT_ID", // Replace with your Azure App Client ID
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  scopes: ["User.Read", "Tasks.ReadWrite"],
};

export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
  graphTodoEndpoint: "https://graph.microsoft.com/v1.0/me/todo/lists",
};
