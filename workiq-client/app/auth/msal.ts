import {
  BrowserCacheLocation,
  PublicClientApplication,
  type Configuration,
} from "@azure/msal-browser";

const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID;
const tenantId = import.meta.env.VITE_ENTRA_TENANT_ID;

if (!clientId || !tenantId) {
  throw new Error(
    "Missing VITE_ENTRA_CLIENT_ID or VITE_ENTRA_TENANT_ID. Configure both values in your .env file",
  );
}

const msalConfig: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: window.location.origin,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: BrowserCacheLocation.SessionStorage,
  },
};

// This module is imported only by entry.client.tsx, never by the server build.
export const msalInstance = new PublicClientApplication(msalConfig);
