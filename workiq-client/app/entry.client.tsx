import { StrictMode, startTransition } from "react";
import { hydrateRoot } from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import { HydratedRouter } from "react-router/dom";
import { msalInstance } from "./auth/msal";

startTransition(() => {
  hydrateRoot(
    document,
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <HydratedRouter />
      </MsalProvider>
    </StrictMode>,
  );
});
