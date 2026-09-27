// Anmeldung für Microsoft Graph.
// Primär Nested App Authentication (NAA): Outlook meldet das Add-in mit dem Outlook-Konto an,
// ohne eigenes Anmeldefenster. Die alten Exchange-Token für Add-ins sind abgeschaltet, NAA ist
// daher der vorgesehene Weg. Falls NAA im jeweiligen Outlook nicht verfügbar ist (Requirement
// Set NestedAppAuth 1.1 fehlt), wird auf eine normale SPA-Anmeldung per Popup ausgewichen.

import {
  createNestablePublicClientApplication,
  createStandardPublicClientApplication,
  InteractionRequiredAuthError,
  type IPublicClientApplication,
} from "@azure/msal-browser";
import { config } from "./config";

export type AuthMode = "naa" | "popup";

let pca: IPublicClientApplication | undefined;
let mode: AuthMode | undefined;

export function naaSupported(): boolean {
  try {
    return Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  } catch {
    return false;
  }
}

async function getClient(): Promise<{ pca: IPublicClientApplication; mode: AuthMode }> {
  if (pca && mode) return { pca, mode };
  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      redirectUri: `${window.location.origin}/taskpane.html`,
    },
    cache: { cacheLocation: "localStorage" as const },
  };
  if (naaSupported()) {
    pca = await createNestablePublicClientApplication(msalConfig);
    mode = "naa";
  } else {
    pca = await createStandardPublicClientApplication(msalConfig);
    mode = "popup";
  }
  return { pca, mode };
}

export async function getAuthMode(): Promise<AuthMode> {
  return (await getClient()).mode;
}

/** Holt ein Zugriffstoken für die angegebenen Graph-Berechtigungen (erst still, dann interaktiv). */
export async function getToken(scopes: string[]): Promise<string> {
  const { pca: client } = await getClient();
  const request = { scopes };
  const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
  try {
    const result = await client.acquireTokenSilent({ ...request, account });
    return result.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError) && account) throw error;
    const result = await client.acquireTokenPopup(request);
    if (result.account) client.setActiveAccount(result.account);
    return result.accessToken;
  }
}
