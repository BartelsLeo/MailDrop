// Konfiguration aus Vite-Umgebungsvariablen (web/.env.local, nicht eingecheckt - Vorlage:
// web/.env.example). Client- und Tenant-ID stammen aus der App-Registrierung, die die IT anlegt.
// Es gibt kein Client-Geheimnis: Anmeldung per Nested App Authentication (SPA, öffentlicher Client).

export const config = {
  clientId: import.meta.env.VITE_CLIENT_ID ?? "",
  tenantId: import.meta.env.VITE_TENANT_ID ?? "organizations",
};

export const scopes = {
  base: ["User.Read"],
  files: ["Files.ReadWrite.All"],
  appFolder: ["Files.ReadWrite.AppFolder"],
  mail: ["Mail.Read"],
  mailShared: ["Mail.Read.Shared"],
  sites: ["Sites.Read.All"],
};

export function isConfigured(): boolean {
  return config.clientId.length > 0;
}
