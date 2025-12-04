# Copilot-Readiness-Assessment-Remediation
Prerequisites:

Install-Module PnP.PowerShell -Scope CurrentUser
Update-Module PnP.PowerShell

API Permission needed:

Sites.FullControl.All
User.Read.All
Scope: Application

Run (app‑only auth):

.\GetAllSites-EEEU-SharingPermissions.ps1 `
  -SiteListCsvPath .\sites.csv `
  -ClientId "<app-id>" `
  -Tenant "<Tenant-Id" `
  -Thumbprint "<cert-thumbprint>"


