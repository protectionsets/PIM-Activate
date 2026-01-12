
# PIM Bulk Activator (WPF GUI) for Microsoft Entra ID

A Windows PowerShell 5.1 + WPF utility that lets **eligible users** quickly **activate** or **deactivate** their own Privileged Identity Management (PIM) role assignments in **Microsoft Entra ID**.

The app presents eligible roles grouped by Administrative Unit (AU) or tenant scope, highlights **already-active** roles in **green**, and disables rows so you don’t accidentally activate them again. It supports bulk actions, includes quick links to Microsoft admin portals, and uses a clean WPF template with optional **Dark Title Bar** and **Mica** system backdrops on Windows 11.

> **Author:** Yoni + Copilot  
> **Script:** `PIM-BulkActivator.ps1`

---

## ✨ Features

- **WPF GUI (PowerShell STA)**: Robust XAML load with status bar, grouped list view, row-highlight triggers, and keyboard shortcuts.
- **Role visibility**: Shows eligible roles as `ROLE – (Tenant wide | AU Name) – MemberType`.
- **Active awareness**: Active roles are **green** in the list and **disabled** from selection.
- **Bulk operations**:
  - **Activate** all selected, non-active roles.
  - **Deactivate** _all_ currently active roles for the signed‑in user.
- **Ticketing fields**: Optional `Ticket number` and `Ticket system` accompany activation requests.
- **Graph-aware**: Reads eligibility and current activations; posts activation/deactivation requests.
- **Quick Admin buttons**: Open M365, Entra, Intune, Defender, Purview, EXO, Teams admin portals in your default browser.
- **Safe UX**: Selection uses row-click or **Spacebar**; active items are non-selectable; report-only status is shown in the log box.
- **Optional Windows 11 flair**: Dark title bar + Mica via DWM attributes.

---

## 🧩 Architecture & Flow

1. **Connect (read scopes)** → load current user (`/me`) and **eligible role schedules**.  
2. **Resolve scope display** → `(Tenant wide)` or AU name (via `/directory/administrativeUnits/{id}`).  
3. **Check active** → read **role assignment schedules** for the principal (assignmentType=`Activated`).  
4. **Mark UI** → green text + disabled rows for active items; grouping by `ScopeDisplay`.
5. **Activate** → submit `selfActivate` requests with justification, duration (1–8h), optional ticket fields.
6. **Deactivate** → submit `selfDeactivate` requests for all currently active roles.
7. **Refresh** → re-query active assignments and re-mark the view.

---

## 🔒 Microsoft Graph permissions & endpoints

### Read scopes
- `User.Read`  
- `RoleEligibilitySchedule.Read.Directory`  
- `RoleAssignmentSchedule.Read.Directory`  
- `AdministrativeUnit.Read.All`

### Write scopes (require **admin consent**)
- `User.Read`  
- `RoleManagement.ReadWrite.Directory`  
- `RoleAssignmentSchedule.ReadWrite.Directory`  
- `AdministrativeUnit.Read.All`

### REST endpoints used (v1.0)
- `GET https://graph.microsoft.com/v1.0/me`
- `GET https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules/filterByCurrentUser(on='principal')?$expand=roleDefinition&$top=999`
- `GET https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentSchedules?$filter=principalId eq '{USERID}' and assignmentType eq 'Activated'&$expand=roleDefinition&$top=999`
- `GET https://graph.microsoft.com/v1.0/directory/administrativeUnits/{id}`
- `POST https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleRequests`  
  - body for **activate**: `action = "selfActivate"`  
  - body for **deactivate**: `action = "selfDeactivate"`

> Admin consent is needed before **write scopes** can be used.

---

## 🖥️ Requirements

- **Windows PowerShell 5.1** (the script self-restarts as **STA** if needed).  
- **.NET Framework** (WPF assemblies: `PresentationCore`, `PresentationFramework`, `WindowsBase`, `System.Xaml`).  
- **Microsoft Graph PowerShell SDK**:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  # or: Install-Module Microsoft.Graph -Repository PSGallery
  ```
- **Microsoft Entra ID account** that is **eligible** for PIM roles (this tool acts for **the current signed-in user only**).

---

## 🚀 Getting started

1. **Install the Graph SDK** (see above).  
2. **Unblock & run the script**:
   ```powershell
   # from an elevated PowerShell 5.1 console on Windows
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   .\PIM-BulkActivator.ps1
   ```
3. Click **Sign in / Load roles** → review eligible roles and active ones (green).  
4. **Select** roles by clicking rows or pressing **Spacebar**.  
5. **Activate selected** (enter justification, duration 1–8 hours, optional ticket data).  
6. Use **Deactivate ALL active** when you want to end all current activations.

> The script will request **write scopes** on first activation/deactivation; if admin consent hasn’t been granted yet, you’ll see a warning.

---

## 🖌️ UI customization (colors & styles)

All button and text colors are defined via **XAML resource brushes** at the top of the `Window.Resources` section:

```xml
<SolidColorBrush x:Key="PanelBrush"        Color="#FFFFFFFF"/>
<SolidColorBrush x:Key="BorderBrush"       Color="#DDDDDD"/>
<SolidColorBrush x:Key="TextBrush"         Color="#111111"/>
<SolidColorBrush x:Key="AccentStrokeBrush" Color="#4F6BED"/>
<SolidColorBrush x:Key="ActiveGreenBrush"  Color="#107C10"/>
```

- **Buttons** use `PanelBrush` (background), `TextBrush` (foreground), and `BorderBrush` (default border).  
- On **hover/press**, the border switches to `AccentStrokeBrush`.  
- **Active roles** in the list use `ActiveGreenBrush` via a `DataTrigger` bound to `IsActive`.

> Change these hex codes to adapt the theme; the global `Button` style applies them across the command rail.

---

## 🧠 Design notes

- **STA enforcement**: If the current runspace isn’t STA, the script **relaunches itself** with `-STA` to ensure WPF works correctly.
- **WPF loading**: Uses `XamlReader.Parse` with a resilient `XmlReader` fallback.
- **Grouping**: Roles are grouped by `ScopeDisplay` (AU name or `(Tenant wide)`), via `CollectionViewSource`.
- **Selection model**: No checkboxes; selection is driven by row click/keyboard, mirrored to `IsChecked` to make bulk actions easy.
- **Active refresh**: After any activation/deactivation, the tool re-queries active schedules and re-marks the list.
- **DWM interop (optional)**: Dark title bar and Mica are applied via `DwmSetWindowAttribute` for Windows 11.

---

## 🔧 Troubleshooting

- **Admin consent required** for write scopes: If you see an **Admin consent needed** warning, have a Global Admin grant consent for the app/scopes, or use your organization’s approved process.
- **No eligible roles**: The tool only surfaces schedules for the **current user**. If you’re not eligible, the list will be empty.
- **Graph throttling**: Large tenants may hit paging; the script handles `@odata.nextLink` for up to `top=999` per page and loops until completion.
- **Execution policy**: Use `-ExecutionPolicy Bypass` if local policy blocks script execution.
- **Display names for AU**: If AU resolution fails, the tool caches and shows the GUID to keep the UI responsive.

---

## 🔐 Security & privacy

- The tool acts **only for the signed-in principal**; it cannot elevate beyond what PIM makes eligible.  
- **Justification** and optional **ticket fields** are sent in the activation request body to Graph; logs are displayed in the GUI only (no file persistence in this script).  
- Consider your local workstation policies for clipboard/log access if you adapt the code to persist logs.

---

## ⚠️ Limitations

- Not intended to manage role assignments for **other users**.  
- **Activation outcomes** depend on PIM settings (approval required, MFA, etc.).  
- Some scenarios (e.g., specific RDP/VDI flows without WebAuthn redirection) are out of scope for this GUI.

---

## 🧪 Development & contributions

- Tested on **Windows 11** with PowerShell **5.1** and the **Microsoft Graph** PowerShell SDK.  
- If you contribute UI changes, please keep to the **resource brush** pattern and avoid hardcoded colors.  
- For logic changes, ensure active/paging handling remains robust and that you refresh the view after any edits.

### Run locally
```powershell
# Optionally pin STA on launch (script will auto-handle if needed)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File .\PIM-BulkActivator.ps1
```

---

## 📄 License

Recommend using **MIT**; adapt as needed for your organization.

```text
MIT License
Copyright (c) 2025 Yoni + Contributors
Permission is hereby granted, free of charge, to any person obtaining a copy...
```

---

## 🙌 Acknowledgements

- Microsoft Graph PowerShell SDK team for robust cmdlets.  
- Community docs and samples around WPF-in-PowerShell patterns.

---

## 📷 (Optional) Screenshots

> Place images under `docs/` and reference them here.

- Eligible roles grouped by AU.  
- Active roles marked in green (disabled rows).  
- Activation pane with justification & duration.

---

## 🧰 File list

- `PIM-BulkActivator.ps1` — main script with WPF XAML, Graph helpers, and UI handlers.  
- `README.md` — this file.

---

## 🗺️ Roadmap ideas

- Add per‑role **duration overrides**.  
- Export audit log to JSON/CSV.  
- Add **theme switch** (light/dark) and compact density mode.

---

## 🔎 Code snippets

Activate selected roles (internals):
```powershell
$res = Request-PIMActivation -UserId $script:SignedInUser.id `
                             -RoleDefinitionId $role.RoleDefinitionId `
                             -DirectoryScopeId $role.DirectoryScopeId `
                             -Justification $just `
                             -Duration $duration `
                             -TicketNumber $ticketNumber `
                             -TicketSystem $ticketSystem
```

Deactivate all active roles (internals):
```powershell
$res = Request-PIMDeactivation -UserId $script:SignedInUser.id `
                               -RoleDefinitionId $a.RoleDefinitionId `
                               -DirectoryScopeId $a.DirectoryScopeId `
                               -Justification "Bulk deactivation from GUI"
```

---

## 📫 Support

Open a GitHub issue with:
- What you tried  
- Console output / screenshots  
- PowerShell & Windows versions  
- Steps to reproduce

