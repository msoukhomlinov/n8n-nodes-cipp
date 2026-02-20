# n8n-nodes-cipp

[![npm version](https://badge.fury.io/js/%40joshuanode%2Fn8n-nodes-cipp.svg)](https://www.npmjs.com/package/@joshuanode/n8n-nodes-cipp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

n8n community node for [CIPP.app](https://cipp.app) - Comprehensive Microsoft 365 multi-tenant management.

![CIPP Node](https://img.shields.io/badge/n8n-Community%20Node-ff6d5a)
![Beta](https://img.shields.io/badge/Status-Beta-orange)

> ⚠️ **Beta Notice**: This node is currently in beta and may not be fully functional yet. Some operations may be incomplete or require adjustments. Use in production at your own risk.
>
> 🤝 **Contributions Welcome!** We welcome bug reports, feature requests, and pull requests. If you encounter issues or have improvements, please open an issue or PR on GitHub.

## Features

This node provides full integration with the CIPP API, enabling automation of:

- **Identity Management** - Users, groups, MFA, devices
- **Tenant Administration** - Alerts, licenses, standards
- **Intune** - Applications, Autopilot, device actions
- **Teams & SharePoint** - Teams, sites, voice numbers, shifts scheduling
- **Security & Compliance** - Defender alerts, incidents
- **Tools** - Breach search, Graph API requests, ExecGraphRequest
- **CIPP System** - Scheduled jobs, backups

### User-Friendly Design

- **Tenant Selector** - Searchable dropdown to select tenants by name
- **Field Picker** - Multi-select for user properties (no need to memorize Graph API field names)
- **Smart Defaults** - Sensible default selections to keep responses fast and small

## Installation

### n8n (Self-hosted)

```bash
npm install @joshuanode/n8n-nodes-cipp
```

Or add to your n8n Docker container:

```bash
# In your Dockerfile
RUN npm install -g @joshuanode/n8n-nodes-cipp
```

### n8n Cloud

Community nodes can be installed via **Settings → Community Nodes → Install**.

## Credentials Setup

1. **Create an Azure AD App Registration** for CIPP API access
2. Configure the following in n8n:
   - **CIPP Instance URL**: Your CIPP deployment URL (e.g., `https://cipp.yourdomain.com`)
   - **Azure AD Tenant ID**: The tenant where your CIPP app registration lives
   - **Application (Client) ID**: From your Azure AD app registration
   - **Client Secret**: Generated from your app registration

For detailed authentication setup, see the [CIPP API Documentation](https://docs.cipp.app/api-documentation/setup-and-authentication).

## Resources & Operations

| Resource           | Operations                                                                                                                                                    |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Tenant**         | Get Many, Get Licenses, Get CSP Licenses, CSP License Action, Clear Cache                                                                                     |
| **User**           | Get Many, Add, Disable, Enable, Reset Password, Reset MFA, Revoke Sessions, Remove, Create TAP, Set Per-User MFA, Send MFA Push, Clear Immutable ID, Offboard |
| **Group**          | Add, Edit Members, Delete, Hide from GAL, Set Delivery Management, Get Many                                                                                   |
| **Device**         | Get Many, Manage, Execute Action, Get Recovery Key, Get LAPS Password                                                                                         |
| **Autopilot**      | Get Many, Assign, Remove, Sync, Get Configurations                                                                                                            |
| **Mailbox**        | Convert, Enable Archive, Set Out of Office, Set Email Forwarding                                                                                              |
| **Alert**          | Add, Get Many, Get Security Alerts, Get Security Incidents, Set Alert Status, Set Incident Status                                                             |
| **Application**    | Get Many, Assign, Remove, Add WinGet/Store/Chocolatey/MSP/Office Apps                                                                                         |
| **Team**           | Add, Get Many, Get Sites, Get Activity, Manage Site Members/Permissions                                                                                       |
| **Teams Shift**    | List/Create/Update/Delete Shifts, Open Shifts, Scheduling Groups, Time Off Reasons; List/Create/Approve/Decline Time Off, Swap Shift & Offer Shift Requests   |
| **Voice**          | Get Phone Numbers, Get Locations, Assign/Unassign Numbers                                                                                                     |
| **Scheduled Item** | Add, Get Many, Remove                                                                                                                                         |
| **Backup**         | Get Many, Run, Restore, Set Auto-Backup                                                                                                                       |
| **Tools**          | Breach Search (Account/Tenant), Exec Graph Request, Graph Request (List), Graph Request (Exec)                                                                |

## Example Usage

### List All Tenants

```
Resource: Tenant
Operation: Get Many
Return All: true
```

### List Users with Sign-In Activity

```
Resource: User
Operation: Get Many
Tenant: Select from dropdown
Fields to Return: Display Name, User Principal Name, Mail, Sign-In Activity
Return All: true
```

### Create a New User

```
Resource: User
Operation: Add
Tenant: Select from dropdown
First Name: John
Last Name: Doe
Domain: contoso.com
```

### Execute Device Action

```
Resource: Device
Operation: Execute Action
Tenant: Select from dropdown
Device ID: <device-guid>
Action: SyncDevice
```

### Custom Graph Request

```
Resource: Tools
Operation: Graph Request (List)
Tenant: Select from dropdown
Endpoint: users
$select: id,displayName,userPrincipalName
$filter: startsWith(displayName,'John')
```

### Teams Shifts (Dedicated Resource)

```
Resource: Teams Shift
Operation: List Shifts
Tenant: Select from dropdown
Team ID: <team-guid>
Filters → Start Date: 2024-03-01T00:00:00Z
Filters → End Date: 2024-03-31T23:59:59Z
```

```
Resource: Teams Shift
Operation: Create Shift
Tenant: Select from dropdown
Team ID: <team-guid>
User ID: <aad-user-id>
Start Date Time: 2024-03-15T08:00:00Z
End Date Time: 2024-03-15T16:00:00Z
Options → Display Name: Morning Shift
Options → Theme: blue
```

> ⚠️ **CIPP-API Requirement**: The Teams Shift resource and the Exec Graph Request tool both use `POST /api/ExecGraphRequest`, which is **not part of the standard CIPP API**. You must be running a custom fork of [CIPP-API](https://github.com/KelvinTegelaar/CIPP-API) that exposes the `ExecGraphRequest` endpoint. Without this, all Teams Shift operations and the Exec Graph Request tool will return a 404 or 400 error.
>
> If your fork uses a different route name (e.g., `/api/GraphRequest`), the `Graph Request (Exec)` tool has a built-in fallback. The dedicated Teams Shift resource does not — it expects `/api/ExecGraphRequest` to exist.

### Graph Request (Exec) — Raw Graph Calls

```
Resource: Tools
Operation: Graph Request (Exec)
Tenant: Select from dropdown
Endpoint: teams/<team-id>/schedule/shifts
Method: POST
Body: {"userId":"<aad-user-id>","schedulingGroupId":"<group-id>","sharedShift":{...}}
```

Notes:

- `Graph Request (Exec)` sends a `POST` to `/api/ExecGraphRequest` and falls back to `/api/GraphRequest` if your fork uses that route name.
- By default, client-side validation requires endpoints matching `teams/{id}/schedule/*` (can be disabled in `Exec Options`).

## Development

```bash
# Install dependencies
npm install

# Build
npm run build

# Lint
npm run lint

# Link for local testing
npm link
```

## Links

- [CIPP.app](https://cipp.app)
- [CIPP Documentation](https://docs.cipp.app)
- [CIPP API Endpoints](https://docs.cipp.app/api-documentation/endpoints/)
- [n8n Community Nodes](https://docs.n8n.io/integrations/community-nodes/)

## License

MIT
