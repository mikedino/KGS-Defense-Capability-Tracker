# KGS Defense Capabilities Tracker

## Summary

KGS Defense Capabilities Tracker is a SharePoint Framework web part used to document, organize, and showcase reusable KGS defense capabilities. The app is intended to help teams quickly answer sales, capture, proposal, and delivery questions such as:

- What does this capability do?
- What kind of solution is it?
- What platform, hosting, connectivity, licensing, backend, and technical requirements are needed?
- Which contracts, past performance records, proposals, and OppNet opportunities are related?
- Who owns or supports the capability?
- What documentation and screenshots are available?

The app provisions its own SharePoint lists/library on first run, then uses those lists as the application data store.

## SharePoint Framework

![SPFx version](https://img.shields.io/badge/SPFx-1.21.1-green.svg)

## Solution

| Solution | Author |
| --- | --- |
| KGS Defense Capabilities Tracker | Mike Landino, Koniag Government Services |

## Core Features

- Capability inventory with searchable and filterable grid/list views.
- Capability detail view with Overview, Supporting Info, Contract, Documentation, and Tagging tabs.
- New/edit capability form with grouped sections and configurable choice fields.
- Related contract management tied to Jamis contract data.
- Document upload and document type tagging.
- Screenshot carousel and PDF export support.
- Excel export for capability grid data.
- PDF capability one-pager and capability book export.
- Admin configuration page for dropdown/choice values.
- Tagging support for:
  - OppNet opportunities
  - Past Performance Library items
  - Proposals Dashboard items

## Provisioned SharePoint Assets

On first run, the app provisions the following hidden SharePoint assets in the current site:

| Asset | Type | Purpose |
| --- | --- | --- |
| `DCTCapabilities` | List | Main capability records |
| `DCTContracts` | List | Contract relationships for capabilities |
| `DCTDocuments` | Document library | Capability documentation and screenshots |
| `DCTConfiguration` | List | App configuration and dropdown values |

The app also creates/uses these SharePoint groups:

| Group | Purpose |
| --- | --- |
| `DCT Admins` | App administrators/owners |
| `DCT Contributors` | Users who can add and edit tracker content |
| `DCT Visitors` | Read-only users |

## First-Run Requirements

The first user to load the app must have sufficient permissions, typically Full Control, on the target SharePoint site. First run performs list/library provisioning, field creation, view setup, security setup, and configuration seeding.

Configuration seed items are only inserted when the configuration list is empty. The app checks for a single existing config item, not every seed item one by one. This keeps startup lightweight and avoids overwriting admin-managed values.

## Configuration Values

Dropdown and choice values are managed through `DCTConfiguration` and the in-app Configuration Management page.

Current seeded configuration categories include:

- Backend
- Capability Status
- Coding Language
- Compliance
- Connectivity
- Customer
- Relevant Partner Tag
- Hosting Environment
- Platform
- Solution Type
- Document Type

If configuration values need to change after first run, use the Configuration Management page or update the SharePoint `DCTConfiguration` list directly. Updating seed files alone will not add values to an environment that already has config items.

## External Data Sources

The tracker reads supporting data from these external SharePoint locations:

| Source | Site/List | Usage |
| --- | --- | --- |
| Jamis contracts | `/sites/Jamis_Data_API` / `ContractEndPoint` | Contract lookup data |
| OppNet | `/sites/OppNet` / `Opportunities` | Opportunity tags |
| Proposals | `/sites/Proposals` / `Past Performance Library` | Past performance tags |
| Proposals | `/sites/Proposals` / `Proposals Dashboard` | Proposal tags |

`ContractEndPoint` is loaded on initial app load only. It is intentionally not refreshed during capability refreshes because it is large and only updated overnight.

## Local Development

### Prerequisites

- Node.js `>=22.14.0 <23.0.0`
- SharePoint Framework toolchain
- Access to the target SharePoint tenant/site
- Permissions to read any configured external source lists

### Install

```bash
npm install
```

### Build

```bash
npm run build
```

### Serve Locally

```bash
gulp serve
```

### Package for Deployment

```bash
npm run package
```

The package script runs a clean build, production bundle, and SharePoint package step.

## Project Structure

Key application areas:

| Path | Purpose |
| --- | --- |
| `src/webparts/dcTracker/components/main.tsx` | Main app shell, routing, filters, exports, and modals |
| `src/webparts/dcTracker/components/data/cfg.ts` | SharePoint list/library/field provisioning configuration |
| `src/webparts/dcTracker/components/data/configSeedItems.ts` | Initial configuration seed values |
| `src/webparts/dcTracker/components/data/ds.ts` | Data source initialization and read helpers |
| `src/webparts/dcTracker/components/services` | Create/edit/delete service wrappers |
| `src/webparts/dcTracker/components/forms` | Capability, contract, and document forms |
| `src/webparts/dcTracker/components/views` | Main views and detail tabs |
| `src/webparts/dcTracker/components/common/props.ts` | Shared TypeScript interfaces |
| `src/webparts/dcTracker/components/common/strings.ts` | List names, external site URLs, shared constants, and colors |

## Operational Notes

- Capability tags are saved as JSON-backed fields on the capability item.
- Document uploads are stored under capability-specific folders in `DCTDocuments`.
- Document type options come from configuration values.
- Status pill colors are centralized in `components/ui/StatusColors.ts`.
- Main grid filters are driven from configuration values.
- The main grid search scans the capability item broadly, including nested values.

## Version History

| Version | Date | Developer | Comments |
| --- | --- | --- | --- |
| 1.0.0.2 | July 25, 2026 | Landino | Initial release |
| 1.0.0.3 | July 27, 2026 | Landino | Fix main search. Add OG and LOB to contract select/form/grid. Remove routes from capability item view |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)
- [SharePoint Framework overview](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Use Microsoft Graph in SPFx](https://learn.microsoft.com/sharepoint/dev/spfx/use-msgraph)
- [Fluent UI React](https://developer.microsoft.com/fluentui)
