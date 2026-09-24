# KGS Defense Capabilities Tracker

## Summary

KGS Defense Capabilities Tracker is a SharePoint Framework web part used to document, organize, and showcase reusable KGS defense capabilities. The app is intended to help teams quickly answer sales, capture, proposal, and delivery questions such as:

- What does this capability do?
- Which primary and secondary capability types categorize it?
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

- Capability inventory with searchable and filterable grid/list and tile views.
- Tier 1 and optional Tier 2 capability classification, with configurable values and combined filtering across either tier.
- Capability detail view with Overview, Technical Info, Tagging, Supporting Contract(s), and Documentation tabs.
- New/edit capability form with grouped sections and configurable choice fields.
- Related contract management tied to normalized JAMIS and CMS contract source data.
- Standalone contract creation and contract detail/edit forms.
- Contract metadata support for contract type, USD contract value, clearance, synonyms, flagging, and place of performance.
- Contract search plus OG, Customer, Partner, and Clearance Level filters.
- Contract-specific capability summaries with link, edit-summary, and unlink actions for each contract/capability relationship.
- Capability document upload and document type tagging.
- Contract document upload, document type tagging, metadata editing, and delete support.
- Screenshot carousel and PDF export support.
- Excel export for capability grid data.
- PDF capability one-pager and capability book export.
- Admin configuration page for dropdown/choice values.
- Dashboard summary view available after the Capabilities and Contracts views.
- Tagging support for:
  - OppNet opportunities
  - Past Performance Library items
  - Proposals Dashboard items

## Provisioned SharePoint Assets

On first run, the app provisions the following hidden SharePoint assets in the current site:

| Asset | Type | Purpose |
| --- | --- | --- |
| `DCTCapabilities` | List | Main capability records |
| `DCTContracts` | List | Main contract records with multi-lookup relationships to capabilities |
| `DCTContractCapabilitySummary` | List | Contract-specific capability summaries by contract/capability pair |
| `DCTDocuments` | Document library | Capability documentation and screenshots |
| `DCTContractDocuments` | Document library | Contract documentation stored in contract-specific folders |
| `DCTConfiguration` | List | App configuration and dropdown values |

The app also creates/uses these SharePoint groups:

| Group | Purpose |
| --- | --- |
| `DCT Admins` | App administrators/owners |
| `DCT Contributors` | Users who can add and edit tracker content |
| `DCT Visitors` | Read-only users |

## First-Run Requirements

The first user to load the app must have sufficient permissions, typically Full Control, on the target SharePoint site. First run performs list/library provisioning, field creation, view setup, security setup, and configuration seeding. The installation dialog is shown to a site owner or site collection administrator only when a configured list, library, field, content type, or security group is missing.

Configuration seed items are inserted only when the configuration list is empty. The app checks for a single existing configuration row rather than comparing the installed values with the seed catalog on every load. Once any row exists, administrators can manage the configuration without the app restoring or overwriting values from the seed file.

## Configuration Values

Dropdown and choice values are managed through `DCTConfiguration` and the in-app Configuration Management page.

Current seeded configuration categories include:

- Backend
- Capability Status
- Coding Language
- Compliance
- Connectivity
- Contract Clearance
- Contract Document Type
- Contract Type
- Customer
- Relevant Partner Tag
- Hosting Environment
- Platform
- Capability Type
- Document Type
- State

If configuration values need to change after first run, use the Configuration Management page or update the SharePoint `DCTConfiguration` list directly. Updating seed files alone will not add values to an environment that already has config items.

## External Data Sources

The tracker reads supporting data from these external SharePoint locations:

| Source | Site/List | Usage |
| --- | --- | --- |
| JAMIS contracts | `/sites/Jamis_Data_API` / `ContractEndPoint` | Primary contract lookup data |
| CMS contracts | `/sites/ContractsProcurementSite` / `CMSArchive` | Supplemental contract lookup data |
| Organization data | `/sites/ContractsProcurementSite` / `OGPresidents` | Operating Group and LOB lookup data |
| OppNet | `/sites/OppNet` / `Opportunities` | Opportunity tags |
| Proposals | `/sites/Proposals` / `Past Performance Library` | Past performance tags |
| Proposals | `/sites/Proposals` / `Proposals Dashboard` | Proposal tags |

JAMIS and CMS contract sources are loaded on initial app load only. They are intentionally not refreshed during in-app capability refreshes because the source lists are larger external datasets and do not need to reload for normal DCT data refreshes.

The contract form uses a normalized combined source collection for lookup/search. JAMIS takes priority when a likely duplicate exists across JAMIS and CMS. CMS `Contract Number` maps to DCT Customer Contract Code, CMS `Project ID` maps to DCT Contract ID, and CMS Operating Group is used to derive LOB through the existing OG lookup data.

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
| `src/webparts/dcTracker/components/services` | Create/edit/delete service wrappers for capabilities, contracts, documents, and relationship summaries |
| `src/webparts/dcTracker/components/forms` | Capability, contract, and document forms |
| `src/webparts/dcTracker/components/views` | Main views and detail tabs |
| `src/webparts/dcTracker/components/common/props.ts` | Shared TypeScript interfaces |
| `src/webparts/dcTracker/components/common/strings.ts` | List names, external site URLs, shared constants, and colors |

## Operational Notes

- Capability tags are saved as JSON-backed fields on the capability item.
- Capability Type is stored as separate Tier 1 and Tier 2 text fields; Tier 2 is optional, and list filtering matches either tier.
- Document uploads are stored under capability-specific folders in `DCTDocuments`.
- Contract document uploads are stored under contract-specific folders in `DCTContractDocuments`.
- Document type, contract document type, contract type, contract clearance, capability type, and state options come from configuration values.
- Contract values are stored as USD currency values and displayed through the shared currency formatter.
- Contract place of performance is entered as City, State, and two-character Country fields. A calculated Location field formats those values for contract grids, detail dialogs, and capability contract cards.
- Contract records can be created independently from the Contracts view or linked while editing a capability.
- Contract duplicate checks compare Contract Title, Customer Contract Code, and Contract ID before save.
- Contract lookup/search combines normalized JAMIS and CMS source rows while preserving DCT contract records as the saved system of record.
- Contract-to-capability relationships remain on the `DCTContracts` multi-lookup so static contract metadata is maintained once per contract.
- Contract-specific capability notes are stored separately in `DCTContractCapabilitySummary`, keyed by contract lookup and capability lookup.
- Contract capability relationships can be created, summarized, edited, and unlinked from the contract details dialog without deleting either source record.
- Flagged contracts are visible to DCT administrators and excluded from non-administrator contract results.
- Status pill colors are centralized in `components/ui/StatusColors.ts`.
- Main grid filters are driven from configuration values.
- The main grid search scans the capability item broadly, including nested values.

## Version History

| Version | Date | Developer | Comments |
| --- | --- | --- | --- |
| 1.0.0.2 | July 25, 2026 | Landino | Initial release |
| 1.0.0.3 | July 27, 2026 | Landino | Fix main search. Add OG and LOB to contract select/form/grid. Remove routes from capability item view |
| 1.0.0.5 | July 28, 2026 | Landino | Add clickable title in grids. Add Tile view. Add POC and Stakeholders to overview. |
| 1.0.0.7 | July 29, 2026 | Landino | Add contract VIEW modal and capability relationship on view. Remove custom permissions on Capabilities and Contracts. Change contract > capability lookup/relationship to be 1:X |
| 2.0.0.1 | August 11, 2026 | Landino | Move dashboard to end, allow creating contract entry, add CMS contracts to dropdown search (combine/normalize with JAMIS), Add contract value & type, add Capability Summary, Add Contract Documents |
| 2.0.1.3 | September 23, 2026 | Landino | Add contract fields Flagged, Clearance Level, Location, Synonyms. Add Capability Type, Synonyms to Capabilities. Adjust cap details tabs - accordion contracts, move columns, condense tagging. Add/adjust filters on both Caps & Contracts grids. |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://learn.microsoft.com/sharepoint/dev/spfx/set-up-your-development-environment)
- [SharePoint Framework overview](https://learn.microsoft.com/sharepoint/dev/spfx/sharepoint-framework-overview)
- [Use Microsoft Graph in SPFx](https://learn.microsoft.com/sharepoint/dev/spfx/use-msgraph)
- [Fluent UI React](https://developer.microsoft.com/fluentui)
