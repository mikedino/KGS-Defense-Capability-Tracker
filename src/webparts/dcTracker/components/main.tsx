import * as React from 'react';
import { useState, useEffect } from "react"
import {
  Stack, Text, DefaultButton, SearchBox, Dropdown, IDropdownOption, Pivot, PivotItem, Spinner, SpinnerSize,
  CommandBar, Dialog, DialogType, DialogFooter, ICommandBarItemProps, mergeStyleSets,
  MessageBar, MessageBarType, Icon,
  IconButton, PrimaryButton
} from "@fluentui/react"
import { DataSource } from './data/ds';
import { formatError } from './common/utils';
import Strings from './common/strings';
import { ICapabilityItem, ICapFormSaveResult, IContractDocumentItem, IContractItem, IDCTrackerProps } from './common/props';
import { customPivotStyles } from './ui/ComponentStyles';
import { CapabilitiesList, ICapabilityContractSummary } from './views/CapabilitiesView';
import { CapForm } from './forms/CapForm';
import { CapabilityService } from './services/CapabilityService';
import styles from './Dct.module.scss';
import { AdminPanel } from './admin/AdminPanel';
import { DocumentService } from './services/DocumentService';
import { CapDetails } from './views/CapDetails';
import { exportToExcel } from './export/ExportToExcel';

import { ThemeProvider } from "@fluentui/react";
import { appTheme } from './ui/theme';
import { ContractsList } from './views/ContractsView';
import { ContractForm } from './forms/ContractForm';
import { ContractService } from './services/ContractService';
import { ContractDetailCard } from './views/contracts/ContractDetailCard';
import { AppDashboard } from './views/Dashboard';
import { Security } from './services/Security';
import { exportCapabilitiesBookPdf } from './export/ExportPdfWrapper';
import { buildPdfBookItems } from './export/ExportPdfUtils';
import { AppHeader } from './ui/AppHeader';
import { HashRouter, useHistory, useLocation } from 'react-router-dom';
import { CapRouteTab, getPathParts, routes } from './routing/routes';
import { Helper } from 'gd-sprest-bs';
import { ContractDocumentsPanel } from './views/contracts/ContractDocumentsPanel';
import { ContractDocumentForm } from './forms/ContractDocumentForm';
import { ContractCapabilityForm } from './forms/ContractCapabilityForm';

interface CustomFile extends File {
  data: ArrayBuffer;
}

const spoTealButtonStyles = {
  root: {
    backgroundColor: Strings.PillStyles.SPOTealColor,
    borderColor: Strings.PillStyles.SPOTealColor,
    color: "#ffffff"
  },
  rootHovered: {
    backgroundColor: Strings.PillStyles.SPOTealColor,
    borderColor: Strings.PillStyles.SPOTealColor,
    color: "#ffffff",
    filter: "brightness(0.9)"
  },
  rootPressed: {
    backgroundColor: Strings.PillStyles.SPOTealColor,
    borderColor: Strings.PillStyles.SPOTealColor,
    color: "#ffffff",
    filter: "brightness(0.82)"
  }
};

const DctContent: React.FC<IDCTrackerProps> = (props) => {
  const history = useHistory();
  const location = useLocation();
  const currentPathRef = React.useRef(location.pathname);
  const previousPathRef = React.useRef<string | undefined>(undefined);

  const [capabilities, setCapabilities] = useState<ICapabilityItem[]>([]);
  const [selectedCap, setSelectedCap] = useState<ICapabilityItem | undefined>(undefined);
  const [showCapForm, setShowCapForm] = useState<boolean>(false);
  const [contracts, setContracts] = useState<IContractItem[]>([]);
  const [selectedContract, setSelectedContract] = useState<IContractItem | undefined>(undefined);
  const [showContractForm, setShowContractForm] = useState<boolean>(false);
  const [isContractEditMode, setIsContractEditMode] = useState<boolean>(false);
  const [contractDocuments, setContractDocuments] = useState<IContractDocumentItem[]>([]);
  const [contractDocumentType, setContractDocumentType] = useState<string | undefined>(undefined);
  const [selectedContractDocument, setSelectedContractDocument] = useState<IContractDocumentItem | undefined>(undefined);
  const [showContractDocUploadDialog, setShowContractDocUploadDialog] = useState<boolean>(false);
  const [showContractDocEditDialog, setShowContractDocEditDialog] = useState<boolean>(false);
  const [showContractDocDeleteDialog, setShowContractDocDeleteDialog] = useState<boolean>(false);
  const [selectedRelationshipCapability, setSelectedRelationshipCapability] = useState<ICapabilityItem | undefined>(undefined);
  const [showContractCapabilityDialog, setShowContractCapabilityDialog] = useState<boolean>(false);
  const [contractCapabilityDialogMode, setContractCapabilityDialogMode] = useState<"link" | "edit">("link");
  const [relationshipCapability, setRelationshipCapability] = useState<ICapabilityItem | undefined>(undefined);
  const [relationshipSummary, setRelationshipSummary] = useState<string>("");
  const [relationshipCapabilityToDelete, setRelationshipCapabilityToDelete] = useState<ICapabilityItem | undefined>(undefined);

  const [searchTerm, setSearchTerm] = useState<string>("");
  const [capStatusFilter, setCapStatusFilter] = useState<string>("all");
  const [capabilityTypeFilter, setCapabilityTypeFilter] = useState<string>("all");
  const [platformFilter, setPlatformFilter] = useState<string>("all");

  const [viewMode, setViewMode] = useState<string>("list");
  const [loading, setLoading] = useState<boolean>(true);

  const [selectedPivot, setSelectedPivot] = useState<string>("caps");
  const [selectedCapTab, setSelectedCapTab] = useState<CapRouteTab>("overview");
  const [showAdminPanel, setShowAdminPanel] = useState<boolean>(false);

  const [showDialog, setShowDialog] = useState(false);
  const [dialogTitle, setDialogTitle] = useState<string>("");
  const [dialogMessage, setDialogMessage] = useState<string>("");
  const [showSpinner, setShowSpinner] = useState(false);
  const [spinnerMessage, setSpinnerMessage] = useState<string>("Loading...");

  const setDialogProps = (title: string, message: string): void => {
    setShowDialog(true);
    setDialogTitle(title);
    setDialogMessage(message);
  };

  const hideDialog = (): void => {
    setShowDialog(false);
  };

  const initDatasource = async (override: boolean): Promise<void> => {
    setSpinnerMessage("Loading data. Please wait...");
    setShowSpinner(true);

    try {
      await DataSource.init(override, props.context);
      const nCapabilities = [...(DataSource.Capabilities ?? [])];
      setCapabilities(nCapabilities);
      const nContracts = [...(DataSource.Contracts ?? [])];
      setContracts(nContracts);
    } finally {
      setShowSpinner(false);
    }
  };

  useEffect(() => {
    const initialize = async (): Promise<void> => {
      try {
        await initDatasource(false);
      } catch (error) {
        const msg = formatError(error);
        console.error("Error Initializing Capability:", msg);
        setDialogProps("Error Initializing Capability", msg);
      } finally {
        setLoading(false);
      }
    };

    initialize().catch((error) => console.error("Unhandled promise rejection:"));
  }, []);

  const handleSelectedCap = (capItem: ICapabilityItem): void => {
    history.push(routes.cap(capItem.Id));
  }

  const handleSelectedContract = (contractItem?: IContractItem): void => {
    if (contractItem?.Id) {
      history.push(routes.contract(contractItem.Id));
    } else {
      setSelectedContract(undefined);
      setShowContractForm(false);
      setIsContractEditMode(false);
    }
  }

  // Load contract documents for the selected DCTContracts item.
  const getContractDocuments = async (contractId: number): Promise<void> => {
    const docs = await DataSource.getDocumentsByContract(contractId);
    setContractDocuments(
      [...docs].sort((a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime())
    );
  };

  const handleCapDetailsBack = (): void => {
    if (previousPathRef.current) {
      history.goBack();
    } else {
      history.push(routes.home);
    }

    initDatasource(true).catch((error) =>
      console.error(`Error refreshing datasource: ${formatError(error)}`)
    );
  };

  const contractSummaryByCapabilityId = React.useMemo<Map<number, ICapabilityContractSummary>>(() => {
    const summaryMap = new Map<number, ICapabilityContractSummary>();

    contracts.forEach((contract) => {
      ContractService.getCapabilityLookups(contract).forEach((capability) => {
        const existing = summaryMap.get(capability.Id) ?? { titles: [], searchText: "" };
        const title = contract.Title?.trim();
        if (title) {
          existing.titles.push(title);
        }

        existing.searchText = [
          existing.searchText,
          contract.Title,
          contract.contractId,
          contract.customerContractCode,
          contract.customer,
          contract.contractPm?.Title,
          contract.partner,
          contract.ogTitle,
          contract.lobTitle
        ]
          .filter(Boolean)
          .join(" ");

        summaryMap.set(capability.Id, existing);
      });
    });

    return summaryMap;
  }, [contracts]);

  const selectedContractCapabilities = React.useMemo<ICapabilityItem[]>(() => {
    if (!selectedContract) return [];

    const relatedCapabilityIds = new Set(
      ContractService.getCapabilityLookups(selectedContract).map((capability) => capability.Id)
    );

    return capabilities
      .filter((capability) => relatedCapabilityIds.has(capability.Id))
      .sort((a, b) => (a.Title ?? "").localeCompare(b.Title ?? ""));
  }, [capabilities, selectedContract]);

  const searchFilteredCapabilities = React.useMemo(() => {
    const search = searchTerm.toLowerCase().trim();

    if (!search) return capabilities;

    // Main grid search intentionally checks only direct capability list fields.
    // Nested people, lookup, and parsed tag objects are skipped here so typing cannot recurse
    // into SharePoint/React object metadata or match hidden relationship data unexpectedly.
    const searchableCapabilityFields: Array<keyof ICapabilityItem> = [
      "Title",
      "synonyms",
      "description",
      "capabilities",
      "link",
      "capStatus",
      "notes",
      "capabilityTypeTier1",
      "capabilityTypeTier2",
      "platform",
      "hostingEnv",
      "connectivity",
      "compliance",
      "licenseReqd",
      "licenseReqmts",
      "extensibility",
      "serverReqmts",
      "codeLanguage",
      "backend",
      "oppNetTagsJson",
      "pastPerformanceTagsJson",
      "proposalTagsJson",
      "Modified"
    ];

    // Normalize direct field values before matching so rich-text fields and casing do not
    // affect search results. Keep this helper primitive-only for easier debugging.
    const toSearchText = (value: unknown): string => {
      if (value === null || value === undefined) return "";
      if (typeof value === "string") return value.replace(/<[^>]*>/g, " ");
      if (typeof value === "number" || typeof value === "boolean") return String(value);
      return "";
    };

    return capabilities.filter((cap) => {
      // Build one searchable string per capability from direct fields only.
      const capabilitySearchText = searchableCapabilityFields
        .map((fieldName) => toSearchText(cap[fieldName]))
        .concat(
          DataSource.getConfigText("capabilityType", cap.capabilityTypeTier1),
          DataSource.getConfigText("capabilityType", cap.capabilityTypeTier2)
        )
        .concat(contractSummaryByCapabilityId.get(cap.Id)?.searchText ?? "")
        .join(" ")
        .toLowerCase();

      return capabilitySearchText.includes(search);
    });
  }, [capabilities, contractSummaryByCapabilityId, searchTerm]);

  /**
   * Final fully filtered capabilities
   */
  const filteredCapabilities = React.useMemo(() => {
    return searchFilteredCapabilities.filter((cap) => {
      const matchesCapStatus = capStatusFilter === "all" || cap.capStatus === capStatusFilter;
      const matchesCapabilityType = capabilityTypeFilter === "all"
        || cap.capabilityTypeTier1 === capabilityTypeFilter
        || cap.capabilityTypeTier2 === capabilityTypeFilter;
      const matchesPlatform = platformFilter === "all" || cap.platform === platformFilter;
      return matchesCapStatus && matchesCapabilityType && matchesPlatform;
    });
  }, [searchFilteredCapabilities, capStatusFilter, capabilityTypeFilter, platformFilter]);

  const handleResetFilters = (): void => {
    setSearchTerm("");
    setCapStatusFilter("all");
    setCapabilityTypeFilter("all");
    setPlatformFilter("all");
  };

  const closeContractForm = (): void => {
    setShowContractForm(false);
    setSelectedContract(undefined);
    setIsContractEditMode(false);
    setShowContractCapabilityDialog(false);
    setRelationshipCapability(undefined);
    setRelationshipSummary("");
    setRelationshipCapabilityToDelete(undefined);
    setSelectedRelationshipCapability(undefined);

    const parts = getPathParts(location.pathname);
    if ((parts[0] ?? "").toLowerCase() === "contracts" && parts[1]) {
      history.push(routes.contracts);
    }
  };

  const filtersActive =
    searchTerm !== "" ||
    capStatusFilter !== "all" ||
    capabilityTypeFilter !== "all" ||
    platformFilter !== "all";

  React.useEffect(() => {
    if (location.pathname !== currentPathRef.current) {
      previousPathRef.current = currentPathRef.current;
      currentPathRef.current = location.pathname;
    }
  }, [location.pathname]);

  React.useEffect(() => {
    const parts = getPathParts(location.pathname);
    const section = (parts[0] ?? "").toLowerCase();

    if (!section || section === "home" || section === "capabilities") {
      setShowAdminPanel(false);
      setSelectedCap(undefined);
          setSelectedContract(undefined);
          setShowContractForm(false);
          setIsContractEditMode(false);
          setSelectedPivot("caps");
          return;
    }

    if (section === "dashboard") {
      setShowAdminPanel(false);
      setSelectedCap(undefined);
      setSelectedContract(undefined);
      setShowContractForm(false);
      setIsContractEditMode(false);
      setSelectedPivot("dashboard");
      return;
    }

    if (section === "admin") {
      if (!Security.IsAdmin) {
        history.replace(routes.home);
        return;
      }

      setSelectedCap(undefined);
      setSelectedContract(undefined);
      setShowContractForm(false);
      setIsContractEditMode(false);
      setShowAdminPanel(true);
      return;
    }

    if (section === "contracts") {
      setShowAdminPanel(false);
      setSelectedCap(undefined);
      setSelectedPivot("contracts");

      const contractId = Number(parts[1]);
      if (Number.isFinite(contractId) && contractId > 0) {
        const contract = contracts.find(c => c.Id === contractId);
        if (contract) {
          setSelectedContract(contract);
          setShowContractForm(true);
          setIsContractEditMode(false);
        }
      } else {
        setSelectedContract(undefined);
        setShowContractForm(false);
        setIsContractEditMode(false);
      }

      return;
    }

    if (section === "caps") {
      setShowAdminPanel(false);
      setShowContractForm(false);
      setIsContractEditMode(false);

      const capId = Number(parts[1]);
      const routeAction = (parts[2] ?? "").toLowerCase();
      if (Number.isFinite(capId) && capId > 0 && routeAction !== "view") {
        history.replace(routes.cap(capId));
      }

      if (Number.isFinite(capId) && capId > 0) {
        const cap = capabilities.find(a => a.Id === capId);
        if (cap) {
          setSelectedCap(cap);
          setSelectedCapTab("overview");
        } else if (capabilities.length) {
          setSelectedCap(undefined);
        }
      } else {
        setSelectedCap(undefined);
      }
    }
  }, [capabilities, contracts, location.pathname]);

  /****
   * Options for dropdown filters
   */
  const capStatusOptions: IDropdownOption[] = React.useMemo(() => {
    return [
      { key: "all", text: "All Capability Statuses" },
      ...DataSource.getConfigOptions("capabilityStatus")
    ];
  }, [capabilities.length, loading]);

  const capabilityTypeOptions: IDropdownOption[] = React.useMemo(() => {
    return [
      { key: "all", text: "All Capability Types" },
      ...DataSource.getConfigOptions("capabilityType")
    ];
  }, [capabilities.length, loading]);

  const platformOptions: IDropdownOption[] = React.useMemo(() => {
    return [
      { key: "all", text: "All Platforms" },
      ...DataSource.getConfigOptions("platform")
    ];
  }, [capabilities.length, loading]);


  const handleExportCapsToExcel = (capabilities: ICapabilityItem[]): void => {
    setSpinnerMessage("Exporting Data Grid...");
    setShowSpinner(true);
    exportToExcel(filteredCapabilities).then(() => {
      setShowSpinner(false);
      setDialogProps("Export Capabilities", "Completed Export! Check your local Downloads folder for your file.");
    }, error => {
      setShowSpinner(false);
      const err = formatError(error);
      setDialogProps("Error Exporting Capabilities", err);
      console.error("Error Exporting Capabilities", err);
    })
  }

  // Export BOOK to PDF handler
  const handleExportCapabilitiesPdf = async (): Promise<void> => {
    setSpinnerMessage("Preparing Capability Summary Book (PDF)...");
    setShowSpinner(true);

    try {
      //const solutionIds: number[] = filteredSolutions.map(s => s.Id);

      // 1 call (chunked internally) to get screenshot metadata for these solutions
      const screenshotDocs = await DataSource.getScreenshotsForBook();

      // Build export payload
      const items = await buildPdfBookItems(
        props.context,
        filteredCapabilities,
        contracts,
        screenshotDocs
      );

      await exportCapabilitiesBookPdf({
        items,
        kgsLogoDataUrl: Strings.Logo,
        fileName: "KGSCapabilities"
      });

      setShowSpinner(false);
      setDialogProps("Export PDF", "Completed Export! Check your local Downloads folder for your file.");
    } catch (error) {
      setShowSpinner(false);
      const err: string = formatError(error);
      setDialogProps("Error Exporting PDF", err);
      console.error("Error Exporting PDF", err);
    }
  };

  const handleRefreshClick = (): void => {
    initDatasource(true).catch((error) => {
      const msg = formatError(error);
      console.error("Error refreshing datasource:", msg);
      setDialogProps("Error Refreshing Data", msg);
    });
  };

  const handleNewCapabilityClick = (): void => {
    setShowCapForm(true);
  };

  const handleNewContractClick = (): void => {
    setSelectedContract(undefined);
    setContractDocuments([]);
    setIsContractEditMode(true);
    setShowContractForm(true);
  };

  const contractDocumentTypeOptions = React.useMemo<IDropdownOption[]>(
    () => DataSource.getConfigOptions("cdocType"),
    [showContractDocUploadDialog]
  );

  React.useEffect(() => {
    if (selectedContract?.Id) {
      getContractDocuments(selectedContract.Id).catch((error) =>
        console.error(`Error loading contract documents: ${formatError(error)}`)
      );
    } else {
      setContractDocuments([]);
    }
  }, [selectedContract?.Id]);

  // Upload a selected file into the selected contract's dedicated document folder.
  const handleAddContractDocument = async (): Promise<void> => {
    if (!selectedContract?.Id || !contractDocumentType) return;

    const file = await Helper.ListForm.showFileDialog();
    if (!file || !file.src) {
      throw new Error("No file selected or invalid file structure.");
    }

    setSpinnerMessage("Uploading contract document...");
    setShowSpinner(true);

    try {
      const folderReady = await DocumentService.ensureContractDocumentFolder(selectedContract.Id);
      if (!folderReady) {
        await DocumentService.createContractFolder(selectedContract.Id);
      }

      const buffer: ArrayBuffer = await file.src.arrayBuffer();
      const customFile: CustomFile = Object.assign(file.src, { data: buffer });
      await DocumentService.uploadContractDocument(selectedContract.Id, file.name, customFile.data, contractDocumentType);
      await getContractDocuments(selectedContract.Id);
      setShowContractDocUploadDialog(false);
      setContractDocumentType(undefined);
    } finally {
      setShowSpinner(false);
    }
  };

  // Save contract document metadata edits and refresh the visible document list.
  const handleEditContractDocument = async (documentItem: IContractDocumentItem): Promise<void> => {
    if (!selectedContract?.Id) return;

    setSpinnerMessage("Updating contract document...");
    setShowSpinner(true);

    try {
      await DocumentService.editContractDocument(documentItem);
      await getContractDocuments(selectedContract.Id);
      setShowContractDocEditDialog(false);
      setSelectedContractDocument(undefined);
    } finally {
      setShowSpinner(false);
    }
  };

  // Delete a contract document and refresh the visible document list.
  const handleDeleteContractDocument = async (documentItem?: IContractDocumentItem): Promise<void> => {
    if (!selectedContract?.Id || !documentItem) return;

    setSpinnerMessage("Deleting contract document...");
    setShowSpinner(true);

    try {
      await DocumentService.deleteContractDocument(documentItem.Id);
      await getContractDocuments(selectedContract.Id);
      setShowContractDocDeleteDialog(false);
      setShowContractDocEditDialog(false);
      setSelectedContractDocument(undefined);
    } finally {
      setShowSpinner(false);
    }
  };

  const refreshSelectedContractRelationships = async (contractId: number): Promise<void> => {
    const refreshedContracts = await DataSource.refreshContracts();
    setContracts([...refreshedContracts]);
    setSelectedContract(refreshedContracts.find((contract) => contract.Id === contractId));
  };

  const openLinkCapabilityDialog = (): void => {
    setContractCapabilityDialogMode("link");
    setRelationshipCapability(undefined);
    setRelationshipSummary("");
    setShowContractCapabilityDialog(true);
  };

  const openEditCapabilitySummaryDialog = (capability: ICapabilityItem): void => {
    setContractCapabilityDialogMode("edit");
    setRelationshipCapability(capability);
    setRelationshipSummary(
      selectedContract
        ? DataSource.getContractCapabilitySummary(selectedContract.Id, capability.Id)?.summary ?? ""
        : ""
    );
    setShowContractCapabilityDialog(true);
  };

  const handleSaveContractCapability = async (): Promise<void> => {
    if (!selectedContract || !relationshipCapability) return;

    setSpinnerMessage(contractCapabilityDialogMode === "link" ? "Linking capability..." : "Updating relationship summary...");
    setShowSpinner(true);

    try {
      if (contractCapabilityDialogMode === "link") {
        await ContractService.linkCapability(
          selectedContract,
          relationshipCapability.Id,
          relationshipCapability.Title,
          relationshipSummary
        );
      } else {
        await ContractService.updateCapabilitySummary(
          selectedContract,
          relationshipCapability.Id,
          relationshipCapability.Title,
          relationshipSummary
        );
      }

      await refreshSelectedContractRelationships(selectedContract.Id);
      setShowContractCapabilityDialog(false);
      setRelationshipCapability(undefined);
      setRelationshipSummary("");
    } catch (error) {
      setDialogProps("Error saving capability relationship", formatError(error));
    } finally {
      setShowSpinner(false);
    }
  };

  const handleUnlinkCapability = async (): Promise<void> => {
    if (!selectedContract || !relationshipCapabilityToDelete) return;

    setSpinnerMessage("Unlinking capability...");
    setShowSpinner(true);

    try {
      await ContractService.unlinkCapability(selectedContract, relationshipCapabilityToDelete.Id);
      await refreshSelectedContractRelationships(selectedContract.Id);
      setRelationshipCapabilityToDelete(undefined);
    } catch (error) {
      setDialogProps("Error unlinking capability", formatError(error));
    } finally {
      setShowSpinner(false);
    }
  };

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: "admin",
      text: "Admin",
      iconProps: { iconName: "Settings" },
      title: "Administer Lookup Data",
      disabled: !Security.IsAdmin,
      onClick: () => history.push(routes.admin)
    },
    {
      key: "export",
      text: "Export",
      iconProps: { iconName: "Download" },
      title: "Export Solution Gallery",
      disabled: Security.IsVisitor,
      subMenuProps: {
        items: [
          {
            key: "exportExcel",
            text: "Export to Excel",
            title: "Export current (filtered) data set to a local Excel document",
            iconProps: { iconName: "ExcelDocument" },
              onClick: () => handleExportCapsToExcel(filteredCapabilities)
          },
          {
            key: "exportPdf",
            text: "Export to PDF (Summary Book)",
            iconProps: { iconName: "PDF" },
            onClick: () => {
              handleExportCapabilitiesPdf()
                .catch((e) => {
                  setDialogProps("Error Exporting Summary Book", formatError(e))
                })
            }
          }
        ]
      }
    },
    {
      key: "refresh",
      text: "Refresh",
      title: "Refresh all Data Sources",
      iconProps: { iconName: "Refresh" },
      onClick: handleRefreshClick
    }
  ];

  // Solution form style override
  const dialogStyles = mergeStyleSets({
    mainOverride: {
      width: '90vw !important',
      maxWidth: '900px !important',
      minWidth: '675px !important',
      height: 'auto'
    },
  });

  if (!props.context || !props.context.pageContext || !props.context.pageContext.web) {
    return (
      <div className={styles.dcTracker}>
        <MessageBar messageBarType={MessageBarType.error}>Error initializing the Tracker. Missing SharePoint context. Please try to refresh the browser.</MessageBar>
      </div>
    )
  }

  if (Security.RoleDisplay === "NoRole") {
    return (
      <div className={styles.dcTracker}>
        <div className={styles.warningBanner}>
          <Icon iconName="Warning" className={styles.warningIcon} aria-hidden />

          <div>
            <Text variant="mediumPlus" block>Access required</Text>

            <Text variant="small" block>
              You don’t currently have a role assigned for the{" "}<strong>{Strings.ProjectName}</strong> capability.
            </Text>

            <Text variant="small" block>
              Please contact the Site Owner or the {" "}<a href="mailto:Justin.White@koniag-gs.com">KGS Custom Apps Support Team</a>{" "} to request access.
            </Text>
          </div>
        </div>
      </div>
    );
  }

  if (loading) {
    return (
      <div className={styles.dcTracker}>
        <Stack className={styles.loadingShell} horizontalAlign="center" verticalAlign="center">
          <Spinner size={SpinnerSize.large} label="Initializing Defense Capabilities Tracker..." />
        </Stack>
      </div>
    )
  }

  return (
    <ThemeProvider theme={appTheme}>
      <div className={styles.dcTracker}>
        {showAdminPanel ? (
          <AdminPanel
            context={props.context}
            onNewCapability={handleNewCapabilityClick}
            onBack={() => {
              history.push(routes.home);
              initDatasource(true).catch((error) =>
                console.error(`Error refreshing datasource: ${formatError(error)}`)
              )
            }}
          />
        ) : selectedCap ? (
          <CapDetails
            capability={selectedCap}
            onBack={handleCapDetailsBack}
            activeTab={selectedCapTab}
            onTabChange={setSelectedCapTab}
            onContractsChanged={setContracts}
            onNewCapability={handleNewCapabilityClick}
            context={props.context}
          />
        ) : (
          <Stack>
            {/* Header */}
            <AppHeader onNewCapability={handleNewCapabilityClick} />

            <div className={styles.mainNavBar}>
              <div className={styles.mainNavTabs}>
                <Pivot
                  styles={customPivotStyles}
                  selectedKey={selectedPivot}
                  onLinkClick={(item) => {
                    if (!item) return;
                    const key = item?.props.itemKey;
                    if (key === "caps") {
                      history.push(routes.capabilities);
                      handleRefreshClick();
                    } else if (key === "contracts") {
                      history.push(routes.contracts);
                    } else {
                      history.push(routes.dashboard);
                    }
                  }}
                  linkFormat="tabs"
                >
                  <PivotItem
                    headerText="Capabilities"
                    title="View all Capabilities"
                    ariaLabel="View all Capabilities"
                    itemKey="caps"
                    itemIcon="ProductCatalog"
                  />                  
                  <PivotItem
                    headerText="Contracts"
                    title="View Supporting Contract Information"
                    ariaLabel="View Supporting Contract Information"
                    itemKey="contracts"
                    itemIcon="CompanyDirectory"
                  />
                  <PivotItem
                    headerText="Dashboard"
                    title="View Dashboard Summary"
                    ariaLabel="View Dashboard Summary"
                    itemKey="dashboard"
                    itemIcon="Chart"
                  />
                </Pivot>
              </div>
              <div className={styles.mainNavCommands}>
                <CommandBar items={commandBarItems} styles={{ root: { padding: 0, background: "transparent" } }} />
              </div>
            </div>

            {/* Main Content */}
            <div className={styles.pageContent}>

              {selectedPivot === "caps" && (
                <Stack tokens={{ childrenGap: 15 }} styles={{ root: { marginTop: 20, marginBottom: 20 } }}>
                  {/* Filters */}
                  <Stack tokens={{ childrenGap: 10 }}>
                    <SearchBox
                      placeholder="Search any capability field..."
                      value={searchTerm}
                      onChange={(_, newValue) => setSearchTerm(newValue || "")}
                      styles={{ root: { width: "100%", maxWidth: 700 } }}
                    />
                    <Stack horizontal tokens={{ childrenGap: 10 }} wrap verticalAlign='center'>
                      <Dropdown
                        placeholder="Filter by Capability Status"
                        options={capStatusOptions}
                        selectedKey={capStatusFilter}
                        onChange={(_, option) => setCapStatusFilter((option?.key as string) ?? "all")}
                        styles={{ dropdown: { width: 180 } }}
                      />
                      <Dropdown
                        placeholder="Filter by Capability Type"
                        options={capabilityTypeOptions}
                        selectedKey={capabilityTypeFilter}
                        onChange={(_, option) => setCapabilityTypeFilter((option?.key as string) ?? "all")}
                        styles={{ root: { width: 425, maxWidth: "100%" }, dropdown: { width: "100%" } }}
                      />
                      <Dropdown
                        placeholder="Filter by Platform"
                        options={platformOptions}
                        selectedKey={platformFilter}
                        onChange={(_, option) => setPlatformFilter((option?.key as string) ?? "all")}
                        styles={{ dropdown: { width: 160 } }}
                      />

                      <IconButton
                        iconProps={{ iconName: "ClearFilter" }}
                        title="Reset filters"
                        ariaLabel="Reset filters"
                        disabled={!filtersActive}
                        onClick={handleResetFilters}
                      />
                    </Stack>
                  </Stack>

                  {/* View Toggle */}
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="medium">
                      Showing {filteredCapabilities.length} of {capabilities.length} capabilities
                    </Text>
                    <Stack horizontal tokens={{ childrenGap: 5 }} verticalAlign='center'>
                      <Text variant="medium">View:</Text>
                      <DefaultButton
                        text="List"
                        title="Switch to List view"
                        iconProps={{ iconName: "List" }}
                        primary={viewMode === "list"}
                        onClick={() => setViewMode("list")}
                      />

                      <DefaultButton
                        text="Tiles"
                        title="Switch to Tile view"
                        iconProps={{ iconName: "GridViewMedium" }}
                        primary={viewMode === "tile"}
                        onClick={() => setViewMode("tile")}
                      />
                    </Stack>
                  </Stack>

                  {loading ? null : (
                    <CapabilitiesList
                      capabilities={filteredCapabilities}
                      contractSummaryByCapabilityId={contractSummaryByCapabilityId}
                      viewMode={viewMode}
                      onSelectCap={handleSelectedCap}
                    />
                  )}

                </Stack>
              )}

              {selectedPivot === "dashboard" && (
                <AppDashboard capabilities={capabilities} />
              )}

              {selectedPivot === "contracts" && (
                <ContractsList
                  contracts={contracts}
                  onSelectContract={handleSelectedContract}
                  onNewContract={!Security.IsVisitor ? handleNewContractClick : undefined}
                />
              )}
            </div>
          </Stack>
        )}

        {/* Capability Form Dialog */}
        <Dialog
          hidden={!showCapForm}
          onDismiss={() => setShowCapForm(false)}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: `Create New Capability`,
            showCloseButton: true
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: dialogStyles.mainOverride }
          }}
        >
          <CapForm
            context={props.context}
            onCancel={() => setShowCapForm(false)}
            onSave={async (result: ICapFormSaveResult) => {
              setSpinnerMessage("Creating new Capacity...");
              setShowSpinner(true);
              try {
                const capResponse = await CapabilityService.create(result.capability)
                //create folder for related docs
                await DocumentService.createCapabilityFolder(capResponse.Id);
                await ContractService.saveForCapability(capResponse.Id, capResponse.Title, result.contracts, result.deletedContractIds);

                setSpinnerMessage("Refreshing data...");

                await initDatasource(true);

                setDialogProps("Success!", "You successfully created a new Capability!");

                setSelectedCap(capResponse);
                setShowCapForm(false);
                setShowSpinner(false);
                history.push(routes.cap(capResponse.Id));

              } catch (error) {
                const fError = formatError(error);
                console.error(`Error creating Capability: ${fError}`);
                setDialogProps("Error creating Capability", fError);
              } finally {
                setShowSpinner(false);
              }
            }}
            onDelete={async () => {
              if (selectedCap) {
                setSpinnerMessage("Deleting Capability Entry...");
                setShowSpinner(true);
                try {
                  await ContractService.removeCapabilityFromAllContracts(selectedCap.Id);
                  await CapabilityService.delete(selectedCap.Id)
                  setCapabilities((prevCaps) => prevCaps.filter(a => a.Id !== selectedCap.Id));
                  setContracts((prevContracts) =>
                    prevContracts
                      .map((contract) => ({
                        ...contract,
                        capability: {
                          results: ContractService.getCapabilityLookups(contract)
                            .filter((capability) => capability.Id !== selectedCap.Id)
                        }
                      }))
                      .filter((contract) => ContractService.getCapabilityLookups(contract).length > 0)
                  );
                  setSelectedCap(undefined);
                  setShowCapForm(false);
                  history.push(routes.capabilities);
                } catch (error) {
                  const fError = formatError(error);
                  console.error(`Error deleting Capability: ${fError}`);
                  setDialogProps("Error deleting Capability", fError);
                } finally {
                  setShowSpinner(false);
                }
              } else {
                setDialogProps("No Capability selected", "No Capability was selected. Please try again.")
                return;
              }
            }}
          />

        </Dialog>

        {/* Contract Form Dialog */}
        <Dialog
          hidden={!showContractForm}
          onDismiss={closeContractForm}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: isContractEditMode
              ? selectedContract ? "Edit Contract" : "New Contract"
              : "Contract Details",
            showCloseButton: true
          }}
          modalProps={{
            isBlocking: true,
            styles: { main: dialogStyles.mainOverride }
          }}
        >
          {selectedContract && !isContractEditMode && (
            <Stack tokens={{ childrenGap: 16 }}>
              <div className={styles.contractModalGrid}>
                <ContractDetailCard
                  contract={selectedContract}
                  eyebrow="Contract"
                  relatedCapabilities={selectedContractCapabilities}
                  documentsContent={(
                    <ContractDocumentsPanel
                      documents={contractDocuments}
                      canEdit={!Security.IsVisitor}
                      onEdit={(documentItem) => {
                        setSelectedContractDocument(documentItem);
                        setShowContractDocEditDialog(true);
                      }}
                      onDelete={(documentItem) => {
                        setSelectedContractDocument(documentItem);
                        setShowContractDocDeleteDialog(true);
                      }}
                      showAddButton={false}
                    />
                  )}
                  onSelectCapability={(capability) => {
                    setShowContractForm(false);
                    setSelectedContract(undefined);
                    setIsContractEditMode(false);
                    history.push(routes.cap(capability.Id));
                  }}
                  onViewCapabilityRelationship={(capability) => setSelectedRelationshipCapability(capability)}
                  onEditCapabilityRelationship={!Security.IsVisitor ? openEditCapabilitySummaryDialog : undefined}
                  onDeleteCapabilityRelationship={!Security.IsVisitor ? setRelationshipCapabilityToDelete : undefined}
                />
              </div>

              <div className={styles.contractViewActions}>
                <div className={styles.contractViewActionsLeft}>
                  {!Security.IsVisitor && (
                    <>
                      <PrimaryButton
                        text="Add Document"
                        iconProps={{ iconName: "Add" }}
                        styles={spoTealButtonStyles}
                        onClick={() => {
                          setContractDocumentType(undefined);
                          setShowContractDocUploadDialog(true);
                        }}
                      />
                      <PrimaryButton
                        text="Link Capability"
                        iconProps={{ iconName: "Link" }}
                        styles={spoTealButtonStyles}
                        onClick={openLinkCapabilityDialog}
                      />
                    </>
                  )}
                </div>

                <div className={styles.contractViewActionsRight}>
                {!Security.IsVisitor && (
                    <PrimaryButton
                      text="Edit"
                      iconProps={{ iconName: "Edit" }}
                      onClick={() => setIsContractEditMode(true)}
                    />
                )}
                <DefaultButton text="Close" onClick={closeContractForm} />
                </div>
              </div>
            </Stack>
          )}

          {isContractEditMode && (
            <Stack tokens={{ childrenGap: 16 }}>
              <ContractForm
                item={selectedContract}
                context={props.context}
                onCancel={() => selectedContract ? setIsContractEditMode(false) : closeContractForm()}
                onSave={async (item) => {
                  const isNewContract = !selectedContract;
                  setSpinnerMessage(isNewContract ? "Creating Contract..." : "Editing Contract...");
                  setShowSpinner(true);

                  try {
                    const savedContract = isNewContract
                      ? await ContractService.create(item)
                      : await ContractService.edit(item);
                    setSpinnerMessage("Refreshing data...");
                    const refreshedContracts = await DataSource.refreshContracts();
                    const refreshedSavedContract = refreshedContracts.find((contract) => contract.Id === savedContract.Id) ?? savedContract;
                    setContracts([...refreshedContracts]);
                    setSelectedContract(refreshedSavedContract);
                    setIsContractEditMode(false);
                    history.push(routes.contract(savedContract.Id));
                    setShowSpinner(false);
                  } catch (err) {
                    console.error(`Error ${isNewContract ? "creating" : "editing"} Contract: ${formatError(err)}`);
                    setShowSpinner(false);
                    setDialogProps(`Error ${isNewContract ? "creating" : "editing"} Contract`, formatError(err));
                  }
                }}
              >
                {selectedContract && (
                  <section className={styles.contractDocumentsCard}>
                    <ContractDocumentsPanel
                      documents={contractDocuments}
                      canEdit={!Security.IsVisitor}
                      onAdd={() => {
                        setContractDocumentType(undefined);
                        setShowContractDocUploadDialog(true);
                      }}
                      onEdit={(documentItem) => {
                        setSelectedContractDocument(documentItem);
                        setShowContractDocEditDialog(true);
                      }}
                      onDelete={(documentItem) => {
                        setSelectedContractDocument(documentItem);
                        setShowContractDocDeleteDialog(true);
                      }}
                    />
                  </section>
                )}
              </ContractForm>
            </Stack>
          )}
        </Dialog>

        {/* Add or edit a contract-capability relationship and its summary. */}
        <Dialog
          hidden={!showContractCapabilityDialog}
          onDismiss={() => setShowContractCapabilityDialog(false)}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: contractCapabilityDialogMode === "link" ? "Link Capability" : "Edit Capability Summary",
            closeButtonAriaLabel: "Close",
            subText: selectedContract?.Title
          }}
          modalProps={{
            isBlocking: true,
            styles: { main: { width: "620px !important", maxWidth: "90vw !important" } }
          }}
        >
          <ContractCapabilityForm
            capabilities={capabilities}
            linkedCapabilityIds={new Set(ContractService.getCapabilityLookups(selectedContract).map((capability) => capability.Id))}
            selectedCapability={relationshipCapability}
            summary={relationshipSummary}
            isEditMode={contractCapabilityDialogMode === "edit"}
            onCapabilityChange={setRelationshipCapability}
            onSummaryChange={setRelationshipSummary}
          />
          <DialogFooter>
            <PrimaryButton
              text={contractCapabilityDialogMode === "link" ? "Link Capability" : "Save Summary"}
              iconProps={{ iconName: contractCapabilityDialogMode === "link" ? "Link" : "Save" }}
              disabled={!relationshipCapability}
              onClick={() => { handleSaveContractCapability().catch(() => undefined); }}
            />
            <DefaultButton text="Cancel" onClick={() => setShowContractCapabilityDialog(false)} />
          </DialogFooter>
        </Dialog>

        {/* Confirm that only the relationship—not either record—will be removed. */}
        <Dialog
          hidden={!relationshipCapabilityToDelete}
          onDismiss={() => setRelationshipCapabilityToDelete(undefined)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Unlink Capability",
            closeButtonAriaLabel: "Cancel",
            subText: relationshipCapabilityToDelete
              ? `Remove the link between ${relationshipCapabilityToDelete.Title} and ${selectedContract?.Title ?? "this contract"}? The capability and contract records will not be deleted.`
              : undefined
          }}
          modalProps={{ isBlocking: true }}
        >
          <DialogFooter>
            <PrimaryButton
              text="Unlink"
              className={styles.deleteButton}
              iconProps={{ iconName: "Delete" }}
              onClick={() => { handleUnlinkCapability().catch(() => undefined); }}
            />
            <DefaultButton text="Cancel" onClick={() => setRelationshipCapabilityToDelete(undefined)} />
          </DialogFooter>
        </Dialog>

        {/* Contract-capability relationship summary dialog */}
        <Dialog
          hidden={!selectedRelationshipCapability}
          onDismiss={() => setSelectedRelationshipCapability(undefined)}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: selectedRelationshipCapability?.Title ?? "Relationship Summary",
            closeButtonAriaLabel: 'Close',
            subText: selectedContract?.Title
          }}
          modalProps={{
            isBlocking: false,
            styles: { main: { width: "560px !important", maxWidth: "90vw !important" } }
          }}
        >
          <Stack tokens={{ childrenGap: 12 }}>
            <div className={styles.contractRelationshipSummary}>
              <span className={styles.contractRelationshipSummaryTitle}>Capability Summary</span>
              <span className={styles.contractRelationshipSummaryText}>
                {selectedContract && selectedRelationshipCapability
                  ? DataSource.getContractCapabilitySummary(selectedContract.Id, selectedRelationshipCapability.Id)?.summary || "No contract-specific summary has been added."
                  : ""}
              </span>
            </div>
          </Stack>
          <DialogFooter>
            {selectedRelationshipCapability && (
              <PrimaryButton
                text="View Capability"
                onClick={() => {
                  const capabilityId = selectedRelationshipCapability.Id;
                  setSelectedRelationshipCapability(undefined);
                  setShowContractForm(false);
                  setSelectedContract(undefined);
                  setIsContractEditMode(false);
                  history.push(routes.cap(capabilityId));
                }}
              />
            )}
            <DefaultButton text="Close" onClick={() => setSelectedRelationshipCapability(undefined)} />
          </DialogFooter>
        </Dialog>

        {/* Contract document upload dialog */}
        <Dialog
          hidden={!showContractDocUploadDialog}
          onDismiss={() => setShowContractDocUploadDialog(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Upload Contract Document",
            closeButtonAriaLabel: 'Close',
            subText: "Select a contract document type before uploading."
          }}
        >
          <Dropdown
            label="Contract Document Type"
            required
            options={contractDocumentTypeOptions}
            selectedKey={contractDocumentType}
            onChange={(_, option?: IDropdownOption) => setContractDocumentType((option?.key as string) ?? undefined)}
            style={{ marginBottom: 20 }}
          />
          <DialogFooter>
            <PrimaryButton
              text="Upload Document"
              iconProps={{ iconName: "Add" }}
              disabled={!contractDocumentType}
              title="Upload a new Contract Document"
              onClick={async () => {
                try {
                  await handleAddContractDocument();
                } catch (error) {
                  setShowContractDocUploadDialog(false);
                  const errorMessage = formatError(error);
                  setDialogProps("Error uploading contract document", errorMessage);
                  console.error("Error uploading contract document", errorMessage);
                }
              }}
            />
            <DefaultButton onClick={() => setShowContractDocUploadDialog(false)} text="Cancel" title="Close Dialog Box" />
          </DialogFooter>
        </Dialog>

        {/* Contract document metadata dialog */}
        <Dialog
          hidden={!showContractDocEditDialog}
          onDismiss={() => setShowContractDocEditDialog(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Edit Contract Document",
            closeButtonAriaLabel: 'Close'
          }}
        >
          {selectedContractDocument && (
            <ContractDocumentForm
              item={selectedContractDocument}
              onSave={handleEditContractDocument}
              onCancel={() => setShowContractDocEditDialog(false)}
              onDelete={(documentItem) => {
                setSelectedContractDocument(documentItem);
                setShowContractDocDeleteDialog(true);
              }}
            />
          )}
        </Dialog>

        {/* Contract document delete confirmation */}
        <Dialog
          hidden={!showContractDocDeleteDialog}
          onDismiss={() => setShowContractDocDeleteDialog(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Delete Contract Document",
            closeButtonAriaLabel: 'Cancel',
            subText: "Are you sure you want to delete this contract document?"
          }}
        >
          <DialogFooter>
            <PrimaryButton
              text="Delete"
              className={styles.deleteButton}
              iconProps={{ iconName: "Delete" }}
              onClick={() => handleDeleteContractDocument(selectedContractDocument)}
              title="Delete this Contract Document"
            />
            <DefaultButton onClick={() => setShowContractDocDeleteDialog(false)} text="Cancel" title="Close Dialog Box" />
          </DialogFooter>
        </Dialog>

        {/* Error dialog */}
        <Dialog
          hidden={!showDialog}
          onDismiss={hideDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: dialogTitle,
            closeButtonAriaLabel: 'Close',
            subText: dialogMessage
          }}
        >
          <DialogFooter>
            <DefaultButton onClick={hideDialog} text="Close" title="Close Dialog Box" />
          </DialogFooter>
        </Dialog>

        {/* Loading Spinner Dialog */}
        <Dialog
          hidden={!showSpinner}
          onDismiss={() => { setShowSpinner(false) }}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Loading...",
            closeButtonAriaLabel: 'Close',
          }}
        >
          <Spinner size={SpinnerSize.large} label={spinnerMessage} />
        </Dialog>

        <Stack horizontalAlign="end" style={{ maxWidth: 1600, marginLeft:"auto", marginRight: "auto" }}>
          <Text variant='xSmall'>Tracker Version: {Strings.Version}</Text>
        </Stack>

      </div>
    </ThemeProvider>
  )
}

const App: React.FC<IDCTrackerProps> = (props) => (
  <HashRouter>
    <DctContent {...props} />
  </HashRouter>
);

export default App;
