import * as React from 'react';
import { useState, useEffect } from "react"
import {
  Stack, Text, DefaultButton, SearchBox, Dropdown, IDropdownOption, Pivot, PivotItem, Spinner, SpinnerSize,
  CommandBar, Dialog, DialogType, DialogFooter, ICommandBarItemProps, mergeStyleSets,
  MessageBar, MessageBarType, Icon,
  IconButton
} from "@fluentui/react"
import { DataSource } from './data/ds';
import { formatError } from './common/utils';
import Strings from './common/strings';
import { ICapabilityItem, ICapFormSaveResult, IContractItem, IDCTrackerProps } from './common/props';
import { customPivotStyles } from './ui/ComponentStyles';
import { CapabilitiesList } from './views/CapabilitiesView';
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
import { AppDashboard } from './views/Dashboard';
import { Security } from './services/Security';
import { exportCapabilitiesBookPdf } from './export/ExportPdfWrapper';
import { buildPdfBookItems } from './export/ExportPdfUtils';
import { AppHeader } from './ui/AppHeader';
import { HashRouter, useHistory, useLocation } from 'react-router-dom';
import { CapRouteTab, getPathParts, routes, tabFromSlug } from './routing/routes';

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

  const [searchTerm, setSearchTerm] = useState<string>("");
  const [capStatusFilter, setCapStatusFilter] = useState<string>("all");
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
    history.push(routes.cap(capItem.Id, "overview"));
  }

  const handleSelectedContract = (contractItem?: IContractItem): void => {
    if (contractItem?.Id) {
      history.push(routes.contract(contractItem.Id));
    } else {
      setSelectedContract(undefined);
      setShowContractForm(false);
    }
  }

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

  const searchFilteredCapabilities = React.useMemo(() => {
    const search = searchTerm.toLowerCase().trim();

    if (!search) return capabilities;

    const toSearchText = (value?: string): string => (value ?? "").replace(/<[^>]*>/g, " ").toLowerCase();

    return capabilities.filter((cap) => {
      return (
        (cap.Title ?? "").toLowerCase().includes(search) ||
        (cap.description ?? "").toLowerCase().includes(search) ||
        toSearchText(cap.capabilities).includes(search) ||
        toSearchText(cap.notes).includes(search) ||
        (cap.link ?? "").toLowerCase().includes(search)
      );
    });
  }, [capabilities, searchTerm]);

  /**
   * Final fully filtered capabilities
   */
  const filteredCapabilities = React.useMemo(() => {
    return searchFilteredCapabilities.filter((cap) => {
      const matchesCapStatus = capStatusFilter === "all" || cap.capStatus === capStatusFilter;
      const matchesPlatform = platformFilter === "all" || cap.platform === platformFilter;
      return matchesCapStatus && matchesPlatform;
    });
  }, [searchFilteredCapabilities, capStatusFilter, platformFilter]);

  const handleResetFilters = (): void => {
    setSearchTerm("");
    setCapStatusFilter("all");
    setPlatformFilter("all");
  };

  const closeContractForm = (): void => {
    setShowContractForm(false);
    setSelectedContract(undefined);

    const parts = getPathParts(location.pathname);
    if ((parts[0] ?? "").toLowerCase() === "contracts" && parts[1]) {
      history.push(routes.contracts);
    }
  };

  const filtersActive =
    searchTerm !== "" ||
    capStatusFilter !== "all" ||
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
      setSelectedPivot("caps");
      return;
    }

    if (section === "dashboard") {
      setShowAdminPanel(false);
      setSelectedCap(undefined);
      setSelectedContract(undefined);
      setShowContractForm(false);
      setSelectedPivot("dashboard");
      return;
    }

    if (section === "admin") {
      setSelectedCap(undefined);
      setSelectedContract(undefined);
      setShowContractForm(false);
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
        }
      } else {
        setSelectedContract(undefined);
        setShowContractForm(false);
      }

      return;
    }

    if (section === "caps") {
      setShowAdminPanel(false);
      setShowContractForm(false);

      const capId = Number(parts[1]);
      const tab = tabFromSlug(parts[2]);
      setSelectedCapTab(tab);

      if (Number.isFinite(capId) && capId > 0) {
        const cap = capabilities.find(a => a.Id === capId);
        if (cap) {
          setSelectedCap(cap);
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
  }, []);

  const platformOptions: IDropdownOption[] = React.useMemo(() => {
    return [
      { key: "all", text: "All Platforms" },
      ...DataSource.getConfigOptions("platform")
    ];
  }, []);


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

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: "newItem",
      text: "New",
      title: "Create a new Record",
      iconProps: { iconName: "Add" },
      disabled: Security.IsVisitor,
      subMenuProps: {
        items: [
          {
            key: "newCap",
            text: "New Capability",
            title: "Create a new Capability",
            onClick: () => { setShowCapForm(true); }
          }
        ],
      },
    },
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
        <Stack horizontalAlign="center" verticalAlign="center" style={{ height: "200px" }}>
          <Spinner size={SpinnerSize.large} label="Initializing Defense Capabilities Tracker..." />
        </Stack>
      </div>
    )
  }

  if (showAdminPanel) {
    return (
      <div className={styles.dcTracker}>
        <AdminPanel
          context={props.context}
          onBack={() => {
            history.push(routes.home);
            initDatasource(true).catch((error) =>
              console.error(`Error refreshing datasource: ${formatError(error)}`)
            )
          }}
        />
      </div>
    )
  }

  return (
    <ThemeProvider theme={appTheme}>
      <div className={styles.dcTracker}>
        {selectedCap ? (
          <CapDetails
            capability={selectedCap}
            onBack={handleCapDetailsBack}
            onHome={() => history.push(routes.home)}
            activeTab={selectedCapTab}
            onTabChange={(tab) => history.push(routes.cap(selectedCap.Id, tab))}
            context={props.context}
          />
        ) : (
          <Stack>
            {/* Header */}
            <AppHeader />

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
                    headerText="Dashboard"
                    title="View Dashboard Summary"
                    ariaLabel="View Dashboard Summary"
                    itemKey="dashboard"
                    itemIcon="Chart"
                  />
                  {Security.IsAdmin && (
                    <PivotItem
                      headerText="Contracts"
                      title="View Supporting Contract Information"
                      ariaLabel="View Supporting Contract Information"
                      itemKey="contracts"
                      itemIcon="CompanyDirectory"
                    />
                  )}
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
                  <Stack horizontal tokens={{ childrenGap: 10 }} wrap verticalAlign='center'>
                    <SearchBox
                      placeholder="Search capabilities, descriptions, notes, links..."
                      value={searchTerm}
                      onChange={(_, newValue) => setSearchTerm(newValue || "")}
                      styles={{ root: { width: 325 } }}
                    />
                    <Dropdown
                      placeholder="Filter by Capability Status"
                      options={capStatusOptions}
                      selectedKey={capStatusFilter}
                      onChange={(event, option) => setCapStatusFilter(option?.key as string)}
                      styles={{ dropdown: { width: 150 } }}
                    />
                    <Dropdown
                      placeholder="Filter by Platform"
                      options={platformOptions}
                      selectedKey={platformFilter}
                      onChange={(event, option) => setPlatformFilter(option?.key as string)}
                      styles={{ dropdown: { width: 220 } }}
                    />

                    <IconButton
                      iconProps={{ iconName: "ClearFilter" }}
                      title="Reset filters"
                      ariaLabel="Reset filters"
                      disabled={!filtersActive}
                      onClick={handleResetFilters}
                    />
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

                      {/**** hide/disable for now ************
                       *  <DefaultButton
                        text="Tiles"
                        title="Switch to Tile view"
                        iconProps={{ iconName: "GridViewMedium" }}
                        primary={viewMode === "tile"}
                        onClick={() => setViewMode("tile")}
                      /> */}
                    </Stack>
                  </Stack>

                  {loading ? null : <CapabilitiesList capabilities={filteredCapabilities} viewMode={viewMode} onSelectCap={handleSelectedCap} />}

                </Stack>
              )}

              {selectedPivot === "dashboard" && (
                <AppDashboard capabilities={capabilities} />
              )}

              {selectedPivot === "contracts" && Security.IsAdmin && (
                <ContractsList contracts={contracts} onSelectContract={handleSelectedContract} />
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
            item={selectedCap}
            context={props.context}
            onCancel={() => setShowCapForm(false)}
            onSave={async (result: ICapFormSaveResult) => {
              setSpinnerMessage("Creating new Capacity...");
              setShowSpinner(true);
              try {
                const capResponse = await CapabilityService.create(result.capability)
                //create folder for related docs
                await DocumentService.createCapabilityFolder(capResponse.Id);
                await ContractService.saveForCapability(capResponse.Id, result.contracts, result.deletedContractIds);

                setSpinnerMessage("Refreshing data...");

                await initDatasource(true);

                setDialogProps("Success!", "You successfully created a new Capability!");

                setSelectedCap(capResponse);
                setShowCapForm(false);
                setShowSpinner(false);
                history.push(routes.cap(capResponse.Id, "overview"));

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
                  const relatedContracts = contracts.filter((contract) => contract.capability?.Id === selectedCap.Id);
                  await Promise.all(relatedContracts.map((contract) => ContractService.delete(contract.Id)));
                  await CapabilityService.delete(selectedCap.Id)
                  setCapabilities((prevCaps) => prevCaps.filter(a => a.Id !== selectedCap.Id));
                  setContracts((prevContracts) => prevContracts.filter(c => c.capability?.Id !== selectedCap.Id));
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
            title: "Edit Contract",
            showCloseButton: true
          }}
          modalProps={{
            isBlocking: true,
            styles: { main: dialogStyles.mainOverride }
          }}
        >
          <ContractForm
            item={selectedContract}
            context={props.context}
            onCancel={closeContractForm}
            onSave={async (item) => {
              if (selectedContract) {
                //edit contract
                setSpinnerMessage("Editing Contract...");
                setShowSpinner(true);

                try {
                  const editContract = await ContractService.edit(item);
                  // Refresh data
                  setSpinnerMessage("Refreshing data...");
                  setContracts((prevContracts) => prevContracts.map((c) => c.Id === selectedContract.Id ? editContract : c));
                  closeContractForm();
                  setShowSpinner(false);
                } catch (err) {
                  console.error(`Error editing Contract: ${formatError(err)}`);
                  setShowSpinner(false);
                  setSelectedContract(undefined);
                  setDialogProps("Error editing Contract", formatError(err));
                }

              } else {
                setDialogProps("No Contract selected", "Contracts must be added from a Capability first.");
              }
            }}
            onDelete={async () => {
              if (selectedContract) {
                setSpinnerMessage("Deleting Contract Entry...");
                setShowSpinner(true);
                try {
                  await ContractService.delete(selectedContract.Id)
                  setContracts((prevContracts) => prevContracts.filter(c => c.Id !== selectedContract.Id));
                  closeContractForm();
                } catch (error) {
                  const fError = formatError(error);
                  console.error(`Error deleting Contract: ${fError}`);
                  setDialogProps("Error deleting Contract", fError);
                } finally {
                  setShowSpinner(false);
                }
              } else {
                setDialogProps("No Contract selected", "No Contract was selected. Please try again.")
                return;
              }
            }}
          />
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

        <Stack horizontalAlign='end' style={{ paddingTop: 10 }}>
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
