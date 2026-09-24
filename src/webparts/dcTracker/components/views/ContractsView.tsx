import * as React from "react";
import { Dropdown, IColumn, IDropdownOption, Icon, IconButton, Link, PrimaryButton, SearchBox, SelectionMode, Stack, Text } from "@fluentui/react";
import PaginatedDetailsList from "../ui/PaginatedDetailsList";
import { formatDate } from "../common/utils";
import styles from "../Dct.module.scss";
import { ContractService } from "../services/ContractService";

import type { IContractItem, IPeoplePickerExtended } from "../common/props";

interface IContractsListProps {
  contracts: IContractItem[];
  onSelectContract: (contract: IContractItem) => void;
  onNewContract?: () => void;
}

const renderPerson = (p?: IPeoplePickerExtended): string => p?.Title ?? "";

const getFilterOptions = (allText: string, values: Array<string | undefined>): IDropdownOption[] => [
  { key: "all", text: allText },
  ...Array.from(new Set(values.map((value) => value?.trim()).filter((value): value is string => Boolean(value))))
    .sort((a, b) => a.localeCompare(b))
    .map((value) => ({ key: value, text: value }))
];

export const ContractsList: React.FunctionComponent<IContractsListProps> = ({ contracts, onSelectContract, onNewContract }) => {
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [ogFilter, setOgFilter] = React.useState<string>("all");
  const [customerFilter, setCustomerFilter] = React.useState<string>("all");
  const [partnerFilter, setPartnerFilter] = React.useState<string>("all");
  const [clearanceFilter, setClearanceFilter] = React.useState<string>("all");
  const [sortColumnKey, setSortColumnKey] = React.useState<string | null>("title");
  const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

  const ogOptions = React.useMemo<IDropdownOption[]>(
    () => getFilterOptions("All OGs", contracts.map((contract) => contract.ogTitle)),
    [contracts]
  );
  const customerOptions = React.useMemo<IDropdownOption[]>(
    () => getFilterOptions("All Customers", contracts.map((contract) => contract.customer)),
    [contracts]
  );
  const partnerOptions = React.useMemo<IDropdownOption[]>(
    () => getFilterOptions("All Partners", contracts.map((contract) => contract.partner)),
    [contracts]
  );
  const clearanceOptions = React.useMemo<IDropdownOption[]>(
    () => getFilterOptions("All Clearance Levels", contracts.map((contract) => contract.clearance)),
    [contracts]
  );

  const filteredContracts = React.useMemo(() => {
    const search = searchTerm.trim().toLowerCase();
    return contracts.filter((contract) => {
      const matchesSearch = !search || (
        (contract.Title ?? "").toLowerCase().includes(search) ||
        (contract.synonyms ?? "").toLowerCase().includes(search) ||
        (contract.contractId ?? "").toLowerCase().includes(search) ||
        (contract.customerContractCode ?? "").toLowerCase().includes(search) ||
        (contract.customer ?? "").toLowerCase().includes(search) ||
        (contract.clearance ?? "").toLowerCase().includes(search) ||
        (contract.ogTitle ?? "").toLowerCase().includes(search) ||
        (contract.lobTitle ?? "").toLowerCase().includes(search) ||
        (contract.location ?? "").toLowerCase().includes(search) ||
        (contract.contractPm?.Title ?? "").toLowerCase().includes(search) ||
        (contract.partner ?? "").toLowerCase().includes(search) ||
        ContractService.getCapabilityLookups(contract).some((capability) =>
          (capability.Title ?? "").toLowerCase().includes(search)
        )
      );

      const matchesOg = ogFilter === "all" || contract.ogTitle === ogFilter;
      const matchesCustomer = customerFilter === "all" || contract.customer === customerFilter;
      const matchesPartner = partnerFilter === "all" || contract.partner === partnerFilter;
      const matchesClearance = clearanceFilter === "all" || contract.clearance === clearanceFilter;

      return matchesSearch && matchesOg && matchesCustomer && matchesPartner && matchesClearance;
    });
  }, [contracts, searchTerm, ogFilter, customerFilter, partnerFilter, clearanceFilter]);

  const filtersActive = searchTerm !== ""
    || ogFilter !== "all"
    || customerFilter !== "all"
    || partnerFilter !== "all"
    || clearanceFilter !== "all";

  const handleResetFilters = (): void => {
    setSearchTerm("");
    setOgFilter("all");
    setCustomerFilter("all");
    setPartnerFilter("all");
    setClearanceFilter("all");
  };

  const sortedContracts = React.useMemo(() => {
    if (!sortColumnKey) return filteredContracts;

    return [...filteredContracts].sort((a, b) => {
      let aVal: string | number = "";
      let bVal: string | number = "";

      switch (sortColumnKey) {
        case "title":
          aVal = (a.Title || "").toLowerCase();
          bVal = (b.Title || "").toLowerCase();
          break;
        case "contractId":
          aVal = (a.contractId || "").toLowerCase();
          bVal = (b.contractId || "").toLowerCase();
          break;
        case "customer":
          aVal = (a.customer || "").toLowerCase();
          bVal = (b.customer || "").toLowerCase();
          break;
        case "location":
          aVal = (a.location || "").toLowerCase();
          bVal = (b.location || "").toLowerCase();
          break;
        case "ogTitle":
          aVal = (a.ogTitle || "").toLowerCase();
          bVal = (b.ogTitle || "").toLowerCase();
          break;
        case "clearance":
          aVal = (a.clearance || "").toLowerCase();
          bVal = (b.clearance || "").toLowerCase();
          break;
        case "startDate":
          aVal = a.startDate ? new Date(a.startDate).getTime() : 0;
          bVal = b.startDate ? new Date(b.startDate).getTime() : 0;
          break;
        case "endDate":
          aVal = a.endDate ? new Date(a.endDate).getTime() : 0;
          bVal = b.endDate ? new Date(b.endDate).getTime() : 0;
          break;
        case "contractPm":
          aVal = (a.contractPm?.Title || "").toLowerCase();
          bVal = (b.contractPm?.Title || "").toLowerCase();
          break;
        default:
          return 0;
      }

      if (aVal < bVal) return isSortedDescending ? 1 : -1;
      if (aVal > bVal) return isSortedDescending ? -1 : 1;
      return 0;
    });
  }, [filteredContracts, sortColumnKey, isSortedDescending]);

  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const newIsSortedDescending = sortColumnKey === column.key ? !isSortedDescending : false;
    setSortColumnKey(column.key);
    setIsSortedDescending(newIsSortedDescending);
  };

  const sortable = (key: string): Pick<IColumn, "isSorted" | "isSortedDescending" | "onColumnClick"> => ({
    isSorted: sortColumnKey === key,
    isSortedDescending,
    onColumnClick
  });

  const columns: IColumn[] = [
    {
      key: "title",
      name: "Contract Title",
      fieldName: "Title",
      minWidth: 180,
      maxWidth: 260,
      isResizable: true,
      ...sortable("title"),
      onRender: (item: IContractItem) => (
        <Stack>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
            {item.isFlagged && (
              <Icon
                iconName="Flag"
                className={styles.contractFlagIconSmall}
                title="Flagged contract"
                ariaLabel="Flagged contract"
              />
            )}
            <Link
              className={styles.listTitleLink}
              onClick={(ev) => {
                ev?.preventDefault();
                ev?.stopPropagation();
                onSelectContract(item);
              }}
            >
              {item.Title}
            </Link>
          </Stack>
          <Text variant="small">{item.contractId || ""}</Text>
        </Stack>
      )
    },
    {
      key: "customerContractCode",
      name: "Customer Contract Code",
      fieldName: "customerContractCode",
      minWidth: 130,
      maxWidth: 170,
      isResizable: true,
      onRender: (item: IContractItem) => <Text>{item.customerContractCode || ""}</Text>
    },
    {
      key: "ogTitle",
      name: "OG",
      fieldName: "ogTitle",
      minWidth: 90,
      maxWidth: 130,
      isResizable: true,
      ...sortable("ogTitle"),
      onRender: (item: IContractItem) => <Text>{item.ogTitle || ""}</Text>
    },
    {
      key: "clearance",
      name: "Clearance Level",
      fieldName: "clearance",
      minWidth: 105,
      maxWidth: 140,
      isResizable: true,
      ...sortable("clearance"),
      onRender: (item: IContractItem) => <Text>{item.clearance || ""}</Text>
    },
    {
      key: "customer",
      name: "Customer",
      fieldName: "customer",
      minWidth: 90,
      maxWidth: 120,
      isResizable: true,
      ...sortable("customer"),
      onRender: (item: IContractItem) => <Text>{item.customer || ""}</Text>
    },
    {
      key: "startDate",
      name: "Start",
      fieldName: "startDate",
      minWidth: 70,
      maxWidth: 90,
      isResizable: true,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      ...sortable("startDate"),
      onRender: (item: IContractItem) => <Text>{item.startDate ? formatDate(item.startDate) : "-"}</Text>
    },
    {
      key: "location",
      name: "Location",
      fieldName: "location",
      minWidth: 130,
      maxWidth: 190,
      isResizable: true,
      ...sortable("location"),
      onRender: (item: IContractItem) => <Text>{item.location || "-"}</Text>
    },
    {
      key: "endDate",
      name: "End",
      fieldName: "endDate",
      minWidth: 70,
      maxWidth: 90,
      isResizable: true,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      ...sortable("endDate"),
      onRender: (item: IContractItem) => <Text>{item.endDate ? formatDate(item.endDate) : "-"}</Text>
    },
    {
      key: "contractPm",
      name: "Contract PM",
      fieldName: "contractPm",
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      ...sortable("contractPm"),
      onRender: (item: IContractItem) => <Text>{renderPerson(item.contractPm)}</Text>
    },
    {
      key: "partner",
      name: "Partner",
      fieldName: "partner",
      minWidth: 90,
      maxWidth: 130,
      isResizable: true,
      onRender: (item: IContractItem) => <Text>{item.partner || ""}</Text>
    },
    {
      key: "capabilityCount",
      name: "Capabilities",
      fieldName: "capabilityCount",
      minWidth: 90,
      maxWidth: 110,
      isResizable: false,
      onRender: (item: IContractItem) => <Text>{ContractService.getCapabilityLookups(item).length}</Text>
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 15 }} styles={{ root: { marginTop: 24 }}}>
      <Stack tokens={{ childrenGap: 10 }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 12 }}>
        <SearchBox
          placeholder="Search contract, location, synonyms, customer, clearance, OG, LOB, PM, or partner..."
          value={searchTerm}
          onChange={(_, newValue) => setSearchTerm(newValue || "")}
          styles={{ root: { width: "100%", maxWidth: 700 } }}
        />

        {onNewContract && (
          <PrimaryButton
            text="New Contract"
            iconProps={{ iconName: "Add" }}
            onClick={onNewContract}
          />
        )}
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 10 }} wrap verticalAlign="center">
          <Dropdown
            placeholder="Filter by OG"
            options={ogOptions}
            selectedKey={ogFilter}
            onChange={(_, option) => setOgFilter((option?.key as string) ?? "all")}
            styles={{ dropdown: { width: 180 } }}
          />
          <Dropdown
            placeholder="Filter by Customer"
            options={customerOptions}
            selectedKey={customerFilter}
            onChange={(_, option) => setCustomerFilter((option?.key as string) ?? "all")}
            styles={{ dropdown: { width: 220 } }}
          />
          <Dropdown
            placeholder="Filter by Partner"
            options={partnerOptions}
            selectedKey={partnerFilter}
            onChange={(_, option) => setPartnerFilter((option?.key as string) ?? "all")}
            styles={{ dropdown: { width: 260 } }}
          />
          <Dropdown
            placeholder="Filter by Clearance Level"
            options={clearanceOptions}
            selectedKey={clearanceFilter}
            onChange={(_, option) => setClearanceFilter((option?.key as string) ?? "all")}
            styles={{ dropdown: { width: 190 } }}
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

      <Text variant="medium">
        Showing {filteredContracts.length} of {contracts.length} contracts
      </Text>

      <PaginatedDetailsList
        items={sortedContracts}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={1}
        isHeaderVisible={true}
        onItemInvoked={(item) => onSelectContract(item as IContractItem)}
        pageSizeOptions={[10, 25, 50]}
        defaultPageSizeOption={25}
        showFirstLastButtons={true}
      />
    </Stack>
  );
};
