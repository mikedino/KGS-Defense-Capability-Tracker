import * as React from "react";
import { Stack, Text, IColumn, SelectionMode, SearchBox } from "@fluentui/react";
import PaginatedDetailsList from "../ui/PaginatedDetailsList";
import { formatDate } from "../utils";
import styles from "../AppTracker.module.scss";

import type { IContractItem, IPeoplePickerExtended } from "../data/props";

interface IContractsListProps {
  contracts: IContractItem[];
  onSelectContract: (contract: IContractItem) => void; // double-click -> open form later
}

const renderPerson = (p?: IPeoplePickerExtended): string => p?.Title ?? "";

const formatMoney = (value?: number): string => {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  return value.toLocaleString(undefined, { style: "currency", currency: "USD" });
};

export const ContractsList: React.FunctionComponent<IContractsListProps> = ({ contracts, onSelectContract }) => {

  // SEARCH
  const [searchTerm, setSearchTerm] = React.useState<string>("");

  // SORTING
  const [sortColumnKey, setSortColumnKey] = React.useState<string | null>("title");
  const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

  const filteredContracts = React.useMemo(() => {
    const search = searchTerm.trim().toLowerCase();
    if (!search) return contracts;

    return contracts.filter((contract) => {
      return (
        (contract.Title ?? "").toLowerCase().includes(search) ||
        (contract.contractTeam ?? "").toLowerCase().includes(search) ||
        (contract.cor?.Title ?? "").toLowerCase().includes(search) ||
        (contract.acor?.Title ?? "").toLowerCase().includes(search) ||
        (contract.ko?.Title ?? "").toLowerCase().includes(search) ||
        (contract.primaryPoc?.Title ?? "").toLowerCase().includes(search)
      );
    });
  }, [contracts, searchTerm]);

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
        case "start":
          aVal = a.start ? new Date(a.start).getTime() : 0;
          bVal = b.start ? new Date(b.start).getTime() : 0;
          break;
        case "end":
          aVal = a.end ? new Date(a.end).getTime() : 0;
          bVal = b.end ? new Date(b.end).getTime() : 0;
          break;
        case "contractValue":
          aVal = typeof a.contractValue === "number" ? a.contractValue : 0;
          bVal = typeof b.contractValue === "number" ? b.contractValue : 0;
          break;
        case "cor":
          aVal = (a.cor?.Title || "").toLowerCase();
          bVal = (b.cor?.Title || "").toLowerCase();
          break;
        case "ko":
          aVal = (a.ko?.Title || "").toLowerCase();
          bVal = (b.ko?.Title || "").toLowerCase();
          break;
        case "primaryPoc":
          aVal = (a.primaryPoc?.Title || "").toLowerCase();
          bVal = (b.primaryPoc?.Title || "").toLowerCase();
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

  const columns: IColumn[] = [
    {
      key: "title",
      name: "Contract",
      fieldName: "Title",
      minWidth: 260,
      maxWidth: 380,
      isResizable: true,
      isSorted: sortColumnKey === "title",
      isSortedDescending,
      onColumnClick,
      onRender: (item: IContractItem) => (
        <Stack>
          <Text variant="medium" style={{ fontWeight: 600 }}>
            {item.Title}
          </Text>
          <Text variant="small">{item.contractTeam || ""}</Text>
        </Stack>
      )
    },
    {
      key: "start",
      name: "Start",
      fieldName: "start",
      minWidth: 110,
      maxWidth: 130,
      isResizable: true,
      isSorted: sortColumnKey === "start",
      isSortedDescending,
      onColumnClick,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      onRender: (item: IContractItem) => <Text>{item.start ? formatDate(item.start) : "-"}</Text>
    },
    {
      key: "end",
      name: "End",
      fieldName: "end",
      minWidth: 110,
      maxWidth: 130,
      isResizable: true,
      isSorted: sortColumnKey === "end",
      isSortedDescending,
      onColumnClick,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      onRender: (item: IContractItem) => <Text>{item.end ? formatDate(item.end) : "-"}</Text>
    },
    {
      key: "contractValue",
      name: "Value",
      fieldName: "contractValue",
      minWidth: 120,
      maxWidth: 160,
      isResizable: true,
      isSorted: sortColumnKey === "contractValue",
      isSortedDescending,
      onColumnClick,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      onRender: (item: IContractItem) => <Text>{formatMoney(item.contractValue)}</Text>
    },
    {
      key: "cor",
      name: "COR",
      fieldName: "cor",
      minWidth: 140,
      maxWidth: 200,
      isResizable: true,
      isSorted: sortColumnKey === "cor",
      isSortedDescending,
      onColumnClick,
      onRender: (item: IContractItem) => <Text>{renderPerson(item.cor)}</Text>
    },
    {
      key: "ko",
      name: "Contracting Officer",
      fieldName: "ko",
      minWidth: 140,
      maxWidth: 200,
      isResizable: true,
      isSorted: sortColumnKey === "ko",
      isSortedDescending,
      onColumnClick,
      onRender: (item: IContractItem) => <Text>{renderPerson(item.ko)}</Text>
    },
    {
      key: "primaryPoc",
      name: "Contract PM",
      fieldName: "primaryPoc",
      minWidth: 160,
      maxWidth: 240,
      isResizable: true,
      isSorted: sortColumnKey === "primaryPoc",
      isSortedDescending,
      onColumnClick,
      onRender: (item: IContractItem) => <Text>{renderPerson(item.primaryPoc)}</Text>
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 24 }}}>
      <SearchBox
        placeholder="Search Contract, Lead, COR, KO, or Contract PM..."
        value={searchTerm}
        onChange={(_, newValue) => setSearchTerm(newValue || "")}
        styles={{ root: { width: 400 } }}
      />

      <PaginatedDetailsList
        items={sortedContracts}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={1}
        isHeaderVisible={true}
        onItemInvoked={(item) => onSelectContract(item as IContractItem)} // double-click opens form later
        pageSizeOptions={[5, 10, 25, 50]}
        defaultPageSizeOption={10}
        showFirstLastButtons={true}
      />
    </Stack>
  );
};
