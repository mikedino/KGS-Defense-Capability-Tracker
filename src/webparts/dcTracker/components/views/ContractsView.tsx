import * as React from "react";
import { IColumn, SearchBox, SelectionMode, Stack, Text } from "@fluentui/react";
import PaginatedDetailsList from "../ui/PaginatedDetailsList";
import { formatDate } from "../common/utils";
import styles from "../Dct.module.scss";

import type { IContractItem, IPeoplePickerExtended } from "../common/props";

interface IContractsListProps {
  contracts: IContractItem[];
  onSelectContract: (contract: IContractItem) => void;
}

const renderPerson = (p?: IPeoplePickerExtended): string => p?.Title ?? "";

export const ContractsList: React.FunctionComponent<IContractsListProps> = ({ contracts, onSelectContract }) => {
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [sortColumnKey, setSortColumnKey] = React.useState<string | null>("title");
  const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

  const uniqueContracts = React.useMemo(() => {
    const contractMap = new Map<string, IContractItem & { capabilityCount?: number }>();

    for (const contract of contracts) {
      const key = contract.contractId || contract.customerContractCode || contract.Title || contract.Id.toString();
      const existing = contractMap.get(key);

      if (existing) {
        existing.capabilityCount = (existing.capabilityCount ?? 1) + 1;
      } else {
        contractMap.set(key, { ...contract, capabilityCount: 1 });
      }
    }

    return Array.from(contractMap.values());
  }, [contracts]);

  const filteredContracts = React.useMemo(() => {
    const search = searchTerm.trim().toLowerCase();
    if (!search) return uniqueContracts;

    return uniqueContracts.filter((contract) => {
      return (
        (contract.Title ?? "").toLowerCase().includes(search) ||
        (contract.contractId ?? "").toLowerCase().includes(search) ||
        (contract.customerContractCode ?? "").toLowerCase().includes(search) ||
        (contract.customer ?? "").toLowerCase().includes(search) ||
        (contract.ogTitle ?? "").toLowerCase().includes(search) ||
        (contract.lobTitle ?? "").toLowerCase().includes(search) ||
        (contract.contractPm?.Title ?? "").toLowerCase().includes(search) ||
        (contract.partner ?? "").toLowerCase().includes(search)
      );
    });
  }, [uniqueContracts, searchTerm]);

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
        case "ogTitle":
          aVal = (a.ogTitle || "").toLowerCase();
          bVal = (b.ogTitle || "").toLowerCase();
          break;
        case "lobTitle":
          aVal = (a.lobTitle || "").toLowerCase();
          bVal = (b.lobTitle || "").toLowerCase();
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
          <Text variant="medium" style={{ fontWeight: 600 }}>{item.Title}</Text>
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
      key: "lobTitle",
      name: "LOB",
      fieldName: "lobTitle",
      minWidth: 90,
      maxWidth: 130,
      isResizable: true,
      ...sortable("lobTitle"),
      onRender: (item: IContractItem) => <Text>{item.lobTitle || ""}</Text>
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
      onRender: (item: IContractItem & { capabilityCount?: number }) => <Text>{item.capabilityCount ?? 1}</Text>
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 24 }}}>
      <SearchBox
        placeholder="Search contract, customer, OG, LOB, PM, or partner..."
        value={searchTerm}
        onChange={(_, newValue) => setSearchTerm(newValue || "")}
        styles={{ root: { width: 460 } }}
      />

      <PaginatedDetailsList
        items={sortedContracts}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={1}
        isHeaderVisible={true}
        onItemInvoked={(item) => onSelectContract(item as IContractItem)}
        pageSizeOptions={[5, 10, 25, 50]}
        defaultPageSizeOption={10}
        showFirstLastButtons={true}
      />
    </Stack>
  );
};
