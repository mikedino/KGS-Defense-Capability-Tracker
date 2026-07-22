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

const renderPeople = (people?: { results: IPeoplePickerExtended[] }): string =>
  people?.results?.map((p) => p.Title).filter(Boolean).join(", ") ?? "";

export const ContractsList: React.FunctionComponent<IContractsListProps> = ({ contracts, onSelectContract }) => {
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [sortColumnKey, setSortColumnKey] = React.useState<string | null>("title");
  const [isSortedDescending, setIsSortedDescending] = React.useState<boolean>(false);

  const filteredContracts = React.useMemo(() => {
    const search = searchTerm.trim().toLowerCase();
    if (!search) return contracts;

    return contracts.filter((contract) => {
      return (
        (contract.Title ?? "").toLowerCase().includes(search) ||
        (contract.contractId ?? "").toLowerCase().includes(search) ||
        (contract.invoice ?? "").toLowerCase().includes(search) ||
        (contract.customerContractCode ?? "").toLowerCase().includes(search) ||
        (contract.customer ?? "").toLowerCase().includes(search) ||
        (contract.contractPm?.Title ?? "").toLowerCase().includes(search) ||
        (contract.primaryPoc?.Title ?? "").toLowerCase().includes(search) ||
        renderPeople(contract.stakeholders).toLowerCase().includes(search) ||
        (contract.partner ?? "").toLowerCase().includes(search)
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
        case "contractId":
          aVal = (a.contractId || "").toLowerCase();
          bVal = (b.contractId || "").toLowerCase();
          break;
        case "invoice":
          aVal = (a.invoice || "").toLowerCase();
          bVal = (b.invoice || "").toLowerCase();
          break;
        case "customer":
          aVal = (a.customer || "").toLowerCase();
          bVal = (b.customer || "").toLowerCase();
          break;
        case "popStart":
          aVal = a.popStart ? new Date(a.popStart).getTime() : 0;
          bVal = b.popStart ? new Date(b.popStart).getTime() : 0;
          break;
        case "popEnd":
          aVal = a.popEnd ? new Date(a.popEnd).getTime() : 0;
          bVal = b.popEnd ? new Date(b.popEnd).getTime() : 0;
          break;
        case "contractPm":
          aVal = (a.contractPm?.Title || "").toLowerCase();
          bVal = (b.contractPm?.Title || "").toLowerCase();
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
      minWidth: 240,
      maxWidth: 360,
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
      key: "invoice",
      name: "Task Order/Invoice ID",
      fieldName: "invoice",
      minWidth: 150,
      maxWidth: 220,
      isResizable: true,
      ...sortable("invoice"),
      onRender: (item: IContractItem) => <Text>{item.invoice || ""}</Text>
    },
    {
      key: "customerContractCode",
      name: "Customer Contract Code",
      fieldName: "customerContractCode",
      minWidth: 160,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: IContractItem) => <Text>{item.customerContractCode || ""}</Text>
    },
    {
      key: "customer",
      name: "Customer",
      fieldName: "customer",
      minWidth: 110,
      maxWidth: 160,
      isResizable: true,
      ...sortable("customer"),
      onRender: (item: IContractItem) => <Text>{item.customer || ""}</Text>
    },
    {
      key: "popStart",
      name: "PoP Start",
      fieldName: "popStart",
      minWidth: 110,
      maxWidth: 130,
      isResizable: true,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      ...sortable("popStart"),
      onRender: (item: IContractItem) => <Text>{item.popStart ? formatDate(item.popStart) : "-"}</Text>
    },
    {
      key: "popEnd",
      name: "PoP End",
      fieldName: "popEnd",
      minWidth: 110,
      maxWidth: 130,
      isResizable: true,
      headerClassName: styles.centeredHeader,
      className: styles.centeredColumn,
      ...sortable("popEnd"),
      onRender: (item: IContractItem) => <Text>{item.popEnd ? formatDate(item.popEnd) : "-"}</Text>
    },
    {
      key: "contractPm",
      name: "Contract PM",
      fieldName: "contractPm",
      minWidth: 160,
      maxWidth: 220,
      isResizable: true,
      ...sortable("contractPm"),
      onRender: (item: IContractItem) => <Text>{renderPerson(item.contractPm)}</Text>
    },
    {
      key: "primaryPoc",
      name: "Capability POC",
      fieldName: "primaryPoc",
      minWidth: 160,
      maxWidth: 220,
      isResizable: true,
      ...sortable("primaryPoc"),
      onRender: (item: IContractItem) => <Text>{renderPerson(item.primaryPoc)}</Text>
    },
    {
      key: "partner",
      name: "Partner",
      fieldName: "partner",
      minWidth: 120,
      maxWidth: 180,
      isResizable: true,
      onRender: (item: IContractItem) => <Text>{item.partner || ""}</Text>
    }
  ];

  return (
    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 24 }}}>
      <SearchBox
        placeholder="Search contract, invoice, customer, POC, stakeholder, or partner..."
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
