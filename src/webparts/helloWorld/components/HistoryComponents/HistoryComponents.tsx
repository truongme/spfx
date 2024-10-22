import React, { useEffect, useState, version } from "react";
import * as pnp from "sp-pnp-js";
import { Stack, IStackTokens, DetailsList, IColumn, DetailsListLayoutMode } from "@fluentui/react";

const stackTokens: IStackTokens = { childrenGap: 20 };

export interface Item {
  idItem: number;
}

interface version {
  status: string;
  editor: string;
  modified: string;
}

export const HistoryComponents: React.FC<Item> = (props) => {

  const [tickets, setTickets] = useState<version[]>([]);

  const fetchItemHistory = async () => {
    try {
      const versions = await pnp.sp.web.lists.getByTitle("IT Support Management").items.getById(props.idItem).versions.get();
      const statusChanges = versions.map((version: any) => ({
        status: version.Status,
        editor: version.Editor?.LookupValue,
        modified: new Date(version.Modified).toLocaleString('vi-VN'),
      }))
      setTickets(statusChanges)
      console.log(versions)
    } catch (error) {
      console.error("Error fetching item history:", error);
    }
  };

  useEffect(() => {
    fetchItemHistory();
  }, []);

  const _columns: IColumn[] = [
    { key: '1', name: 'Status', fieldName: 'status', minWidth: 100, maxWidth: 100, isResizable: true },
    { key: '2', name: 'User', fieldName: 'editor', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: '3', name: 'Time', fieldName: 'modified', minWidth: 100, maxWidth: 200, isResizable: true },
  ];

  return (
    <Stack tokens={stackTokens}>
      <DetailsList
        items={tickets}
        columns={_columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionPreservedOnEmptyClick={true}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="select row"
      />
    </Stack>
  );
};
