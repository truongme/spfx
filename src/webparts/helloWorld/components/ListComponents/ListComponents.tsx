import React, { useState, useEffect } from 'react';
import { Announced } from '@fluentui/react/lib/Announced';
import { TextField, ITextFieldStyles } from '@fluentui/react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from '@fluentui/react/lib/DetailsList';
import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyles, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { Text } from '@fluentui/react/lib/Text';
import * as pnp from 'sp-pnp-js';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { useBoolean } from '@fluentui/react-hooks';
import { Stack, IStackTokens, Icon, IContextualMenuProps, ContextualMenu, Checkbox, Persona, PersonaSize } from '@fluentui/react';
import { DefaultButton, IconButton, PrimaryButton } from '@fluentui/react/lib/Button';
import "./index.scss"
import { Column } from "../HelloWorld"
import { NewTicket } from '../CreateComponents/NewTicket';
import { EditForm } from '../EditTicketCom/EditForm';
import { AddField } from '../AddFieldComponents/AddField';
import * as XLSX from 'xlsx';


export interface User {
  Id: number;
  Title: string;
  UserPrincipalName: string;
}

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

const iconClass = mergeStyles({
  fontSize: 30,
});

const classNames = mergeStyleSets({
  deepSkyBlue: [{ color: 'white' }, iconClass],
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

const stackTokens: IStackTokens = { childrenGap: 10 };

export interface List {
  name: string;
  column: Column[];
}

export interface Ticket {
  Id: number;
  Title: string;
}

const ListComponents: React.FC<List> = (props) => {
  const [tickets, setTickets] = useState<Ticket[]>([]);
  const name = props.name;
  const columns: string[] = props.column.map(column => column.title);
  columns.push("Title")
  columns.push("Id")

  const [ticketsCopy, setTicketsCopy] = useState<Ticket[]>([]);

  const [menuTarget, setMenuTarget] = useState<HTMLElement | null>(null);
  const [isMenuVisible, setMenuVisible] = useState(false);

  const [isOpenNew, { setTrue: openPanelNew, setFalse: dismissPanelNew }] = useBoolean(false);
  const [isOpenEdit, { setTrue: openPanelEdit, setFalse: dismissPanelEdit }] = useBoolean(false);
  const [isOpenFilter, { setTrue: openPanelFilter, setFalse: dismissPanelFilter }] = useBoolean(false);
  const [isOpenAddField, { setTrue: openPanelAddField, setFalse: dismissPanelAddField }] = useBoolean(false);

  const [selectionDetails, setSelectionDetails] = useState<string>('');
  const [selectedTicket, setSelectedTicket] = useState<Ticket | null>(null);

  const [isSortedDescending, setIsSortedDescending] = useState<boolean>(true);

  const [arrStatus, setArrStatus] = useState<string[]>([]);
  const [arrStatusApply, setArrStatusApply] = useState<string[]>([]);
  const [selectColumn, setSelectColumn] = useState<string>("");

  const [users, setUsers] = useState<User[]>([]);
  const [showNew, setShowNew] = useState<boolean>(false);
  const [showAddField, setShowAddField] = useState<boolean>(false);

  const getColumnValues = (columnKey: keyof Ticket) => {
    const values = ticketsCopy.map(ticket => ticket[columnKey]);
    return Array.from(new Set(values));
  };

  const menuProps: IContextualMenuProps = {
    shouldFocusOnMount: true,
    shouldFocusOnContainer: true,
    items: [
      {
        key: 'sortaz',
        text: 'A to Z',
        onClick: () => {
          setIsSortedDescending(false)
          handleSort(selectColumn as keyof Ticket, false);
          setMenuVisible(false);
        }
      },
      {
        key: 'sortza',
        text: 'Z to A',
        onClick: () => {
          setIsSortedDescending(true)
          handleSort(selectColumn as keyof Ticket, true);
          setMenuVisible(false);
        }
      },
      {
        key: 'filter',
        text: 'Filter by',
        onClick: () => {
          openPanelFilter();
          setMenuVisible(false);
        }
      },
      {
        key: 'group',
        text: `Group by ${selectColumn}`,
        onClick: () => {
          setMenuVisible(false);
          handleGroup(selectColumn as keyof Ticket)
        }
      },
      {
        key: 'field',
        text: 'Column settings',
        onClick: () => {
          openPanelAddField();
          setMenuVisible(false);
          setShowAddField(true)
        }
      },
    ],
  };

  const handleGroup = (columnName: keyof Ticket) => {
    const groupedItems = ticketsCopy.reduce((groups: { [key: string]: Ticket[] }, ticket) => {
      const key = String(ticket[columnName]);
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(ticket);
      return groups;
    }, {});

    const groupedArray = Object.keys(groupedItems).map(key => ({
      key,
      items: groupedItems[key]
    }));

    setTickets(groupedArray.flatMap(group => group.items));
  }

  const handleSort = (columnName: keyof Ticket, isSortedDescending: boolean) => {
    const sortedItems = [...tickets].sort((a, b) => {
      if (a[columnName] < b[columnName]) {
        return isSortedDescending ? 1 : -1;
      }
      if (a[columnName] > b[columnName]) {
        return isSortedDescending ? -1 : 1;
      }
      return 0;
    });
    setTickets(sortedItems);
  }

  const handleStatusChange = (item: string, isChecked: boolean | undefined) => {
    setArrStatus(prev =>
      isChecked ? [...prev, item] : prev.filter(s => s !== item)
    );
  };

  const filterByStatus = () => {
    setArrStatusApply(arrStatus);
    dismissPanelFilter();
    if (arrStatus.length > 0) {
      const filteredTickets = ticketsCopy.filter((ticket: any) => {
        return arrStatus.includes(String(ticket[selectColumn]));
      });
      setTickets(filteredTickets);
    } else {
      setTickets(ticketsCopy);
    }
  };

  const handleDeleteStatus = (item: string) => {
    const newArrStatus = arrStatus.filter(status => status !== item);
    setArrStatus(newArrStatus);
    setArrStatusApply(newArrStatus);
    const filteredTickets = newArrStatus.length > 0 ? ticketsCopy.filter((ticket: any) => newArrStatus.includes(String(ticket[selectColumn]))) : ticketsCopy;
    setTickets(filteredTickets);
  };

  const _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const newColumns: IColumn[] = _columns.slice();
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    setSelectColumn(currColumn.fieldName as keyof Ticket);

    if (event.currentTarget !== menuTarget) {
      setMenuVisible(false);
      setMenuTarget(event.currentTarget);
      setMenuVisible(true);
    } else {
      setMenuVisible(!isMenuVisible);
    }
  };

  const closePanelFilter = () => {
    setArrStatus(arrStatusApply)
    dismissPanelFilter();
  };

  const closePanelNew = () => {
    dismissPanelNew();
    fetchTickets();
  };

  const closePanelFecth = () => {
    dismissPanelEdit();
    fetchTickets();
  };

  const closePanelAddField = () => {
    dismissPanelAddField();
    fetchTickets();
  };

  const _selection = new Selection({
    onSelectionChanged: () => {
      const selectedItems = _selection.getSelection();
      setSelectionDetails(_getSelectionDetails());
      if (selectedItems.length) {
        setSelectedTicket(selectedItems[0] as Ticket);
      } else {
        setSelectedTicket(null);
      }
    },
  });

  const _columns: IColumn[] = [
    {
      key: '0',
      name: 'Actions',
      minWidth: 100,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: Ticket) => (
        <div>
          <IconButton
            iconProps={{ iconName: 'Edit' }}
            onClick={() => {
              setSelectedTicket(item);
              openPanelEdit();
            }}
            title="Edit"
            ariaLabel="Edit"
            style={{ marginRight: '8px' }}
            className='edit'
          />
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            onClick={() => deleteTicket(item.Id)}
            title="Delete"
            ariaLabel="Delete"
            className='delete'
          />
        </div>
      ),
      // onRenderHeader: (props, defaultRender) => (
      //   <div className='flex'>
      //     {defaultRender && defaultRender(props)}
      //   </div>
      // ),
    },
    {
      key: '1', name: 'Category', fieldName: 'Title', minWidth: 100, maxWidth: 150, isResizable: true, onColumnClick: _onColumnClick, isSorted: selectColumn == 'Title',
      isSortedDescending: isSortedDescending, sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A', isPadded: true,
      onRenderHeader: (props, defaultRender) => (
        <div className='flex'>
          {defaultRender && defaultRender(props)}
        </div>
      ),
    },
  ];

  const numberArr = columns.length;

  for (let i = 0; i < numberArr - 2; i++) {
    const dynamicColumn: IColumn = {
      key: `${i + 2}`,
      name: columns[i].slice(-2) === "Id" ? columns[i].slice(0, -2) : columns[i],
      fieldName: columns[i],
      minWidth: 125,
      maxWidth: 175,
      isResizable: true,
      onRender: (item: any) => {
        const fieldName = columns[i];
        if (fieldName.slice(-2) === 'Id') {
          if (fieldName.slice(0, -2) !== 'Lookup') {
            const idValue = item[fieldName];
            const user = users.find(user => user.Id === idValue);
            if (user) {
              return <Persona
                text={user.Title}
                imageUrl={`https://nitecovietnam.sharepoint.com/sites/English-Philips/truong-dev/_layouts/15/userphoto.aspx?size=L&accountname=${user.UserPrincipalName}`}
                size={PersonaSize.size24}
                showSecondaryText={false}
              />
            }
          }
          // else {
          //   return <div>{lookup[i]}</div>;
          // }
        } else {
          return <div>{item[fieldName]}</div>;
        }
      },
      onColumnClick: _onColumnClick,
      isSorted: selectColumn === columns[i], isSortedDescending: isSortedDescending, sortAscendingAriaLabel: 'Sorted A to Z',
      sortDescendingAriaLabel: 'Sorted Z to A', isPadded: true,
      onRenderHeader: (props, defaultRender) => (
        <div className='flex'>
          {defaultRender &&
            defaultRender(props)
          }
        </div>
      ),
    };

    _columns.push(dynamicColumn);
  }

  const _getSelectionDetails = (): string => {
    const selectionCount = _selection.getSelectedCount();
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (_selection.getSelection()[0] as Ticket).Title;
      default:
        return `${selectionCount} items selected`;
    }
  };

  const _onFilter = (event: React.ChangeEvent<HTMLInputElement>) => {
    const input = event.target.value.toLowerCase();
    const filter = ticketsCopy.filter(ticket =>
      ticket.Title.toLowerCase().includes(input)
    );
    setTickets(filter);
  };

  const _onItemInvoked = (item: Ticket): void => {
    setSelectedTicket(item);
  };

  function arraysEqual(arr1: string[], arr2: string[]) {
    if (arr1.length !== arr2.length) {
      return false;
    }
    for (let i = 0; i < arr1.length; i++) {
      if (arr1[i] !== arr2[i]) {
        return false;
      }
    }
    return true;
  }

  const handleExport = () => {
    const items = tickets;
    const worksheet = XLSX.utils.json_to_sheet(items);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, 'Export.xlsx');
  }

  const handleNewTicket = () => {
    setShowNew(true)
    openPanelNew()
  }

  const handleAddField = () => {
    setSelectColumn("")
    setShowAddField(true)
    openPanelAddField()
  }

  const deleteTicket = async (id: number) => {
    try {
      const confirm = window.confirm("Bạn có muốn xóa ticket này không?");
      if (confirm) {
        await pnp.sp.web.lists.getByTitle(name).items.getById(id).delete();
        setTickets(tickets.filter(ticket => ticket.Id !== id));
      }
    } catch (error) {
      console.error('Error deleting ticket:', error);
    }
  };


  const fetchUsers = async () => {
    try {
      const response = await pnp.sp.web.siteGroups.getById(92).users.get();
      setUsers(response);
    } catch (error) {
      console.error("Error fetching tickets:", error);
    }
  };
  const fetchTickets = async () => {
    try {
      const response = await pnp.sp.web.lists.getByTitle(name).items.select(...columns).get();
      setTickets(response);
      setTicketsCopy(response)
    } catch (error) {
      console.error('Error fetching tickets:', error);
    }
  };

  useEffect(() => {
    fetchTickets()
    fetchUsers()
  }, []);

  return (
    <div>
      <div className='info'>
        <div className='header-title'>
          <div className='icon-ctn'>
            <Icon aria-label="" iconName="AddOnlineMeeting" className={classNames.deepSkyBlue} />
          </div>
          <h3>IT Support Management</h3>
        </div>
        <div style={{ marginTop: '8px' }}>
          <PrimaryButton text="+ Add new " onClick={() => handleNewTicket()} allowDisabledFocus style={{ marginRight: '8px' }} />
          <DefaultButton text="+ Add column" onClick={() => handleAddField()} allowDisabledFocus style={{ marginRight: '8px' }} />
          <DefaultButton text="Export" onClick={() => handleExport()} allowDisabledFocus />
        </div>
        <div className={exampleChildClass}>{selectionDetails}</div>
        <Text>
          Note: While focusing a row, pressing enter or double-clicking will execute onItemInvoked.
        </Text>
        <Announced message={selectionDetails} />
        <TextField
          className={exampleChildClass}
          label="Filter by Category:"
          onChange={_onFilter}
          styles={textFieldStyles}
        />
        <Announced message={`Number of items after filter applied: ${tickets.length}.`} />
      </div>
      <div className='filter'>
        {arrStatusApply.length > 0 &&
          <h4>{selectColumn}</h4>
        }
        {arrStatusApply.map(item => (
          <div className={`status-ctn ${item}`} >
            <p className={`text-status`}>{item}</p>
            <Icon iconName="Cancel" onClick={() => handleDeleteStatus(String(item))} style={{ cursor: "pointer" }} />
          </div>
        ))}
      </div>
      <MarqueeSelection selection={_selection} className='table'>
        <DetailsList
          items={tickets}
          columns={_columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selection={_selection}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="select row"
          onItemInvoked={_onItemInvoked}
          className='table'
        />
      </MarqueeSelection>
      {/* <EditForm /> */}
      {showNew && (
        <Panel
          headerText={`New Ticket`}
          isOpen={isOpenNew}
          onDismiss={closePanelNew}
          closeButtonAriaLabel="Close"
          type={PanelType.medium}
        >
          <NewTicket
            name={name}
          />
        </Panel>
      )}
      {selectedTicket && (
        <Panel
          headerText={`Edit Ticket ${selectedTicket.Title}`}
          isOpen={isOpenEdit}
          onDismiss={closePanelFecth}
          closeButtonAriaLabel="Close"
          type={PanelType.medium}
        >
          <EditForm
            idTicket={selectedTicket.Id}
            name={name}
          />
        </Panel>
      )}
      {showAddField && (

        <Panel
          headerText={selectColumn === "" ? `Create a column` : `Edit Column`}
          isOpen={isOpenAddField}
          onDismiss={closePanelAddField}
          closeButtonAriaLabel="Close"
          type={PanelType.smallFixedFar}
        >
          <AddField
            name={name}
            nameColumn={selectColumn}
          />
        </Panel>
      )}
      <Panel
        headerText={`Filter Ticket by Status`}
        isOpen={isOpenFilter}
        onDismiss={closePanelFilter}
        closeButtonAriaLabel="Close"
      >
        <Stack tokens={stackTokens} className='mb-10 mt-10'>
          {getColumnValues(selectColumn as keyof Ticket).map((item) => (
            <Checkbox
              key={String(item)}
              label={String(item)}
              checked={arrStatus.includes(String(item))}
              onChange={(e, isChecked) => handleStatusChange(String(item), isChecked)}
            />
          ))
          }
        </Stack>
        <Stack horizontal tokens={stackTokens}>
          <PrimaryButton text="Apply" disabled={arraysEqual(arrStatusApply, arrStatus)} allowDisabledFocus onClick={filterByStatus} />
          <DefaultButton text="Clear all" allowDisabledFocus onClick={() => setArrStatus([])} />
        </Stack>
      </Panel>
      {isMenuVisible && (
        <ContextualMenu
          items={menuProps.items}
          target={menuTarget}
          onDismiss={menuProps.onDismiss}
          shouldFocusOnMount={true}
        />
      )}
    </div>
  );
};

export default ListComponents;
