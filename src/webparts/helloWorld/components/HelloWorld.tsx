import React, { useEffect, useState } from "react";
import "./HelloWorld.module.scss";
import { IHelloWorldProps } from "./IHelloWorldProps";
import * as pnp from 'sp-pnp-js';
import { Checkbox, Dropdown, IDropdownOption, PrimaryButton } from "@fluentui/react";
import ListComponents from "./ListComponents/ListComponents";

export interface Column {
  title: string;
}

export interface OptionColor extends IDropdownOption {
  color?: string;
}

export const colors: string[] = ["#D4E7F6", "#CAF0CC", "#FFEBC0", "#FABBC3", "#C3F8F9", "#C3CAF9"]

const HelloWorld: React.FC<IHelloWorldProps> = () => {

  const [name, setName] = useState<string>();
  const [showColumn, setShowColumn] = useState<boolean>(false);
  const [showList, setShowList] = useState<boolean>(false);
  const [columns, setColumns] = useState<Column[]>([]);
  const [columnsSubmit, setColumnsSubmit] = useState<Column[]>([]);
  const [options, setOptions] = useState<OptionColor[]>([]);

  const handleClickSubmit = () => {
    if (!columnsSubmit || columnsSubmit.length === 0) {
      alert("Hãy chọn ít nhất 1 cột để hiển thị list")
      return
    }
    setShowList(true);
  };

  const fetchList = async () => {
    try {
      const response = await pnp.sp.web.lists.filter('BaseTemplate eq 100').get();
      const options: OptionColor[] = response.map((a: any, index: number) => ({
        key: a.Title,
        text: a.Title,
        color: colors[index % colors.length]
      }));
      setOptions(options);
    } catch (error) {
      console.error("Error fetching lists:", error);
    }
  };

  const handleClick = async (title: string) => {
    try {
      setColumns([])
      setShowList(false);
      setColumnsSubmit([]);
      if (name !== title) {
        setName(title);
        const response = await pnp.sp.web.lists.getByTitle(title).fields.filter("CanBeDeleted eq true").get();
        const data = response.map((item: any) => ({
          title: item.Title,
          type: item.TypeDisplayName
        }));
        data.forEach((item: any) => {
          if (item.type === "Person or Group" || item.type === "Lookup") {
            item.title += 'Id';
          }
        });
        setColumns(data);
        setShowColumn(true);
      } else {
        setName('');
        setShowColumn(false);
      }
    } catch (error) {
      console.error("Error fetching columns:", error);
    }
  };

  const handleSelected = (column: Column) => {
    setShowList(false);
    const columnIndex = columnsSubmit.findIndex(col => col.title === column.title);
    if (columnIndex === -1) {
      setColumnsSubmit([...columnsSubmit, column]);
    } else {
      const updatedColumns = columnsSubmit.filter(col => col.title !== column.title);
      setColumnsSubmit(updatedColumns);
    }
  };

  useEffect(() => {
    fetchList();
  }, []);

  return (
    <div>
      <h2>Danh sách các list</h2>
      <Dropdown
        placeholder="Chọn 1 list để hiển thị"
        options={options}
        onChange={(e, option) => {
          if (option) {
            handleClick(option.key as string);
          }
        }}
        onRenderOption={(option) => {
          if (!option) return null;
          const colorOption = option as OptionColor;
          return (
            <div style={{ backgroundColor: colorOption.color, padding: '2px 8px', borderRadius: '15px', margin: '5px 0px' }}>
              {option.text}
            </div>
          );
        }}
        onRenderTitle={(option: any) => {
          if (!option) return null;
          return (
            <span style={{ backgroundColor: option[0].color, padding: '2px 8px', borderRadius: '15px', margin: '5px 0px' }}>
              {option[0].text}
            </span>
          );
        }}
        className="mt-10"
      />
      {showColumn && name && (
        <div>
          <h3>Danh sách các cột có trong list</h3>
          <div className="flex list-ctn" style={{ padding: '7px', backgroundColor: '#f0f0f0', border: '2px solid #498205', display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '10px' }}>
            {columns.map((column) => (
              <div key={column.title} style={{ paddingRight: '7px' }}>
                <Checkbox
                  label={column.title.slice(-2) === "Id" ? column.title.slice(0, -2) : column.title}
                  onChange={() => handleSelected(column)}
                />
              </div>
            ))}
          </div>
          <div style={{ margin: '15px 0px 8px 0px' }}>
            <PrimaryButton onClick={handleClickSubmit}>
              Hiển thị
            </PrimaryButton>
          </div>
        </div>
      )}
      {showList && name && (
        <ListComponents name={name} column={columnsSubmit} />
      )}
    </div>
  );
};

export default HelloWorld;
