import {
    DatePicker, DayOfWeek, defaultDatePickerStrings, Dropdown
    , Persona, PersonaSize, PrimaryButton, Spinner, SpinnerSize, Stack, TextField
} from "@fluentui/react";
import React, { useEffect, useState } from "react";
import * as pnp from 'sp-pnp-js';
import { OptionColor } from "../HelloWorld";
import { checkRequireValue, Column, fetchColumns, fetchTickets, fetchUsers } from "../Function/function";
import { Ticket } from "../ListComponents/ListComponents";

export interface Name {
    name: string
}

export const NewTicket: React.FC<Name> = (props) => {

    const [tickets, setTickets] = useState<Ticket[]>([]);
    const [columns, setColumns] = useState<Column[]>([]);
    const [options, setOptions] = useState<OptionColor[]>([]);
    const [spinner, setSpinner] = useState<boolean>(true)
    const [checkSubmit, setCheckSubmit] = useState<boolean>(false)
    const [error, setError] = useState<string>("")

    const handleTextChange = (title: string, newValue: string | undefined) => {
        setError("")
        setColumns(columns.map(col =>
            col.title === title ? { ...col, value: newValue || "" } : col
        ));
    };

    const handleDateChange = (title: string, date: Date | null | undefined) => {
        setError("")
        setColumns(columns.map(col =>
            col.title === title ? { ...col, value: date ?? null } : col
        ));
    };

    const handleDropdownChange = (title: string, option: OptionColor | undefined) => {
        setError("")
        setColumns(columns.map(col =>
            col.title === title ? { ...col, value: option ? option.key : undefined } : col
        ));
    };

    const handleSubmit = async () => {
        try {

            setCheckSubmit(true)

            const submit = true;

            if (columns.some(item => checkRequireValue(item, submit))) return

            for (const column of columns) {
                if (column.unique && column.value) {
                    const titles = tickets.map((a: any) => a[column.title]);
                    const isUnique = !titles.includes(column.value);
                    if (!isUnique) {
                        setError(`Giá trị ${column.title} phải là duy nhất. Vui lòng nhập giá trị khác.`)
                        return;
                    }
                }
            }

            const columnSubmit = columns.reduce((arr, column) => {
                if (column.value !== undefined && column.value !== null) {
                    if (column.type === "Person or Group") {
                        column.title = column.title + "Id";
                    }
                    arr[column.title] = column.value;
                }
                return arr;
            }, {} as Record<string, any>);

            await pnp.sp.web.lists.getByTitle(props.name).items.add({ ...columnSubmit });
            alert("Thêm ticket mới thành công");

        } catch (error) {
            console.error("Error adding item:", error);
        }
    };

    useEffect(() => {
        const loadData = async () => {
            try {

                const [users, tickets, columns] = await Promise.all([
                    fetchUsers(),
                    fetchTickets(props.name),
                    fetchColumns(props.name)
                ]);

                console.log(tickets)

                setOptions(users);
                setTickets(tickets);
                setColumns(columns);


                setSpinner(false);

            } catch (error) {
                console.error("Error in data loading:", error);
            }
        };

        loadData();
    }, [props.name]);


    return (
        <div>
            {columns.map((item: Column) => {
                switch (item.type) {
                    case "Single line of text":
                        return (
                            <TextField
                                key={item.title}
                                label={item.title}
                                required={item.require}
                                value={item.value || ""}
                                placeholder={item.description}
                                onChange={(e, newValue) => handleTextChange(item.title, newValue)}
                                errorMessage={checkRequireValue(item, checkSubmit)}
                            />
                        );
                    case "Multiple lines of text":
                        return (
                            <TextField
                                key={item.title}
                                label={item.title}
                                multiline
                                placeholder={item.description}
                                required={item.require}
                                value={item.value || ""}
                                onChange={(e, newValue) => handleTextChange(item.title, newValue)}
                                errorMessage={checkRequireValue(item, checkSubmit)}

                            />
                        );
                    case "Date and Time":
                        return (
                            <DatePicker
                                key={item.title}
                                label={item.title}
                                firstDayOfWeek={DayOfWeek.Monday}
                                showWeekNumbers={true}
                                firstWeekOfYear={1}
                                showMonthPickerAsOverlay={true}
                                placeholder={item.description}
                                ariaLabel="Select a date"
                                strings={defaultDatePickerStrings}
                                value={item.value || null}
                                onSelectDate={(date) => handleDateChange(item.title, date)}
                            />
                        );
                    case "Person or Group":
                        return (
                            <Dropdown
                                key={item.title}
                                placeholder={item.description}
                                label={item.title}
                                options={options}
                                required={item.require}
                                selectedKey={item.value || undefined}
                                onChange={(e, option) => handleDropdownChange(item.title, option)}
                                errorMessage={checkRequireValue(item, checkSubmit)}
                                onRenderOption={(option) => {
                                    if (!option) return null;
                                    return (
                                        <div>
                                            <Persona
                                                text={option.text}
                                                imageUrl={`https://nitecovietnam.sharepoint.com/sites/English-Philips/truong-dev/_layouts/15/userphoto.aspx?size=L&accountname=${option.key}`}
                                                size={PersonaSize.size24}
                                                showSecondaryText={false}
                                            />
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
                            />
                        );
                    case "Choice":
                        return (
                            <Dropdown
                                key={item.title}
                                placeholder={item.description}
                                label={item.title}
                                options={item.option}
                                required={item.require}
                                selectedKey={item.value || undefined}
                                onChange={(e, option) => handleDropdownChange(item.title, option)}
                                errorMessage={checkRequireValue(item, checkSubmit)}
                                onRenderOption={(option: any) => {
                                    return (
                                        <div style={{ backgroundColor: option.color, padding: '2px 8px', borderRadius: '15px', margin: '5px 0px' }}>
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
                            />
                        );
                    default:
                        return <>{item.title}</>;
                }
            })}
            {checkSubmit && error !== "" &&
                <p style={{ backgroundColor: "#FABBC3", padding: "5px 8px", borderRadius: '2px' }}>{error}</p>
            }
            {spinner ? (
                <Stack style={{ margin: '10px' }}>
                    <Spinner size={SpinnerSize.medium} />
                </Stack>
            ) : (
                <Stack style={{ marginTop: '10px' }}>
                    <PrimaryButton text="+ Add item" onClick={handleSubmit} />
                </Stack>
            )}
        </div>
    );
};
