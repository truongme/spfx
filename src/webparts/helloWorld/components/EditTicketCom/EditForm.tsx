import { Dropdown, IDropdownOption, Persona, PersonaSize, PrimaryButton, Spinner, SpinnerSize, Stack, TextField } from "@fluentui/react";
import React, { useEffect, useState } from "react";
import * as pnp from "sp-pnp-js";
import './index.scss'
import { colors } from "../HelloWorld";
import { checkRequireValue, Column, fetchTickets, fetchUsers } from "../Function/function";
import { Ticket } from "../ListComponents/ListComponents";
// import { tickets } from "../ListComponents/ListComponents";

interface Props {
    name: string
    idTicket: number
}

export const EditForm: React.FC<Props> = (props) => {

    const [ticket, setTicket] = useState<Column[]>([]);
    const [options, setOptions] = useState<IDropdownOption[]>([]);
    const [spinner, setSpinner] = useState<boolean>(true)
    const [check, setCheck] = useState<boolean>(false)
    const [error, setError] = useState<string>("")
    const [tickets, setTickets] = useState<Ticket[]>([]);


    const handleTextChange = (title: string, newValue: string | undefined) => {
        setError("")
        setTicket(ticket.map(col =>
            col.title === title ? { ...col, value: newValue || "" } : col
        ));
    };

    // const handleDateChange = (title: string, date: Date | null | undefined) => {
    //     setTicket(ticket.map(col =>
    //         col.title === title ? { ...col, value: date ?? null } : col
    //     ));
    // };

    const handleDropdownChange = (title: string, option: IDropdownOption | undefined) => {
        setError("")
        setTicket(ticket.map(col =>
            col.title === title ? { ...col, value: option ? option.key : undefined } : col
        ));
    };

    const handleSubmit = async () => {
        try {

            setCheck(true)

            const submit = true;

            if (ticket.some(item => checkRequireValue(item, submit))) return;

            let count = true;

            for (const item of ticket) {
                if (item.value !== item.initialValue) {
                    count = false
                    if (item.unique && item.value) {
                        const titles = tickets.map((a: any) => a[item.title]);
                        const isUnique = !titles.includes(item.value);
                        if (!isUnique) {
                            setError(`Giá trị ${item.title} phải là duy nhất. Vui lòng nhập giá trị khác.`)
                            return;
                        }
                    }
                }
            }

            if (count === true) {
                alert(`Các giá trị của các trường chưa có sự thay đổi.`);
                return;
            }

            // chuyển arr sang obj
            const columnSubmit = ticket.reduce((arr, column) => {
                if (column.value !== undefined && column.value !== null) {
                    arr[column.title] = column.value;
                }
                return arr;
            }, {} as Record<string, any>);

            await pnp.sp.web.lists.getByTitle(props.name).items.getById(props.idTicket).update({ ...columnSubmit });

            alert("Sửa ticket thành công");
            fetchTicket()
            fetchData()

        } catch (error) {
            console.error("Error adding item:", error);
        }
    };

    const fetchTicket = async () => {
        try {
            // Khởi tạo cột đầu tiên
            const title: Column = {
                title: "Title",
                type: "Single line of text",
                value: "",
                option: [],
                require: true,
                unique: true,
                readOnly: true,
                description: ""
            };

            // lấy danh sách các cột trong list
            const response = await pnp.sp.web.lists.getByTitle(props.name).fields.filter("CanBeDeleted eq true").get();
            const data: Column[] = response.map((item: any) => ({
                title: item.Title,
                type: item.TypeDisplayName,
                require: item.Required,
                unique: item.EnforceUniqueValues,
                readOnly: item.ReadOnlyField
            }));

            // Xử lý với các trường là person
            for (const item of data) {
                if (item.type === "Person or Group") {
                    item.title = item.title + "Id"
                }
                if (item.type === "Choice") {
                    const response = await pnp.sp.web.lists.getByTitle(props.name).fields.getByInternalNameOrTitle(item.title).get();
                    const options: IDropdownOption[] = response.Choices.map((a: any, index: number) => ({
                        key: a,
                        text: a,
                        color: colors[index % colors.length]
                    }));
                    item.option = options;
                }
            }
            data.unshift(title);

            // Lấy thông tin của ticket
            const ticket = await pnp.sp.web.lists.getByTitle(props.name).items.getById(props.idTicket).get();

            // map dữ liệu của ticket với các cột để hiển thị
            const updatedData = data.map(column => {
                if (column.title in ticket) {
                    return {
                        ...column,
                        value: ticket[column.title],
                        initialValue: ticket[column.title]
                    };
                }
                return column;
            });

            setTicket(updatedData)
            setSpinner(false)

        } catch (error) {
            console.error("Error fetching IdProps:", error);
        }
    };

    const fetchData = async () => {
        try {
            const response = await fetchTickets(props.name);
            setTickets(response)
        } catch (error) {
            console.error("Error fetching Data:", error);
        }
    };

    const loadUsers = async () => {
        try {
            const userOptions = await fetchUsers();
            setOptions(userOptions);
        } catch (error) {
            console.error("Error fetching users in Component1:", error);
        }
    };


    useEffect(() => {
        loadUsers();
        fetchData();
        fetchTicket();
    }, [])

    return (
        <div>
            {ticket.map((item: any) => {
                switch (item.type) {
                    case "Single line of text":
                        return (
                            <TextField
                                key={item.title}
                                label={item.title}
                                disabled={item.readOnly}
                                required={item.require} value={item.value || ""}
                                onChange={(e, newValue) => handleTextChange(item.title, newValue)}
                                errorMessage={checkRequireValue(item, check)}
                                placeholder={item.description}
                            />
                        );
                    case "Multiple lines of text":
                        return (
                            <TextField
                                key={item.title}
                                label={item.title}
                                disabled={item.readOnly}
                                multiline
                                required={item.require} value={item.value || ""}
                                onChange={(e, newValue) => handleTextChange(item.title, newValue)}
                                errorMessage={checkRequireValue(item, check)}
                                placeholder={item.description}
                            />
                        );
                    // case "Date and Time":
                    //     return (
                    //         <DatePicker
                    //             key={item.title}
                    //             label={item.title}
                    //             firstDayOfWeek={DayOfWeek.Monday}
                    //             showWeekNumbers={true}
                    //             firstWeekOfYear={1}
                    //             showMonthPickerAsOverlay={true}
                    //             placeholder="Select a date..."
                    //             ariaLabel="Select a date"
                    //             strings={defaultDatePickerStrings}
                    //             value={item.value || null}
                    //         // onSelectDate={(date) => handleDateChange(item.title, date)}
                    //         />
                    //     );
                    case "Person or Group":
                        return (
                            <Dropdown
                                key={item.title}
                                placeholder={item.description}
                                label={item.title.slice(0, -2)}
                                disabled={item.readOnly}
                                options={options}
                                required={item.require} selectedKey={item.value || undefined}
                                onChange={(e, option) => handleDropdownChange(item.title, option)}
                                errorMessage={checkRequireValue(item, check)}
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
                            />
                        );
                    case "Choice":
                        return (
                            <Dropdown
                                key={item.title}
                                placeholder={item.description}

                                label={item.title}
                                disabled={item.readOnly}
                                options={item.option}
                                required={item.require} selectedKey={item.value || undefined}
                                onChange={(e, option) => handleDropdownChange(item.title, option)}
                                errorMessage={checkRequireValue(item, check)}
                                onRenderOption={(option: any) => {
                                    if (!option) return null;
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

            {check && error !== "" &&
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
