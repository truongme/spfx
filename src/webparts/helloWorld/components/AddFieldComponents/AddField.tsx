import { DefaultButton, Dropdown, IDropdownOption, IToggleStyles, PrimaryButton, Stack, TextField, Toggle } from "@fluentui/react";
import React, { useEffect, useState } from "react";
import { useForm, Controller } from "react-hook-form";
import * as pnp from 'sp-pnp-js';

interface Props {
    name: string;
    nameColumn: string;
}

interface Field {
    name: string;
    description: string;
    type: string;
    defaultValue: any;
    required: boolean;
    unique: boolean;
}

const toggleStyles: Partial<IToggleStyles> = { root: { margin: '10px 0' } };

export const AddField: React.FC<Props> = (props) => {

    if (props.nameColumn.slice(-2) === "Id") {
        props.nameColumn = props.nameColumn.slice(0, -2)
    }

    const { control, handleSubmit, reset, formState: { errors } } = useForm<Field>();
    const [options, setOptions] = useState<IDropdownOption[]>([]);

    const columnTypes: string[] = [
        "Single line of text",
        "Multiple lines of text",
        "Choice",
        "Number",
        "Currency",
        "Date and Time",
        "Lookup",
        "Yes/No",
        "Person or Group",
        "Hyperlink",
        "Managed Metadata",
        "Location",
        "Image"
    ];

    const onSubmit = async (field: Field) => {
        try {
            if (props.nameColumn !== "") {
                await pnp.sp.web.lists.getByTitle(props.name).fields.getByTitle(props.nameColumn).update({
                    Description: field.description,
                    DefaultValue: field.defaultValue,
                    EnforceUniqueValues: field.unique,
                    Required: field.required,

                });
                alert("Cập nhật column thành công!")
            }
            else {
                await pnp.sp.web.lists.getByTitle(props.name).fields.addText(field.name, 225, {
                    Description: field.description,
                    DefaultValue: field.defaultValue,
                    EnforceUniqueValues: field.unique,
                    Required: field.required,
                });
                alert("Thêm mới column thành công!")
            }
        } catch (error) {
            console.error("Error adding column:", error);
        }
    };

    const handleDelete = async () => {
        try {
            const confirm = window.confirm("Bạn có muốn xóa ticket này không?");
            if (confirm) {
                await pnp.sp.web.lists.getByTitle(props.name).fields.getByTitle(props.nameColumn).delete();
                alert("Xoá column thành công!")
            }
        } catch (error) {
            console.error("Error adding column:", error);
        }
    };

    const loadOptions = () => {
        try {
            const typeOptions: IDropdownOption[] = columnTypes.map((item) => ({
                key: item,
                text: item
            }));
            setOptions(typeOptions);
        } catch (error) {
            console.error("Error loading options:", error);
        }
    };

    const fetchColumn = async () => {
        try {
            const response = await pnp.sp.web.lists.getByTitle(props.name).fields.getByInternalNameOrTitle(props.nameColumn).get();
            const data = {
                name: response.Title,
                description: response.Description,
                type: response.TypeDisplayName,
                defaultValue: response.DefaultValue,
                unique: response.EnforceUniqueValues ?? false,
                required: response.Required ?? false,
            };

            reset(data);
        } catch (error) {
            console.error("Error fetching column:", error);
        }
    };

    useEffect(() => {
        loadOptions();
        if (props.nameColumn !== "") fetchColumn();
    }, [props.nameColumn]);

    return (
        <div>
            <form onSubmit={handleSubmit(onSubmit)}>
                <Controller
                    name="name"
                    control={control}
                    defaultValue=""
                    disabled={props.nameColumn !== ""}
                    rules={{ required: 'Name is required' }}
                    render={({ field }) => (
                        <TextField
                            {...field}
                            label='Name'
                            required
                            errorMessage={errors.name ? errors.name.message : undefined}
                        />
                    )}
                />
                <Controller
                    name="description"
                    control={control}
                    defaultValue=""
                    render={({ field }) => (
                        <TextField
                            {...field}
                            label="Description"
                            multiline
                        />
                    )}
                />
                <Controller
                    name="type"
                    control={control}
                    rules={{ required: 'Type is required' }}
                    render={({ field }) => (
                        <Dropdown
                            {...field}
                            placeholder="Select column type"
                            label="Type"
                            disabled={props.nameColumn !== ""}
                            options={options}
                            required
                            onChange={(e, option) => field.onChange(option?.key as string)}
                            selectedKey={field.value}
                            errorMessage={errors.type ? errors.type.message : undefined}
                        />
                    )}
                />
                <Controller
                    name="defaultValue"
                    control={control}
                    defaultValue=""
                    render={({ field }) => (
                        <TextField
                            {...field}
                            label='Default value'
                            placeholder="Enter a default value"
                        />
                    )}
                />
                <Controller
                    name="required"
                    control={control}
                    render={({ field }) => (
                        <Toggle
                            {...field}
                            styles={toggleStyles}
                            label="Require that this column contains information"
                            checked={field.value ?? false}
                            onChange={(e, checked) => field.onChange(checked)}
                        />
                    )}
                />
                <Controller
                    name="unique"
                    control={control}
                    render={({ field }) => (
                        <Toggle
                            {...field}
                            styles={toggleStyles}
                            label="Enforce unique values"
                            checked={field.value ?? false}
                            onChange={(e, checked) => field.onChange(checked)}
                        />
                    )}
                />
                <Stack horizontal style={{ marginTop: '10px' }}>
                    <PrimaryButton text="Save" allowDisabledFocus type="submit" style={{ marginRight: '8px' }} />
                    {props.nameColumn !== "" &&
                        <DefaultButton text="Delete" allowDisabledFocus onClick={() => handleDelete()} />
                    }
                </Stack>
            </form>
        </div>
    );
};
