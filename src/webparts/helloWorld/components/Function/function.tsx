import { IDropdownOption } from "@fluentui/react";
import * as pnp from "sp-pnp-js";
import { colors, OptionColor } from "../HelloWorld";

export interface Column {
    title: string
    type: string
    value: any
    option: OptionColor[]
    require: boolean
    unique: boolean
    readOnly: boolean
    initialValue?: any
    description: string
}

export const fetchUsers = async () => {
    try {
        const response = await pnp.sp.web.siteGroups.getById(92).users.get();
        const userOptions: IDropdownOption[] = response.map((item: any) => ({
            key: item.Id,
            text: item.Title
        }));
        return userOptions;
    } catch (error) {
        console.error("Error fetching users:", error);
        throw error;
    }
};

export const checkRequireValue = (item: any, check: boolean) => {
    if (check === true) {
        if (item.require === true && (item.value == null || item.value === "")) {
            return "Vui lòng nhập thông tin vào trường này!";
        }
    }
}

export const fetchTickets = async (name: string) => {
    try {
        const response = await pnp.sp.web.lists.getByTitle(name).items.get();
        return response;
    } catch (error) {
        console.error('Error fetching tickets:', error);
        throw error;
    }
};

export const fetchColumns = async (name: string) => {
    try {

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

        const response = await pnp.sp.web.lists.getByTitle(name).fields.filter("CanBeDeleted eq true").get();

        const data: Column[] = response.map((item: any) => ({
            title: item.Title,
            type: item.TypeDisplayName,
            require: item.Required,
            unique: item.EnforceUniqueValues,
            description: item.Description
        }));

        for (const item of data) {
            if (item.type === "Choice") {
                const response = await pnp.sp.web.lists.getByTitle(name).fields.getByInternalNameOrTitle(item.title).get();
                const options: OptionColor[] = response.Choices.map((a: any, index: number) => ({
                    key: a,
                    text: a,
                    color: colors[index % colors.length]
                }));
                item.option = options;
            }
        }

        data.unshift(title);

        return data

    } catch (error) {
        console.error("Error fetching user:", error);
        throw error;
    }
};