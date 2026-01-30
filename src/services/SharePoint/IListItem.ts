/* eslint-disable @typescript-eslint/no-explicit-any */

export interface IListItem {
    Id: number;
    [index: string]: any;
}

export interface IListItemCollection {
    value: IListItem[];
}