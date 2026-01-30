/* eslint-disable @typescript-eslint/no-explicit-any */
import { IListItem } from "../../../services/SharePoint/IListItem";

export interface IListViewPromotesState {
    items: IListItem[],
    filtereditems: IListItem[],
    loading: boolean;
    error: string;
    sortBy: string;
    filterBy: {
        field: string,
        isActive: boolean,
        displayname: string,
        options: string[]
    }[];
    activefilters: {
        field: string,
        options: string[]
    }[],
    sortOrder: 'asc' | 'desc';
    currentPage: number;
    itemsPerPage: number;
}