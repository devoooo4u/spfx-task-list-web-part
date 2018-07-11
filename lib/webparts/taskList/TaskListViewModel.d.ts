import 'datatables.net-responsive';
import { ITaskListWebPartProps } from './TaskListWebPart';
import { Web } from 'sp-pnp-js';
export interface ITaskListBindingContext extends ITaskListWebPartProps {
    shouter: KnockoutSubscribable<{}>;
}
export interface IListViewHeader {
    Title: string;
}
export interface IListViewHeaders {
    value: IListViewHeader[];
}
export interface IDict {
    key: string;
    value: string;
    url: string;
}
export declare class ISiteUser {
    Id: string;
    UserTitle: string;
}
export default class TaskListViewModel {
    selectedList: KnockoutObservable<string>;
    selectedView: KnockoutObservable<string>;
    private BaseUrl;
    viewName: KnockoutObservable<string>;
    tblViewName: KnockoutObservable<string>;
    listViewHeaders: KnockoutObservableArray<IListViewHeader>;
    private taskListItems;
    meetingMinutestypeId: string;
    taskListClass: string;
    containerClass: string;
    rowClass: string;
    columnClass: string;
    titleClass: string;
    subTitleClass: string;
    descriptionClass: string;
    buttonClass: string;
    labelClass: string;
    constructor(bindings: ITaskListBindingContext);
    /**
     * DisplayListView
     */
    DisplayListView(web: Web, listId: string, viewId: string, userColl: Array<ISiteUser>): void;
    /**
     * LoadAllSiteUsers
     */
    LoadAllSiteUsers(): void;
}
