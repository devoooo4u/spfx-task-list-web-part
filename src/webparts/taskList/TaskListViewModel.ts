import * as $ from 'jquery';
import 'datatables.net-responsive';
import * as ko from 'knockout';
import styles from './TaskList.module.scss';
import { ITaskListWebPartProps } from './TaskListWebPart';
import pnp, { Web, List, ListEnsureResult, ItemAddResult } from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

require('./TaskList.scss');
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

export class ISiteUser {
  Id: string;
  UserTitle: string;
}

export default class TaskListViewModel {
  public selectedList: KnockoutObservable<string> = ko.observable('');
  public selectedView: KnockoutObservable<string> = ko.observable('');
  private BaseUrl: KnockoutObservable<string> = ko.observable('');
  public viewName: KnockoutObservable<string> = ko.observable('');
  public tblViewName: KnockoutObservable<string> = ko.observable('');
  public listViewHeaders: KnockoutObservableArray<IListViewHeader> = ko.observableArray([]);
  private taskListItems: KnockoutObservableArray<any> = ko.observableArray([]);
  public meetingMinutestypeId: string = "";

  public taskListClass: string = styles.taskList;
  public containerClass: string = styles.container;
  public rowClass: string = styles.row;
  public columnClass: string = styles.column;
  public titleClass: string = styles.title;
  public subTitleClass: string = styles.subTitle;
  public descriptionClass: string = styles.description;
  public buttonClass: string = styles.button;
  public labelClass: string = styles.label;

  constructor(bindings: ITaskListBindingContext) {
    SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Lora:400,400i,700,700i");
    SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700");
    SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
    SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.13/css/jquery.dataTables.min.css');
    SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.1.1/css/responsive.dataTables.min.css');

    this.selectedList(bindings.listTitle);
    this.selectedView(bindings.listViewName);

    var url = bindings.context.pageContext.web.absoluteUrl;
    this.BaseUrl(url);

    bindings.shouter.subscribe((value: string) => {
      this.selectedList(value);
    }, this, 'listTitle');
    bindings.shouter.subscribe((value: string) => {
      this.selectedView(value);
    }, this, 'listViewName');

    // Call of the wild
    if (this.selectedList() && this.selectedList() !== '' && this.selectedView() && this.selectedView() !== '') {
      this.LoadAllSiteUsers();
    }
    else {
      $('#spnMessage').css('display', 'block');
    }
  }

  /**
   * DisplayListView
   */
  public DisplayListView(web: Web, listId: string, viewId: string, userColl: Array<ISiteUser>) {
    web.contentTypes.select("Name, Id/StringValue").expand("Id").get().then(contentTypes => {
      contentTypes.forEach(contentType => {
        if(contentType["Name"] == "Meeting Minutes"){
          this.meetingMinutestypeId = contentType.Id.StringValue;
        }
      });
    });
    web.lists.getById(listId).views.getById(viewId).fields.get().then(val => {
      val.Items.forEach(element => {
        if(element == "MM_x0020_Due_x0020_Date"){
          element = "Due Date";
        }
        let vh: IListViewHeader = {
          Title: element,
        };
        this.listViewHeaders.push(vh);
      });
    }).then(a1=> {
      let webUrl = ko.unwrap(this.BaseUrl());
      let _viewXML: string = "";
      web.lists.getById(listId).views.getById(viewId).get().then(xx => {
        _viewXML = xx.ListViewXml;
        _viewXML = _viewXML.replace('<FieldRef Name="LinkTitle" />', '<FieldRef Name="LinkTitle" /><FieldRef Name="ContentTypeId" />');
        if(webUrl.indexOf("Executive")<0 && webUrl.indexOf("Director")<0 && webUrl.indexOf("NCC")<0 && webUrl.indexOf("FAP")<0)
        {
          _viewXML = _viewXML.replace('<FieldRef Name="DueDate" />', '<FieldRef Name="DueDate" /><FieldRef Name="MM_x0020_Due_x0020_Date" />');
        }
      })
      .then(a2 => {
          web.lists.getById(listId).getItemsByCAMLQuery({ ViewXml: _viewXML }).then(items => {
          items.forEach(element => {
            let itmValue: Array<IDict> = new Array<IDict>();
              this.listViewHeaders().forEach(x => {
              if (x.Title == "AssignedTo") {
                //this is to get user info based on user id
                var result = userColl.filter(IUser => IUser.Id == element['AssignedToId']);
                if (result && result.length !== 0) {
                  itmValue.push({ key: x.Title, value: result[0]['UserTitle'], url: null });
                }
                else {
                  itmValue.push({ key: x.Title, value: '', url: null });
                }
              }
              else {
                if (x.Title == "Predecessors") {
                  itmValue.push({ key: x.Title, value: element['PredecessorsId'], url: null });
                }
                else {
                  if (x.Title == "WorkflowLink") {
                    if(element['WorkflowLink']) {
                      itmValue.push({ key: x.Title, value: element['WorkflowLink']['Description'], url: element['WorkflowLink']['Url'] });
                    }else {
                      itmValue.push({ key: x.Title, value: "", url: "" });
                    }
                  }
                  else {
                    if (x.Title == "LinkTitle") {
                      let actionUrl = "";
                      if(element["ContentTypeId"].indexOf(this.meetingMinutestypeId)>-1){
                        actionUrl = this.BaseUrl() + '/Lists/Tasks/EditForm.aspx?ID=' + element['Id'] + '&Source=' + this.BaseUrl() + '/SitePages/Dashboard.aspx';
                      }
                      else{
                        actionUrl = this.BaseUrl() + '/_layouts/15/WrkTaskIP.aspx?List=' + this.selectedList() + '&ID=' + element['Id'] + '&Source=' + this.BaseUrl() + '/SitePages/Dashboard.aspx' + '&ContentTypeId=' + element['ContentTypeId'];
                      }
                      itmValue.push({ key: x.Title, value: element['Title'], url: actionUrl });
                    }
                    else {
                      if (x.Title == "DueDate") {
                        var dt;
                        if(element["ContentTypeId"].indexOf(this.meetingMinutestypeId)>-1) {
                          dt = new Date(element['MM_x0020_Due_x0020_Date']);
                        } else {
                          dt = new Date(element['DueDate']);
                        }
                        itmValue.push({ key: x.Title, value: dt.toDateString(), url: null });
                      }
                      else if(x.Title == "Due Date"){
                        dt = new Date(element['MM_x0020_Due_x0020_Date']);
                        itmValue.push({ key: "Due Date", value: dt.toDateString(), url: null });
                      }else {
                        itmValue.push({ key: x.Title, value: element[x.Title], url: null });
                      }
                    }
                  }
                }
              }
            });
            this.taskListItems.push(itmValue);
          });
          let tableId: string = '#tbl' + this.tblViewName();
          $(tableId).DataTable({
            responsive: true,
            "lengthMenu": [[5, 10, 25], [5, 10, 25]],
            initComplete: () => {
            }
          });
          $('#span' + this.tblViewName()).css('display', 'none');
        });
      });
    });
  }

  /**
   * LoadAllSiteUsers
   */
  public LoadAllSiteUsers() {
    var web = new Web(this.BaseUrl());

    web.siteUsers.get().then(u => {
      let siteUsersCollection: Array<ISiteUser> = new Array<ISiteUser>();
      u.forEach(el => {
        let iSiteUser: ISiteUser = new ISiteUser();
        iSiteUser.Id = el.Id;
        iSiteUser.UserTitle = el.Title;
        siteUsersCollection.push(iSiteUser);
      });
      return siteUsersCollection;
    }).then(_userColl => {
      let ListID = this.selectedList();
      let ViewID = this.selectedView();
      // get view Name
      web.lists.getById(ListID).views.getById(ViewID).get().then(v => {
        this.viewName(v.Title);
        this.tblViewName(v.Title.replace(/ /g, ''));
      });
      //
      this.DisplayListView(web, ListID, ViewID, _userColl);
    });
  }
}
