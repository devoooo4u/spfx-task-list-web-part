"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var $ = require("jquery");
require("datatables.net-responsive");
var ko = require("knockout");
var TaskList_module_scss_1 = require("./TaskList.module.scss");
var sp_pnp_js_1 = require("sp-pnp-js");
var sp_loader_1 = require("@microsoft/sp-loader");
require('./TaskList.scss');
var ISiteUser = (function () {
    function ISiteUser() {
    }
    return ISiteUser;
}());
exports.ISiteUser = ISiteUser;
var TaskListViewModel = (function () {
    function TaskListViewModel(bindings) {
        var _this = this;
        this.selectedList = ko.observable('');
        this.selectedView = ko.observable('');
        this.BaseUrl = ko.observable('');
        this.viewName = ko.observable('');
        this.tblViewName = ko.observable('');
        this.listViewHeaders = ko.observableArray([]);
        this.taskListItems = ko.observableArray([]);
        this.meetingMinutestypeId = "";
        this.taskListClass = TaskList_module_scss_1.default.taskList;
        this.containerClass = TaskList_module_scss_1.default.container;
        this.rowClass = TaskList_module_scss_1.default.row;
        this.columnClass = TaskList_module_scss_1.default.column;
        this.titleClass = TaskList_module_scss_1.default.title;
        this.subTitleClass = TaskList_module_scss_1.default.subTitle;
        this.descriptionClass = TaskList_module_scss_1.default.description;
        this.buttonClass = TaskList_module_scss_1.default.button;
        this.labelClass = TaskList_module_scss_1.default.label;
        sp_loader_1.SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Lora:400,400i,700,700i");
        sp_loader_1.SPComponentLoader.loadCss("https://fonts.googleapis.com/css?family=Open+Sans:300,400,600,700");
        sp_loader_1.SPComponentLoader.loadCss("https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css");
        sp_loader_1.SPComponentLoader.loadCss('https://cdn.datatables.net/1.10.13/css/jquery.dataTables.min.css');
        sp_loader_1.SPComponentLoader.loadCss('https://cdn.datatables.net/responsive/2.1.1/css/responsive.dataTables.min.css');
        this.selectedList(bindings.listTitle);
        this.selectedView(bindings.listViewName);
        var url = bindings.context.pageContext.web.absoluteUrl;
        this.BaseUrl(url);
        bindings.shouter.subscribe(function (value) {
            _this.selectedList(value);
        }, this, 'listTitle');
        bindings.shouter.subscribe(function (value) {
            _this.selectedView(value);
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
    TaskListViewModel.prototype.DisplayListView = function (web, listId, viewId, userColl) {
        var _this = this;
        web.contentTypes.select("Name, Id/StringValue").expand("Id").get().then(function (contentTypes) {
            contentTypes.forEach(function (contentType) {
                if (contentType["Name"] == "Meeting Minutes") {
                    _this.meetingMinutestypeId = contentType.Id.StringValue;
                }
            });
        });
        web.lists.getById(listId).views.getById(viewId).fields.get().then(function (val) {
            val.Items.forEach(function (element) {
                if (element == "MM_x0020_Due_x0020_Date") {
                    element = "Due Date";
                }
                var vh = {
                    Title: element,
                };
                _this.listViewHeaders.push(vh);
            });
        }).then(function (a1) {
            var webUrl = ko.unwrap(_this.BaseUrl());
            var _viewXML = "";
            web.lists.getById(listId).views.getById(viewId).get().then(function (xx) {
                _viewXML = xx.ListViewXml;
                _viewXML = _viewXML.replace('<FieldRef Name="LinkTitle" />', '<FieldRef Name="LinkTitle" /><FieldRef Name="ContentTypeId" />');
                if (webUrl.indexOf("Executive") < 0 && webUrl.indexOf("Director") < 0 && webUrl.indexOf("NCC") < 0 && webUrl.indexOf("FAP") < 0) {
                    _viewXML = _viewXML.replace('<FieldRef Name="DueDate" />', '<FieldRef Name="DueDate" /><FieldRef Name="MM_x0020_Due_x0020_Date" />');
                }
            })
                .then(function (a2) {
                web.lists.getById(listId).getItemsByCAMLQuery({ ViewXml: _viewXML }).then(function (items) {
                    items.forEach(function (element) {
                        var itmValue = new Array();
                        _this.listViewHeaders().forEach(function (x) {
                            if (x.Title == "AssignedTo") {
                                //this is to get user info based on user id
                                var result = userColl.filter(function (IUser) { return IUser.Id == element['AssignedToId']; });
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
                                        if (element['WorkflowLink']) {
                                            itmValue.push({ key: x.Title, value: element['WorkflowLink']['Description'], url: element['WorkflowLink']['Url'] });
                                        }
                                        else {
                                            itmValue.push({ key: x.Title, value: "", url: "" });
                                        }
                                    }
                                    else {
                                        if (x.Title == "LinkTitle") {
                                            var actionUrl = "";
                                            if (element["ContentTypeId"].indexOf(_this.meetingMinutestypeId) > -1) {
                                                actionUrl = _this.BaseUrl() + '/Lists/Tasks/EditForm.aspx?ID=' + element['Id'] + '&Source=' + _this.BaseUrl() + '/SitePages/Dashboard.aspx';
                                            }
                                            else {
                                                actionUrl = _this.BaseUrl() + '/_layouts/15/WrkTaskIP.aspx?List=' + _this.selectedList() + '&ID=' + element['Id'] + '&Source=' + _this.BaseUrl() + '/SitePages/Dashboard.aspx' + '&ContentTypeId=' + element['ContentTypeId'];
                                            }
                                            itmValue.push({ key: x.Title, value: element['Title'], url: actionUrl });
                                        }
                                        else {
                                            if (x.Title == "DueDate") {
                                                var dt;
                                                if (element["ContentTypeId"].indexOf(_this.meetingMinutestypeId) > -1) {
                                                    dt = new Date(element['MM_x0020_Due_x0020_Date']);
                                                }
                                                else {
                                                    dt = new Date(element['DueDate']);
                                                }
                                                itmValue.push({ key: x.Title, value: dt.toDateString(), url: null });
                                            }
                                            else if (x.Title == "Due Date") {
                                                dt = new Date(element['MM_x0020_Due_x0020_Date']);
                                                itmValue.push({ key: "Due Date", value: dt.toDateString(), url: null });
                                            }
                                            else {
                                                itmValue.push({ key: x.Title, value: element[x.Title], url: null });
                                            }
                                        }
                                    }
                                }
                            }
                        });
                        _this.taskListItems.push(itmValue);
                    });
                    var tableId = '#tbl' + _this.tblViewName();
                    $(tableId).DataTable({
                        responsive: true,
                        "lengthMenu": [[5, 10, 25], [5, 10, 25]],
                        initComplete: function () {
                        }
                    });
                    $('#span' + _this.tblViewName()).css('display', 'none');
                });
            });
        });
    };
    /**
     * LoadAllSiteUsers
     */
    TaskListViewModel.prototype.LoadAllSiteUsers = function () {
        var _this = this;
        var web = new sp_pnp_js_1.Web(this.BaseUrl());
        web.siteUsers.get().then(function (u) {
            var siteUsersCollection = new Array();
            u.forEach(function (el) {
                var iSiteUser = new ISiteUser();
                iSiteUser.Id = el.Id;
                iSiteUser.UserTitle = el.Title;
                siteUsersCollection.push(iSiteUser);
            });
            return siteUsersCollection;
        }).then(function (_userColl) {
            var ListID = _this.selectedList();
            var ViewID = _this.selectedView();
            // get view Name
            web.lists.getById(ListID).views.getById(ViewID).get().then(function (v) {
                _this.viewName(v.Title);
                _this.tblViewName(v.Title.replace(/ /g, ''));
            });
            //
            _this.DisplayListView(web, ListID, ViewID, _userColl);
        });
    };
    return TaskListViewModel;
}());
exports.default = TaskListViewModel;

//# sourceMappingURL=TaskListViewModel.js.map
