"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
var ko = require("knockout");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_pnp_js_1 = require("sp-pnp-js");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var strings = require("TaskListWebPartStrings");
var TaskListViewModel_1 = require("./TaskListViewModel");
var sp_http_1 = require("@microsoft/sp-http");
var _instance = 0;
var TaskListWebPart = (function (_super) {
    __extends(TaskListWebPart, _super);
    function TaskListWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._koListTitle = ko.observable('');
        _this._koListViewName = ko.observable('');
        _this.taskListsDropdownDisabled = true;
        _this.listViewDropdownDisabled = true;
        // ...
        /**
         * Shouter is used to communicate between web part and view model.
         */
        _this._shouter = new ko.subscribable();
        _this._self = _this;
        return _this;
    }
    TaskListWebPart.prototype.ApplyChanges = function () {
        location.reload();
    };
    TaskListWebPart.prototype.loadLists = function () {
        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=(Hidden eq false)and(BaseTemplate eq 107)", sp_http_1.SPHttpClient.configurations.v1) // sending the request to SharePoint REST API
            .then(function (response) {
            return response.json();
        }).then(function (response) {
            var asd = [];
            var bsd = response.value;
            bsd.forEach(function (e) {
                asd.push({ key: e.Id.toString(), text: e.Title.toString() });
            });
            // this.listDropDownOptions.push({ key: e.Id.toString(), text: e.Title.toString() });
            return asd;
        });
    };
    // ...
    TaskListWebPart.prototype.loadItems = function () {
        var wp = this;
        if (!this.properties.listTitle) {
            // resolve to empty options since no list has been selected
            return Promise.resolve();
        }
        else {
            var listId = this.properties.listTitle;
            return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + '/_api/web/lists(guid\'' + listId + '\')/views', sp_http_1.SPHttpClient.configurations.v1) // requesting views from SharePoint REST API
                .then(function (response) {
                return response.json();
            })
                .then(function (response) {
                var _asd = [];
                var _bsd = response.value;
                _bsd.forEach(function (e) {
                    _asd.push({ key: e.Id.toString(), text: e.Title.toString() });
                });
                // this.listDropDownOptions.push({ key: e.Id.toString(), text: e.Title.toString() });
                return _asd;
            });
        }
    };
    // ...
    TaskListWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        var _this = this;
        this.taskListsDropdownDisabled = !this.taskLists;
        this.listViewDropdownDisabled = !this.properties.listTitle || !this.listViews;
        if (this.taskLists) {
            return;
        }
        //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'options');
        this.loadLists()
            .then(function (listOptions) {
            _this.taskLists = listOptions;
            _this.taskListsDropdownDisabled = false;
            _this.context.propertyPane.refresh();
            return _this.loadItems();
        })
            .then(function (itemOptions) {
            _this.listViews = itemOptions;
            _this.listViewDropdownDisabled = !_this.properties.listTitle;
            _this.context.propertyPane.refresh();
            _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
            _this.render();
        });
    };
    // ...
    // ...
    TaskListWebPart.prototype.onPropertyPaneFieldChanged = function (propertyPath, oldValue, newValue) {
        var _this = this;
        if (propertyPath === 'listTitle' &&
            newValue) {
            // push new list value
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
            // get previously selected item
            var previousItem = this.properties.listViewName;
            // reset selected item
            this.properties.listViewName = undefined;
            // push new item value
            this.onPropertyPaneFieldChanged('listViewName', previousItem, this.properties.listViewName);
            // disable item selector until new items are loaded
            this.listViewDropdownDisabled = true;
            // refresh the item selector control by repainting the property pane
            this.context.propertyPane.refresh();
            // communicate loading items
            //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'items');
            this.loadItems()
                .then(function (itemOptions) {
                // store items
                _this.listViews = itemOptions;
                // enable item selector
                _this.listViewDropdownDisabled = false;
                // clear status indicator
                _this.context.statusRenderer.clearLoadingIndicator(_this.domElement);
                // re-render the web part as clearing the loading indicator removes the web part body
                _this.render();
                // refresh the item selector control by repainting the property pane
                _this.context.propertyPane.refresh();
            });
        }
        else {
            _super.prototype.onPropertyPaneFieldChanged.call(this, propertyPath, oldValue, newValue);
        }
    };
    /**
     * Initialize the web part.
     */
    TaskListWebPart.prototype.onInit = function () {
        var _this = this;
        console.log('Oninit');
        this._id = _instance++;
        var tagName = "ComponentElement-" + this._id;
        this._componentElement = this._createComponentElement(tagName);
        this._registerComponent(tagName);
        // When web part description is changed, notify view model to update.   
        this._koListTitle.subscribe(function (newValue) {
            _this._shouter.notifySubscribers(newValue, 'listTitle');
        });
        this._koListViewName.subscribe(function (newValue) {
            _this._shouter.notifySubscribers(newValue, 'listViewName');
        });
        var bindings = {
            listTitle: this.properties.listTitle,
            listViewName: this.properties.listViewName,
            applyButton: this.properties.applyButton,
            shouter: this._shouter,
            context: this.context
        };
        ko.applyBindings(bindings, this._componentElement);
        return _super.prototype.onInit.call(this).then(function (_) {
            sp_pnp_js_1.default.setup({
                spfxContext: _this.context
            });
        });
    };
    TaskListWebPart.prototype.render = function () {
        if (!this.renderedOnce) {
            this.domElement.appendChild(this._componentElement);
        }
        this._koListTitle(this.properties.listTitle);
        this._koListViewName(this.properties.listViewName);
    };
    TaskListWebPart.prototype._createComponentElement = function (tagName) {
        var componentElement = document.createElement('div');
        componentElement.setAttribute('data-bind', "component: { name: \"" + tagName + "\", params: $data }");
        return componentElement;
    };
    TaskListWebPart.prototype._registerComponent = function (tagName) {
        ko.components.register(tagName, {
            viewModel: TaskListViewModel_1.default,
            template: require('./TaskList.template.html'),
            synchronous: false
        });
    };
    Object.defineProperty(TaskListWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    TaskListWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                sp_webpart_base_1.PropertyPaneLabel('labelField', {
                                    text: 'Please select List and List view'
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('listTitle', {
                                    label: 'select the task list',
                                    options: this.taskLists,
                                    disabled: this.taskListsDropdownDisabled
                                }),
                                sp_webpart_base_1.PropertyPaneDropdown('listViewName', {
                                    label: 'select the task list view',
                                    options: this.listViews,
                                    disabled: this.listViewDropdownDisabled
                                }),
                                sp_webpart_base_1.PropertyPaneButton('applyButton', {
                                    text: 'Apply',
                                    disabled: false,
                                    buttonType: sp_webpart_base_1.PropertyPaneButtonType.Primary,
                                    onClick: this.ApplyChanges
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return TaskListWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = TaskListWebPart;

//# sourceMappingURL=TaskListWebPart.js.map
