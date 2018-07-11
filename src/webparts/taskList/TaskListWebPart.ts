import * as ko from 'knockout';
import { Version } from '@microsoft/sp-core-library';
import pnp from 'sp-pnp-js';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneLabel,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneButton,
  IPropertyPaneButtonProps,
  PropertyPaneButtonType
} from '@microsoft/sp-webpart-base';

import * as strings from 'TaskListWebPartStrings';
import TaskListViewModel, { ITaskListBindingContext } from './TaskListViewModel';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

let _instance: number = 0;

export interface ITaskListWebPartProps {
  listTitle: string;
  listViewName: string;
  applyButton: any;
  context: any;
}

export default class TaskListWebPart extends BaseClientSideWebPart<ITaskListWebPartProps> {
  private _id: number;
  private _componentElement: HTMLElement;
  private _koListTitle: KnockoutObservable<string> = ko.observable('');
  private _koListViewName: KnockoutObservable<string> = ko.observable('');

  private taskLists: IPropertyPaneDropdownOption[];
  private taskListsDropdownDisabled: boolean = true;
  private listViews: IPropertyPaneDropdownOption[];
  private listViewDropdownDisabled: boolean = true;

  private ApplyChanges():void{     
    location.reload();
  }

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=(Hidden eq false)and(BaseTemplate eq 107)`, SPHttpClient.configurations.v1) // sending the request to SharePoint REST API
    .then((response: SPHttpClientResponse) => { // httpClient.get method returns a response object where json method creates a Promise of getting result
      return response.json();
    }).then((response: any) => { // response is an ISPLists object  
      var asd = [];
      var bsd = response.value;
      bsd.forEach(e => {
        asd.push({ key: e.Id.toString(), text: e.Title.toString() });
      });
     // this.listDropDownOptions.push({ key: e.Id.toString(), text: e.Title.toString() });
      return asd; 
    });
 }

   // ...
   private loadItems(): Promise<IPropertyPaneDropdownOption[]> {
    const wp: TaskListWebPart = this;
    if (!this.properties.listTitle) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }      
   else {
    var listId = this.properties.listTitle;
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + '/_api/web/lists(guid\'' + listId + '\')/views', SPHttpClient.configurations.v1) // requesting views from SharePoint REST API
      .then((response: SPHttpClientResponse) => { // httpClient.get method returns a response object where json method creates a Promise of getting result
        return response.json();
      })
      .then((response: any) => { // response is an ISPViews object        
        var _asd = [];
        var _bsd = response.value;
        _bsd.forEach(e => {
          _asd.push({ key: e.Id.toString(), text: e.Title.toString() });
        });
       // this.listDropDownOptions.push({ key: e.Id.toString(), text: e.Title.toString() });
        return _asd; 
      });
    }
  }

      // ...
      protected onPropertyPaneConfigurationStart(): void {
        this.taskListsDropdownDisabled = !this.taskLists;
        this.listViewDropdownDisabled = !this.properties.listTitle || !this.listViews;
    
        if (this.taskLists) {
          return;
        }
    
        //this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'options');
        
            this.loadLists()
              .then((listOptions: IPropertyPaneDropdownOption[]): Promise<IPropertyPaneDropdownOption[]> => {
                this.taskLists = listOptions;
                this.taskListsDropdownDisabled = false;
                this.context.propertyPane.refresh();
                return this.loadItems();
              })
              .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
                this.listViews = itemOptions;
                this.listViewDropdownDisabled = !this.properties.listTitle;
                this.context.propertyPane.refresh();
                this.context.statusRenderer.clearLoadingIndicator(this.domElement);
                this.render();
              });
      }
      // ...
  
      // ...
    protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
      if (propertyPath === 'listTitle' &&
          newValue) {
        // push new list value
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        // get previously selected item
        const previousItem: string = this.properties.listViewName;
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
          .then((itemOptions: IPropertyPaneDropdownOption[]): void => {
            // store items
            this.listViews = itemOptions;
            // enable item selector
            this.listViewDropdownDisabled = false;
            // clear status indicator
            this.context.statusRenderer.clearLoadingIndicator(this.domElement);
            // re-render the web part as clearing the loading indicator removes the web part body
            this.render();
            // refresh the item selector control by repainting the property pane
            this.context.propertyPane.refresh();          
          });
      }
      else {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      }
    }
    // ...

  /**
   * Shouter is used to communicate between web part and view model.
   */
  private _shouter: KnockoutSubscribable<{}> = new ko.subscribable();
  private _self = this;

  /**
   * Initialize the web part.
   */
  protected onInit(): Promise<void> {
    console.log('Oninit');
    this._id = _instance++;

    const tagName: string = `ComponentElement-${this._id}`;
    this._componentElement = this._createComponentElement(tagName);
    this._registerComponent(tagName);

    // When web part description is changed, notify view model to update.   
    this._koListTitle.subscribe((newValue: string) => {
      this._shouter.notifySubscribers(newValue, 'listTitle');
    });
    this._koListViewName.subscribe((newValue: string) => {
      this._shouter.notifySubscribers(newValue, 'listViewName');
    });

    const bindings: ITaskListBindingContext = {      
      listTitle: this.properties.listTitle,
      listViewName: this.properties.listViewName,
      applyButton: this.properties.applyButton,
      shouter: this._shouter,
      context: this.context
    };

    ko.applyBindings(bindings, this._componentElement);

    return super.onInit().then(_ => {
      pnp.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    if (!this.renderedOnce) {
      this.domElement.appendChild(this._componentElement);
    }

    this._koListTitle(this.properties.listTitle);
    this._koListViewName(this.properties.listViewName);  
  }

  private _createComponentElement(tagName: string): HTMLElement {
    const componentElement: HTMLElement = document.createElement('div');
    componentElement.setAttribute('data-bind', `component: { name: "${tagName}", params: $data }`);
    return componentElement;
  }

  private _registerComponent(tagName: string): void {
    ko.components.register(
      tagName,
      {
        viewModel: TaskListViewModel,
        template: require('./TaskList.template.html'),
        synchronous: false
      }
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneLabel('labelField', {
                  text: 'Please select List and List view'
                }),
                PropertyPaneDropdown('listTitle', {
                  label: 'select the task list',
                  options: this.taskLists,
                  disabled: this.taskListsDropdownDisabled
                }),               
                PropertyPaneDropdown('listViewName', {
                  label: 'select the task list view',
                  options: this.listViews,
                  disabled: this.listViewDropdownDisabled
                }),
                PropertyPaneButton('applyButton', {
                  text: 'Apply',
                  disabled: false,
                  buttonType:PropertyPaneButtonType.Primary,
                  onClick: this.ApplyChanges
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
