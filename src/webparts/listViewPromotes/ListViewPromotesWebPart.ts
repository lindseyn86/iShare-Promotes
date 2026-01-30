/* eslint-disable @typescript-eslint/no-explicit-any */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Environment, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import ListViewPromotes from './components/ListViewPromotes';
import { IListViewPromotesProps } from './components/IListViewPromotesProps';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';

import SharePointService from '../../services/SharePoint/spService';

export interface IListViewPromotesWebPartProps {
  webpartTitle: string;
  description: string;
  listId: string;
  employeeName: string;
  employeeLastName: string;
  imageURL: string;
  newCareerLevel: string;
  promotionLevel: string;
  solutionFunction: string;
  region: string;
  country: string;
  office: string;
  peopleSearchURL: string;
  showoverview: boolean;
  overview: string;
  sortFields: string[];
  filterFields: string[];
}

export default class ListViewPromotesWebPart extends BaseClientSideWebPart<IListViewPromotesWebPartProps> {

  //list options state
  private listOptions: IPropertyPaneDropdownOption[];
  private listOptionsLoading: boolean = false;

  // field options state
  private fieldOptions: IPropertyPaneDropdownOption[];
  private fieldOptionsLoading: boolean = false;
  
  public render(): void {
    const element: React.ReactElement<IListViewPromotesProps> = React.createElement(
      ListViewPromotes,
      {
        webpartTitle: this.properties.webpartTitle,
        description: this.properties.description,
        listId: this.properties.listId,
        employeeName: this.properties.employeeName,
        employeeLastName: this.properties.employeeLastName,
        imageURL: this.properties.imageURL,
        newCareerLevel: this.properties.newCareerLevel,
        promotionLevel: this.properties.promotionLevel,
        solutionFunction: this.properties.solutionFunction,
        region: this.properties.region,
        country: this.properties.country,
        office: this.properties.office,
        peopleSearchURL: this.properties.peopleSearchURL,
        showoverview: this.properties.showoverview,
        overview: this.properties.overview,
        sortFields: this.properties.sortFields,
        filterFields: this.properties.filterFields
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      SharePointService.setup(this.context, Environment.type);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: ''
          },
          groups: [
            {
              groupName: 'Web Part Settings',
              groupFields: [
                PropertyPaneTextField('webpartTitle', {
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('description', {
                  label: 'Web Part Description'
                })
              ]
            },
            {
              groupName: 'List Settings',
              groupFields: [
                PropertyPaneDropdown('listId', {
                  label: 'List',
                  options: this.listOptions,
                  disabled: this.listOptionsLoading
                }),
                PropertyPaneDropdown('employeeName', {
                  label: 'Employee Name',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('employeeLastName', {
                  label: 'Employee Last Name',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('imageURL', {
                  label: 'Image URL',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('newCareerLevel', {
                  label: 'New Career Level',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('promotionLevel', {
                  label: 'Promotion Level',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('solutionFunction', {
                  label: 'Solution / Function',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('region', {
                  label: 'Region',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('country', {
                  label: 'Country',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('office', {
                  label: 'Office',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('peopleSearchURL', {
                  label: 'People Search',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneDropdown('overview', {
                  label: 'Overview',
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading
                }),
                PropertyPaneCheckbox('showoverview', {
                  checked: false,
                  text: 'Show Overview'
                })
              ]
            },
            {
              groupName: 'Sort and Filter Settings',
              groupFields: [
                PropertyFieldMultiSelect('sortFields', {
                  key: 'sortFields',
                  label: "Sort By",
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading,
                  selectedKeys: this.properties.sortFields
                }),
                PropertyFieldMultiSelect('filterFields', {
                  key: 'filterFields',
                  label: "Filter By",
                  options: this.fieldOptions,
                  disabled: this.fieldOptionsLoading,
                  selectedKeys: this.properties.filterFields
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getLists(): Promise<IPropertyPaneDropdownOption[]> {
    this.listOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getLists().then(lists => {
      this.listOptionsLoading = false;
      this.context.propertyPane.refresh();
  
      return lists.value.map(list => {
        return {
          key: list.Id,
          text: list.Title
        };
      });
    });
  }

  public getFields(): Promise<any> {
    //no list selected
    if(!this.properties.listId) return Promise.resolve();
    
    this.fieldOptionsLoading = true;
    this.context.propertyPane.refresh();

    return SharePointService.getListFields(this.properties.listId).then(fields => {
      this.fieldOptionsLoading = false;
      this.context.propertyPane.refresh();
  
      return fields.value.map(field => {
        return {
          key: `${field.InternalName}+${field.Title}`,
          text: `${field.Title} (${field.TypeAsString})`
        };
      });
    });
  }

  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    await this.getLists().then(listOptions => {
      this.listOptions = listOptions;
      this.context.propertyPane.refresh();
    }).then(async () => {
      await this.getFields().then(fieldOptions => {
        this.fieldOptions = fieldOptions;
        this.context.propertyPane.refresh();
      });
    });
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): Promise<void> {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();

    if(propertyPath === 'listId' && newValue) {
      this.properties.employeeName = '';
      this.properties.employeeLastName = '';
      this.properties.imageURL = '';
      this.properties.newCareerLevel = '';
      this.properties.solutionFunction = '';
      this.properties.region = '';
      this.properties.country = '';
      this.properties.office = '';
      this.properties.peopleSearchURL = '';
      this.properties.showoverview = false;
      this.properties.overview = '';
      this.properties.sortFields = [];
      this.properties.filterFields = [];

      await this.getFields().then(fieldOptions => {
        this.fieldOptions = fieldOptions;
        this.context.propertyPane.refresh();
      });
    }
  }
}
