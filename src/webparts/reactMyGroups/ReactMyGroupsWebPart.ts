import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneChoiceGroup } from "@microsoft/sp-property-pane";
import GroupService from '../../services/GroupService';
import * as strings from 'ReactMyGroupsWebPartStrings';
import { ReactMyGroups, IReactMyGroupsProps } from './components';

export interface IReactMyGroupsWebPartProps {
  titleEn: string;
  titleFr: string;
  layout: string;
  sort: string;
  numberPerPage: number;
}

export default class ReactMyGroupsWebPart extends BaseClientSideWebPart<IReactMyGroupsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactMyGroupsProps > = React.createElement(
      ReactMyGroups,
      {
        titleEn: this.properties.titleEn,
        titleFr: this.properties.titleFr,
        layout: this.properties.layout,
        sort: this.properties.sort,
        numberPerPage: this.properties.numberPerPage,
        spHttpClient: this.context.spHttpClient,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(() => {
      GroupService.setup(this.context);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const { layout }  = this.properties;
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('titleEn', {
                  label: strings.setTitleEn
                }),
                PropertyPaneTextField('titleFr', {
                  label: strings.setTitleFr
                }),
                PropertyPaneTextField('numberPerPage', {
                  label: strings.setPageNum
                }),
                PropertyPaneChoiceGroup("layout", {
                  label: strings.setLayoutOpt,
                  options: [
                    {
                      key: "Grid",
                      text: strings.gridIcon,
                      iconProps: { officeFabricIconFontName: "GridViewSmall"},
                      checked: layout === "Grid" ? true : false,

                    },
                    {
                      key: "Compact",
                      text: strings.compactIcon,
                      iconProps: { officeFabricIconFontName: "BulletedList2"},
                      checked: layout === "Compact" ? true : false
                    }
                  ]
                }),
                PropertyPaneChoiceGroup("sort", {
                  label: strings.setSortOpt,
                  options: [
                    {
                      key: "DateCreation",
                      text: strings.dateCreation,
                      checked: layout === "DateCreation" ? true : false,

                    },
                    {
                      key: "Alphabetical",
                      text: strings.alphabetical,
                      checked: layout === "Alphabetical" ? true : false
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
