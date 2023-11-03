import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneChoiceGroup, PropertyPaneToggle, PropertyPaneDropdown } from "@microsoft/sp-property-pane";
import GroupService from '../../services/GroupService';
import { SelectLanguage } from "./components/SelectLanguage";
import { ReactMyGroups, IReactMyGroupsProps } from './components';
import { ThemeProvider, ThemeChangedEventArgs, IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IReactMyGroupsWebPartProps {
  seeAllLink: string;
  createCommLink: string;
  titleEn: string;
  titleFr: string;
  layout: string;
  sort: string;
  toggleSeeAll: boolean;
  numberPerPage: number;
  themeVariant: IReadonlyTheme | undefined;
  prefLang: string;
}

export default class ReactMyGroupsWebPart extends BaseClientSideWebPart<IReactMyGroupsWebPartProps> {
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme;

  private strings: IReactMyGroupsWebPartStrings;

  public updateWebPart= async () => {
    this.context.propertyPane.refresh();
    this.render();
  }

   public render(): void {
    const element: React.ReactElement<IReactMyGroupsProps > = React.createElement(
      ReactMyGroups,
      {
        seeAllLink: this.properties.seeAllLink,
        createCommLink: this.properties.createCommLink,
        titleEn: this.properties.titleEn,
        titleFr: this.properties.titleFr,
        layout: this.properties.layout,
        sort: this.properties.sort,
        numberPerPage: this.properties.numberPerPage,
        toggleSeeAll: this.properties.toggleSeeAll,
        spHttpClient: this.context.spHttpClient,
        themeVariant: this._themeVariant,
        prefLang: this.properties.prefLang,
        updateWebPart:this.updateWebPart
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this.strings = SelectLanguage(this.properties.prefLang);

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);
    return super.onInit().then(() => {
      GroupService.setup(this.context);
    });
  }

  /**
 * Update the current theme variant reference and re-render.
 *
 * @param args The new theme
 */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const { layout }  = this.properties;
    let numberPerPageOption: any;
        // if toggleSeeAll is true desable numberperpage
        if (this.properties.toggleSeeAll) {
          numberPerPageOption = PropertyPaneTextField('numberPerPage', {
            label: this.strings.setPageNum,
            disabled: true
          })
        } else {
          numberPerPageOption =   PropertyPaneTextField('numberPerPage', {
            label: this.strings.setPageNum,
            disabled: false
          })
        }
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown('prefLang', {
                  label: 'Preferred Language',
                  options: [
                    { key: 'account', text: 'Account' },
                    { key: 'en-us', text: 'English' },
                    { key: 'fr-fr', text: 'Français' }
                  ],
                  selectedKey: this.strings.userLang,
                }),
                PropertyPaneTextField('seeAllLink', {
                  label: this.strings.seeAllLink
                }),
                PropertyPaneTextField('createCommLink', {
                  label: this.strings.createCommLink
                }),
                PropertyPaneTextField('titleEn', {
                  label: this.strings.setTitleEn
                }),
                PropertyPaneTextField('titleFr', {
                  label: this.strings.setTitleFr
                }),
                PropertyPaneToggle('toggleSeeAll', {
                  key: 'toggleSeeAll',
                  label: this.strings.seeAllToggle,
                  checked: false,
                  onText: this.strings.seeAllOn,
                  offText: this.strings.seeAllOff,
                }),
                numberPerPageOption,
                PropertyPaneChoiceGroup("layout", {
                  label: this.strings.setLayoutOpt,
                  options: [
                    {
                      key: "Grid",
                      text: this.strings.gridIcon,
                      iconProps: { officeFabricIconFontName: "GridViewSmall"},
                      checked: layout === "Grid" ? true : false,

                    },
                    {
                      key: "Compact",
                      text: this.strings.compactIcon,
                      iconProps: { officeFabricIconFontName: "BulletedList2"},
                      checked: layout === "Compact" ? true : false
                    },
                    {
                      key: "List",
                      text: this.strings.ListIcon,
                      iconProps: { officeFabricIconFontName: "ViewList"},
                      checked: layout === "List" ? true : false
                    }
                  ]
                }),
                PropertyPaneChoiceGroup("sort", {
                  label: this.strings.setSortOpt,
                  options: [
                    {
                      key: "DateCreation",
                      text: this.strings.dateCreation,
                      checked: layout === "DateCreation" ? true : false,

                    },
                    {
                      key: "Alphabetical",
                      text: this.strings.alphabetical,
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
