import { SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface IReactMyGroupsProps {
  seeAllLink:string;
  titleEn: string;
  titleFr: string;
  layout: string;
  sort: string;
  numberPerPage: number;
  spHttpClient: SPHttpClient;
  themeVariant: IReadonlyTheme | undefined;
}
