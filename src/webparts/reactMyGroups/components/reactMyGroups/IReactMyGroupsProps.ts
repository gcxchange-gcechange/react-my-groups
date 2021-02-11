import { SPHttpClient } from '@microsoft/sp-http'; 
export interface IReactMyGroupsProps {
  titleEn: string;
  titleFr: string;
  layout: string;
  sort: string;
  numberPerPage: number;
  spHttpClient: SPHttpClient;  
}
