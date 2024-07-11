export interface IGroup {
  id: string;
  displayName: string;
  url?: string;
  siteId: string;
}

export interface IGroupCollection {
  value: IGroup[];
}
