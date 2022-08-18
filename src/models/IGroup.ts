export interface IGroup {
  id: string;
  displayName: string;
  description: string;
  url?: string;
}

export interface IGroupCollection {
  value: IGroup[];
}
