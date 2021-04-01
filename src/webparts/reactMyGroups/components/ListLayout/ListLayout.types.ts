import {  ISize } from 'office-ui-fabric-react/lib/Utilities';
import { IListProps } from 'office-ui-fabric-react/lib/List';

export interface IListLayoutProps {
  sort?: string;
  /**
   * Accessible text for the grid layout
   */
  ariaLabel?: string;

  /**
   * The array of items to display
   */
  items: any[];

  /**
   * In case you want to override the underlying list
   */
  listProps?: Partial<IListProps>;

  /**
   * The method to render each cell item
   */
   onRenderListItem: (item: any, finalSize: ISize, isCompact: boolean) => JSX.Element;

}

export interface IListLayoutState {
  isLoading?: boolean;
}
