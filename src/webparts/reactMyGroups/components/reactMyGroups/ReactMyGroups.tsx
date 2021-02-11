import * as React from 'react';
import styles from './ReactMyGroups.module.scss';
import { IReactMyGroupsProps } from './IReactMyGroupsProps';
import GroupService from '../../../../services/GroupService';
import { IReactMyGroupsState } from './IReactMyGroupsState';
import { GroupList } from '../GroupList';
import { Spinner, ISize, GroupShowAll } from 'office-ui-fabric-react';
import { GridLayout } from '../GridList';
import * as strings from 'ReactMyGroupsWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'; 
import { DefaultButton, PrimaryButton,CommandBarButton } from 'office-ui-fabric-react/lib/Button';

export class ReactMyGroups extends React.Component<IReactMyGroupsProps, IReactMyGroupsState> {

  constructor(props: IReactMyGroupsProps) {
    super(props);

    this.state = {
      groups: [],
      isLoading: true,
      currentPage: 1,
      pagelimit: 0,
      showless: false
    };
  }

  public render(): React.ReactElement<IReactMyGroupsProps> {
    let myData=[];
    (this.props.sort == "DateCreation") ?  myData = [].concat(this.state.groups).sort((a, b) => a.createdDateTime < b.createdDateTime ? 1 : -1) : myData = [].concat(this.state.groups).sort((a, b) => a.displayName < b.displayName ? 1 : -1);
    let pagedItems: any[] = myData;
    const totalItems: number = pagedItems.length;
    let showPages: boolean = false;

    const maxEvents: number = this.state.pagelimit;
    const { currentPage } = this.state;

    if (true && totalItems > 0 && totalItems > maxEvents) {

      const pageStartAt: number = maxEvents * (currentPage - 1);
      const pageEndAt: number = (maxEvents * currentPage);

      pagedItems = pagedItems.slice(pageStartAt, pageEndAt);
      showPages = true;
     } 

    return (
      <div className={ styles.reactMyGroups }>
        <div className={styles.title} role="heading" aria-level={2}>{(strings.userLang == "FR" ? this.props.titleFr :this.props.titleEn )} </div>      
          {this.state.isLoading ?
            <Spinner label="Loading sites..." />
                : 
                <div>
                  {this.props.layout == 'Compact' ?
                    <GroupList groups={pagedItems} onRenderItem={(item: any, index: number) => this._onRenderItem(item, index)}/>
                  :
                    <GridLayout sort={this.props.sort} items={pagedItems} onRenderGridItem={(item: any, finalSize: ISize, isCompact: boolean) => this._onRenderGridItem(item, finalSize, isCompact)}/>
                  }
                  <div>
                    {showPages &&
                      <DefaultButton  className={styles.buttonLink} text={strings.showmore} onClick={this.ShowAll} />
                    }
                    {this.state.showless &&
                      <DefaultButton  className={styles.buttonLink} text={strings.showless} onClick={this.ShowLess} />
                    }
                  </div>
                </div>
          }
      </div>
    );
  }

  public componentDidMount (): void {
    this._getGroups();
    this.setState({
      pagelimit: this.props.numberPerPage
    })
  }

  public _getGroups = (): void => {
    GroupService.getGroups().then(groups => {
      this.setState({
        groups: groups
      });
      this._getGroupLinks(groups);
    });
  }

  public _getGroupLinks = (groups: any): void => {
    groups.map(groupItem => (
      GroupService.getGroupLinks(groupItem).then(groupurl => {
        this.setState(prevState => ({
          groups: prevState.groups.map(group => group.id === groupItem.id ? {...group, url: groupurl.value} : group)
        }));
      })
    ));
    this._getGroupThumbnails(groups);
  }

  public _getGroupThumbnails = (groups: any): void => {
    groups.map(groupItem => (
      GroupService.getGroupThumbnails(groupItem).then(grouptb => {
        this.setState(prevState => ({
          groups: prevState.groups.map(group => group.id === groupItem.id ? {...group, thumbnail: grouptb, color: "#0078d4"} : group)
        }));
      })
    ));
    this.setState({
      isLoading: false
    });
  }

  private _onRenderItem = (item: any, index: number): JSX.Element => {
    return (
      <div className={styles.compactContainer}>
        <a className={styles.compactA} href={item.url}>
          <div className={styles.compactWrapper}>
            <img className={styles.compactBanner} src={item.thumbnail} alt={`${strings.altImgLogo} ${item.displayName}`}/>
            <div className={styles.compactDetails}>
              <div className={styles.compactTitle}>{item.displayName}</div>
            </div>
          </div>
        </a>
      </div>
    );
  }

  private _onRenderGridItem = (item: any, finalSize: ISize, isCompact: boolean): JSX.Element => {

    return (
        <div className={styles.siteCard}>
            <a href={item.url}>
              <div className={styles.cardBanner}>
                <div className={styles.topBanner} style={{backgroundColor: item.color}}></div>
                <img className={styles.bannerImg} src={item.thumbnail} alt={`${strings.altImgLogo} ${item.displayName}`} />
                <div className={styles.cardTitle}>{item.displayName}</div>
              </div>
            </a>
          </div>
    );
  }

   private _onPageUpdate = (pageNumber: number): void => {
    this.setState({
      currentPage: pageNumber
    });
  }

  private ShowAll= (): void =>{
    if(this.state.pagelimit != 0){
      this.setState({
        pagelimit:999,
        showless: true
      })
    }
  }

  private ShowLess= (): void =>{
    if(this.state.pagelimit != 0){
      this.setState({
        pagelimit:this.props.numberPerPage,
        showless: false
      })
    }
  }
}
