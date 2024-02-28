import * as React from 'react';
import styles from './ReactMyGroups.module.scss';
import { IReactMyGroupsProps } from './IReactMyGroupsProps';
import GroupService from '../../../../services/GroupService';
import { IReactMyGroupsState } from './IReactMyGroupsState';
import { GroupList } from '../GroupList';
import { Spinner, ISize } from 'office-ui-fabric-react';
import { GridLayout } from '../GridList';
import { ListLayout } from '../ListLayout';
//import * as strings from 'ReactMyGroupsWebPartStrings';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Paging } from '../../components/paging';
import { SelectLanguage } from '../SelectLanguage';

export class ReactMyGroups extends React.Component<IReactMyGroupsProps, IReactMyGroupsState> {

  constructor(props: IReactMyGroupsProps) {
    super(props);

    this.state = {
      groups: [],
      isLoading: true,
      currentPage: 1,
      pagelimit: 0,
      showless: false,
      pageSeeAll: false
    };
  }

  public strings = SelectLanguage(this.props.prefLang);
  public async componentDidUpdate (prevProps:IReactMyGroupsProps):Promise<void>{
    // if (prevProps.prefLang !== this.props.prefLang) {
    //   this.strings = SelectLanguage(this.props.prefLang);
    //    this.props.updateWebPart();
    // }
  this.props.updateWebPart();
  }

  public render(): React.ReactElement<IReactMyGroupsProps> {
    let myData=[];
    (this.props.sort === "DateCreation") ?  myData = [].concat(this.state.groups).sort((a, b) => a.createdDateTime < b.createdDateTime ? 1 : -1) : myData = [].concat(this.state.groups).sort((a, b) => a.displayName < b.displayName ? 1 : -1);
    let pagedItems: any[] = myData;
    const totalItems: number = pagedItems.length;

    let maxEvents: number = this.props.numberPerPage;
    const { currentPage } = this.state;

    //if on see all page, only show 20 at the time
    if(this.props.toggleSeeAll){
      maxEvents = 20;
    }
    if (true && totalItems > 0 && totalItems > maxEvents) {

      const pageStartAt: number = maxEvents * (currentPage - 1);
      const pageEndAt: number = (maxEvents * currentPage);

      pagedItems = pagedItems.slice(pageStartAt, pageEndAt);
    }

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    return (
      <div className={styles.reactMyGroups} style={{ backgroundColor: semanticColors.bodyBackground }}>
        <div className={styles.title} role="heading" aria-level={2}>{(this.strings.userLang === "FR" ? this.props.titleFr :this.props.titleEn )} </div>
        {(this.props.toggleSeeAll === false && !!this.props.seeAllLink) && <a aria-label={this.strings.seeAllLabel} href={this.props.seeAllLink}><div className={styles.seeAll}>{this.strings.seeAll}</div></a>}
        <div className={styles.createComm}><Icon iconName="Add" className={styles.addIcon} /><a href={this.props.createCommLink}>{this.strings.createComm}</a></div>
          {this.state.isLoading ?
    <Spinner label={this.strings.loadingState} />
    :
    <div>
    <div className={styles.groupsContainer}>
      {this.props.layout === 'Compact' ?
        <GroupList groups={pagedItems} onRenderItem={(item: any, index: number) => this._onRenderItem(item, index)}/>
      :
        this.props.layout === 'Grid' ?
        <GridLayout sort={this.props.sort} items={pagedItems} onRenderGridItem={(item: any, finalSize: ISize, isCompact: boolean) => this._onRenderGridItem(item, finalSize, isCompact)}/>
        :
          <ListLayout sort={this.props.sort} items={pagedItems} onRenderListItem={(item: any, finalSize: ISize, isCompact: boolean) => this._onRenderListItem(item, finalSize, isCompact)}/>
      }
                  {this.props.toggleSeeAll ?
                  <div>
                    <Paging
                      showPageNumber={true}
                      currentPage={currentPage}
                      itemsCountPerPage={20}
                      totalItems={totalItems}
                      onPageUpdate={this._onPageUpdate}
                      nextButtonLabel={this.strings.pagNext}
                      previousButtonLabel={this.strings.pagPrev}
                    />
                  </div> : ""
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
    });
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
    GroupService.getGroupLinksBatch(groups).then(groupUrls => {
      this.setState(prevState => ({
        groups: prevState.groups.map(group => group.id !== null ? {...group, url: groupUrls[group.id]} : group)
      }));
    });

    this._getGroupMembers(groups);
  }

  public _getGroupMembers = (groups: any): void => {
    GroupService.getGroupMembersBatch(groups).then(groupMembers => {
      this.setState(prevState => ({
        groups: prevState.groups.map(group => group.id !== null ? {...group, members: groupMembers[group.id]} : group)
      }));
    });

    this._getGroupThumbnails(groups);
  }

  public _getGroupThumbnails = (groups: any): void => {
    GroupService.getGroupThumbnailsBatch(groups).then(grouptbs => {
      this.setState(prevState => ({
        groups: prevState.groups.map(group => group.id !== null ? {...group, thumbnail: "data:image/jpeg;base64," + grouptbs[group.id], color: "#0078d4"} : group)
      }));
    });

    this.setState({
      isLoading: false
    });
  }

  private _onRenderItem = (item: any, index: number): JSX.Element => {
    return (
      <div className={styles.compactContainer}>
        <a className={styles.compactA} href={item.url}>
          <div className={styles.compactWrapper}>
            <img className={styles.compactBanner} src={item.thumbnail} alt={`${this.strings.altImgLogo} ${item.displayName}`}/>
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
          <div className={styles.topBanner} style={{backgroundColor: item.color}} />
          <img className={styles.bannerImg} src={item.thumbnail} alt={`${this.strings.altImgLogo} ${item.displayName}`} />
          <h3 className={styles.cardTitle}>{item.displayName}</h3>
          <p
          className={styles.cardDescription}
          aria-label={item.description}
        >
          {item.description}
        </p>
        </div>
      </a>
    </div>
    );
  }

  private _onRenderListItem = (item: any, finalSize: ISize, isCompact: boolean): JSX.Element => {

    return (
        <div className={styles.siteCardList}>
        <a className="community-list-item" href={item.url}>
           <div className={styles.cardBannerList}>
                <div className={styles.articleFlex} style={{'width':'60px'}}>
                   <img className={styles.bannerImgList} src={item.thumbnail} alt={`${this.strings.altImgLogo} ${item.displayName}`} />
                </div>
                <div className={`${styles.articleFlex} ${styles.secondSection}`}>
                  <div className={styles.cardTitleList}>{item.displayName}</div>
                  <div className={styles.cardDescription}>{item.description}</div>
                <div className={styles.members}>{item.members} {this.strings.members}</div>
              </div>
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
}
