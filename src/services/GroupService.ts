import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGroup, IGroupCollection } from "../models";


export class GroupServiceManager {
  public context: WebPartContext;

  public setup(context: WebPartContext): void {
    this.context = context;
  }

  public getGroups(): Promise<MicrosoftGraph.Group[]> {
    return new Promise<MicrosoftGraph.Group[]>((resolve, reject) => {
      const responseResults: MicrosoftGraph.Group[] = [];
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(
                "/me/memberOf/$/microsoft.graph.group?$filter=groupTypes/any(a:a eq 'unified')"
              )
              .get((error: any, groups: IGroupCollection, rawResponse: any) => {
                responseResults.push(...groups.value);

                this.context.msGraphClientFactory
                  .getClient("3")
                  .then((client: MSGraphClientV3) => {
                    client
                      .api("/me/ownedObjects/$/microsoft.graph.group")
                      .get(
                        (
                          error: any,
                          groups2: IGroupCollection,
                          rawResponse: any
                        ) => {
                          groups2.value.forEach(function (value) {
                            let foundDuplicate: boolean = false;

                            responseResults.forEach(function (value2) {
                              if (value.id === value2.id) {
                                foundDuplicate = true;
                              }
                            });

                            if (!foundDuplicate) {
                              responseResults.push(value);
                            }
                          });

                          resolve(responseResults);
                        }
                      );
                  });
              });
          });
      } catch (error) {
        console.error("ERROR-" + error);
      }
    });
  }

  public getOwnedGroups(): Promise<MicrosoftGraph.Group[]> {
    return new Promise<MicrosoftGraph.Group[]>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api("/me/ownedObjects/$/microsoft.graph.group") // ?$filter=groupTypes/any(a:a eq 'unified')
              .get((error: any, groups: IGroupCollection, rawResponse: any) => {
                //console.log("OWNED GROUP "+JSON.stringify(groups))
                resolve(groups.value);
              });
          });
      } catch (error) {
        console.error("ERROR-" + error);
      }
    });
  }

  public getGroupLinks(groups: IGroup): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(`/groups/${groups.id}/sites/root/weburl`)
              .get((error: any, group: any, rawResponse: any) => {
                resolve(group);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupDetailsBatch(groups: IGroup[]): Promise<any> {
    const x = typeof groups;
    console.log(x);
    const requestBody = { requests: [] };
    requestBody.requests = groups.map((group) => ({
      id: group.id,
      method: "GET",
      url: `/groups/${group.id}/sites/root/?$select=id,webUrl,lastModifiedDateTime`,
    }));

    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(`/$batch`)
              .post(requestBody, (error: any, responseObject: any) => {
                const linksResponseContent = {};
                responseObject.responses.forEach(
                  (response) =>
                    (linksResponseContent[response.id] = response.body)
                );

                resolve(linksResponseContent);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupMembers(groups: IGroup): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(
                `/groups/${groups.id}/members/$count?ConsistencyLevel=eventual`
              )
              .get((error: any, group: any, rawResponse: any) => {
                resolve(group);
                console.log("MEMBERS " + JSON.stringify(group));
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupMembersBatch(groups: IGroup[]): Promise<any> {
    const x = typeof groups;
    console.log(x);
    const requestBody = { requests: [] };
    requestBody.requests = groups.map((group) => ({
      id: group.id,
      method: "GET",
      url: `/groups/${group.id}/members/$count?ConsistencyLevel=eventual`,
    }));

    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(`/$batch`)
              .post(requestBody, (error: any, responseObject: any) => {
                const membersResponseContent = {};
                responseObject.responses.forEach(
                  (response) =>
                    (membersResponseContent[response.id] = response.body)
                );

                resolve(membersResponseContent);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupThumbnails(groups: IGroup): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(`/groups/${groups.id}/photos/48x48/$value`)
              //.responseType('blob')
              .get((error: any, group: any, rawResponse: any) => {
                resolve(window.URL.createObjectURL(group));
              });
          });
      } catch (error) {
        console.error("ERROR " + error);
      }
    });
  }

  public getGroupThumbnailsBatch(groups: IGroup[]): Promise<any> {
    const x = typeof groups;
    console.log(x);

    const requestBody = { requests: [] };
    requestBody.requests = groups.map((group) => ({
      id: group.id,
      method: "GET",
      url: `/groups/${group.id}/photos/48x48/$value`,
    }));

    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(`/$batch`)
              .post(requestBody, (error: any, responseObject: any) => {
                const thumbnailsResponseContent = {};
                responseObject.responses.forEach(
                  (response) =>
                    (thumbnailsResponseContent[response.id] = response.body)
                );

                resolve(thumbnailsResponseContent);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }

  public getGroupViewsBatch(groups: IGroup[]): Promise<any> {
    const x = typeof groups;
    console.log(x);
    const requestBody = { requests: [] };
    requestBody.requests = groups.map((group) => ({
      id: group.id,
      method: "GET",
      url: `/sites/${group.siteId}/analytics/lastsevendays/access/actionCount`,
    }));

    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient("3")
          .then((client: MSGraphClientV3) => {
            client
              .api(`/$batch`)
              .post(requestBody, (error: any, responseObject: any) => {
                const viewsResponseContent = {};
                responseObject.responses.forEach(
                  (response) =>
                    (viewsResponseContent[response.id] = response.body.value)
                );

                resolve(viewsResponseContent);
              });
          });
      } catch (error) {
        console.error(error);
      }
    });
  }
}

const GroupService = new GroupServiceManager();
export default GroupService;

