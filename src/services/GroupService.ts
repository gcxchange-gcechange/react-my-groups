import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IGroup, IGroupCollection } from "../models";
import { GraphRequest } from "@microsoft/microsoft-graph-client";


export class GroupServiceManager {
  public context: WebPartContext;

  public setup(context: WebPartContext): void {
    this.context = context;
  }

  public getGroups(): Promise<MicrosoftGraph.Group[]> {
    return new Promise<MicrosoftGraph.Group[]>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api("/me/memberOf/$/microsoft.graph.group?$filter=groupTypes/any(a:a eq 'unified')")
          .get((error: any, groups: IGroupCollection, rawResponse: any) => {
           // console.log("GROUP "+JSON.stringify(groups))
            resolve(groups.value);
          });
        });
      } catch(error) {
        console.error("ERROR-"+error);
      }
    });
  }

  public getGroupLinks(groups: IGroup): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api(`/groups/${groups.id}/sites/root/weburl`)
          .get((error: any, group: any, rawResponse: any) => {
            resolve(group);
          });
        });
      } catch(error) {
        console.error(error);
      }
    });
  }

  public getGroupLinksBatch(groups: IGroup[]): Promise<MicrosoftGraph.Group[]> {
    let requestBody = new MicrosoftGraph.BatchRequestContent(
      groups.map( (group) => <MicrosoftGraph.BatchRequestStep>{
        id: group.id,
        request: new Request(`/groups/${group.id}/sites/root/weburl`, {
          method: "GET"
        })
      })
    );

    return new Promise<MicrosoftGraph.Group[]>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api(`/$batch`)
          .post( requestBody.getContent(), (error: any, rawResponse: any) => {
            let linksResponseContent = new MicrosoftGraph.BatchResponseContent(rawResponse);

            resolve(groups.map(group => group.id !== null ? {...group, url: linksResponseContent.getResponseById(group.id).json().body.value} : group));
          });
        });
      } catch(error) {
        console.error(error);
      }
    });
  }

  public getGroupMembers(groups: IGroup): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api(`/groups/${groups.id}/members/$count?ConsistencyLevel=eventual`)
          .get((error: any, group: any, rawResponse: any) => {
            resolve(group);
            console.log("MEMBERS "+JSON.stringify(group))

          });
        });
      } catch(error) {
        console.error(error);
      }
    });
  }

  public getGroupMembersBatch(groups: IGroup[]): Promise<MicrosoftGraph.Group[]> {
    let requestBody = new MicrosoftGraph.BatchRequestContent(
      groups.map( (group) => <MicrosoftGraph.BatchRequestStep>{
        id: group.id,
        request: new Request(`/groups/${group.id}/members/$count?ConsistencyLevel=eventual`, {
          method: "GET"
        })
      })
    );

    return new Promise<MicrosoftGraph.Group[]>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api(`/$batch`)
          .post( requestBody.getContent(), (error: any, rawResponse: any) => {
            let membersResponseContent = new MicrosoftGraph.BatchResponseContent(rawResponse);

            resolve(groups.map(group => group.id !== null ? {...group, members: membersResponseContent.getResponseById(group.id).json().body.value} : group));
          });
        });
      } catch(error) {
        console.error(error);
      }
    });
  }

  public getGroupThumbnails(groups: IGroup): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api(`/groups/${groups.id}/photos/48x48/$value`)
          .responseType('blob')
          .get((error: any, group: any, rawResponse: any) => {
            resolve(window.URL.createObjectURL(group));
          });
        });
      } catch(error) {
        console.error("ERROR "+error);
      }
    });
  }

  public getGroupThumbnailsBatch(groups: IGroup[]): Promise<MicrosoftGraph.Group[]> {
    let requestBody = new MicrosoftGraph.BatchRequestContent(
      groups.map( (group) => <MicrosoftGraph.BatchRequestStep>{
        id: group.id,
        request: new Request(`/groups/${group.id}/photos/48x48/$value`, {
          method: "GET"
        })
      })
    );

    return new Promise<MicrosoftGraph.Group[]>((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient) => {
          client
          .api(`/$batch`)
          .post( requestBody.getContent(), (error: any, rawResponse: any) => {
            let thumbnailsResponseContent = new MicrosoftGraph.BatchResponseContent(rawResponse);

            resolve(groups.map(group => group.id !== null ? {...group, thumbnail: thumbnailsResponseContent.getResponseById(group.id).json().body.value, color: "#0078d4"} : group));
          });
        });
      } catch(error) {
        console.error(error);
      }
    });
  }
}

const GroupService = new GroupServiceManager();
export default GroupService;

