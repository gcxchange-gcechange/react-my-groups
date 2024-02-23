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

      try {
        this.context.msGraphClientFactory.getClient('3')
        .then((client: MSGraphClientV3) => {
          client.api("/me/memberOf/$/microsoft.graph.group?$filter=groupTypes/any(a:a eq 'unified')")
          .get((error: any, groups: IGroupCollection, rawResponse: any) => {


            this.context.msGraphClientFactory.getClient('3')
            .then((client: MSGraphClientV3) => {
              client.api("/me/ownedObjects/$/microsoft.graph.group")
              .get((error: any, groups2: IGroupCollection, rawResponse: any) => {

                const responseResults: any[] = groups.value.concat(groups2.value);

                const uniqueValues = responseResults.filter(
                  (obj, index, self) =>
                    index === self.findIndex((innerObj) => innerObj.id === obj.id)
                );

                resolve(uniqueValues);
              });
            });
          });
        });
      } catch(error) {
        console.error("ERROR-"+error);
      }
    });
  }

  public getGroupDetailsBatch(group: any): Promise<any> {
    const requestBody = {
      requests: [
        {
          id: "1",
          method: "GET",
          url: `/groups/${group.id}/sites/root/weburl`,
        },
        {
          id: "2",
          method: "GET",
          url: `/groups/${group.id}/members/$count?ConsistencyLevel=eventual`
        },
        {
          id: "3",
          method: "GET",
          url: `/groups/${group.id}/photos/48x48/$value`
        },

      ],
    };
    return new Promise((resolve, reject) => {
      try {
        this.context.msGraphClientFactory
          .getClient('3')
          .then((client: MSGraphClientV3):void => {
            client
              .api(`/$batch`)
              .post(requestBody, (error: any, responseObject: any) => {

                if (error) {
                  Promise.reject(error);
                }
                const responseContent = {};

                responseObject.responses.forEach((response) => {

                  if (response.status === 200) {
                    responseContent[response.id] = response.body;
                  } else if (response.status === 403 || response.status === 404) {
                    return null;
                  }
                });

                resolve(responseContent);
              });
          });
      } catch (error) {
        reject(error);
        console.error(error);
      }
    });
  }

}

const GroupService = new GroupServiceManager();
export default GroupService;

