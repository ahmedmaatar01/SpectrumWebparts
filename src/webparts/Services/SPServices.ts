

import { WebPartContext } from "@microsoft/sp-webpart-base";

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";
export class SPOperations {


 /**
   * GetListItemEntityTypeFullName
   * context: WebPartContext
   **/
 private async GetListItemEntityTypeFullName(context: WebPartContext, listTitle: string): Promise<string> {
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')?$select=ListItemEntityTypeFullName`;
    try {
      const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const data = await response.json();
      return data.ListItemEntityTypeFullName;
    } catch (error) {
      console.error('Error fetching ListItemEntityTypeFullName:', error);
      throw error;
    }
  }

  /**
   * UpdateListItem
   * context: WebPartContext
   **/
  public async UpdateListItem(context: WebPartContext, listTitle: string, itemId: number, updatedFields: any): Promise<void> {
    const entityTypeFullName = await this.GetListItemEntityTypeFullName(context, listTitle);
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;
    const spHttpClientOptions: ISPHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({
        "__metadata": { "type": entityTypeFullName },
        ...updatedFields
      })
    };
    console.log('Update URL:', url);
    console.log('Update Options:', spHttpClientOptions);
    try {
      const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions);
      if (!response.ok) {
        const errorResponse = await response.json();
        console.error('Error updating item:', errorResponse);
        throw new Error(`Error updating item: ${errorResponse.error.message}`);
      }
    } catch (error) {
      console.error('Error updating item:', error);
      throw error;
    }
  }


    /**
    GetAllList
    context: WebpartContext
    **/
    public GetAllList(context: WebPartContext): Promise<IDropdownOption[]> {
        let restApiurl: string =
            context.pageContext.web.absoluteUrl + "/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false";
        var listTitles: IDropdownOption[] = [];
        return new Promise<IDropdownOption[]>(async (resolve, reject) => {
            context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1).then(
                (response: SPHttpClientResponse) => {
                    response.json().then((results: any) => {

                        results.value.map((result: any) => {
                            listTitles.push({
                                key: result.Title,
                                text: result.Title,
                            });
                        });
                    });
                    resolve(listTitles);
                },
                (error: any): void => {
                    reject("error: " + error);
                }
            );
        });

    }
    public async GetUserById(context: WebPartContext, userId: number): Promise<any> {
        const userApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${userId})`;
        try {
            const response = await context.spHttpClient.get(userApiUrl, SPHttpClient.configurations.v1);
            const userData = await response.json();
            return userData;
        } catch (error) {
            console.error('Error fetching user details:', error);
            throw error;
        }
    }
    /**
  GetDocLibItems
  context: WebpartContext
  **/
    public GetDocLibItems(context: WebPartContext, title: string, directory: string): Promise<SPListItem[]> {
        let dir_link = "";
        console.log("directory from service : " + directory)
        if (!directory) {
            dir_link = context.pageContext.web.serverRelativeUrl + "/" + title
        } else {
            dir_link = directory
        }
        let restApiurl: string =
            `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${title}')/items?$select=*,FileDirRef,FileLeafRef&$filter=FileDirRef eq '${dir_link}'`;
        console.log("API URL : " + restApiurl)
        var listItems: SPListItem[] = [];
        return new Promise<SPListItem[]>(async (resolve, reject) => {
            context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1).then(
                (response: SPHttpClientResponse) => {
                    response.json().then((results: any) => {

                        listItems = results.value;
                        resolve(listItems);
                    });
                    console.log("listItems");
                    console.log(listItems);

                },
                (error: any): void => {
                    reject("error: " + error);
                }
            );
        });

    }

    /**
  GetFileAndFolderCounts
  context: WebpartContext
  **/
    public async GetFileAndFolderCounts(context: WebPartContext, listTitle: string): Promise<{ fileCount: number, folderCount: number }> {
        const restApiurl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=FileSystemObjectType`;

        try {
            const response = await context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1);
            const results = await response.json();

            let fileCount = 0;
            let folderCount = 0;

            results.value.forEach((item: any) => {
                if (item.FileSystemObjectType === 0) {
                    fileCount++;
                } else if (item.FileSystemObjectType === 1) {
                    folderCount++;
                }
            });

            return { fileCount, folderCount };
        } catch (error) {
            console.error('Error fetching file and folder counts:', error);
            throw error;
        }
    }

    /**
 * GetListColumns
 * context: WebpartContext
 **/
    public GetListColumns(context: WebPartContext, title: string): Promise<SPListColumn[]> {
        let restApiurl: string =
            `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${title}')/fields?$filter=Hidden eq false`;

        var columns: SPListColumn[] = [];
        return new Promise<SPListColumn[]>(async (resolve, reject) => {
            context.spHttpClient.get(restApiurl, SPHttpClient.configurations.v1).then(
                (response: SPHttpClientResponse) => {
                    response.json().then((results: any) => {
                        results.value.map((column: any) => {
                            // Only process columns where ReadOnlyField is false
                            if (!column.ReadOnlyField) {
                                let choices: string[] | undefined;
                                // Parse SchemaXml to get choices for choice fields
                                if (column.SchemaXml && column.SchemaXml.indexOf('<CHOICES>') !== -1) {
                                    const start = column.SchemaXml.indexOf('<CHOICES>') + '<CHOICES>'.length;
                                    const end = column.SchemaXml.indexOf('</CHOICES>');
                                    const choicesXml = column.SchemaXml.substring(start, end);
                                    choices = choicesXml.split('<CHOICE>').map((choice: string) => choice.replace('</CHOICE>', ''));
                                }

                                columns.push({
                                    id: column.Id,
                                    title: column.Title,
                                    type: column.TypeAsString,
                                    internalName: column.InternalName,
                                    description: column.Description || '',
                                    required: column.Required || false,
                                    readOnly: column.ReadOnlyField || false,
                                    fieldTypeKind: column.FieldTypeKind || 0,
                                    choices: choices,
                                    lookupField: column.LookupField || undefined
                                });
                            }
                        });
                        resolve(columns);
                    });
                },
                (error: any): void => {
                    reject("error: " + error);
                }
            );
        });
    }
}


export interface SPListItem {
    [key: string]: any; // Dynamically fetch properties

}
export interface SPListColumn {
    id: string;
    title: string;
    type: string;
    internalName: string;
    description: string;
    required: boolean;
    readOnly: boolean;
    fieldTypeKind: number;
    choices: string[] | undefined;
    lookupField: string | undefined;
}
