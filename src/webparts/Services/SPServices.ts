

import { WebPartContext } from "@microsoft/sp-webpart-base";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { IDropdownOption } from "office-ui-fabric-react";
export class SPOperations {

    public async GetUsers(context: WebPartContext): Promise<{ id: number, title: string }[]> {
        const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/siteusers`;
        try {
            const response = await context.spHttpClient.get(restApiUrl, SPHttpClient.configurations.v1);
            const data = await response.json();
            console.log("==========users from service========")
            console.log(data)
            return data.value.map((user: any) => ({ id: user.Id, title: user.Title }));
        } catch (error) {
            console.error('Error fetching users:', error);
            throw error;
        }
    }
    public DeleteListItem(context: WebPartContext, listTitle: string, itemId: number): Promise<string> {
        const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;
    
        return new Promise<string>((resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, {
                headers: {
                    Accept: "application/json;odata=nometadata",
                    "Content-Type": "application/json;odata=nometadata",
                    "odata-version": "",
                    "IF-MATCH": "*",
                    "X-HTTP-METHOD": "DELETE",
                }
            }).then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    resolve(`Item with ID ${itemId} deleted successfully`);
                } else {
                    reject(`Error deleting item: ${response.statusText}`);
                }
            }, (error: any) => {
                reject(`Error deleting item: ${error}`);
            });
        });
    }
    

    /**
     * Update a list item
     * @param context The web part context
     * @param listTitle The title of the list
     * @param itemId The ID of the item to update
     * @param item The updated item data
     */
    public UpdateListItemFields(context: WebPartContext, listTitle: string, itemId: number, fieldsToUpdate: { [key: string]: any }): Promise<string> {
        const restApiUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${listTitle}')/items(${itemId})`;
        const body = JSON.stringify(fieldsToUpdate);
    
        return new Promise<string>((resolve, reject) => {
            context.spHttpClient.post(restApiUrl, SPHttpClient.configurations.v1, {
                headers: {
                    "Accept": "application/json;odata=nometadata",
                    "Content-Type": "application/json;odata=nometadata",
                    "odata-version": "",
                    "IF-MATCH": "*",
                    "X-HTTP-METHOD": "MERGE",
                },
                body: body,
            })
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    resolve(`Item with ID ${itemId} updated successfully`);
                } else {
                    reject(`Error updating item: ${response.statusText}`);
                }
            }, (error: any) => {
                reject(`Error updating item: ${error}`);
            });
        });
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
                        console.log("+-+-+-+-+-+-+-+listItems From Service+-+-+-+-+-+-+-+");
                        console.log(results.value);
                    });
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
