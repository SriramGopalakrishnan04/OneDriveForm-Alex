import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

// This is the list in which we will hold OneDrive migration requests
const listName = "One Drive Migration";
const completedListName = "Completed One Drive Migrations";

function getItemsFromCompletedListByUser(webPartContext: WebPartContext, userName: string): Promise<SPHttpClientResponse> {
    // Create the CAML query
    const queryText: string = `<View>
                                    <Query>
                                        <Where>
                                            <Eq>
                                                <FieldRef Name="PorJnumber" />
                                                <Value Type="Text">${userName}</Value>
                                            </Eq>
                                        </Where>
                                        <OrderBy>
                                            <FieldRef Name="ID" />
                                        </OrderBy>
                                    </Query>
                                </View>`;

    // Create the request options to use the CAML query
    const options: ISPHttpClientOptions = {
        headers: { 'odata-version': '3.0' },
        body: `{'query': {
            '__metadata': {'type': 'SP.CamlQuery'},
            'ViewXml': '${queryText}'
        }}`
    };

    // Make the request URL
    let requestUrl = webPartContext.pageContext.web.absoluteUrl.concat(`/_api/web/Lists/GetByTitle('${completedListName}')/GetItems`);

    return new Promise((res, rej) => {

        // Send the request to check if the user has already filled out the form.
        webPartContext.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, options)
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    res(response);
                } else {
                    rej(response);
                }
            }).catch((response: SPHttpClientResponse) => {
                rej(response);
            });
    });
}

function getListItemsByUserId(webPartContext: WebPartContext, userId: string): Promise<SPHttpClientResponse> {
    // Create the CAML query
    const queryText: string = `<View>
                                    <Query>
                                        <Where>
                                            <Eq>
                                                <FieldRef Name="Author" LookupId="TRUE" />
                                                <Value Type="User">${userId}</Value>
                                            </Eq>
                                        </Where>
                                        <OrderBy>
                                            <FieldRef Name="ID" />
                                        </OrderBy>
                                    </Query>
                                </View>`;

    // Create the request options to use the CAML query
    const options: ISPHttpClientOptions = {
        headers: { 'odata-version': '3.0' },
        body: `{'query': {
            '__metadata': {'type': 'SP.CamlQuery'},
            'ViewXml': '${queryText}'
        }}`
    };

    // Make the request URL
    let requestUrl = webPartContext.pageContext.web.absoluteUrl.concat(`/_api/web/Lists/GetByTitle('${listName}')/GetItems`);

    return new Promise((res, rej) => {

        // Send the request to check if the user has already filled out the form.
        webPartContext.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, options)
            .then((response: SPHttpClientResponse) => {
                if (response.ok) {
                    res(response);
                } else {
                    rej(response);
                }
            }).catch((response: SPHttpClientResponse) => {
                rej(response);
            });
    });
}

function createOneDriveMigrationRequest(webPartContext: WebPartContext, userId: string, typeName: string): Promise<SPHttpClientResponse> {
    let userLogin = webPartContext.pageContext.user.loginName.replace('i:0#.f|membership|', '').split("@")[0];


    const body: string = JSON.stringify({
        '__metadata': {
            'type': typeName
        },
        'Title': `${webPartContext.pageContext.user.displayName}`,
        'PorJNumber': userLogin
    });

    let requestUrl = webPartContext.pageContext.web.absoluteUrl.concat(`/_api/web/Lists/GetByTitle('${listName}')/items`);

    return new Promise((res, rej) => {
        webPartContext.spHttpClient.post(requestUrl,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=verbose',
                    'odata-version': ''
                },
                body: body
            }).then((result) => {
                result.json().then(val => {
                    res(val);
                });
            }, (error) => {
                error.json().then(val => {
                    rej(val);
                });
            });
    });
}

function getCurrentUserLookupId(userLogin: string, siteUrl: string, spHttpClient: SPHttpClient): Promise<SPHttpClientResponse> {

    const payload: string = JSON.stringify({
        'logonName': userLogin // i:0#.f|membership|firstname.lastname@contoso.onmicrosoft.com      
    });

    var postData: ISPHttpClientOptions = {
        body: payload
    };

    var endPoint = `${siteUrl}/_api/web/ensureuser`;

    return new Promise((resolve, reject) => {
        spHttpClient.post(endPoint,
            SPHttpClient.configurations.v1,
            postData)
            .then(
                (response: SPHttpClientResponse) => {
                    resolve(response);
                },
                (error: SPHttpClientResponse) => {
                    reject(error);
                }
            );
    });
}


function getListItemEntityTypeName(siteUrl: string, spHttpClient: SPHttpClient): Promise<SPHttpClientResponse> {
    return new Promise<SPHttpClientResponse>((resolve, reject) => {
    //   if (this.listItemEntityTypeName) {
    //     resolve(this.listItemEntityTypeName);
    //     return;
    //   }

      spHttpClient.get(`${siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
        .then((response: SPHttpClientResponse) => {
          resolve(response);
        }, (error: any): void => {
          reject(error);
        });
    });
  }


export {
    getListItemsByUserId,
    createOneDriveMigrationRequest,
    getCurrentUserLookupId,
    getListItemEntityTypeName,
    getItemsFromCompletedListByUser
};