//import { SPHttpClient } from '@pnp/sp';
import constant from './Constant';
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from "@microsoft/sp-http";
export default class CommonUtility {

    /**
     *Get Default Site Collection Path. */
    public siteCollectionPath = '/sites/' + constant.SiteCollUrlName;

   /**
    * Get Default Tenant Name from window url.
    */
    public tenantURL() {
        try {
            // let currentPageURL = window.location.href;
            // let tenantURL = currentPageURL.split('com/')[0];
            // tenantURL = tenantURL + 'com';
            // return tenantURL;
            let currentPageURL = window.location.href;
            var tenantURL = "https://"+ currentPageURL.substr(8, currentPageURL.length).split('/')[0];
            return tenantURL;
        } catch (error) {
            console.log("tenantURL: " + error);
            return '';
        }
    }
    public TenantUrl = `${this.tenantURL()}${this.siteCollectionPath}`;

    /**
     * Check wheather the new site collection exist or not.
     * @param siteUrl 
     */
    public _checkSiteCollectionExist(siteUrl: string, context): Promise<any> {
        var checkSiteTenantName = `${this.tenantURL()}/_api/SP.Site.Exists`;
        return new Promise((resolve, reject) => {
            var client = context.spHttpClient;
            client.post(checkSiteTenantName, SPHttpClient.configurations.v1, {
                body: JSON.stringify({
                    url: siteUrl
                }),
                headers: {
                    "accept": "application/json;",
                }, credentials: "same-origin"
            }).then(d => {
                d.json().then((v: any) => {
                    resolve(v.value);
                });
            }).catch(d => {
                reject(d);
            });
        });
    }

    /**
     * Remove space in between tags
     * @param htmlContain 
     */
    public extractContent(htmlContain) {
        var span = document.createElement('span');
        span.innerHTML = htmlContain;
        return span.textContent || span.innerText;
      }

   /**
    * This function to get data from List or Library on 'GET' request.
    * @param url 
    */
    public _getRequest(url: string): any {        
        try {
            return fetch(url, {
                headers: { Accept: 'application/json;odata=verbose' },
                credentials: "same-origin"
            }).then((response) => {
                if (response.status >= 200 && response.status < 400) {
                    return response.json();
                }
                else {
                    return response.json();
                }
            }).catch(error =>
                console.error("getRequest: " + error));
        } catch (error) {
            console.log("getRequest: " + error);
        }        
    }

    /**
     * This function to post data in List or Library on 'POST' request.
     * @param url 
     * @param postBody 
     * @param xMethod 
     */
    public _postRequest(url: string, postBody, xMethod): any {
        try {
            //GET FormDigestValue
            return this._getValues().then((token) => {
                return fetch(url, {
                    headers: {
                        Accept: 'application/json;odata=verbose',
                        "Content-Type": 'application/json;odata=verbose',
                        "X-RequestDigest": token.d.GetContextWebInformation.FormDigestValue,
                        "X-Http-Method": xMethod,
                        'IF-MATCH': '*'
                    },
                    method: 'POST',
                    body: postBody,
                    credentials: "same-origin"
                }).then((response) => {
                    if (response.status <= 204 && response.status >= 200) {
                        //resolve(response);
                        if (response.status == 204 || xMethod == 'DELETE') {
                            return 'success';
                        } else {
                            return response.json();
                        }
                    }
                    else {
                        return response.json();
                    }
                }, (err) => {
                    console.log(err);
                }).catch(error =>
                    console.error("postRequest: " + error));
            }, (error) => {
                console.log("getValues => postRequest" + error);
            });
        } catch (error) {
            console.log("postRequest: " + error);
        }
    }

    /**
     *This function is to get FormDigestValue for X-RequestDigest. 
    */
    public _getValues(): any {
        try {
            var url = `${this.TenantUrl}/_api/contextinfo`;
            return fetch(url, {
                method: "POST",
                headers: { Accept: "application/json;odata=verbose" },
                credentials: "same-origin"
            }).then((response) => {
                return response.json();
            });
        } catch (error) {
            console.log("getValues: " + error);
        }
    }

    /**
     * Giving READ permission to Everyone of site collection.
     * @param LoginName 
     */
    public async _ensureUser(siteGrpId) {
        try {
            var payload = JSON.stringify({
                '__metadata': {
                    'type': 'SP.User'
                },
                'LoginName': constant.everyOneLoginName
            });
            var sitegrpURL = `${this.TenantUrl}/_api/web/sitegroups(${siteGrpId})/users`;            
            await this._postRequest(sitegrpURL, payload, 'POST').then((data) => {
                return data;
            });
        } catch (error) {
            console.log('ensureUser : ' + error);
        }
    }

    /**
     * Check wheather user is admin or not.
     */
    public _checkUserIsSiteAdmin() {
        try {
            var userIsAdminURL = `${this.tenantURL()}/_api/web/currentUser/issiteadmin`;
            return this._getRequest(userIsAdminURL).then((responseData) => {
                return responseData.d.IsSiteAdmin;
            });
        } catch (error) {
            console.log('_checkUserIsSiteAdmin: ' + error);
        }
    }

    /**
     * Create New Team Site.
     * @param context 
     */
    public async _createNewTeamSite(context) {
        try {
            var client = context.spHttpClient;
            let opt = {
                headers: {
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(constant.postBody),
                credentials: "same-origin"
            };
            var newTeamSiteUrl = `${this.tenantURL()}/_api/GroupSiteManager/CreateGroupEx`;
            return client.post(newTeamSiteUrl, SPHttpClient.configurations.v1, opt)
                .then((response: SPHttpClientResponse) => {
                    return response.json();
                }).then(result => {
                    return result;
                });
        } catch (error) {
            console.log('createNewTeamSite : ' + error);
        }
    }

    /**
     * Give FullControl permission to Everyone user in DocsnodeText list.
     */
    public _addFullControlPermission() {
        try {
            var endPointUrl = `${this.TenantUrl}/_api/web/lists/getByTitle('${constant.DocsNodeTextName}')/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`;
            return this._postRequest(endPointUrl, '', 'POST').then(() => {
                return this.getGroupID().then((grpID) => {
                    if (grpID != null) {
                        var endPointUrlRoleAssignment = `${this.TenantUrl}/_api/web/lists/getByTitle('${constant.DocsNodeTextName}')/roleassignments/addroleassignment(principalid=${grpID},roleDefId=1073741829)`;
                        return this._postRequest(endPointUrlRoleAssignment, '', 'POST').then(() => {
                            this._ensureUser(grpID).then((data) => {
                                return data;
                            });
                        });
                    } else {
                        return grpID;
                    }
                });
            });
        } catch (error) {
            console.log('_addFullControlPermission :' + error);
        }
    }

    /**
     * Create new site group in newly created Team site.
     */
    public async _createNewSiteGrp() {
        try {
            var siteGrpUrl = `${this.TenantUrl}/_api/Web/SiteGroups`;
            var JsonString = JSON.stringify({
                __metadata: { 'type': 'SP.Group' },
                Title: constant.newSiteGroupName
            });
            await this._postRequest(siteGrpUrl, JsonString, 'POST').then((data) => {
                return data;
            });
        } catch (error) {
            console.log('_createNewSiteGrp :' + error);
        }
    }

    /**
     * Get New site group ID.
     */
    public getGroupID() {
        try {
            var grpIDURL = `${this.TenantUrl}/_api/web/sitegroups?$select=id,*&$filter=LoginName eq '${constant.newSiteGroupName}'`;
            return this._getRequest(grpIDURL).then((responseData) => {
                var resultData = responseData.d.results;
                if (resultData.length > 0) {
                    return resultData[0].Id;
                } else {
                    return null;
                }
            });
        } catch (error) {
            console.log('getUserID : ' + error);
        }
    }

    /**
     * Add Item Level Read Permission to Everyone User.
     * @param itemID 
     */
    public _addItemLevelPermission(itemID) {
        try {
            var endPointUrl = `${this.TenantUrl}/_api/web/lists/getByTitle('${constant.DocsNodeTextName}')/items(${itemID})/breakroleinheritance(copyRoleAssignments=true, clearSubscopes=true)`;
            return this._postRequest(endPointUrl, '', 'POST').then(() => {
                return this.getGroupID().then((grpID) => {
                    if (grpID != null) {
                        var removePermissionURL = `${this.TenantUrl}/_api/web/lists/getByTitle('${constant.DocsNodeTextName}')/items(${itemID})/roleassignments/getbyprincipalid(${grpID})`;
                        return this._postRequest(removePermissionURL, '', 'DELETE').then(() => {
                            return this.getUserID().then((userID) => {
                                var endPointUrlRoleAssignment = `${this.TenantUrl}/_api/web/lists/getByTitle('${constant.DocsNodeTextName}')/items(${itemID})/roleassignments/addroleassignment(principalid=${userID},roleDefId=1073741826)`;
                                return this._postRequest(endPointUrlRoleAssignment, '', 'POST').then((responsedata) => {
                                    return responsedata;
                                });
                            });
                        });
                    } else {
                        console.log('item level permission not assigned');
                        return grpID;
                    }
                });
            });
        } catch (error) {
            console.log('_addItemLevelPermission :' + error);
        }
    }

    /**
     * Get Everyone user Id.
     */
    public getUserID() {
        try {
            var userIDURL = `${this.TenantUrl}/_api/web/siteusers?$select=id,*&$filter=LoginName eq '${constant.everyOneLoginName}'`;
            return this._getRequest(userIDURL).then((responseData) => {
                var resultData = responseData.d.results;
                return resultData[0].Id;
            });
        } catch (error) {
            console.log('getUserID : ' + error);
        }
    }

    /**
     * Remove Site group.
     * @param id 
     */
    public async _removeNewSiteGrp(id) {        
        var siteGrpUrl = `${this.TenantUrl}/_api/Web/SiteGroups/removebyid(${id})`;
        await this._postRequest(siteGrpUrl, '', 'POST').then((data) => {
            return data;
        });
    }
}