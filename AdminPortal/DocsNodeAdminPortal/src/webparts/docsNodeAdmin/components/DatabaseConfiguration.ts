//Importing files and objects creation
import CommonUtility from "./CommonUtility";
import constant from "./Constant";
import { ISPHttpClientOptions, SPHttpClientResponse, SPHttpClient } from "@microsoft/sp-http";

/**
 *Creating object of existing files. */
const CU: CommonUtility = new CommonUtility();

/**
 *Array for List items is to be inserted by default in configuration list. */
let dbConfigArrData = [];
dbConfigArrData.push({ assetTitle: constant.TemplateTitle, sourceLocation: constant.DefaultDocsNodeTemplatesLibraryName, sourceListURL: '', sourceListGUID: '' });
dbConfigArrData.push({ assetTitle: constant.ImageTitle, sourceLocation: constant.DefaultDocsNodePictureName, sourceListURL: '', sourceListGUID: '' });
dbConfigArrData.push({ assetTitle: constant.SlideTitle, sourceLocation: constant.DocsNodeSlidesName, sourceListURL: '', sourceListGUID: '' });
dbConfigArrData.push({ assetTitle: constant.TextSnippetTitle, sourceLocation: constant.DocsNodeTextName, sourceListURL: '', sourceListGUID: '' });
dbConfigArrData.push({ assetTitle: constant.CategoryTitle, sourceLocation: constant.DocsNodeCategoriesName, sourceListURL: '', sourceListGUID: '' });

export default class DatabaseConfiguration {

    /**
     * This function return array to create default List and Library on webpart load. */
    public _defaultListOrLibraryArray(): any {
        let dbConfigArr = [];
        dbConfigArr.push({ listname: constant.DocsNodeConfigurationName, type: constant.NewList, BaseTemplate: 100 });
        dbConfigArr.push({ listname: constant.DocsNodeCategoriesName, type: constant.NewList, BaseTemplate: 100 });
        dbConfigArr.push({ listname: constant.DocsNodePlaceHolderName, type: constant.NewList, BaseTemplate: 100 });
        dbConfigArr.push({ listname: constant.DocsNodeSlidesName, type: constant.NewLibrary, BaseTemplate: 101 });
        dbConfigArr.push({ listname: constant.DocsNodeTextName, type: constant.NewList, BaseTemplate: 100 });
        dbConfigArr.push({ listname: constant.DefaultDocsNodePictureName, type: constant.NewLibrary, BaseTemplate: 109 });
        dbConfigArr.push({ listname: constant.DocsNodeProductLogoName, type: constant.NewLibrary, BaseTemplate: 109 });
        dbConfigArr.push({ listname: constant.DefaultDocsNodeTemplatesLibraryName, type: constant.NewLibrary, BaseTemplate: 101 });
        //dbConfigArr.push({ listname: constant.DocsNodeSiteConfiguratonName, type: constant.NewList, BaseTemplate: 100 });
        return dbConfigArr;
    }

    /**
     * This function gets the data from configuration list. */
    public async _dynamicListOrLibraryArray() {
        try {
            //Get request from configuration list
            await this._getDocsNodeConfigurationName().then(async (responseData) => {
                let dynConfigDataArr = [];
                if (responseData.length > 0) {
                    //Binding configuration list data
                    dynConfigDataArr = this._bindingConfigData(responseData);
                    //Check the All Dynamic List and Library are exist
                    //If not, then this would create  new one
                    if(dynConfigDataArr.length > 0){
                        await this._checkListExistsOrNot(dynConfigDataArr).then(async (data) => {
                            //Check the columns of All Dynamic List created are exist 
                            //If not, then this would create columns for that List or Library
                            await this._checkForColumnExistence(dynConfigDataArr);
                        });
                    }else{
                        console.log('dynConfigDataArr : no items found');
                    }                   
                } else {
                    console.log('_dynamicListOrLibraryArray: no items found');
                }
            });
        } catch (error) {
            console.log('_dynamicListOrLibraryArray: ' + error);
        }
    }

    /**
     * Binding configuration list data.
     * @param responseData 
     */
    public _bindingConfigData(responseData) {
        let dynConfigArr = [];
        if(responseData.length > 0){
            for (var i = 0; i < responseData.length; i++) {
                var singleResultData = responseData[i];
                if (singleResultData.assetTitle == constant.TemplateTitle) {
                    constant.DocsNodeTemplatesLibraryName = singleResultData.sourceLocation;
                    dynConfigArr.push({ listname: singleResultData.sourceLocation, type: constant.NewLibrary, BaseTemplate: 101 });
                }            
                if (singleResultData.assetTitle == constant.ImageTitle) {
                    constant.DocsNodePictureName = singleResultData.sourceLocation;
                    dynConfigArr.push({ listname: singleResultData.sourceLocation, type: constant.NewLibrary, BaseTemplate: 109 });
                }            
            }
            return dynConfigArr;
        }else{
            return dynConfigArr;
        }        
    }

    /**
     * Check weather List or Library is exist or Not.
     * @param dbConfigArr 
     */
    public async _checkListExistsOrNot(dbConfigArr) {
        try {
            if(dbConfigArr.length > 0){
                var filterLstName = '';
                filterLstName += '(';
                for(var j = 0; j < dbConfigArr.length; j++){
                    var lstTitle = dbConfigArr[j].listname;                                      
                    if(j ==  dbConfigArr.length - 1){
                        filterLstName += `(Title eq '${lstTitle}')`;
                    }else{
                        filterLstName += `(Title eq '${lstTitle}') or`;
                    }                   
                }
                filterLstName += ')';
                var listURL = `${CU.TenantUrl}/_api/Web/Lists?$filter=${filterLstName}`;
                await CU._getRequest(listURL)
                        .then(async (data: any) => {
                            var resultData = data.d.results;
                            if (resultData.length > 0) {
                                for(var i = 0; i < dbConfigArr.length; i++){
                                    var lstObject = dbConfigArr[i];
                                    var listName = dbConfigArr[i].listname;
                                    var lstFlag = false;
                                    for(var k = 0; k < resultData.length; k++){
                                        var lstName = resultData[k].Title;
                                        if(lstName == listName){                                            
                                            lstFlag = true;
                                        }                                        
                                    }
                                    if(lstFlag == false){
                                        await this._creatingListOrLibrary(lstObject);
                                    }
                                }                                
                                if (lstName != constant.DocsNodeSlidesName && lstName != constant.DefaultDocsNodePictureName 
                                    && lstName != constant.DocsNodeProductLogoName && lstName != constant.DefaultDocsNodeTemplatesLibraryName 
                                    && lstName != constant.DocsNodeTextName && lstName != constant.DocsNodeCategoriesName 
                                    && lstName != constant.DocsNodeConfigurationName && lstName != constant.DocsNodePlaceHolderName) {
                                    
                                }
                            }
                            else {
                                for(var x = 0; x < dbConfigArr.length; x++){
                                    var listObject = dbConfigArr[x];                             
                                    await this._creatingListOrLibrary(listObject);
                                }
                            }
                        });
                return ("Success");
            }else{
                return ("Fail");
            }            
        } catch (error) {
            console.log("checkListExistsOrNot: " + error);
            return ("Fail");
        }
    }

    public async _creatingListOrLibrary(arry){
        var spDefaultMetadata = {
            List: JSON.stringify({
                BaseTemplate: arry.BaseTemplate,
                __metadata: { type: "SP.List" },
                Title: arry.listname
            }),
            Document: JSON.stringify({
                __metadata: { type: "SP.List" },
                AllowContentTypes: true,
                BaseTemplate: arry.BaseTemplate,
                ContentTypesEnabled: true,
                Title: arry.listname
            })
        };
        if (arry.type == constant.NewList) {
            //Creating New List
            await this._createNewListOrLibrary(spDefaultMetadata.List);
            //Giving full control to everyone user in DocsnodeText list                                
            if(arry.listname == constant.DocsNodeTextName){
                await CU._addFullControlPermission(); 
            }
        }
        else {
            //Creating New Library
            await this._createNewListOrLibrary(spDefaultMetadata.Document);
        }
    }

    /**
     * Creating new List or Library.
     * @param listDetails 
     */
    public async _createNewListOrLibrary(listDetails: any) {
        try {
            var newListOrLibraryUrl = `${CU.TenantUrl}/_api/web/lists/`;         
            await CU._postRequest(newListOrLibraryUrl, listDetails, '')
                .then(() => {                             
                });
        } catch (error) {
            console.log("createNewList: " + error);
        }
    }

    /**
     * Insert data into configuration list
     * @param context 
     */
    public async _insertConfigData(context) {
        try {
            //Check if items exist or not in configuration list
            var isRecordExist = await this._checkForConfigDataExistence();
            if (!isRecordExist) {
                //Upload default product logo
                await this._uploadProdLOGO(context);
                for (var i = 0; i < dbConfigArrData.length; i++) {
                    var resultData = dbConfigArrData[i];
                    resultData.sourceListGUID = await this._getListGUID(resultData.sourceLocation);
                    resultData.sourceListURL = CU.siteCollectionPath;                    
                    var commomJSON = '';
                    var addConfigDataUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeConfigurationName}')/items`;
                    commomJSON = JSON.stringify({
                        __metadata: { 'type': await this._getListItemEntityTypeFullName(constant.DocsNodeConfigurationName) },
                        Title: resultData.assetTitle,
                        ConfigAssestTitle: resultData.assetTitle,
                        ConfigSourceList: resultData.sourceLocation,
                        ConfigSourceListPath: resultData.sourceListURL,
                        ConfigSourceListGUID: resultData.sourceListGUID
                    });                    
                    await CU._postRequest(addConfigDataUrl, commomJSON, '');
                }
            }
        }
        catch (error) {
            console.log("insert config data error: " + error);
        }
    }

    /**
     * Check the Items exist or not in configuration List. */
    public async _checkForConfigDataExistence() {
        try{
            var checkConfigDataUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeConfigurationName}')/items`;
            //Get request
            return CU._getRequest(checkConfigDataUrl).then((responseData) => {
                if (responseData.d.results.length > 0)
                    return true;
                else
                    return false;
            });
        }catch(error){
            console.log('_checkForConfigDataExistence :'+ error);
        }        
    }

    /**
     * Upload product logo to Site assest list of site collection
     * @param context 
     */
    public async _uploadProdLOGO(context) {
        var logo = String(require('../images/logo.png'));
        var itemURL = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeProductLogoName}')/items?$select=FileRef,ID,${constant.InternalLinkFilename}&$filter=substringof('ProductLogo',FileRef)`;
        //Get request
        await CU._getRequest(itemURL).then(async (responseData) => {
            if (responseData.d.results.length == 0) {
                await fetch(logo)
                    .then(res => res.blob())
                    .then(async blob => {
                        const file = new File([blob], 'ProductLogo.png', blob);
                        //Post request for uploading product logo
                        await this._uploadProductLogo(file, context);
                        //insert data here
                        var commomJSON = '';
                        var addLogoItemUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeConfigurationName}')/items`;
                        commomJSON = JSON.stringify({
                            __metadata: { 'type': await this._getListItemEntityTypeFullName(constant.DocsNodeConfigurationName) },
                            Title: 'Product Logo',
                            ConfigAssestTitle: 'Product Logo',
                            ConfigSourceList: constant.DocsNodeProductLogoName,
                            ConfigSourceListPath: CU.siteCollectionPath,
                            ConfigSourceListGUID: await this._getListGUID(constant.DocsNodeProductLogoName)
                        });
                        await CU._postRequest(addLogoItemUrl, commomJSON, '');
                    });
            }
        }, (error) => {
            console.log(error);
        }).catch((error) => {
            console.log('_uploadProdLOGO :' + error);
        });
    }
    
    /**
     * Creating Column Array for List and Library.
     * @param dbConfigArr 
     */
    public async _checkForColumnExistence(dbConfigArr) {
        try {
            for (var i = 0; i < dbConfigArr.length; i++) {
                var listDisplayName = dbConfigArr[i].listname;
                switch (listDisplayName) {
                    case constant.DocsNodeCategoriesName:
                        let docsNodeCatColm = [];
                        docsNodeCatColm.push({ listName: listDisplayName, columnName: constant.CategoryName, FieldTypeKind: 2, EnforceUniqueValues: true, Indexed: true });
                        docsNodeCatColm.push({ listName: listDisplayName, columnName: constant.CategoryParentId, FieldTypeKind: 9, EnforceUniqueValues: false, Indexed: false });
                        docsNodeCatColm.push({ listName: listDisplayName, columnName: constant.CategoryType, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodeCatColm.push({ listName: listDisplayName, columnName: constant.CategoryLevel, FieldTypeKind: 9, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeCatColm);
                        break;
                    case constant.DocsNodeSlidesName:
                        let docsNodeSlideColm = [];
                        docsNodeSlideColm.push({ listName: listDisplayName, columnName: constant.SlidesCategoryName, FieldTypeKind: 7, EnforceUniqueValues: false, Indexed: false });
                        docsNodeSlideColm.push({ listName: listDisplayName, columnName: constant.SlidesDiscriptionName, FieldTypeKind: 3, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeSlideColm);
                        break;
                    case constant.DocsNodePictureName:
                        let docsNodeImgColm = [];
                        docsNodeImgColm.push({ listName: listDisplayName, columnName: constant.ImageCategoryName, FieldTypeKind: 7, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeImgColm);
                        break;
                    case constant.DocsNodeTextName:
                        let docsNodeTxtColm = [];
                        docsNodeTxtColm.push({ listName: listDisplayName, columnName: constant.TextSnippetName, FieldTypeKind: 3, EnforceUniqueValues: false, Indexed: false });
                        docsNodeTxtColm.push({ listName: listDisplayName, columnName: constant.TextCategoryName, FieldTypeKind: 7, EnforceUniqueValues: false, Indexed: false });
                        docsNodeTxtColm.push({ listName: listDisplayName, columnName: constant.TextSnippetDiscriptionName, FieldTypeKind: 3, EnforceUniqueValues: false, Indexed: false });                        
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeTxtColm);
                        break;
                    case constant.DocsNodePlaceHolderName:
                        let docsNodePlcholdColm = [];
                        docsNodePlcholdColm.push({ listName: listDisplayName, columnName: constant.PlaceHolderName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodePlcholdColm.push({ listName: listDisplayName, columnName: constant.SiteCollectionUrlName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodePlcholdColm.push({ listName: listDisplayName, columnName: constant.SubSiteUrlName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodePlcholdColm.push({ listName: listDisplayName, columnName: constant.ListUrlName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodePlcholdColm.push({ listName: listDisplayName, columnName: constant.ListFieldName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodePlcholdColm.push({ listName: listDisplayName, columnName: constant.PlaceHolderDiscrip, FieldTypeKind: 3, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodePlcholdColm);
                        break;
                    case constant.DocsNodeConfigurationName:
                        let docsNodeConfigColm = [];
                        docsNodeConfigColm.push({ listName: listDisplayName, columnName: constant.ConfigAssetTitleName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodeConfigColm.push({ listName: listDisplayName, columnName: constant.ConfigSourceListName, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodeConfigColm.push({ listName: listDisplayName, columnName: constant.ConfigSourceListPathUrl, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        docsNodeConfigColm.push({ listName: listDisplayName, columnName: constant.ConfigSourceListGUID, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                        //Check Column exist in List or Library
                        await this._checkColumn(docsNodeConfigColm);
                        break;
                    // case constant.DocsNodeSiteConfiguratonName:
                    //     let docsNodeSiteConfigColm = [];
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.SiteCollectionNameColumn, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.SubSiteNameColumn, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.DocumentLibraryNameColumn, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.FolderNameColumn, FieldTypeKind: 2, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.SiteCollectionUrlColumn, FieldTypeKind: 11, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.subsiteUrlColumn, FieldTypeKind: 11, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.DocumentLibraryURLColumn, FieldTypeKind: 11, EnforceUniqueValues: false, Indexed: false });
                    //     docsNodeSiteConfigColm.push({ listName: listDisplayName, columnName: constant.FolderURLColumn, FieldTypeKind: 11, EnforceUniqueValues: false, Indexed: false });
                    //     //Check Column exist in List or Library
                    //     await this._checkColumn(docsNodeSiteConfigColm);
                    //     break;
                    default:
                        break;
                }
            }
        } catch (error) {
            console.log("checkForColumnExistence: " + error);
        }
    }

    /**
    * Check wheather Column exist or not in List or Library.
    * @param columnData 
    */
    public async _checkColumn(columnData: any) {
        try {
            if(columnData.length > 0){                
                var filter = '';
                filter += '(';
                for (var i = 0; i < columnData.length; i++) {
                    var columnList = columnData[i];
                    if(i == columnData.length - 1){
                        filter += `(InternalName eq '${columnList.columnName}')`;
                    }else{
                        filter += `(InternalName eq '${columnList.columnName}') or `;
                    }                    
                }
                filter += ')'; 
                var fieldURL = `${CU.tenantURL()}${CU.siteCollectionPath}/_api/Web/Lists/GetByTitle('${columnData[0].listName}')/Fields?select=Title&$filter=${filter}`;
                    //GET request
                    await CU._getRequest(fieldURL)
                        .then(async (data: any) => {
                            var resultData = data.d.results;
                            for (var k = 0; k < columnData.length; k++) {
                                var coluData = columnData[k];
                                var flag = false;
                                if (resultData.length > 0) {
                                    for (var j = 0; j < resultData.length; j++) {
                                        var columnResult = resultData[j];
                                        if (columnResult.InternalName == coluData.columnName) {                                          
                                            flag = true;
                                        }
                                    }
                                    if (flag == false) {
                                        //Creating new column
                                        await this._createNewColumn(coluData);
                                    }
                                }
                                else {
                                    await this._createNewColumn(coluData);
                                }
                            }
                        });
            }            
        } catch (error) {
            console.log("checkColumn:" + error);
        }
    }

    /**
     * Create New column for List or Library.
     * @param columnData 
     */
    public async _createNewColumn(columnData: any) {
        try {
            var colData = null;
            colData = JSON.stringify({
                __metadata: { 'type': 'SP.Field' },
                Description: 'Created From DocsNode',
                Title: columnData.columnName,
                FieldTypeKind: columnData.FieldTypeKind,
                EnforceUniqueValues: columnData.EnforceUniqueValues,
                Indexed: columnData.Indexed
            });
            if (columnData.columnName == constant.SiteCollectionUrlColumn || columnData.columnName == constant.SubSiteUrlName
                || columnData.columnName == constant.DocumentLibraryURLColumn || columnData.columnName == constant.FolderURLColumn) {
                colData = JSON.stringify({
                    __metadata: { 'type': 'SP.FieldUrl' },
                    FieldTypeKind: columnData.FieldTypeKind,
                    Title: columnData.columnName,
                    DisplayFormat: 1
                });
            }
            if (columnData.FieldTypeKind != 7) {
                var columnFieldURL = `${CU.TenantUrl}/_api/web/lists/GetByTitle('${columnData.listName}')/Fields`;
                //POST request
                await CU._postRequest(columnFieldURL, colData, '')
                    .then(() => {                                             
                    }, (error) => {
                        console.log(error);
                    });
            } else {
                //Creating lookup column in List or Library
                await this._createLookedUpCol(columnData);
            }
        } catch (error) {
            console.log("createNewColumn: " + error);
        }
    }

    /**
     * Create New lookup column in List or Library.
     * @param columnData 
     */
    public async _createLookedUpCol(columnData: any) {
        try {
            //Get GUID of Existing List or Library
            await this._getListGUID(constant.DocsNodeCategoriesName).then(async (listId) => {
                var JSONVAR = "{ 'parameters': { '__metadata': { 'type': 'SP.FieldCreationInformation' }, 'FieldTypeKind': 7,'Title': '" + columnData.columnName + "', 'LookupListId': '" + listId + "' ,'LookupFieldName': '" + constant.CategoryName + "' } }";
                var addColFieldURL = `${CU.TenantUrl}/_api/web/lists/GetByTitle('${columnData.listName}')/Fields/addfield`;
                //POST request
                await CU._postRequest(addColFieldURL, JSONVAR, '')
                    .then(() => {                        
                    },(err)=>{
                        console.log('_createLookedUpCol _postRequest:'+ err);
                    });
            });
        } catch (error) {
            console.log('createLookedUpCol : ' + error);
        }
    }

    /**
     * Get GUID of List or Library.
     * @param listName 
     */
    public _getListGUID(listName) {
        try {
            var guidURL = `${CU.TenantUrl}/_api/web/lists?$filter=title eq ('${listName}')`;
            return CU._getRequest(guidURL).then((data) => {
                return data.d.results[0].Id;
            }).catch((error) => {
                console.log("_getListGUID after fetch: " + error);
            });
        }
        catch (error) {
            console.log("_getListGUID: " + error);
        }
    }

    /**
     *This function is use get all items from DocsNodeSlide Library. */
    public _getDocsNodeSlidesName() {
        try {
            var DocsNodeSlidesArrayItems = [];
            var getSlideURL = `${CU.TenantUrl}/_api/web/Lists/getbytitle('${constant.DocsNodeSlidesName}')/items?$select=ID,${constant.InternalLinkFilename},DocIcon,${constant.SlidesCategoryName}/${constant.Title},${constant.SlidesCategoryName}/${constant.CategoryName},FileLeafRef,FileRef,ContentTypeId,ContentType/Id,ContentType/Name,*&$expand=${constant.SlidesCategoryName},File,ContentType`;
            //GET request
            return CU._getRequest(getSlideURL)
                .then((ResponseData) => {
                    var resultData = ResponseData.d.results;
                    if (resultData.length > 0) {
                        for (var i = 0; i < resultData.length; i++) {
                            var results = resultData[i];
                            if (results.ContentType.Name != 'Folder' && results.DocIcon == 'pptx') {
                                DocsNodeSlidesArrayItems.push({ Key: results.ID, Title: results[constant.Title], FileName: results[constant.InternalLinkFilename], Category: (results[constant.SlidesCategoryName][constant.CategoryName]  != null ? results[constant.SlidesCategoryName][constant.CategoryName] : ''), FileType: this._getIcon(results.DocIcon), LinkURL: results.File.LinkingUri });
                            }
                        }
                    }
                    return DocsNodeSlidesArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodeSlidesName inside getRequest : ' + responseError);
                    return DocsNodeSlidesArrayItems;
                });
        }
        catch (error) {
            console.log('getDocsNodeSlidesName : ' + error);
        }
    }

    /**
     *This function is use get all items from DocsNodeTextSnippet List. */
    public _getDocsNodeTextSnippetName() {
        try {
            var DocsNodeTextSnippetArrayItems = [];
            var getTxtSnippetURL = `${CU.TenantUrl}/_api/web/Lists/getbytitle('${constant.DocsNodeTextName}')/items?$select=ID,${constant.Title},${constant.TextSnippetName},${constant.TextCategoryName}/${constant.Title},${constant.TextCategoryName}/${constant.CategoryName}&$expand=${constant.TextCategoryName}`;
            //GET request
            return CU._getRequest(getTxtSnippetURL)
                .then((ResponseData) => {
                    var resultData = ResponseData.d.results;
                    if (resultData.length > 0) {
                        for (var i = 0; i < resultData.length; i++) {
                            var results = resultData[i];
                            DocsNodeTextSnippetArrayItems.push({ Key: results.ID, Title: results[constant.Title], FileName: results[constant.Title], Category: (results[constant.TextCategoryName][constant.CategoryName] != undefined ? results[constant.TextCategoryName][constant.CategoryName]  : ''), FileType: '' });
                        }
                    }
                    return DocsNodeTextSnippetArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodeTextSnippetName inside getRequest : ' + responseError);
                    return DocsNodeTextSnippetArrayItems;
                });
        }
        catch (error) {
            console.log('getDocsNodeTextSnippetName : ' + error);
        }
    }

    /**
     *This function is use get all items from DocsNodePicture Library. */
    public _getDocsNodePictureName() {
        try {
            var DocsNodePictureArrayItems = [];
            var getImageURL = `${CU.TenantUrl}/_api/web/Lists/getbytitle('${constant.DocsNodePictureName}')/items?$select=ID,Title,LinkFilenameNoMenu,DocIcon,EncodedAbsUrl,${constant.ImageCategoryName}/${constant.Title},${constant.ImageCategoryName}/${constant.CategoryName}&$expand=${constant.ImageCategoryName}`;
            //GET request
            return CU._getRequest(getImageURL)
                .then((ResponseData) => {
                    var resultData = ResponseData.d.results;
                    if (resultData.length > 0) {
                        for (var i = 0; i < resultData.length; i++) {
                            var results = resultData[i];
                            DocsNodePictureArrayItems.push({ Key: results.ID, Title: results[constant.Title], FileName: results.LinkFilenameNoMenu, Category: (results[constant.ImageCategoryName][constant.CategoryName] != undefined ? results[constant.ImageCategoryName][constant.CategoryName] : ''), FileType: this._getIcon(results.DocIcon), LinkURL: results.EncodedAbsUrl });
                        }
                    }
                    return DocsNodePictureArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodePictureName inside getRequest : ' + responseError);
                    return DocsNodePictureArrayItems;
                });
        } catch (error) {
            console.log('getDocsNodePictureName : ' + error);
        }
    }

    /**
     *This function is use get all items from DocsNodePlaceHolder List. */
    public _getDocsNodePlaceHolderName() {
        try {
            var DocsNodePlaceHolderArrayItems = [];
            var getPlaceholderURL = `${CU.TenantUrl}/_api/web/Lists/getbytitle('${constant.DocsNodePlaceHolderName}')/items?$select=ID,*`;
            //GET request
            return CU._getRequest(getPlaceholderURL)
                .then((ResponseData) => {
                    var resultData = ResponseData.d.results;
                    if (resultData.length > 0) {
                        for (var i = 0; i < resultData.length; i++) {
                            var result = resultData[i];
                            DocsNodePlaceHolderArrayItems.push({ Key: result.ID, Title: result[constant.Title], PlaceholderName: result[constant.PlaceHolderName], FileName: result[constant.PlaceHolderName] });
                        }
                    }
                    return DocsNodePlaceHolderArrayItems;
                }, (responseError) => {
                    console.log('getDocsNodePlaceHolderName inside getRequest : ' + responseError);
                    return DocsNodePlaceHolderArrayItems;
                });
        } catch (error) {
            console.log('getDocsNodePlaceHolderName : ' + error);
        }
    }

    /**
     *This function is use get all items from DocsNodeConfiguration Library. */
    public _getDocsNodeConfigurationName() {
        try {
            var DocsNodeConfigArrayItems = [];
            var getImageURL = `${CU.TenantUrl}/_api/web/Lists/getbytitle('${constant.DocsNodeConfigurationName}')/items?$select=ID,*`;
            //GET request
            return CU._getRequest(getImageURL)
                .then((ResponseData) => {
                    var resultData = ResponseData.d.results;
                    if (resultData.length > 0) {
                        for (var i = 0; i < resultData.length; i++) {
                            var singleResultData = resultData[i];
                            DocsNodeConfigArrayItems.push({ Key: singleResultData.ID, assetTitle: singleResultData[constant.ConfigAssetTitleName], sourceLocation: singleResultData[constant.ConfigSourceListName], sourceListGUID: singleResultData[constant.ConfigSourceListGUID], sourceListURL: singleResultData[constant.ConfigSourceListPathUrl] });
                        }
                    }
                    return DocsNodeConfigArrayItems;
                }, (responseError) => {
                    console.log('_getDocsNodeConfigurationName inside getRequest : ' + responseError);
                    return DocsNodeConfigArrayItems;
                });
        } catch (error) {
            console.log('_getDocsNodeConfigurationName : ' + error);
        }
    }

    /**
     * This function is use get all items from DocsNodeCategory List.
     * @param filterCategory 
     */
    public _getDocsNodeCategoriesName(filterCategory) {
        try {
            var DocsNodeCategoriesArrayItems = [];
            var DocsNodeParentCategoriesArrayItems = [];
            var DocsNodeCategoriesItemsData = [];
            var filter = '';
            if (filterCategory != '') {
                filter = `${constant.CategoryType} eq '${filterCategory}'`;
            } else {
                filter = `${constant.CategoryType} ne 'null'`;
            }
            var getCategoryURL = `${CU.tenantURL()}${CU.siteCollectionPath}/_api/web/Lists/getbytitle('${constant.DocsNodeCategoriesName}')/items?$select=ID,${constant.InternalLinkFilename},DocIcon,Title,*&$filter=${filter}`;
            //GET request
            return CU._getRequest(getCategoryURL)
                .then(async (ResponseData) => {
                    var _catArray = [];
                    var resultData = ResponseData.d.results;
                    if(filterCategory == ''){
                        var slideCategoryUrl = `${CU.tenantURL()}${CU.siteCollectionPath}/_api/web/Lists/getbytitle('${constant.DocsNodeSlidesName}')/items?$select=ID,Title,${constant.SlidesCategoryName}/${constant.Title},${constant.SlidesCategoryName}/${constant.CategoryName}&$expand=${constant.SlidesCategoryName},File`;
                        await CU._getRequest(slideCategoryUrl).then((slideData) => {
                            var slideResponseData = slideData.d.results;
                            if (slideResponseData.length > 0) {
                                for (var g = 0; g < slideResponseData.length; g++) {
                                    var slideResults = slideResponseData[g];
                                    _catArray.push(slideResults.SlidesCategory.Category != null ? slideResults.SlidesCategory.Category : '');
                                }
                            }
                        });
                        var imageCategoryUrl = `${CU.tenantURL()}${CU.siteCollectionPath}/_api/web/Lists/getbytitle('${constant.DocsNodePictureName}')/items?$select=ID,Title,${constant.ImageCategoryName}/${constant.Title},${constant.ImageCategoryName}/${constant.CategoryName}&$expand=${constant.ImageCategoryName}`;
                        await CU._getRequest(imageCategoryUrl).then((imageData) => {
                            var imageResponseData = imageData.d.results;
                            if (imageResponseData.length > 0) {
                                for (var e = 0; e < imageResponseData.length; e++) {
                                    var imageResults = imageResponseData[e];
                                    _catArray.push((imageResults.ImageCategory.Category != undefined ? imageResults.ImageCategory.Category : ''));
                                }
                            }
                        });
                        var txtCategoryUrl = `${CU.tenantURL()}${CU.siteCollectionPath}/_api/web/Lists/getbytitle('${constant.DocsNodeTextName}')/items?$select=ID,Title,${constant.TextCategoryName}/${constant.Title},${constant.TextCategoryName}/${constant.CategoryName}&$expand=${constant.TextCategoryName}`;
                        await CU._getRequest(txtCategoryUrl).then((textData) => {
                            var textSniResponseData = textData.d.results;
                            if (textSniResponseData.length > 0) {
                                for (var f = 0; f < textSniResponseData.length; f++) {
                                    var txtResults = textSniResponseData[f];
                                    _catArray.push(txtResults.TextCategory.Category != null ? txtResults.TextCategory.Category : '');
                                }
                            }
                        });
                    }                    
                    if (resultData.length > 0) {
                        for (var j = 0; j < resultData.length; j++) {
                            var rsltData = resultData[j];
                            if (rsltData.ParentCategory == null) {
                                DocsNodeCategoriesArrayItems.push({ key: rsltData.ID, text: rsltData.Category, data: false });
                                resultData.filter((item) => {
                                    if (item.ParentCategory == rsltData.ID) {
                                        DocsNodeCategoriesArrayItems.push({ key: item.ID, text: item.Category, data: true });
                                    }
                                });
                            }
                        }
                        if(filterCategory == ''){
                            for (var i = 0; i < resultData.length; i++) {
                                var responseResultData = ResponseData.d.results[i];
                                if (responseResultData.ParentCategory == null) {
                                    DocsNodeParentCategoriesArrayItems.push({ key: responseResultData.ID, text: responseResultData.Category, CategoryType: responseResultData.CategoryType });
                                }
                            }                        
                            resultData.map((item) => {
                                var flag = false;
                                DocsNodeParentCategoriesArrayItems.map((itemParent) => {
                                    if (itemParent.key == item.ParentCategory) {
                                        DocsNodeCategoriesItemsData.push({ Key: item.ID, FileName: item.Category, Title: item.Category, CategoryType: (item.CategoryType != undefined ? item.CategoryType : null), ParentCategory: itemParent.text, Category: (item.Category != undefined ? item.Category : null), FileType: this._getIcon(item.DocIcon) });
                                        flag = true;
                                    }
                                });
                                if (flag == false) {
                                    DocsNodeCategoriesItemsData.push({ Key: item.ID, FileName: item.Category, Title: item.Category, CategoryType: (item.CategoryType != undefined ? item.CategoryType : null), ParentCategory: (item.ParentCategory != null ? item.ParentCategory : ''), Category: (item.Category != undefined ? item.Category : null), FileType: this._getIcon(item.DocIcon) });
                                }
                            });
                            for (var z = 0; z < DocsNodeCategoriesItemsData.length; z++) {
                                DocsNodeCategoriesItemsData.map((item) => {
                                    if (DocsNodeCategoriesItemsData[z].Title == item.ParentCategory) {
                                        DocsNodeCategoriesItemsData[z]["DeleteEditFlag"] = false;
                                    }
                                });
                                _catArray.map((item) => {
                                    if (DocsNodeCategoriesItemsData[z].Title == item) {
                                        DocsNodeCategoriesItemsData[z]["DeleteEditFlag"] = false;
                                    }
                                });
                            }
                        }                        
                    }
                    return ({ DocsNodeCategoriesArrayItems, DocsNodeParentCategoriesArrayItems, DocsNodeCategoriesItemsData });
                }, (responseError) => {
                    console.log('getDocsNodeCategoriesName inside getRequest : ' + responseError);
                    return ({ DocsNodeCategoriesArrayItems, DocsNodeParentCategoriesArrayItems });
                });
        } catch (error) {
            console.log('getDocsNodeCategoriesName : ' + error);
        }
    }

    /**
     * Upload File or image and adding or editing item in Library.
     * @param uploadFileObj 
     * @param titleValue 
     * @param discriValue 
     * @param key 
     * @param filename 
     * @param context 
     * @param ListDisplayName 
     */
    public async _uploadFiles(uploadFileObj, titleValue, discriValue, key, filename, context, ListDisplayName) {
        try {
            if (uploadFileObj != '') {
                var file = uploadFileObj;
                if (file != undefined || file != null) {
                    let spOpts: ISPHttpClientOptions = {
                        headers: {
                            "Accept": "application/json",
                            "Content-Type": "application/json"
                        },
                        body: file,
                        credentials: "same-origin"
                    };
                    var fileName = file.name;
                    fileName = encodeURIComponent("" + fileName + "");
                    var url = `${CU.TenantUrl}/_api/Web/Lists/getByTitle('${ListDisplayName}')/RootFolder/Files/Add(url='${fileName}', overwrite=true)`;
                    //POST call
                    return context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then(async (response: SPHttpClientResponse) => {
                        return response.json();
                    }).then(async (responseJSON: JSON) => {
                        //updating columns values of this item
                        return this._updateLibraryColValue(responseJSON['Name'], titleValue, discriValue, key, ListDisplayName);
                    });
                }
            } else {
                //updating columns values of this item
                return await this._updateLibraryColValue(filename, titleValue, discriValue, key, ListDisplayName);
            }
        }
        catch (error) {
            console.log('uploadFiles : ' + error);
        }
    }

    /**
     * Upload product logo to site assets library.
     * @param uploadFileObj 
     * @param context 
     */
    public async _uploadProductLogo(uploadFileObj, context) {
        try {
            if (uploadFileObj != '') {
                //await this._deleteProductLogo();
                var file = uploadFileObj;
                //this.uploadFile(file);
                if (file != undefined || file != null) {
                    let spOpts: ISPHttpClientOptions = {
                        headers: {
                            "Accept": "application/json",
                            "Content-Type": "application/json"
                        },
                        body: file,
                        credentials: "same-origin"
                    };
                    var fileName = file.name;
                    var url =`${CU.TenantUrl}/_api/Web/Lists/getByTitle('${constant.DocsNodeProductLogoName}')/RootFolder/Files/Add(url='${fileName}')`;
                    //POST call
                    return context.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts).then(async (response: SPHttpClientResponse) => {
                        return response.json();
                    });
                }
            } else {
                console.log('_uploadProductLogo inside');
            }
        }
        catch (error) {
            console.log('_uploadProductLogo : ' + error);
        }
    }

    /**
     * Updating column properties of existing item.
     * @param responseName 
     * @param titleValue 
     * @param discriValue 
     * @param key 
     * @param ListDisplayName 
     */
    public _updateLibraryColValue(responseName, titleValue, discriValue, key, ListDisplayName) {
        try {
            var itemURL = '';
            var FileName = encodeURIComponent("" + responseName + "");
            itemURL = `${CU.TenantUrl}/_api/web/lists/getbytitle('${ListDisplayName}')/items?$select=FileRef,ID,${constant.InternalLinkFilename}&$filter=substringof('${FileName}',FileRef)`;
            //GET request
            return CU._getRequest(itemURL).then(async (responseData) => {
                var resultData = responseData.d.results;
                if(resultData.length > 0){
                    var addDocColumnUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${ListDisplayName}')/items(${resultData[0].ID})`;
                    var commomJSON = null;
                    if (key == '') {
                        key = null;
                    }
                    switch (ListDisplayName) {
                        case constant.DocsNodeSlidesName:
                            commomJSON = JSON.stringify({
                                __metadata: { 'type': await this._getListItemEntityTypeFullName(ListDisplayName) },
                                SlidesDiscription: discriValue,
                                Title: titleValue,
                                SlidesCategoryId: key
                            });
                            break;
                        case constant.DocsNodePictureName:
                            commomJSON = JSON.stringify({
                                __metadata: { 'type': await this._getListItemEntityTypeFullName(ListDisplayName) },
                                Title: titleValue,
                                ImageCategoryId: key,
                                Description: discriValue
                            });
                            break;
                        default:
                            break;
                    }
                    //POST request           
                    return CU._postRequest(addDocColumnUrl, commomJSON, 'MERGE').then((data) => {
                        return data;
                    });
                }                
            });
        } catch (error) {
            console.log("postRequest: " + error);
        }
    }

    /**
     *Get Product logo from site assets. */
    public _getProductLogo() {
        try {
            var responseResult = '';
            var itemURL = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeProductLogoName}')/items?$select=FileRef,ID,${constant.InternalLinkFilename},EncodedAbsThumbnailUrl,EncodedAbsUrl,*`;
            //GET request
            return CU._getRequest(itemURL).then(async (responseData) => {
                responseResult = responseData.d.results;
                if (responseResult.length > 0) {
                    var EncodeURL = '';
                    var itemID = '';
                    for (var i = 0; i < responseResult.length; i++) {
                        var results = responseResult[i];                        
                        EncodeURL = results['EncodedAbsUrl'];
                        itemID = results['ID'];
                        return ({ EncodeURL, itemID });                      
                    }
                } else {
                    return responseResult;
                }
            });
        } catch (error) {
            console.log("_getProductLogo: " + error);
        }
    }

    /**
     * DELETE Product logo from site assets.
     * @param id 
     */
    public _deleteProductLogo(id) {
        try {
            var itemURL = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeProductLogoName}')/items(${id})`;
            //POST request           
            return CU._postRequest(itemURL, '', 'DELETE').then((data) => {
                return data;
            });
        } catch (error) {
            console.log("_deleteProductLogo: " + error);
        }
    }

    /**
     * Adding or editing List items.
     * @param titleValue 
     * @param textsnippetORSiteColl 
     * @param discriValueORsubsite 
     * @param keyORplaceholderDiscrip 
     * @param ListDisplayName 
     * @param flag 
     * @param itemID 
     * @param catgyTypeORList 
     * @param ParentLevelORFieldName 
     */
    public async _updateListItem(titleValue, textsnippetORSiteColl, discriValueORsubsite, keyORplaceholderDiscrip, ListDisplayName, flag, itemID, catgyTypeORList, ParentLevelORFieldName) {
        try {
            var xMethod = '';
            var url = '';
            var commomJSON = '';
            if (flag) {//Edit item url
                url = `${CU.TenantUrl}/_api/web/lists/getbytitle('${ListDisplayName}')/items(${itemID})`;
            } else {// Add item url
                url = `${CU.TenantUrl}/_api/web/lists/getbytitle('${ListDisplayName}')/items`;
            }
            if (keyORplaceholderDiscrip == '') {
                keyORplaceholderDiscrip = null;
            }
            if (ListDisplayName == constant.DocsNodeTextName) {
                commomJSON = JSON.stringify({
                    __metadata: { 'type': await this._getListItemEntityTypeFullName(constant.DocsNodeTextName) },
                    Title: titleValue,
                    TextSnippet: textsnippetORSiteColl,
                    TSDiscription: discriValueORsubsite,
                    TextCategoryId: keyORplaceholderDiscrip
                });
            } else if (ListDisplayName == constant.DocsNodeCategoriesName) {
                commomJSON = JSON.stringify({
                    __metadata: { 'type': await this._getListItemEntityTypeFullName(constant.DocsNodeCategoriesName) },
                    Title: titleValue,
                    Category: titleValue,
                    ParentCategory: keyORplaceholderDiscrip,
                    CategoryType: catgyTypeORList,
                    CategoryLevel: ParentLevelORFieldName
                });
            } else {
                commomJSON = JSON.stringify({
                    __metadata: { 'type': await this._getListItemEntityTypeFullName(constant.DocsNodePlaceHolderName) },
                    Title: titleValue,
                    Placeholder: titleValue,
                    SiteCollectionUrl: textsnippetORSiteColl,
                    SubSiteUrl: discriValueORsubsite,
                    ListUrl: catgyTypeORList,
                    ListField: ParentLevelORFieldName,
                    PlaceholderDiscription: keyORplaceholderDiscrip
                });
            }
            if (flag) {
                xMethod = 'MERGE';
            }
            //POST request
            return CU._postRequest(url, commomJSON, xMethod).then(async (data)=> {
                if(data.d != null){
                    if(ListDisplayName == constant.DocsNodeTextName && flag == false){
                        await CU._addItemLevelPermission(data.d.ID).then((result)=>{
                            return result;
                        });
                    }else{
                        return data;
                    } 
                }else{
                    return null;
                }                               
            },(error)=>{
                console.log('_updateListItem _postRequest :'+ error);
            });
        } catch (error) {
            console.log('_updateListItem : ' + error);
        }
    }

    /**
     * Get the item for editing for Library or List.
     * @param itemResult 
     * @param listname 
     */
    public _getLibraryItemToEdit(itemResult, listname) {
        try {
            var getListOrLibDataUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${listname}')/items(${itemResult.Key})?$select=ID,${constant.InternalLinkFilename},DocIcon,${constant.Title},*`;
            //GET request
            return CU._getRequest(getListOrLibDataUrl)
                .then((responseData) => {
                    var editItemArray = [];
                    var responseData = responseData.d;
                    switch (listname) {
                        case constant.DocsNodeSlidesName:
                            editItemArray.push({ Name: responseData[constant.InternalLinkFilename], Title: responseData[constant.Title], Discription: responseData[constant.SlidesDiscriptionName], CategoryKey: responseData.SlidesCategoryId });
                            break;
                        case constant.DocsNodePictureName:
                            editItemArray.push({ Name: responseData[constant.InternalLinkFilename], Title: responseData[constant.Title], Discription: responseData.Description, CategoryKey: responseData.ImageCategoryId });
                            break;
                        case constant.DocsNodeTextName:
                            editItemArray.push({ Name: responseData[constant.Title], Title: responseData.Title, Discription: responseData[constant.TextSnippetDiscriptionName], TextSnippet: responseData[constant.TextSnippetName], CategoryKey: responseData.TextCategoryId });
                            break;
                        case constant.DocsNodeCategoriesName:
                            editItemArray.push({ Name: responseData[constant.CategoryName], Title: responseData[constant.CategoryName], CategoryKey: responseData[constant.CategoryParentId], CategoryType: responseData[constant.CategoryType] });
                            break;
                        case constant.DocsNodePlaceHolderName:
                            editItemArray.push({ PlaceHolderName: responseData[constant.PlaceHolderName], Discription: responseData[constant.PlaceHolderDiscrip], SiteCollUrl: responseData[constant.SiteCollectionUrlName], SubsiteUrl: responseData[constant.SubSiteUrlName], ListUrl: responseData[constant.ListUrlName], Listfield: responseData[constant.ListFieldName] });
                            break;
                        default:
                            break;
                    }
                    return editItemArray;
                });
        }
        catch (error) {
            console.log('_getLibraryItemToEdit : ' + error);
        }
    }

    /**
     * Delete the item from Library or List.
     * @param itemResult 
     * @param listname 
     */
    public _deleteListItem(itemResult, listname) {
        try {
            var deleteItmeUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${listname}')/items(${itemResult.Key})`;
            //POST request
            return CU._postRequest(deleteItmeUrl, '', 'DELETE')
                .then((responseData) => {
                    if (responseData == 'success') {
                        return responseData;
                    } else {
                        return responseData.error.message.value;
                    }
                });
        } catch (error) {
            console.log('_deleteListItem: ' + error);
        }
    }

    /**
     * Add new item in configuration list.
     * @param newResultData 
     */
    public async _addNewConfigurationListData(newResultData) {
        try {
            if (newResultData) {
                for (var i = 0; i < newResultData.length; i++) {
                    var resultData = newResultData[i];
                    var configURL = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeConfigurationName}')/items?$filter=${constant.ConfigAssetTitleName} eq '${resultData.assetTitle}'`;
                    //GET request
                    await CU._getRequest(configURL).then(async (responseData) => {
                        var results = responseData.d.results;
                        if(results.length > 0){                            
                            var configUrlMerge = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodeConfigurationName}')/items(${results[0].ID})`;
                            var commomJSON = JSON.stringify({
                                __metadata: { 'type': await this._getListItemEntityTypeFullName(constant.DocsNodeConfigurationName) },
                                ConfigSourceList: resultData.sourceLocation,
                                ConfigSourceListPath: CU.siteCollectionPath,
                                ConfigSourceListGUID: await this._getListGUID(resultData.sourceLocation)
                            });
                            //POST request
                            await CU._postRequest(configUrlMerge, commomJSON, 'MERGE').then(() => {                                
                            });
                        }else{
                            console.log('_addNewConfigurationListData no data found:');
                        }
                    });
                }
                return 'success';
            }
        } catch (error) {
            console.log('_addNewConfigurationListData :' + error);
        }
    }

    /**
     * Get icons for document.
     * @param IconType 
     */
    public _getIcon(IconType: string): string {
        if (IconType != null) {
            switch (IconType) {
                case 'docx':
                    return String(require('../images/icon-docx.png'));
                case 'xlsx':
                    return String(require('../images/icon-xlsx.png'));
                case 'ppt':
                case 'pptx':
                    return String(require('../images/icon-ppt.png'));
                case 'pdf':
                    return String(require('../images/icon-pdf.png'));
                case 'jpg':
                case 'jpeg':
                case 'gif':
                case 'bmp':
                case 'png':
                    return String(require('../images/icon-img.png'));
                case 'txt':
                    return String(require('../images/icon-text.png'));
                default:
                    return String(require('../images/icon-other.png'));
            }
        }
    }

    /**
     * Check wheather the placeholder name is exist or not in placeholder list.
     * @param placeholdername 
     */
    public _placeHolderExistOrNot(placeholdername): any {
        try {
            var jason = JSON.stringify({
                "query": {
                    "__metadata": { "type": "SP.CamlQuery" },
                    "ViewXml": "<View><Query><Where><Eq><FieldRef Name='Placeholder' /><Value Type='Text'>" + placeholdername + "</Value></Eq></Where></Where></Query></View>"
                }
            });
            var getPlaceholdeItemUrl = `${CU.TenantUrl}/_api/web/lists/getbytitle('${constant.DocsNodePlaceHolderName}')/getitems`;
            //POST request
            return CU._postRequest(getPlaceholdeItemUrl, jason, '').then((responseData) => {
                var result = responseData.d.results;
                if (result.length > 0) {
                    return 'Exist';
                } else {
                    return result;
                }
            });
        } catch (error) {
            console.log('placeHolderExistOrNot : ' + error);
        }
    }

    /**
     *Get all Site collection from tenant. */
    public _getAllSiteCollections(): any {
        let listAllSC;
        try {
            let tempArray = CU.tenantURL().split('.');
            let mySitePath = `${tempArray[0]}-my.${tempArray[1]}.${tempArray[2]}/personal`; //"https://binaryrepublik516-my.sharepoint.com/personal";
            let url = `${CU.tenantURL()}/_api/search/query?querytext='NOT Path:${mySitePath}/* contentclass:sts_site'&rowLimit=499&TrimDuplicates=false`;
            //GET request
            return CU._getRequest(url).then((data) => {
                var result = [];
                listAllSC = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
                for (let rslt of listAllSC) {
                    if (rslt.Cells.results[6].Value.indexOf("/portals/") < 0) {
                        result.push({
                            Title: rslt.Cells.results[3].Value,
                            Path: rslt.Cells.results[6].Value
                        });
                    }
                }
                return result;
            });
        }
        catch (error) {
            console.log("getAllSiteCollections: " + error);
        }
    }

    /**
     * Get all list from site collection or subsite.
     * @param ssURL 
     * @param baseTempId 
     */
    public _getAllList(ssURL, baseTempId): any {
        let listAllLists;
        try {
            let getListUrl = `${ssURL}/_api/Web/Lists?$filter=(BaseTemplate eq ${baseTempId})`;
            //GET request
            return CU._getRequest(getListUrl).then((data) => {
                listAllLists = data.d.results;
                return listAllLists;
            });
        } catch (error) {
            console.log("getAllList: " + error);
        }
    }

    /**
     * Get all fields from list.
     * @param ListUrl 
     * @param ListName 
     */
    public _getfields(ListUrl, ListName, disListName) {
        let fields = [];
        try {
            let url = `${CU.tenantURL()}${ListUrl}/_api/web/lists/getbytitle('${ListName}')/fields?$expand=columns&filter=readOnly eq false`;
            //GET request
            return CU._getRequest(url).then((data) => {
                var dataItem = data.d.results;
                for (var i = 0; i < dataItem.length; i++) {
                    var results = dataItem[i];
                    if (results.ReadOnlyField === false && results.InternalName !== "ContentType" && results.InternalName !== "Attachments"
                        && results.InternalName !== "Order" && results.InternalName !== "FileLeafRef" && results.InternalName !== "MetaInfo" && results.TypeDisplayName !== "Person or Group"
                        && results.TypeDisplayName !== "Date and Time" && results.TypeDisplayName !== "Lookup") {
                        fields.push({ key: results.InternalName, text: results.Title, internalName: results.InternalName });
                    }
                }
                return fields;
            });
        } catch (error) {
            console.log('_getfields: ' + error);
        }
    }
    // public _getfields(ListUrl, ListName) : any {
    //     let fields = [];
    //     var dfdFldr = $.Deferred();
    //     try {
    //         let url = `${CU.tenantURL()}${ListUrl}/_api/web/lists/getbytitle('${ListName}')/fields?$expand=columns&filter=readOnly eq false`;
    //         //GET request
    //         CU._getRequest(url).done((data) => {
    //             var dataItem = data.d.results;
    //             for (var i = 0; i < dataItem.length; i++) {
    //                 var results = dataItem[i];
    //                 if (results.ReadOnlyField === false && results.InternalName !== "ContentType" && results.InternalName !== "Attachments"
    //                     && results.InternalName !== "Order" && results.InternalName !== "FileLeafRef" && results.InternalName !== "MetaInfo" && results.TypeDisplayName !== "Person or Group"
    //                     && results.TypeDisplayName !== "Date and Time" && results.TypeDisplayName !== "Lookup") {
    //                     fields.push({ key: results.InternalName, text: results.Title, internalName: results.InternalName });
    //                 }
    //             }
    //             dfdFldr.resolve(fields);
    //         });
    //     } catch (error) {
    //         console.log('_getfields: ' + error);
    //         dfdFldr.reject(error);
    //     }
    //     return dfdFldr.promise();
    // }

    /**
     * Create EntityTypeName of list and library.
     * @param name 
     */
    public _GetItemTypeForListName(name) {
        return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
    }

    /**
     * Get list entity full name for list operations.
     * @param listName 
     */
    public _getListItemEntityTypeFullName(listName): any {
        const one = new Promise<any>(async (resolve, reject) => {
            try {
                var getAllListURL = `${CU.TenantUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`;
                return fetch(getAllListURL, {
                    headers: { Accept: 'application/json;odata=verbose' },
                    credentials: "same-origin"
                }).then((response) => {
                    return response.json();
                }, (errorMsg) => {
                    console.log("getListItemEntityTypeFullName: ", errorMsg);
                }).then((resposeJson) => {
                    resolve(resposeJson.d.ListItemEntityTypeFullName);
                }).catch((responseError) => {
                    return [];
                });
            }
            catch (error) {
                console.log('getListItemEntityTypeFullName: ' + error);
            }
        });
        return one;
    }

    /**
     * Binding list data.
     * @param doclibdata 
     */
    public _bindingAllList(doclibdata) {
        var listOption = [];
        for (var i = 0; i < doclibdata.length; i++) {
            var results = doclibdata[i];
            var NEntityName = "";
            if ((results.EntityTypeName.indexOf("_x005f_") != -1) || (results.EntityTypeName.indexOf("_x0020_") != -1)) {
                NEntityName = results.EntityTypeName.toString();
                NEntityName = NEntityName.replace(new RegExp('_x0020_', 'g'), '%20').replace(new RegExp('_x005f_', 'g'), '_');
            }
            else {
                NEntityName = results.EntityTypeName;
            }
            if (doclibdata.length == 1) {
                listOption.push({ key: results.EntityTypeName, text: results.Title, value: results.ParentWebUrl, internalName: NEntityName });
            }
            else {
                listOption.push({ key: results.EntityTypeName, text: results.Title, value: results.ParentWebUrl, internalName: NEntityName });
            }
        }
        return listOption;
    }
}