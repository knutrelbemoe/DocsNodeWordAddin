var SPURL = "";
var GraphAPIToken = '';
var SPToken = '';
var templateServerRelURL = "/sites/DocsNodeAdmin/";
var TextSnippetLibraryDisplayName = "DocsNodeText";
var TemplateLibraryDisplayName = "";
var ImageListGUID = "";
var CategoryListGUID = "";
var SnippetColumn = "TextSnippet";
var textSnippetList = [];
var twoBreadCrumbArr = [];
var oneBreadCrumbArr = [];
var imageMaxID = 0;
var imageParentMaxID = 0;
var textMaxID = 0;
var textParentMaxID = 0;
var imageItem = 20;
var textItem = 20;
var PlaceHolderListName = 'DocsNodePlaceHolder';
var BRTemplatesJS = window.BRTemplatesJS || {};
var imageSearchQuery = "";
var textSearchQuery = "";
var textLazyLoad = false;

var snippetSearch = "";
var imgSearch = "";
var TokenArray;
var textCatAvailable = false;
var imageListSite = '';

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $("body").on("click", ".clickme a", function () {
                $('.clickme a').removeClass('activelink');
                $(this).addClass('activelink');
                var tabid = $(this).data('tag');
                $('.list').removeClass('active').addClass('hide');
                $('#' + tabid).addClass('active').removeClass('hide');
                $('#ManageTabsContent').show();
                if (tabid == 'two') {
                    $('#btnRefresh').show();
                    $(".tabSearchBox").show();
                    $('.placeholderMain').hide();
                    $('#secondMainContent').css('display', 'block');
                    $('#mainContent').css('display', 'none');
                    $('#liAddTxtSni').css('display', 'none');
                    $("#txtSearch").attr("placeholder", "Search Corporate Images");
                    $("#txtSearch").val(imgSearch);
                    if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == "0" || $("#hdnCategory").val() == undefined) {
                        $("#two ul.categoryList").show();
                        $("#two ul.categoryItems").show();
                        $("#two ul.childCategoryList").hide();
                        $("#two ul.childCategoryItems").hide();
                    }
                    else {

                        $('#txtSearch').val("");
                        textSearchQuery = "";
                        imageSearchQuery = "";

                        imageMaxID = 0;
                        var $_category = $("#hdnCategory").val();
                        var $_categoryName = $("#hdnCategoryName").val();
                        LoadCategoryandItemsOnFolderClick($_category, $_categoryName, false);
                    }
                }
                if (tabid == 'one') {
                    $('#alertMsg').hide();
                    $('#btnRefresh').show();
                    $(".tabSearchBox").show();
                    $('.placeholderMain').hide();
                    $('#secondMainContent').css('display', 'block');
                    $('#mainContent').css('display', 'none');
                    $('#liAddTxtSni').css('display', 'inline-block');
                    $("#txtSearch").attr("placeholder", "Search Text Snippet");
                    $("#txtSearch").val(snippetSearch);
                    if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined) {
                        $("#one ul.categoryItems").show();
                        $("#ulCat").show();
                        $("#one ul.childTextCategory").hide();
                        $("#one ul.childTextItems").hide();
                    }
                    else {

                        $('#txtSearch').val("");
                        textSearchQuery = "";
                        imageSearchQuery = "";

                        textMaxID = 0;
                        var $_category = $("#hdnTextCategory").val();
                        var $_categoryName = $("#hdnTextCategoryName").val();
                        loadAllCategories($_category, $_categoryName, false);
                    }
                }
                if (tabid == "four") {
                    $(".tabSearchBox").hide();
                    $('#btnRefresh').show();
                    $('#liAddTxtSni').css('display', 'none');
                    $('#secondMainContent').css('display', 'block');
                    $('#mainContent').css('display', 'none');
                    $('.placeholderMain').show();
                    LoadPlaceHolderItemsOnPageLoad();
                }
                if (tabid == 'zero') {
                    $('#secondMainContent').css('display', 'none');
                    $('#mainContent').css('display', 'block');
                    $('#liAddTxtSni').css('display', 'none')
                    $('#btnRefresh').hide();
                    $('#ManageTabsContent').hide();
                }
            });

            $("body").on("click", ".txtSnippet", function () {
                var snippetID = parseInt($(this).find(".lblId").text());
                LoadCategories(TemplateLibraryDisplayName, SnippetColumn, "ID eq " + snippetID).then(function (data) {
                    if (data.results.length > 0) {
                        insertParagraphs(data.results[0].TextSnippet);
                    }
                });
            });

            function insertParagraphs(desc) {
                Word.run(function (context) {
                    var selectionRange = context.document.getSelection();
                    selectionRange.insertHtml(desc, Word.InsertLocation.before);
                    return context.sync();
                })
                    .catch(function (error) {
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });
            }

            $("body").on("click", "#two .liTitle", function () {
                imageMaxID = 0;
                var $_this = $(this).attr("attr-category");
                var catName = $(this).text();
                LoadCategoryandItemsOnFolderClick($_this, catName, true);
            });

            $("body").on("click", "#one .liTitle", function () {
                $('#alertMsg').hide();
                textMaxID = 0;
                var $_this = $(this);
                var $_category = $_this.attr("attr-category");
                var catName = $(this).text();
                loadAllCategories(parseInt($_category), catName, true);
            });

            $('.list.SideContentBox').on('scroll', function () {
                if ($(this).scrollTop() + $(this).innerHeight() >= $(this)[0].scrollHeight) {
                    if ($('.list.SideContentBox.active').attr("id") == "two") // Image section
                    {
                        loadImagesOnScroll();
                    }
                    if ($('.list.SideContentBox.active').attr("id") == "one") // Image section
                    {
                        loadTextOnScroll();
                    }
                }
            });

            //on keydown, clear the countdown 
            $('#txtSearch').on('keypress', function (e) {
                if (e.which == 13) {

                    if ($(".SideContentBox.active").attr("id") == "one") {
                        snippetSearch = $("#txtSearch").val();
                        var $_category = $("#hdnTextCategory").val();
                        if ($("#txtSearch").val() != "") {
                            searchTextSnippet($("#txtSearch").val().trim(), $_category);
                        }
                        else {
                            clearSearchText();
                        }
                    }

                    if ($(".SideContentBox.active").attr("id") == "two") {
                        imgSearch = $("#txtSearch").val();
                        var $_category = $("#hdnCategory").val();
                        if ($("#txtSearch").val() != "") {
                            seachImages($("#txtSearch").val().trim(), $_category);
                        }
                        else {
                            clearSearchText();
                        }
                    }
                }
            });

            $("body").on("click", "#liPlaceHolder", function () {
                $(".tabSearchBox").hide();
                $('#liAddTxtSni').css('display', 'none');
                LoadPlaceHolderItemsOnPageLoad();
            });

            $("body").on("click", "#liAddTxtSni", function () {
                openTextSnippetDialogBox();
            });
            $("body").on("click", "#clearSearch", function () {
                $('.list.SideContentBox.active').scrollTop(0);
                setTimeout(function () {
                    refresh();
                }, 100);
            });

            $("body").on("click", "#btnRefresh", function () {
                $('.list.SideContentBox.active').scrollTop(0);
                setTimeout(function () {
                    refresh();
                }, 100);
            });

            $('#SPPlaceholder').on('change', getFieldsFromList);

            $("#btnCreatePlaceHolder").click(function () {
                if ($('#SPPlaceholder').val() !== '0') {
                    $("#dvPlacVal").hide();
                    createPlaceholder($("#SPPlaceholder").val(), $('#SPPlaceholder option:selected').attr('listName'), $('option:selected', $("#SPPlaceholder")).attr('fieldName'), $("#SPFieldValue").val());
                }
                else {
                    $("#dvPlacVal").show();
                }
            });
            $('.cancel_btn').on('click', function () {
                closeTextSnippetDialogBox();
            });
            $('.alert_msg_close_btn').on('click', function () {
                $('#alertMsg').hide();
                $('#requiredMsg').hide();
            });
        });
    };
})();
BRTemplatesJS.Config = function () {
    var utility = new BRDocsNodeJS.postTokens();
    TokenArray = utility.callFunction();
    SPToken = TokenArray[0];
    GraphAPIToken = TokenArray[1];
    SPURL = TokenArray[2];

    setInterval(function () {
        TokenArray = utility.callFunction();
        SPToken = TokenArray[0];
        GraphAPIToken = TokenArray[1];
        SPURL = TokenArray[2];
        console.log("Refresh Tokens template");
    }, 540001);
    GetConfigurations();
};
function GetConfigurations() {

    var url = SPURL + templateServerRelURL + "/_api/web/lists/getbytitle('DocsNodeConfiguration')/items?$select=ConfigAssestTitle,ConfigSourceList,ConfigSourceListGUID,ConfigSourceListPath,*";
    $.ajax({
        url: url,
        method: "GET",
        async: false,
        headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken }
    }).then(function (result) {

        var catList = _.filter(result.d.results, function (itm) { return itm.ConfigAssestTitle == 'Category List' });
        var textList = _.filter(result.d.results, function (itm) { return itm.ConfigAssestTitle == 'Text Snippet List' });
        var imgList = _.filter(result.d.results, function (itm) { return itm.ConfigAssestTitle == 'Images Library' });

        CategoryListGUID = catList[0].ConfigSourceListGUID;
        TemplateLibraryDisplayName = textList[0].ConfigSourceListGUID;
        ImageListGUID = imgList[0].ConfigSourceListGUID;
        imageListSite = imgList[0].ConfigSourceListPath;

        loadAllCategories(0, "", false);
        LoadCategoryandItemsOnPageLoad();

    }).fail(function (data) {
        console.log(JSON.stringify(data));
    });
}

function refresh() {
    //it means it is a text snippet or images
    if ($(".list.SideContentBox.active").attr("id") == "one" || $(".list.SideContentBox.active").attr("id") == "two") {
        $('#txtSearch').val("");
        textSearchQuery = "";
        imageSearchQuery = "";
        clearSearchText();
        $('#alertMsg').hide();
        $('#richTxtSnippet').val('');
    }
    else {
        $("#dvPlacVal").hide();
        LoadPlaceHolderItemsOnPageLoad();
        //loadPlaceholderControls();
    }
}

function getFieldsFromList() {
    try {
        var placeholderName = $("#SPPlaceholder").val();
        $('#SPFieldValue').find('option:not(:first)').remove();
        var selected = $('#SPPlaceholder option:selected');
        var sitecollectionurl = selected.attr('siteColl');
        var subsiteUrl = selected.attr('subSiteColl');
        var listname = selected.attr('listName');
        var fieldname = selected.attr('fieldName');
        var url = '';
        if (subsiteUrl != 'null') {
            url = SPURL + subsiteUrl + "/_api/web/lists/getbytitle('" + listname + "')/items?expand=fields(select=" + fieldname + ")";
        } else {
            url = sitecollectionurl + "/_api/web/lists/getbytitle('" + listname + "')/items?expand=fields(select=" + fieldname + ")";
        }
        callAjaxGet(url).done(function (response) {
            var dd = response.d.results;
            var result = _.sortBy(dd, function (i) {
                return i[fieldname].toLowerCase();
            });
            var listOfItems = "<option value='0'>--Select--</option>";
            if (response) {
                for (var i = 0; i < result.length; i++) {
                    listOfItems += "<option key='" + result[i].ID + "'>" + result[i][fieldname] + "</option>";
                }
                $('#SPFieldValue').html(listOfItems);
            }
        }, function (error) {
            console.log('getFieldsFromList after rest call : ' + error);
        });
    } catch (error) {
        console.log('getFieldsFromList :' + error);
    }
}

//Place Holder Items Load on Page Load.
function LoadPlaceHolderItemsOnPageLoad() {
    try {
        $('#SPPlaceholder').find('option:not(:first)').remove();
        $('#SPFieldValue').find('option:not(:first)').remove();

        var url = SPURL + templateServerRelURL + "_api/web/lists/getbytitle('" + PlaceHolderListName + "')/items";
        callAjaxGet(url).done(function (response) {
            var dd = response.d.results;
            var result = _.sortBy(dd, function (i) {
                if (i.Placeholder != null)
                    return i.Placeholder.toLowerCase();
            });
            var listOfLibrary = "<option value='0'>--Select--</option>";
            if (response) {
                for (var i = 0; i < result.length; i++) {
                    if (result[i].Placeholder != null)
                        listOfLibrary += "<option siteColl='" + result[i].SiteCollectionUrl + "' subSiteColl='" + result[i].SubSiteUrl + "' listName='" + result[i].ListUrl + "' fieldName='" + result[i].ListField + "' value='" + result[i].Placeholder + "'>" + result[i].Placeholder + "</option>";
                }
            }
            $('#SPPlaceholder').html(listOfLibrary);
        });
    }
    catch (error) {
        console.log('LoadPlaceHolderItemsOnPageLoad : ' + error);
    }

}

//Create new placeholder
function createPlaceholder(placeholdername, listname, fieldName, setvalue) {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {
        var range = context.document.getSelection();
        var myContentControl = range.insertContentControl();
        myContentControl.tag = SPURL + "¤" + $('#SPPlaceholder option:selected').attr('subSiteColl') + "¤" + listname + "¤" + fieldName;
        myContentControl.title = placeholdername;
        myContentControl.insertText('[' + listname + ' : ' + placeholdername + ']', 'Replace');
        myContentControl.load("tag", "title", "id", "placeholderText", "font", "text");
        //myContentControl.cannotEdit = true;
        return context.sync().then(function () {
            //loadPlaceholderControls();
            Word.run(function (context) {
                var contentControls = context.document.contentControls;
                context.load(contentControls, ["tag", "title", "id"]);//["id","tag"]
                return context.sync().then(function () {
                    var numberOfItem = contentControls.items.length;
                    var selectedItem = $('#SPFieldValue').val() == "0" ? "" : $('#SPFieldValue').val();
                    if (numberOfItem && numberOfItem > 0) {
                        for (var i = 0; i < numberOfItem; i++) {
                            if (contentControls.items[i].title == placeholdername) {
                                contentControls.items[i].insertText(selectedItem, 'replace');
                            }
                        }
                        return context.sync()
                            .then(function () {

                            });
                    }
                    else {
                        console.log('No content control found.');
                    }
                });
            });
        });
    }).catch(function (error) {
        console.log('Error: ' + JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });
}

/**
 * Close Text Snippet Dialog Box.
 */
function closeTextSnippetDialogBox() {
    $('#requiredMsg').hide();
    $('#richTxtSnippet').val('');
    $('#txtSniTitle').val('');
    $('#txtSniDescript').val('');
    // $('#txtSniCat').find('option:not(:first)').remove();    // Commented by Amartya on 10-01-2020 for Category
    $('#SaveSnippetModal').css('display', 'none');
}

/**
 * Open Text Snippet Dialog Box.
 */
function openTextSnippetDialogBox() {
    try {
        Word.run(function (context) {
            var range = context.document.getSelection(); // Create a range proxy object for the current selection.
            context.load(range);
            // Synchronize the document state by executing the queued commands,and return a promise to indicate task completion.
            return context.sync().then(function () {
                if (range.isEmpty) { //Check if the selection is empty    
                    $('#alertMsg').show();
                    $('#waranigMsg').text('Please select the text for text snippet.');
                    $('#richTxtSnippet').val('');
                } else {
                    var html = range.getHtml();
                    return context.sync().then(function () {
                        var txtLine = html.value;
                        if (txtLine.indexOf('<img') > -1 && localStorage.platform == 'PC') {
                            $('#alertMsg').show();
                            $('#waranigMsg').text('Please remove image/s from selected text and Try Again!!');
                            $('#SaveSnippetModal').css('display', 'none');
                            $('#requiredMsg').hide();
                        } else {
                            $('#alertMsg').hide();
                            $('#SaveSnippetModal').css('display', 'block');
                            $('#requiredMsg').hide();
                            $('#richTxtSnippet').val(txtLine);
                            // getTextSnippetCategory();   // Commented by Amartya on 10-01-2020 for Category
                            $('.save_btn').off('click');
                            $('.save_btn').on('click', function () {
                                addNewTextSnippet();
                            });
                        }
                    });
                }
            });
        });
    } catch (error) {
        console.log('openTextSnippetDialogBox : ' + error);
    }
}

/**
 * Add New Text Snippet in List with item level Permission.
 */
function addNewTextSnippet() {
    try {
        var txtSnippetTitle = $('#txtSniTitle').val();
        var txtSnippetDescript = $('#txtSniDescript').val();
        // var txtSnippetCategory = $('#txtSniCat option:selected').attr('key')    // Commented by Amartya on 10-01-2020 for Category
        txtSnippetTitle = txtSnippetTitle != '' ? txtSnippetTitle.trim() : '';
        txtSnippetDescript = txtSnippetDescript != '' ? txtSnippetDescript.trim() : '';
        var addTxtURL = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items";
        var tokenURL = SPURL + templateServerRelURL;
        // if (txtSnippetCategory == '') {
        //     txtSnippetCategory = null;
        // }    // Commented by Amartya on 10-01-2020 for Category
        if (txtSnippetTitle != '') {
            $('#spinner').show();
            $('#requiredMsg').hide();
            _getListItemEntityTypeFullName(TextSnippetLibraryDisplayName).then(function (metaData) {
                var JSONString = JSON.stringify({
                    __metadata: { 'type': metaData },
                    Title: txtSnippetTitle,
                    TextSnippet: $('#richTxtSnippet').val(),
                    TSDiscription: txtSnippetDescript,
                    // TextCategoryId: txtSnippetCategory       // Commented by Amartya on 10-01-2020 for Category
                });
                getValues(tokenURL).then(function (token) {
                    _postRequest(addTxtURL, JSONString, token, 'POST').then(function (resultsData) {
                        var responseData = resultsData.d;
                        if (responseData != null) {
                            var itemID = responseData.Id;
                            var endPointUrl = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items(" + itemID + ")/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)";
                            _postRequest(endPointUrl, '', token, 'POST').then(function (breakPremission) {
                                var endPointUrlRoleAssignment = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items(" + itemID + ")/roleassignments/addroleassignment(principalid=3,roleDefId=1073741829)";
                                _postRequest(endPointUrlRoleAssignment, '', token, 'POST').then(function (responsedata) {
                                    $('#spinner').hide();
                                    closeTextSnippetDialogBox();
                                    refresh();
                                });
                            });
                        }
                    });
                });
            });
        } else {
            $('#requiredMsg').show();
        }
    } catch (error) {
        console.log('addNewTextSnippet : ' + error);
    }
}

/**
 * Get Text Snippet Category from Category list.
 */
function getTextSnippetCategory() {
    var dfd = $.Deferred();
    try {
        var filter = "(CategoryType eq 'Text Snippet')";
        if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
            loadDataFromSharePoint(CategoryListGUID, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', filter).then(function (result) {
                var resultData = result.results;
                var listOfLibrary = "<option key='' value='0'>--Select--</option>";
                if (resultData.length > 0) {
                    for (var i = 0; i < resultData.length; i++) {
                        var responseData = resultData[i];
                        listOfLibrary += "<option key='" + responseData.ID + "' categoryName='" + responseData.Category + "' value='" + i + 1 + "'>" + responseData.Category + "</option>";
                    }
                    $('#txtSniCat').html(listOfLibrary);
                    dfd.resolve(listOfLibrary);
                } else {
                    listOfLibrary = "<option key='' value='1' disabled>No data found</option>"
                    $('#txtSniCat').html(listOfLibrary);
                    dfd.resolve(listOfLibrary);
                }
            });
        }
    } catch (error) {
        console.log('getTextSnippetCategory :' + error);
    }
    return dfd.promise();

}

//Load default placeholder if there in file
function loadPlaceholderControls() {
    Word.run(function (context) {
        var contentControls = context.document.contentControls;
        context.load(contentControls, ["tag", "title", "id", "placeholderText", "text"]);
        return context.sync().then(function () {
            var numberOfItem = contentControls.items.length;
            var tags = [];
            var lstName = [];
            if (numberOfItem && numberOfItem > 0) {
                var listOfPlaceHolder = "";
                var existListName = $("#listOfPlaceHolder").find('li');
                for (var j = 0; j < numberOfItem; j++) {
                    var items = contentControls.items[j];
                    var listDisplayName = items.tag.split("¤")[2];
                    lstName.push({ listname: listDisplayName, tag: items.tag, title: items.title });
                }
                var itemUnique = removeDuplicates(lstName, 'listname');

                if (itemUnique.length > 0) {
                    for (var k = 0; k < itemUnique.length; k++) {
                        var item = itemUnique[k];
                        if (item.title !== null && item.title !== '') {
                            listOfPlaceHolder += generatePlaceholderLI(item);
                        }
                    }
                }
                for (var i = 0; i < numberOfItem; i++) {
                    var item = contentControls.items[i];
                    //contentControls.items[i].cannotEdit = false
                    if (tags.indexOf(item.tag) === -1) {
                        tags.push(item.tag);
                        //var listDisplayName = item.tag.split("¤")[2];
                        //if (jobsUnique.length != 1) {
                        //    for (var k = 0; k < jobsUnique.length; k++) {
                        //        if (item.title !== null && item.title !== '' && listDisplayName != jobsUnique[k].listname) {
                        //            listOfPlaceHolder += generatePlaceholderLI(item);
                        //        }
                        //    }  
                        //} else {
                        //    if (item.title !== null && item.title !== '' && listDisplayName != lstName[k]) {
                        //        listOfPlaceHolder += generatePlaceholderLI(item);
                        //    }
                        //}
                    }
                }
                $("#listOfPlaceHolder").html(listOfPlaceHolder);
                return context.sync()
                    .then(function () {
                    });
            }
            else {
                console.log('No content control found.');
                $("#listOfPlaceHolder").html('');
            }
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
function removeDuplicates(myArr, prop) {
    return myArr.filter(function (obj, pos, arr) {
        return arr.map(function (mapObj) {
            return mapObj[prop]
        }).indexOf(obj[prop]) === pos;
    });
}
function generatePlaceholderLI(item) {
    var listName = item.tag.split("¤")[2];
    var liElement = '<li listName="' + listName + '" tag="' + item.tag + '"  class="list-group-item"><div class="inline icons"><a href="#" tag="' + item.tag + '" class="btn btn-primary documentText" >' + listName + '</a></div><div class="inline icons" style="float: right;"><i class="fa fa-plus"></i><i data-toggle="modal" data-target="#basicModal" class="fa fa-close"></i></div></li>';
    return liElement;
}

function clearSearchText() {
    if ($(".SideContentBox.active").attr("id") == "one") {
        snippetSearch = "";
        var $_category = $("#hdnTextCategory").val();
        var $_categoryName = $("#hdnTextCategoryName").val();

        if ($_category == "" || $_category == undefined || $_category == "0") {
            textParentMaxID = 0;
            $("#one ul.categoryItems").html("");
            $("#ulCat").html("");
            loadAllCategories(0, "", false);
        }
        else {
            textMaxID = 0;
            $("#one ul.childTextCategory").html("");
            $("#one ul.childTextItems").html("");
            loadAllCategories($_category, $_categoryName, false);
        }
    }

    if ($(".SideContentBox.active").attr("id") == "two") {
        imgSearch = "";
        var $_category = $("#hdnCategory").val();
        var $_categoryName = $("#hdnCategoryName").val();

        if ($_category == "" || $_category == undefined || $_category == "0") {
            imageParentMaxID = 0;
            $("#two ul.categoryList").html("");
            $("#two ul.categoryItems").html("");
            LoadCategoryandItemsOnPageLoad(0);
        }
        else {
            imageMaxID = 0;
            $("#two ul.childCategoryList").html("");
            $("#two ul.childCategoryItems").html("");
            LoadCategoryandItemsOnFolderClick($_category, $_categoryName, false);
        }
    }
}

function seachImages(searchParam, categoryId) {

   // $("#WaitDialog").show();

    if (categoryId != "" && categoryId != "0") {
        imageMaxID = 0;

        $(".childCategoryList").html("");
        $(".childCategoryItems").html("");

        $(".childCategoryList").hide();
        $(".childCategoryItems").show();
        $(".categoryList").hide();
        $("#two ul.categoryItems").hide();

        var categoryList = [];
        categoryList.push(categoryId);
        if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
            loadDataFromSharePoint(CategoryListGUID, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (ParentCategory eq " + categoryId + "))").then(function (data) {
                $.each(data.results, function (index, value) {
                    categoryList.push(value.ID);
                });

                var searchString = "substringof('" + searchParam + "',Title) and ";
                var cateStr = "";
                if (categoryList.length > 1) {
                    cateStr += "("
                    for (var i = 0; i < categoryList.length; i++) {
                        if (i + 1 == categoryList.length)
                            cateStr += "(ImageCategory/ID eq " + categoryList[i] + ")";
                        else
                            cateStr += "(ImageCategory/ID eq " + categoryList[i] + ") or ";
                    }
                    cateStr += ")"

                    imageSearchQuery = "(" + searchString + cateStr + ")";
                }
                else {
                    imageSearchQuery = "(" + searchString + "(ImageCategory/ID eq " + categoryId + "))";
                }

                imageSearchAjax($("#two ul.childCategoryItems"));

            });
        }

    }
    else {
        imageSearchQuery = "(substringof('" + searchParam + "',Title))";
        imageParentMaxID = 0;
        $(".categoryList").html("");
        $("#two ul.categoryItems").html("");

        $(".childCategoryList").hide();
        $(".childCategoryItems").hide();

        $(".categoryList").hide();
        $("#two ul.categoryItems").show();

        imageSearchAjax($("#two ul.categoryItems"));
    }
}

function searchTextSnippet(searchParam, categoryId) {
  //  $("#WaitDialog").show();
    textCatAvailable = false;
    if (categoryId != "" && categoryId != "0") {

        textMaxID = 0;
        $("#one ul.childTextCategory").html("");
        $("#one ul.childTextItems").html("");

        $("#ulCat").hide();
        $("#one ul.categoryItems").hide();
        $("#one ul.childTextCategory").hide();
        $("#one ul.childTextItems").show();

        var categoryList = [];
        categoryList.push(categoryId);
        if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
            loadDataFromSharePoint(CategoryListGUID, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Text Snippet') and (ParentCategory eq " + categoryId + "))").then(function (data) {
                $.each(data.results, function (index, value) {
                    categoryList.push(value.ID);
                });

                var searchString = "substringof('" + searchParam + "',Title) and ";
                var cateStr = "";
                if (categoryList.length > 1) {
                    cateStr += "("
                    for (var i = 0; i < categoryList.length; i++) {
                        if (i + 1 == categoryList.length)
                            cateStr += "(TextCategoryId eq " + categoryList[i] + ")";
                        else
                            cateStr += "(TextCategoryId eq " + categoryList[i] + ") or ";
                    }
                    cateStr += ")"

                    textSearchQuery = "(" + searchString + cateStr + ")";
                }
                else {
                    textSearchQuery = "(" + searchString + "(TextCategoryId eq " + categoryId + "))";
                }

                textSearchAjax($("#one ul.childTextItems"));

            });
        }

    }
    else {
        textSearchQuery = "(substringof('" + searchParam + "',Title))";
        textParentMaxID = 0;
        $("#ulCat").html("");
        $("#one ul.categoryItems").html("");

        $("#ulCat").show();
        $("#one ul.categoryItems").hide();
        $("#one ul.childTextCategory").hide();
        $("#one ul.childTextItems").hide();

        textSearchAjax($("#ulCat"));
    }
}

function imageSearchAjax(container) {
//if(imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
    loadDataFromSharePoint(ImageListGUID, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl', imageSearchQuery, "image",true).then(function (data) {

        if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == "0" || $("#hdnCategory").val() == undefined) {
            imageParentMaxID = _.max(_.pluck(data.results, "ID"));
        }
        else
            imageMaxID = _.max(_.pluck(data.results, "ID"));

        if (data.results.length > 0) {
            $.each(data.results, function (index, image) {

                toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                    var base64result = "data:image/png;base64, " + dataUrl;
                    container.append("<li title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img src='" + base64result + "'>" +
                        "</div></a></li>");
                });
            });
        }
        else {
            container.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
        }
        $("#WaitDialog").hide();
    });
//}
}

function textSearchAjax(container) {
    if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
        loadDataFromSharePoint(TemplateLibraryDisplayName, 'ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory', textSearchQuery, 'text').then(function (data) {
            if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined) {
                textParentMaxID = _.max(_.pluck(data.results, "ID"));
            }
            else
                textMaxID = _.max(_.pluck(data.results, "ID"));

            bindTextSnippetHTML(data.results, container);
            $("#WaitDialog").hide();
        });
    }
}

function loadAllCategories(categoryId, $_categoryName, $_handleBreadcrumb) {

    if (categoryId == 0) {
        $("#one ul.breadcrumb").hide();
    }
    if ($_handleBreadcrumb) {
        if (categoryId != 0) {
            //Added for Breadcrumb by Darshana
            oneBreadCrumbArr.push({ id: categoryId, category: $_categoryName });
            if (oneBreadCrumbArr.length > 0) {
                $("#one ul.breadcrumb").css('display', 'inline-block');
            }
        }
        else {
            $("#one ul.breadcrumb").hide();
        }
        var $_breaCrumbUL = $("#one ul.breadcrumb");
        $_breaCrumbUL.html("");
        // Append to breadcrumb
        for (var key in oneBreadCrumbArr) {
            if (oneBreadCrumbArr.hasOwnProperty(key)) {
                var catID = key == 0 ? 0 : oneBreadCrumbArr[key - 1].id;
                $_breaCrumbUL.append('<li title="' + oneBreadCrumbArr[key].category + '"><a class="ms-Breadcrumb-itemLink" data-id="' + key + '" attr-category="' + catID + '"><i class="ms-Breadcrumb-chevron ms-Icon ms-Icon--ChevronLeft"></i><label class="ms-Breadcrumb-itemLink">' + oneBreadCrumbArr[key].category + '</label></li>');

                $(document).on("click", "#one ul.breadcrumb li", function (event) {
                    $('#alertMsg').hide();
                    $("#txtSearch").val("");
                    textSearchQuery = "";
                    var index = parseInt($(this).find("a").attr("data-id"));
                    if (index < oneBreadCrumbArr.length) {
                        oneBreadCrumbArr = oneBreadCrumbArr.splice(0, oneBreadCrumbArr.length - (oneBreadCrumbArr.length - index));
                        if (oneBreadCrumbArr.length == 0) {
                            $("#one ul.breadcrumb").hide();
                        }
                        var sli = $("#one ul.breadcrumb li a").filter(function () {

                            return $(this).attr("data-id") >= index;
                        });
                        if ($(sli).length > 0) $(sli).parent().remove();
                        if (parseInt($(this).find("a").attr("data-id")) - 1 == -1) {
                            oneBreadCrumbArr = [];
                            var $_breaCrumbUL = $("#one ul.breadcrumb");
                            $_breaCrumbUL.html("");

                            $("#hdnTextCategory").val("");
                            $("#hdnTextCategoryName").val("");

                            $("#ulCat").show();
                            $("#one ul.categoryItems").show();
                            $("#one ul.childTextCategory").hide();
                            $("#one ul.childTextItems").hide();

                        } else {
                            textMaxID = 0;
                            $("#hdnTextCategory").val($(this).find("a").attr("attr-category"));
                            $("#hdnTextCategoryName").val(oneBreadCrumbArr[index - 1].category);
                            loadcategoryandSnippets($(this).find("a").attr("attr-category"));
                        }
                    }
                });
            }
        }
    }

    $("#hdnTextCategory").val(categoryId);
    $("#hdnTextCategoryName").val($_categoryName);
    loadcategoryandSnippets(categoryId);
    //By Darshana
}

function loadcategoryandSnippets(categoryId) {

   // $("#WaitDialog").show();

    var catContainer = "";
    var itemContainer = "";
    var filterCondition = "";

    if (categoryId == 0) {

        catContainer = $("#one ul.categoryItems");
        itemContainer = $("#ulCat");

        $("#one ul.categoryItems").show();
        $("#ulCat").show();
        $("#one ul.childTextCategory").hide();
        $("#one ul.childTextItems").hide();

        filterCondition = "((CategoryType eq 'Text Snippet') and (CategoryLevel eq 0))";
    }
    else {
        $("#one ul.childTextCategory").html("");
        $("#one ul.childTextItems").html("");

        catContainer = $("#one ul.childTextCategory");
        itemContainer = $("#one ul.childTextItems");

        $("#one ul.categoryItems").hide();
        $("#ulCat").hide();
        $("#one ul.childTextCategory").show();
        $("#one ul.childTextItems").show();

        filterCondition = "((CategoryType eq 'Text Snippet') and (ParentCategory eq " + categoryId + "))";
    }

    loadTextSnippetAjax(filterCondition, catContainer, itemContainer, categoryId);
}

function loadTextSnippetAjax(filterCondition, catContainer, itemContainer, categoryId) {

    new Promise(function (callback, reject) {
        textCatAvailable = false;
        LoadCategories(CategoryListGUID, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', filterCondition).then(function (data) {
            if (data.results.length > 0)
                textCatAvailable = true;
            $.each(data.results, function (index, value) {
                var category = value.Category;
                var str = "<li>" +
                    "<a href='#' class='liTitle' attr-category='" + value.Id + "'>" +
                    "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                    "<span class='fileName' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 25) + "..." : category) + "</span>" +
                    "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                    "</a >" +
                    "</li>"
                catContainer.append(str);
            });
            $("#WaitDialog").hide();
            callback(1);
        });

    }).then(function (result, reject) {
        getTextSnippet(categoryId).then(function (data) {
            if (categoryId == 0)
                textParentMaxID = _.max(_.pluck(data.results, "ID"));
            else
                textMaxID = _.max(_.pluck(data.results, "ID"));
            bindTextSnippetHTML(data.results, itemContainer);
        });
    }).then(function (result) {

    });
}

function LoadCategoryandItemsOnPageLoad() {

    twoBreadCrumbArr = [];
    var $_breaCrumbUL = $("#two ul.breadcrumb");
    $_breaCrumbUL.html("");
    $("#two ul.breadcrumb").hide();


    $("#two ul.categoryList").html("");
    $("#two ul.categoryItems").html("");

    $("#two ul.categoryList").show();
    $("#two ul.categoryItems").show();

    $("#two ul.childCategoryList").hide();
    $("#two ul.childCategoryItems").hide();
   // $("#WaitDialog").show();
    var $_ul = $("#two ul.categoryList");
    var isCategoryAvailable = false;
    if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
        new Promise(function (callback, reject) {

            loadDataFromSharePoint(CategoryListGUID, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (CategoryLevel eq 0))").then(function (data) {
                if (data.results.length > 0)
                    isCategoryAvailable = true;

                $.each(data.results, function (index, value) {
                    var category = value.Category;
                    var str = "<li>" +
                        "<a href='#' class='liTitle' attr-category='" + value.ID + "'>" +
                        "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                        "<span class='fileName' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 25) + "..." : category) + "</span>" +
                        "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                        "</a >" +
                        "</li>"
                    $_ul.addClass("root");
                    $_ul.append(str);
                });
                callback(isCategoryAvailable);
                $("#WaitDialog").hide();
            });
        }).then(function (result) {
            var $_ul = $("#two ul.categoryItems");
            loadDataFromSharePoint(ImageListGUID, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory', "(ImageCategory eq null)", "image").then(function (data) {
                imageParentMaxID = _.max(_.pluck(data.results, "ID"));
                if (data.results.length == 0 && !result) {
                    $_ul.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
                }
                else {
                    $.each(data.results, function (index, image) {
                        toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                            var base64result = "data:image/png;base64, " + dataUrl;
                            $_ul.append("<li title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img src='" + base64result + "'>" +
                                "</div></a></li>");
                        });
                    });
                }
            });

        });
    } else {
        console.log("When images are not in DocsNodeAdmin");
         var $_ul = $("#two ul.categoryItems");
            loadDataFromSharePoint(ImageListGUID, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl', "", "image").then(function (data) {
                imageParentMaxID = _.max(_.pluck(data.results, "ID"));
                if (data.results.length == 0 && !result) {
                    $_ul.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
                }
                else {
                    $.each(data.results, function (index, image) {
                        toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                            var base64result = "data:image/png;base64, " + dataUrl;
                            $_ul.append("<li title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img src='" + base64result + "'>" +
                                "</div></a></li>");
                        });
                    });
                }
                $("#WaitDialog").hide();
            });
    }
}

function loadImagesOnScroll() {
    var $_category = $("#hdnCategory").val();
    var $_ul = "";
    var filter = ""

    if (imageSearchQuery != "") {
        filter = imageSearchQuery;
        if ($_category == "" || $_category == undefined || $_category == "0") {
            $_ul = $("#two ul.categoryItems");
        }
        else {
            $_ul = $("#two ul.childCategoryItems");
        }
    }
    else {
        if ($_category == "" || $_category == undefined || $_category == "0") {
            filter = "(ImageCategory eq null)";
            $_ul = $("#two ul.categoryItems");
        }
        else {
            filter = "(ImageCategory/ID eq " + $_category + ")";
            $_ul = $("#two ul.childCategoryItems");
        }
    }

    //if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
        loadDataFromSharePoint(ImageListGUID, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl', filter, "image").then(function (data) {
            if ($_category == "" || $_category == undefined || $_category == "0") {
                imageParentMaxID = _.max(_.pluck(data.results, "ID"));
            }
            else {
                imageMaxID = _.max(_.pluck(data.results, "ID"));
            }

            $.each(data.results, function (index, image) {
                toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                    var base64result = "data:image/png;base64, " + dataUrl;
                    $_ul.append("<li title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img src='" + base64result + "'>" +
                        "</div></a></li>");
                });
            });
        });
    //}
}

function loadTextOnScroll() {
    textCatAvailable = true;
    var $_category = $("#hdnTextCategory").val();
    var $_ul = "";
    var filter = "";
    var cat = 0;

    if (textSearchQuery != "") {
        filter = textSearchQuery;
    }

    if ($_category == "" || $_category == undefined || $_category == "0") {
        $_ul = $("#ulCat");
    }
    else {
        cat = $_category;
        $_ul = $("#one ul.childTextItems");
    }

    getTextSnippetForLazyLoading(cat, filter).then(function (data) {
        if (cat == 0)
            textParentMaxID = _.max(_.pluck(data.results, "ID"));
        else
            textMaxID = _.max(_.pluck(data.results, "ID"));
        bindTextSnippetHTML(data.results, $_ul);
    });
}

function LoadCategoryandItemsOnFolderClick($_this, $_categoryName, $_handleBreadcrumb) {

    $("#two ul.childCategoryList").html("");
    $("#two ul.childCategoryItems").html("");

    var $_category = $_this;
    //Added for Breadcrumb by Darshana
    if ($_handleBreadcrumb) {
        twoBreadCrumbArr.push({ id: $_category, category: $_categoryName });
        if (twoBreadCrumbArr.length > 0) {
            $("#two ul.breadcrumb").css('display', 'inline-block');
        }
        var $_breaCrumbUL = $("#two ul.breadcrumb");
        $_breaCrumbUL.html("");
        // Append to breadcrumb
        for (var key in twoBreadCrumbArr) {
            if (twoBreadCrumbArr.hasOwnProperty(key)) {
                var catID = key == 0 ? 0 : twoBreadCrumbArr[key - 1].id;
                $_breaCrumbUL.append('<li title="' + twoBreadCrumbArr[key].category + '"><a class="ms-Breadcrumb-itemLink" data-id="' + key + '" attr-category="' + catID + '"><i class="ms-Breadcrumb-chevron ms-Icon ms-Icon--ChevronLeft"></i><label class="ms-Breadcrumb-itemLink">' + twoBreadCrumbArr[key].category + '</label></li>');

                $(document).on("click", "#two ul.breadcrumb li", function (event) {
                    $("#txtSearch").val("");
                    imageSearchQuery = "";
                    var index = parseInt($(this).find("a").attr("data-id"));
                    if (index < twoBreadCrumbArr.length) {
                        twoBreadCrumbArr = twoBreadCrumbArr.splice(0, twoBreadCrumbArr.length - (twoBreadCrumbArr.length - index));
                        if (twoBreadCrumbArr.length == 0) {
                            $("#two ul.breadcrumb").hide();
                        }
                        var sli = $("#two ul.breadcrumb li a").filter(function () {

                            return $(this).attr("data-id") >= index;
                        });
                        if ($(sli).length > 0) $(sli).parent().remove();
                        if (parseInt($(this).find("a").attr("data-id")) - 1 == -1) {
                            $("#hdnCategory").val("");
                            $("#hdnCategoryName").val("");

                            $("#two ul.categoryList").show();
                            $("#two ul.categoryItems").show();
                            $("#two ul.childCategoryList").hide();
                            $("#two ul.childCategoryItems").hide();

                        } else {
                            imageMaxID = 0;
                            $("#hdnCategory").val($(this).find("a").attr("attr-category"));
                            $("#hdnCategoryName").val(twoBreadCrumbArr[index - 1].category);
                            loadImagesByCategory($(this).find("a").attr("attr-category"));
                        }
                    }
                });
            }
        }
    }
    //By Darshana

    loadImagesByCategory($_category);
   // $("#WaitDialog").show();
    $("#hdnCategory").val($_category);
    $("#hdnCategoryName").val($_categoryName);
}

function loadImagesByCategory($_category) {
    imageMaxID = 0;

    $("#two ul.categoryList").hide();
    $("#two ul.categoryItems").hide();

    $("#two ul.childCategoryList").show();
    $("#two ul.childCategoryItems").show();
    new Promise(function (callback, reject) {
        var $_ul = $("#two ul.childCategoryList");
        $_ul.html("");
        var isCategoryAvailable = false;
        //if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
            loadDataFromSharePoint(CategoryListGUID, 'ID,Title,Category,ParentCategory,CategoryType,CategoryLevel', "((CategoryType eq 'Images') and (ParentCategory eq " + $_category + "))").then(function (data) {
                if (data.results.length > 0)
                    isCategoryAvailable = true;

                $.each(data.results, function (index, value) {
                    var category = value.Category;
                    var str = "<li>" +
                        "<a href='#' class='liTitle' attr-category='" + value.ID + "'>" +
                        "<span class='file-icon'><i class='ms-Icon ms-Icon--OpenFolderHorizontal' aria-hidden='true'></i></span>" +
                        "<span class='fileName' title='" + category + "'>" + (category.length > 20 ? category.substr(0, 25) + "..." : category) + "</span>" +
                        "<span class='fileEnter-icon'><i class='ms-Icon ms-Icon--ChevronRightSmall' aria-hidden='true'></i></span>" +
                        "</a >" +
                        "</li>";
                    $_ul.append(str);
                });
                $("#WaitDialog").hide();
                callback(isCategoryAvailable);
            });
        //}
    }).then(function (result) {

        var $_ul = $("#two ul.childCategoryItems");
        $_ul.html("");
        if (imageListSite.toLowerCase() == '/sites/docsnodeadmin/' || imageListSite.toLowerCase() == '/sites/docsnodeadmin') {
        loadDataFromSharePoint(ImageListGUID, 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory', "(ImageCategory/ID eq " + $_category + ")", "image").then(function (data) {
            imageMaxID = _.max(_.pluck(data.results, "ID"));
            if (data.results.length == 0 && !result) {
                $_ul.append("<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>");
            }
            else {
                $.each(data.results, function (index, image) {
                    toDataURL(image.EncodedAbsThumbnailUrl.toString(), function (dataUrl) {
                        var base64result = "data:image/png;base64, " + dataUrl;
                        $_ul.append("<li title='" + (image.Title != null ? image.Title : image.LinkFilenameNoMenu) + "' class='liIteamImg' onClick=\"insertImage('" + image.EncodedAbsUrl.toString() + "')\"><a href='#'><div class='liInnerImage'><img src='" + base64result + "'>" +
                            "</div></a></li>");
                    });
                });
            }
        });
    }
    
    });
}

//used for categories and Get Snippet by ID
function LoadCategories(category_name, columns, filters) {
    var dfd = $.Deferred();
    var url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";

    try {
        callAjaxGet(url).done(function (data) {
            dfd.resolve(data.d);
        });
    } catch (err) {
        dfd.reject(err);
    }

    return dfd;
}

function insertImage(fileURL) {
    toDataURL(fileURL, function (base64result) {
        Office.context.document.setSelectedDataAsync(base64result, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    console.log(asyncResult.error.message);
                }
            });
    });
}

function toDataURL(fileURL, callback) {
    fileURL = fileURL.replace(SPURL, "");
    var url = SPURL + imageListSite + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
    var xhr = new window.XMLHttpRequest();
    xhr.open("GET", url, true);
    xhr.setRequestHeader("Accept", "application/json; odata=verbose");
    xhr.setRequestHeader("Authorization", "Bearer " + SPToken);
    //Now set response type
    xhr.responseType = 'arraybuffer';
    xhr.addEventListener('load', function () {
        if (xhr.status === 200) {
            var sampleBytes = new Uint8Array(xhr.response);
            var blob = new Blob([sampleBytes], { type: "image/jpeg" });
            var reader = new FileReader();
            reader.onload = function () {
                var dataUrl = reader.result;
                var base64 = dataUrl.split(',')[1];
                callback(base64);
            };
            reader.readAsDataURL(blob);
        }
    })
    xhr.send();
}

function getTextSnippet(categoryId) {
    var dfd = $.Deferred();


    var cat = "";

    if (categoryId == 0)
        cat = "null";
    else
        cat = categoryId;

    var url = "";

    if (categoryId == 0)// It means parent level
        url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory&$filter=(TextCategoryId eq " + cat + " and ID gt " + textParentMaxID + ")&$top=" + textItem;
    else
        url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory&$filter=(TextCategoryId eq " + cat + " and ID gt " + textMaxID + ")&$top=" + textItem;

    try {
        callAjaxGet(url).done(function (data) {
            dfd.resolve(data.d);
        });
    } catch (err) {
        console.log(err);
        dfd.reject(err);
    }

    return dfd;
}

function getTextSnippetForLazyLoading(categoryId, filter) {
    var dfd = $.Deferred();
    var cat = "";
    if (categoryId == 0)
        cat = "null";
    else
        cat = categoryId;

    var url = "";

    if (categoryId == 0)// It means parent level
    {
        if (filter == "")
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory&$filter=(TextCategoryId eq " + cat + " and ID gt " + textParentMaxID + ")&$top=" + textItem;
        else
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory&$filter=(" + filter + " and ID gt " + textParentMaxID + ")&$top=" + textItem;
    }
    else {
        if (filter == "") {
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory&$filter=(TextCategoryId eq " + cat + " and ID gt " + textMaxID + ")&$top=" + textItem;
        }
        else {
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Title,TextCategory/Id,TextCategory/Title,*&$expand=TextCategory&$filter=(" + filter + " and ID gt " + textMaxID + ")&$top=" + textItem;
        }
    }
    try {
        callAjaxGet(url).done(function (data) {
            dfd.resolve(data.d);

        });
    } catch (err) {
        console.log(err);
        dfd.reject(err);
    }

    return dfd;
}

function loadDataFromSharePoint(category_name, columns, filters, type,search) {
    var dfd = $.Deferred();
    var url = "";
    
    // type is not undefined means lazy loading should be done.
    if (type != undefined) {
        if (type == "image") {
            if (imageListSite.toLowerCase() == "/sites/docsnodeadmin" || imageListSite.toLowerCase() == "/sites/docsnodeadmin/") {
                columns = 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl,ImageCategory/ID,ImageCategory/Title&$expand=ImageCategory';
                //filters = '(ImageCategory eq null)';
                if (imageListSite.toLowerCase() == "/sites/docsnodeadmin") {
                    imageListSite = imageListSite + "/";
                }
            }
            else {
                columns = 'ID,Title,LinkFilenameNoMenu,EncodedAbsThumbnailUrl,EncodedAbsUrl';
                if (!search) 
                filters = '';
            }
           
            if ($("#hdnCategory").val() == "" || $("#hdnCategory").val() == "0" || $("#hdnCategory").val() == undefined)// It means parent level
                url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + (filters ? (filters+" and"):'') + " ID gt " + imageParentMaxID + ")&$top=" + imageItem;
            else
                url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + imageMaxID + ")&$top=" + imageItem;
        }
        else if (type == "text") {
            if ($("#hdnTextCategory").val() == "" || $("#hdnTextCategory").val() == "0" || $("#hdnTextCategory").val() == undefined)
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + textParentMaxID + ")&$top=" + textItem;
            else
                url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + " and ID gt " + textMaxID + ")&$top=" + textItem;
        }
    }
    else {
        if (imageListSite.toLowerCase() == "/sites/docsnodeadmin" || imageListSite.toLowerCase() == "/sites/docsnodeadmin/") {
            url = SPURL + templateServerRelURL + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";
        }
        else {
            url = SPURL + imageListSite + "_api/web/lists(guid'" + category_name + "')/items?$select=" + columns + "&$filter=(" + filters + ")";
        }
    }

    try {
        callAjaxGet(url).done(function (data) {
            dfd.resolve(data.d);
        });
    } catch (err) {
        dfd.reject(err);
    }

    return dfd;
}

function callAjaxGet(url) {
    var dfdGET = $.Deferred();
    try {
        if (SPToken) {
            $.ajax({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                success: function (data) {
                    dfdGET.resolve(data);
                },
                error: function (data) {
                    console.log(data);
                    dfdGET.reject(data);
                }
            });
        }
    } catch (error) {
        console.log("callAjaxGet: " + error);
        dfdGET.reject(error);
    }
    return dfdGET.promise();
}

function _postRequest(url, JSONString, token, MethodType) {
    var dfd = $.Deferred();
    try {
        if (SPToken) {
            $.ajax({
                url: url,
                headers: {
                    "Accept": 'application/json;odata=verbose',
                    "Content-Type": 'application/json;odata=verbose',
                    "X-RequestDigest": token,
                    "IF-MATCH": "*",
                    'Authorization': 'Bearer ' + SPToken,
                },
                type: MethodType,
                data: JSONString,
                success: function (data) {
                    dfd.resolve(data);
                },
                error: function (error) {
                    console.log(error);
                    dfd.reject(error);
                }
            });
        }
    } catch (error) {
        console.log("postRequest: " + error);
    }
    return dfd.promise();
}

function getValues(tokenURl) {
    var dfdReqDig = $.Deferred();
    try {
        if (SPToken) {
            $.ajax({
                url: tokenURl + "_api/contextinfo",
                method: "POST",
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Authorization': 'Bearer ' + SPToken,
                },
                success: function (data) {
                    dfdReqDig.resolve(data.d.GetContextWebInformation.FormDigestValue);
                },
                error: function (err) {
                    console.log("getValues: " + err);
                    dfdReqDig.reject(err);
                }
            });
        }
    } catch (error) {
        console.log("getValues: " + error);
    }
    return dfdReqDig.promise();
}

/**
 * Get List Item Entity Name
 * @param {any} listName
 */
function _getListItemEntityTypeFullName(listName) {
    var dfd = $.Deferred();
    try {

        var getAllListURL = SPURL + templateServerRelURL + "_api/web/lists/getbytitle('" + listName + "')?$select=ListItemEntityTypeFullName";
        callAjaxGet(getAllListURL).then(function (responseData) {
            dfd.resolve(responseData.d.ListItemEntityTypeFullName);
        });
    }
    catch (error) {
        console.log('getListItemEntityTypeFullName: ' + error);
        dfd.reject('error');
    }
    return dfd.promise();
}

function bindTextSnippetHTML(data, itemContainer) {
    textSnippetList = [];
    for (var i = 0; i < data.length; i++) {
        textSnippetList.push({ "Id": data[i].Id, "Title": data[i].Title, "Category": data[i].TextCategory.Title, "Desc": data[i].TextSnippet == null ? "" : data[i].TextSnippet, "ShortDesc": data[i].TSDiscription == null ? "" : data[i].TSDiscription, "CategoryId": data[i].TextCategory.Id });
    }

    var catLi = '';
    if (textSnippetList.length == 0 && !textCatAvailable) {
        catLi = "<div style='margin-top: 100px;color: red;text-align: center;'><label>No Records Found</label></div>";
    }
    else {
        for (var i = 0; i < textSnippetList.length; i++) {
            catLi += "<li title='" + textSnippetList[i].Title + "'>";
            catLi += '<a href="#" class="txtSnippet">';
            catLi += '<label class="lblId" style="display:none">' + textSnippetList[i].Id + '</label>'
            catLi += '<span class="SubSicon"><i class="ms-Icon ms-Icon--TextDocument" aria-hidden="true"></i></span>';
            catLi += '<span class="SubScontent">';
            catLi += '<h3>' + textSnippetList[i].Title + '</h3>';
            catLi += '<p>' + textSnippetList[i].ShortDesc + '</p>';
            catLi += '</span>';
            catLi += '</a>';
            catLi += '</li>';
        }
    }
    itemContainer.append(catLi);
}