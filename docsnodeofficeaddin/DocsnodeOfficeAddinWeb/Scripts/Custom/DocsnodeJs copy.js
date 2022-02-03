"use strict";
var isRoot = true;
var platform;
var BRDocsNodeJS = window.BRDocsNodeJS || {};
var currentTemplateView = "Box";
var gSelectedView = "All%20Documents";
var filterViewArrayList = "";
var filteredData = "";
var RowResult = "";
var SPToken = "";
var GraphAPIToken = "";
var SPURL = "";
var listOfSiteCollectionsArray = [];
var listOfSitesArray = [];
var listOfDocLibsArray = [];
var listOfFoldersArray = [];

var GET_PIN_LOCATION = 'https://docsnode-functions.azurewebsites.net/api/GetPinnedLocations?code=iqlCa0FCvh1eWVgQoly608fCMqEOKqV3mMaQ17yPPS3XEKuspSsjEQ==';
var GET_USER_MSTEAMS = 'https://docsnode-functions.azurewebsites.net/api/GetTeams?code=3hKh9BoXHZM8Ij2kxWInDcJKUaaoAMXvLPkzdCBone90S2d5JHyyvQ==';
var GET_MSTEAM_CHANNELS = 'https://docsnode-functions.azurewebsites.net/api/GetChannel?code=jEPLF3N5jqoqju%2FCoad4aMKuYAy6NHOiqZNTixyHIMtpcBygqg1tIg==';
var GET_TEAMS_WEBURL = "https://docsnode-functions.azurewebsites.net/api/GetTeamsWebUrl?code=Zt3jsbF0UOQoYRDxrMlHKy0sVvxVauX9H22flV1Hu8e0NDWQ1n/88g==";
var GET_CHANNEL_FOLDER = "https://docsnode-functions.azurewebsites.net/api/GetChanelFolder?code=4oc8YxJAaz/nM/LLAuAuAoaZwII3VJXB3YvD6I8ftLyJT7iMuQUuLA==";
var GET_CHANNEL_ALLFOLDER = "https://docsnode-functions.azurewebsites.net/api/GetAllFoldersAtOnce?code=sQYvNv8e5NjPCRDLlIHjyt6h3lRaxqMGa3jgF9pQVHQ4un4oohtgzQ==";
var GET_CURRENT_USER = "https://docsnode-functions.azurewebsites.net/api/GetCurrentUser?code=cud/dWUIUwYJwzGQt3RAkc5BoxwOsvA60r5lRNBr8Ay8XnHzU2FryQ==";
var GET_TEAM_LIBRARY_NAMES = "https://docsnode-functions.azurewebsites.net/api/GetLibraryInternalName?code=OQJoYyaS0VSkJA1tAvV3a3T1/BP9v1a3a5bvlhCHwMhrzPNjzAHeUg==";
var GET_USER_ONEDRIVE = 'https://docsnode-functions.azurewebsites.net/api/GetLoginUsrOneDrive?code=6BP8pKlpxDCWAuNHDkKuDg8MyuFt4DU23GEtN9DP4baPWxS5bwnTeg==';
var GET_USER_ONEDRIVE_HIERARCHY = 'https://docsnode-functions.azurewebsites.net/api/GetUserOndriveFolderHierarchi?code=KTTD7/KTaAodruz51RQe1e6AqqaZX/MpPT5cZAvZ8pvhm6vwdEszZg==';
var DELETE_PINNED_LOCATION = 'https://docsnode-functions.azurewebsites.net/api/DeletePinItem?code=hr/rImI651kFIO268J/X9JOoLZNK3Ijlj8V/1fU/OhkfAul2bqzVww==';
var GET_LOCATION_DETAILS_URL = 'https://docsnode-functions.azurewebsites.net/api/GetLibraryFolderFiles?code=AJz4fp7zj8E2CZ1bWaElj32675EQbP0tFiTvpIG2d9165TOPl6myKQ==';
var GET_DEFAULT_LOCATION_URL = 'https://docsnode-functions.azurewebsites.net/api/GetDefaultView?code=JtcFoDYy8sKVH9UFXxIeG4KzmWEUJU7mrESLN4UmlqVYDmBmUHIaHA==';
var SET_DEFAULT_LOCATION_URL = 'https://docsnode-functions.azurewebsites.net/api/SetDefaultView?code=DbozssXppGlNDbMyhB2wwjrV9NpTjDCGd7BH5Kzdo75HhroqHsYhhQ==';
var GET_TEAM_CHANNEL_TAB_URL = "https://docsnode-functions.azurewebsites.net/api/GetTab?code=AxH/B1aXrKmPhA3fakfSZyqYufHidOXTx3nSrJ2gRAuB/MlA2pVahQ==";

var type = ".doc";

$(document).ready(function () {
    platform = localStorage.platform;
    $('#btnRefresh').hide();
    var token = new BRDocsNodeJS.Tokens();

    setInterval(function () {
        $.getAuthContext.acquireToken(SPURL, function (error, token) {
            if (error || !token) {
                console.log(error);
                var authContext = new AuthenticationContext();
                authContext.login();
            } else {
                SPToken = token;
                $.getAuthContext.acquireToken(config.endpoints.graphApiUrl, function (error, graphtoken) {
                    if (error || !graphtoken) {
                        console.log(error);
                        var authContext = new AuthenticationContext();
                        authContext.login();
                    } else {
                        GraphAPIToken = graphtoken;
                    }
                });
                console.log("Refresh Tokens");
            }
        });

    }, 540000);

    $(".lib-section").css('display', 'none');
});
BRDocsNodeJS.postTokens = function () {
    this.callFunction = function () {
        var Array = [SPToken, GraphAPIToken, SPURL];
        return Array;
    }
}
BRDocsNodeJS.Tokens = function () {
    $.getAuthContext.acquireToken("https://graph.microsoft.com", function (error, graphtoken) {
        if (error || !graphtoken) {
            console.log(error);
            showErrorMessage("Token Error:- " + JSON.stringify(error));
            var authContext = new AuthenticationContext();
            authContext.login();
        } else {
            GraphAPIToken = graphtoken;
            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/root";

            $.ajax({
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json");
                },
                type: "GET",
                url: GraphAPI,
                dataType: "json",
                async: false,
                headers: {
                    'Authorization': 'Bearer ' + graphtoken,
                }
            }).then(function (response) {
                if (response) {
                    SPURL = response.webUrl;

                    $.getAuthContext.acquireToken(SPURL, function (error, token) {
                        if (error || !token) {
                            console.log(error);
                            showErrorMessage("Token Error:- " + JSON.stringify(error));
                            var authContext = new AuthenticationContext();
                            authContext.login();
                        } else {
                            SPToken = token;
                            var docsNode = new BRDocsNodeJS.load();
                            docsNode.GetConfigurations();
                            docsNode.init();
                            var utility = new BRTemplatesJS.Config();
                            $('#createFile').on("click", docsNode.CreateNewTemplate);
                        }
                    });
                }
            }).fail(function (error) {
                console.log('error in gettings source tenant site and web id. error:- ' + error.responseText);
            });
        }
    });
}
BRDocsNodeJS.load = function () {
    var $searchResultsDiv = $('#myTabContent');
    var DestinationWebRelativeUrl = "";
    var TemplateLibraryDisplayName = "";
    var TemplateLibraryDisplayNamee = "";
    var sitePinnedLocations = "DocsNodePinnedLocations";
    var ConfigurationListName = "Configuration";
    var templateServerRelURL = "/sites/DocsNodeAdmin/" //Knut Env    
    var docsNodeFilterExtention = ".doc";
    var docsNodeNewFileExtention = ".docx";
    var docsNodeListTemplateLogo = "word-logo.png";
    var destinationListId = "";
    var destinationSiteIdWebID = "";
    var currentDocId = "";
    var selectedTemplateName = "";
    var currentTenantUrl = "";
    var createdDocName = "";
    var currentLibrary = "";
    var arryOfColumnAndItem = [];
    var filterColumns = [];
    var currentPage = null;
    var Filecount = 0;
    var flag = 0;
    var len;
    var TotalPages;
    var oldfile = "";
    var destinationServerRelativeUrl = '';
    var getdocumentUrlsString = "";
    var boxViewString = "";
    var gboxViewhtml = "";
    var DefaultPageFlag = false;
    var IsPinnedLocation = false;
    var pinnedString = "";
    var saveType = "";

    openWaitDialog();

    this.GetConfigurations = function () {
        var webUrl = SPURL;
        var url = webUrl + templateServerRelURL + "/_api/web/lists/getbytitle('DocsNodeConfiguration')/items?$select=ConfigAssestTitle,ConfigSourceList,ConfigSourceListGUID,ConfigSourceListPath";
        $.ajax({
            url: url,
            method: "GET",
            async: false,
            headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken }
        }).then(function (result) {
            for (var i = 0; i < result.d.results.length; i++) {
                if (result.d.results[i].ConfigAssestTitle == "Template Library") {
                    TemplateLibraryDisplayName = result.d.results[i].ConfigSourceListGUID;
                    TemplateLibraryDisplayNamee = result.d.results[i].ConfigSourceList;
                    getallViews();
                    break;
                }
            }
        }).fail(function (data) {
            console.log(JSON.stringify(data));
        });
    }
    //This function is used for toggle view
    function toggleView() {
        try {
            $("#ViewUL").hide();
            $("#filterUL").hide();
            $('#txtTemplateSearch').val("");
            if (currentTemplateView == "List") {
                $("#noDataFoundLbl").hide();
                $('#DocTemplatesBoxView').hide();
                $('#listOfTemplate').show();
                $('#grdvw').removeClass("selectView");
                $('#lstvw').addClass("selectView");
            }
            else {
                currentTemplateView = "Box";
                $("#noDataFoundLbl").hide();
                $('#DocTemplatesBoxView').show();
                $('#listOfTemplate').hide();
                $('#lstvw').removeClass("selectView");
                $('#grdvw').addClass("selectView");
            }
            if (platform == "OfficeOnline") {
                getalltemp(filteredData);
            }
            else {
                getallClientTemp(filteredData);
            }
            $("#listOfTemplate").find("input:checked").each(function () { $(this).prop('checked', false) });
            $("#DocTemplatesBoxView").find("input:checked").each(function () { $(this).prop('checked', false) })
            $("#previewbtn").attr('disabled', 'disabled');
            $("#previewbtn").css('background', '');
            $("#previewbtn").css('color', '');
            $("#previewbtn").css('cursor', 'default');
            $("#nextbtn").attr('disabled', 'disabled');
            $("#nextbtn").css('background-color', '');
            $("#nextbtn").css('color', '');
            $("#nextbtn").css('cursor', 'default');
            $('li:contains("' + gSelectedView + '")').addClass('selectView');
        }
        catch (e) {
            console.log("toggleView: " + e);
        }
    }

    // Manage screens
    function showfirststep() {
        $("#myTabContent1").css('display', 'block');
        $(".new_default_page").css('display', 'none');
    }
    function showDefaultScreen() {
        if (DefaultPageFlag == true) {
            $("#myTabContent1").css('display', 'block');
            $(".lib-section").css('display', 'none');
        }
        else {
            $("#myTabContent1").css('display', 'block');
            $(".lib-section").css('display', 'none');
        }
    }
    function showSavePanel() {
        showLastPage();
        $(".lib-section").css('display', 'none');
        $("#third_step").css('display', 'block');
        $("#successMessage").css('display', 'none');
        $('#txtNewFileName').focus();
    }
    function showLastPage() {
        $('#popupsave').css('background-color', '#04aba3');
        $('#popupsave').css('pointer-events', 'auto');
        $("#third_step").find("input").removeAttr("disabled");
        $("#third_step").find("input").val("");
        $("#DocumentUrls").html("");

        $(".new_default_page").css('display', 'none');
        $("#third_step").css('display', 'block');
        $("#successMessage").css('display', 'none');
        $('#txtNewFileName').focus();
    }
    function showPreview() {
        $("#myTabContent1").css('display', 'none');
        $("#preview_step").css('display', 'block');
    }
    function showDefaultPageScreen() {
        if (DefaultPageFlag == true) {
            $("#myTabContent1").css('display', 'none');
            $(".lib-section").css('display', 'block');
            $("#preview_step").css('display', 'none');
        }
        else {
            $("#myTabContent1").css('display', 'none');
            $(".lib-section").css('display', 'inline-block');
            $("#preview_step").css('display', 'none');
            PreviousPage();
        }
        $("#defaultPreviousbtn").off('click');
        $("#defaultPreviousbtn").on('click', showfirststep);
        $('#defaultCreatebtn').off("click");
        $('#defaultCreatebtn').on("click", function () {
            showLastPage();
            _showCreatePopup();
        });
    }
    function cancelClick() {
        $("#PinnedLocationMsg").html("");
        $("#preview_step").css('display', 'none');
        $("#third_step").css('display', 'none');
        $('#alertMessage').css('display', 'none');
        $('.alert-msg').css('display', 'none');
        $('.permissionalert-msg').css('display', 'none');
        currentPage = null;
        TotalPages = null;

        if (DefaultPageFlag == true) {
            $("#myTabContent1").css('display', 'none');
            $(".lib-section").css('display', 'block');
            PreviousPage();
        }
        else {
            $("#myTabContent1").css('display', 'none');
            $(".lib-section").css('display', 'block');
            PreviousPage();
        }
        if ($("#pinnedcheckbox").prop('checked') == true) {
            getSPPinnedLocations();
        }
    }
    function PreviousPage() {
        $("#previous").off("click");
        $("#previous").on("click", showDefaultScreen);
        $("#previous").css('background', '#04aba3');
        $("#previous").css('color', '#ffffff');
        $("#previous").removeAttr('disabled');
        $("#previous").css('cursor', 'pointer');
    }
    function showSecondScreen() {
        try {
            if (($("#listOfTemplate").find("input:checked").length > 0) || ($("#DocTemplatesBoxView").find("input:checked").length > 0)) {
                $(".new_default_page").css('display', 'none');
                $("#preview_step").css('display', 'none');
                $(".lib-section").css('display', 'block');
                PreviousPage();
            }
        }
        catch (error) {
            console.log("showSecondScreen: " + error);
        }
    };

    // save template
    this.CreateNewTemplate = function () {
        $('.alert-msg').css('display', 'none');
        $('.permissionalert-msg').css('display', 'none');

        showSavePanel();
        _showCreatePopup();
    };
    function _showCreatePopup() {
        try {
            handlePageInSavePopup(currentTemplateView);
        } catch (error) {
            console.log("_showCreatePopup: " + error);
        }
    }
    function handlePageInSavePopup(currentView) {
        currentTemplateView = currentView;
        if (currentPage == null) {
            if (currentTemplateView == "List") {
                len = $("#listOfTemplate").find("input:checked").length;
                oldfile = $("#listOfTemplate").find("input:checked").eq(Filecount).parent().attr("documentTitle");
                destinationServerRelativeUrl = $("#listOfTemplate").find("input:checked").eq(Filecount).parent().attr("serverrelativeURL");
            }
            else {
                len = $("#DocTemplatesBoxView").find("input:checked").length;
                oldfile = $("#DocTemplatesBoxView").find("input:checked").eq(Filecount).parent().attr("documentTitle");
                destinationServerRelativeUrl = $("#DocTemplatesBoxView").find("input:checked").eq(Filecount).parent().attr("serverrelativeURL");
            }
            currentPage = 1;
            TotalPages = len;
            flag = TotalPages;
            $('.Contentss')[0].innerHTML = "<b>Template Name:</b> " + oldfile;
            $('#page')[0].innerText = (Filecount + 1) + " of " + TotalPages;
            $('#popupsave').css('display', 'none');
            $('#popupnext').css('display', 'block');
            $("#popupsave").off('click');
            $("#popupsave").on('click', function () {
                createDocumentInDestLib_Treeview(true);
            });
            $("#popupnext").off('click');
            $("#popupnext").on('click', function () {
                createDocumentInDestLib_Treeview(false);
            });
        }
        if ((Filecount) == TotalPages - 1) {
            $('#popupnext').css('display', 'none');
            $('#popupsave').css('display', 'block');
        }
    }
    function handleChange() {
        var selected = [];
        $("#listOfTemplate").find("input:checked").each(function () {
            selected.push($(this));
        });
        $("#DocTemplatesBoxView").find("input:checked").each(function () {
            selected.push($(this));
        });
        if (selected.length == 1) {
            $("#previewbtn").removeAttr('disabled');
            $("#previewbtn").css('background', '#04aba3');
            $("#previewbtn").css('color', '#ffffff');
            $("#nextbtn").removeAttr('disabled');
            $("#nextbtn").css('background-color', '#04aba3');
            $("#nextbtn").css('cursor', 'pointer');
            $("#nextbtn").css('color', '#ffffff');
            $("#previewbtn").css('cursor', 'pointer');
        }
        else if (selected.length > 1) {
            $("#previewbtn").attr('disabled', 'disabled');
            $("#previewbtn").css('background', '');
            $("#previewbtn").css('color', '');
            $("#previewbtn").css('cursor', 'default');
        }
        else {
            $("#previewbtn").attr('disabled', 'disabled');
            $("#previewbtn").css('background', '');
            $("#previewbtn").css('color', '');
            $("#previewbtn").css('cursor', 'default');
            $("#nextbtn").attr('disabled', "disabled");
            $("#nextbtn").css('color', '');
            $("#nextbtn").css('background-color', '');
            $("#nextbtn").css('cursor', 'default');
        }
    }

    this.init = function () {
        openWaitDialog();
        // treeview
        $('#SPAllTreeView').hide();
        $('#SPPinnedAll').hide();
        $('#myTabContent2').on("click", "#btnConnectToSP", getListOfLibraryFromWeb);
        $('#nextbtn').on("click", showDefaultPageScreen);
        $('#nextbtn2').on("click", showDefaultPageScreen);
        $('#btnCancel').on("click", cancelClick);
        $('#skipbtn').on('click', showSecondScreen);
        $("#viewbtndropdown").off("mouseover");
        $("#viewbtndropdown").on("mouseover", function () {
            $("#filterUL").hide();
            $("#ViewUL").show();
            $("#ViewUL").css('display', 'block');
            $(document).click(function (event) {
                //if you click on anything except the modal itself or the "open modal" link, close the modal
                if ($(event.target).closest("#ViewUL").length) {
                    $('#txtTemplateSearch').val("");
                    $("#ViewUL").css('display', 'none');
                }
            });
            var mousehover = false;
            $("#ViewUL").mouseleave(function () { mousehover = true; });
            $("#ViewUL").mouseenter(function () { mousehover = false; });
            $(".input-group").hover(function () {
                if (mousehover) {
                    $("#ViewUL").hide();
                }
            });
            $(".list-item-sec").hover(function () {
                if (mousehover) {
                    $("#ViewUL").hide();
                }
            });
            $('#c-shareddrive').hover(function () {
                if (mousehover) {
                    $("#ViewUL").hide();
                }
            });
            $(".list-item-sec").click(function () {
                if (mousehover) {
                    $("#ViewUL").hide();
                }
            });
        });

        $('#txtTemplateSearch').bind('keyup', function () {
            $("#noDataFoundLbl").hide();
            $("#ViewUL").hide();
            var searchString = $(this).val();
            if (currentTemplateView == "List") {
                $("#listOfTemplate li").each(function (index, value) {
                    $('#listOfTemplate').css('display', 'block');
                    $('.viewtempdropbtn').attr('disabled', false);
                    $('.filterBtn').attr('disabled', false);
                    var currentName = $(value).text();
                    if (currentName.toUpperCase().indexOf(searchString.toUpperCase()) > -1) {
                        $(value).show();
                    } else {
                        $(value).hide();
                    }
                });
                if ($("#listOfTemplate li:visible").length === 0) {
                    $("#noDataFoundLbl").show();
                    $('#listOfTemplate').css('display', 'none');
                    $('.viewtempdropbtn').attr('disabled', true);
                    $('.filterBtn').attr('disabled', true);
                }
            }
            else {
                $("#DocTemplatesBoxView li").each(function (index, value) {
                    $('#DocTemplatesBoxView').css('display', 'block');
                    $('.viewtempdropbtn').attr('disabled', false);
                    $('.filterBtn').attr('disabled', false);
                    var currentName = $(value).text();
                    if (currentName.toUpperCase().indexOf(searchString.toUpperCase()) > -1) {
                        $(value).show();
                    } else {
                        $(value).hide();
                    }
                });
                if ($("#DocTemplatesBoxView li:visible").length === 0) {
                    $("#noDataFoundLbl").show();
                    $('#DocTemplatesBoxView').css('display', 'none');
                    $('.viewtempdropbtn').attr('disabled', true);
                    $('.filterBtn').attr('disabled', true);
                }
            }
        });

        getColumnFieldName(gSelectedView);

        // get templates

        getListofTemplateFromSourceList();

        //Treeview  --> show more/ show less events
        $(".SPTreeViewMore").on('click', function (e) {
            if ($('.treeshowmore').css('display') == 'block') {
                $('#SPFavTreeView').hide();
                $('#SPAllTreeView').show();
                $('.treeshowmore').hide();
                $('.treeshowless').show();

                // get All SiteCollections - Render
                if (listOfSiteCollectionsArray != null && listOfSiteCollectionsArray.length > 0) {
                    getAllSiteCollections_Treeview_Render();
                }
            }
            else {
                $('#SPAllTreeView').hide();
                $('#SPFavTreeView').show();
                $('.treeshowmore').show();
                $('.treeshowless').hide();
            }

            $("span").removeClass("treeselected");
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
            $("li").removeClass("treeselected");
        });

        //Pinned Location --> show more/ show less events
        $(".SPPinnedMore").on('click', function (e) {
            if ($('.pinshowmore').css('display') == 'block') {
                $('#SPPinned').hide();
                $('#SPPinnedAll').show();
                $('.pinshowmore').hide();
                $('.pinshowless').show();
            }
            else {
                $('#SPPinned').show();
                $('#SPPinnedAll').hide();
                $('.pinshowmore').show();
                $('.pinshowless').hide();
            }

            $("span").removeClass("treeselected");
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
            $("li").removeClass("treeselected");
        });

        $("#nextbtn").click(function (e) {
            checkPinnedLocationListExist();
            if (listOfSiteCollectionsArray == null || listOfSiteCollectionsArray.length == 0) {

                // get Pinned Locations
                getSPPinnedLocations();

                // get Fav SiteCollections
                getFavSites_Treeview_Render();

            }
        });
        $("#nextbtn2").click(function (e) {
            checkPinnedLocationListExist();
            if (listOfSiteCollectionsArray == null || listOfSiteCollectionsArray.length == 0) {

                // get Pinned Locations
                getSPPinnedLocations();

                // get Fav SiteCollections
                getFavSites_Treeview_Render();
            }
        });

        //Treeview expand/collapse class
        var toggler = document.getElementsByClassName("caretCustom");
        for (var i = 0; i < toggler.length; i++) {
            toggler[i].addEventListener("click", function () {
                this.parentElement.querySelector(".active").classList.toggle("nested");
                this.classList.toggle("caret-down");
            });
        }
    };

    function pinnedLocations_click(e) {
        $("span").removeClass("treeselected");
        $("li").removeClass("treeselected");
        $(e).addClass("treeselected");

        saveType = "pinned";

        $('#createFile').removeAttr('disabled');
        $("#createFile").css('background-color', '#04aba3');
        $("#createFile").css('cursor', 'pointer');
        $("#createFile").css('color', '#ffffff');
    }

    /////////////// get All Fav SiteCollections - Render/////////////////////
    function getFavSites_Treeview_Render() {
        var favData = [];
        var treeViewHTMLFav = "";
        var hasNoFav = true;

        getAllSiteCollection_Treeview(SPURL).then(function (sdata) {
            if (listOfSiteCollectionsArray != null && listOfSiteCollectionsArray.length > 0) {

                var sitesCollectionArray = listOfSiteCollectionsArray.filter(function (objFav) {
                    return objFav.siteUrl
                });

                getMyFollowedSites().done(function (favResults) {
                    for (var i = 0; i < sitesCollectionArray.length; i++) {
                        for (var j = 0; j < favResults.length; j++) {
                            if (sitesCollectionArray[i].siteUrl === favResults[j].Url) {
                                favData.push(favResults[j]);
                            }
                        }
                    }
                    treeViewHTMLFav = "<ul id='treeviewUL'>";
                    for (var i = 0; i < favData.length; i++) {
                        hasNoFav = false;
                        var siteKey = favData[i].Title + "_" + i;

                        treeViewHTMLFav += " <li id='" + favData[i].Title + "'><span class='caretCustom caret-down treeSpan' ><div class='type' hidden>sitecollection</div>"
                        treeViewHTMLFav += " <div class='level' hidden> " + 0 + "</div > <div class='sitekey' hidden> " + siteKey + "</div> <div class='siteurl' hidden> " + favData[i].Url + "</div> "
                        treeViewHTMLFav += " <div class='sitetitle' hidden> " + favData[i].Title + "</div > <div class='siteId' hidden> " + favData[i].ItemReference.SiteId + "</div>"
                        treeViewHTMLFav += " <a href= '#' > <i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i><span class='singleLineEllipse'>" + favData[i].Title + "</span></a ></span > "; // main site collection li level -1 open
                        treeViewHTMLFav += " <ul class='active nested' id='" + siteKey + "'>"; // ul level - 2 open
                        treeViewHTMLFav += " </ul>"; // Site shared documents
                        treeViewHTMLFav += " </li>"; // Site  
                    }
                    treeViewHTMLFav += "</ul>"; //treeviewUL

                    if (hasNoFav) {
                        treeViewHTMLFav = "<p> No favorite sites are found..!! </p>";
                    }

                    $('#SPFavTreeView').html(treeViewHTMLFav);
                    $(".treeSpan").off('click');
                    $(".treeSpan").on('click', function () {
                        getAllSites_Treeview_click(this);
                    });

                    closeWaitDialog();
                }).fail(function (jqXHR, textStatus) {
                    closeWaitDialog();
                });
            }
        });
    }

    /////////////// get All SiteCollections - Render /////////////////////
    function getAllSiteCollections_Treeview_Render() {
        if (listOfSiteCollectionsArray == null || listOfSiteCollectionsArray.length == 0) {
            getAllSiteCollection_Treeview(SPURL).then(function (sdata) {
                treeviewBind();
            });
        }
        else {
            treeviewBind();
        }
    }

    ////////// Treeview Bind //////////////
    function treeviewBind() {
        if (listOfSiteCollectionsArray != null && listOfSiteCollectionsArray.length > 0) {
            var treeViewHTML = "<ul id='treeviewUL'>";

            listOfSiteCollectionsArray.map(function (sitecollection, index) {
                treeViewHTML += " <li id='" + sitecollection.siteTitle + "'><span class='caretCustom caret-down treeSpan' ><div class='type' hidden>sitecollection</div><div class='level' hidden> " + 0 + "</div>"
                treeViewHTML += " <div class='sitekey' hidden> " + sitecollection.siteKey + "</div > <div class='siteurl' hidden> " + sitecollection.siteUrl + "</div> <div class='sitetitle' hidden> " + sitecollection.siteTitle + "</div>"
                treeViewHTML += " <div class='siteId' hidden> " + sitecollection.siteId + "</div > <a href='#'><i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i><span class='singleLineEllipse'>" + sitecollection.siteTitle + "</span></a></span > "; // main site collection li level -1 open
                treeViewHTML += " <ul class='active nested' id='" + sitecollection.siteKey + "'>"; // ul level - 2 open
                treeViewHTML += " </ul>"; // Site shared documents
                treeViewHTML += " </li>"; // Site  
            });

            treeViewHTML += "</ul>"; //treeviewUL
            $('#SPTreeView').html(treeViewHTML);

            $(".treeSpan").click(function () {
                getAllSites_Treeview_click(this);
            });
        }
    }

    /////////////// get All SubSites Of SiteCollection - render/////////////////////
    function getAllSites_Treeview_click(e) {
        if ($(e).siblings(".nested").length > 0) {

            $("span").removeClass("treeselected");
            $("li").removeClass("treeselected");
            saveType = "";
            $(e).addClass("treeselected");

            var siteUrl = $(e).find('.siteurl').text().trim();
            var siteTitle = $(e).find('.sitetitle').text().trim();
            var siteKey = $(e).find('.sitekey').text().trim();
            var siteId = $(e).find('.siteId').text().trim();
            var level = $(e).find('.level').text().trim();
            var tenantName = SPURL.substr(8, SPURL.length);
            var selectedSiteURL = siteUrl.split("/sites/")[1];
            var cutmAttr = "";
            if (siteUrl.indexOf("/sites/") < 0 && tenantName != siteUrl.substr(8, SPURL.length)) {
                selectedSiteURL = siteUrl.substring(siteUrl.lastIndexOf("/") + 1, siteUrl.length);
                cutmAttr = "rootsubsites";
            }
            var rootSiteURL = siteUrl.split(tenantName)[1];
            var selectedSiteURLForLib = siteUrl.split("/sites/")[1] == undefined ? rootSiteURL : selectedSiteURL;
            isRoot = siteUrl.indexOf("/sites/") < 0 ? true : false;
            level = parseInt(level) + 1;

            getAllsubSitesOfSiteCollection_Treeview(selectedSiteURL, siteUrl, isRoot, cutmAttr, siteKey, level, siteId)
                .then(function (favData) {

                    var liHTML = "";
                    var parentsiteKey = "";

                    if (listOfSitesArray.length > 0) {

                        parentsiteKey = listOfSitesArray[0].parentSiteKey;

                        liHTML += " <li id='" + parentsiteKey + "'><a href='#' class='treeDocLib'><div class='appWebUrl' hidden> " + siteUrl + "</div><div class='listName' hidden> " + level + "</div><div class='selectedLibURL' hidden> " + level + "</div><i class='ms-Icon ms-Icon--FabricDocLibrary' aria-hidden='true'></i></a></li>"; // main site collection shared doc li level - 2 open/close                    

                        for (var i = 0; i < listOfSitesArray.length; i++) {
                            if (listOfSitesArray[i].hasSubSite) {
                                var siteName = listOfSitesArray[i].SubSiteName.trim();
                                var siteURL = listOfSitesArray[i].SubSiteURL.trim();
                                var siteKey = listOfSitesArray[i].SiteKey;
                                var siteId = listOfSitesArray[i].siteId;
                                var ParentSiteURL = listOfSitesArray[i].ParentSiteURL;
                                liHTML += " <li id='" + siteName + "'><span class='caretCustom caret-down treeSpan1' ><div class='type' hidden>site</div><div class='level' hidden> " + level + "</div><div class='sitekey' hidden> " + siteKey + "</div><div class='siteurl' hidden> " + siteURL + "</div><div class='sitetitle' hidden> " + siteName + "</div><div class='siteId' hidden> " + siteId + "</div> <a href='#'><i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i><span class='singleLineEllipse'>" + siteName + "</span></a></span > "; // subsite level li - 2 open
                                liHTML += " <ul class='active nested' id='" + siteKey + "'>"; // Sub Site 1 ul level - 3 open  
                                liHTML += " </ul>"; // Sub Site 1 - Shared Documents - ul level - 3 close
                                liHTML += " </li>"; // Sub Site 1 li   level - 2 close 
                            }
                        }
                        if (parentsiteKey != "") {
                            var skey = parentsiteKey.trim();
                            $("ul[id='" + skey + "']").html(liHTML);
                            $("ul[id='" + skey + "']").removeClass("nested");
                            $("ul[id='" + skey + "']").prev('span').removeClass("caret-down");
                            $("ul[id='" + skey + "']").prev('span').addClass("expanded");

                            $(document).off("click", ".treeSpan1");
                            $(document).on("click", ".treeSpan1", function (event) {
                                getAllSites_Treeview_click($(this));
                            });

                            $(document).off("click", ".treeDocLib");
                            $(document).on("click", ".treeDocLib", function (event) {
                                getAllFoldersFromLibrary_Treeview_click($(this));
                            });

                        }
                    }
                    listOfSitesArray = [];
                });
        }
        else {

            var docLibKey = $(e).find('.sitekey').text().trim();

            if (docLibKey != "") {
                $("ul[id='" + docLibKey + "']").addClass("nested");
                $("ul[id='" + docLibKey + "']").prev('span').addClass("caret-down");
            }
        }
    }

    // get All Folders From Document Library - render
    function getAllFoldersFromLibrary_Treeview_click(e) {
        if ($(e).siblings(".nested").length > 0) {

            saveType = "";
            $("span").removeClass("treeselected");
            $("li").removeClass("treeselected");
            $(e).addClass("treeselected");

            var selectedLibURL = $(e).find('.selectedLibURL').text().trim();
            var listName = $(e).find('.listName').text().trim();
            var displayName = $(e).find('a').text().trim();
            var appWebUrl = $(e).find('.appWebUrl').text().trim();
            var docLibKey = $(e).find('.docLibKey').text();
            var siteName = "";

            var type = $(e).find('.type').text();
            if (type == "documentlibrary" || type == "folder") {
                $('#createFile').removeAttr('disabled');
                $("#createFile").css('background-color', '#04aba3');
                $("#createFile").css('cursor', 'pointer');
                $("#createFile").css('color', '#ffffff');
            }
            else {
                $("#createFile").attr('disabled', 'disabled');
                $("#createFile").css('background-color', '');
                $("#createFile").css('cursor', 'default');
            }

            getAllFoldersFromLibrary_Treeview(appWebUrl, displayName, selectedLibURL, siteName).then(function (folData) {

                if (listOfFoldersArray.length > 0) {
                    var liHTML = "";
                    for (var i = 0; i < listOfFoldersArray.length; i++) {
                        var folderName = listOfFoldersArray[i].folderName.trim();
                        var folderURL = listOfFoldersArray[i].folderURL.trim();
                        var folderKey = folderName + "_" + i;;

                        liHTML += " <li id='" + folderKey + "'><span class='treeFolder'><div class='appWebUrl' hidden> " + appWebUrl + "</div><div class='folderName' hidden> " + folderName + "</div><div class='type' hidden>folder</div><div class='folderURL' hidden> " + folderURL + "</div><div class='folderKey' hidden> " + folderKey + "</div><a href='#'><i class='ms-Icon ms-Icon--FabricFolderFill' aria-hidden='true'></i><span class='singleLineEllipse'>" + folderName + "</span></a></span>"; // Sub Site 1 - Shared Documents 1 //level - 3 open/close
                        liHTML += " <ul class='active ' id='" + folderKey + "'>"; // folder ul level - open 
                        liHTML += " </ul>"; // folder - ul - close
                        liHTML += " </li>";
                    }

                    listOfFoldersArray = [];

                    if (docLibKey != "") {
                        var skey = docLibKey.trim();
                        $("ul[id='" + skey + "']").html(liHTML);
                        $("ul[id='" + skey + "']").removeClass("nested");
                        $("ul[id='" + skey + "']").prev('span').removeClass("caret-down");

                        $(document).off("click", ".treeFolder");
                        $(document).on("click", ".treeFolder", function (event) {
                            Folder_Treeview_click($(this));
                        });

                    }
                }
                else {
                    var txt = docLibKey.trim();
                    $("div.docLibKey:contains(" + txt + ")").parent().removeClass("caretCustom");
                    $("div.docLibKey:contains(" + txt + ")").parent().removeClass("caret-down");

                    $("span").removeClass("treeselected");
                    $("li").removeClass("treeselected");
                    saveType = "";
                    $("div.docLibKey:contains(" + txt + ")").parent().addClass("treeselected");
                }
            });
        }
        else {

            var docLibKey = $(e).find('.docLibKey').text().trim();

            if (docLibKey != "") {
                $("ul[id='" + docLibKey + "']").addClass("nested");
                $("ul[id='" + docLibKey + "']").prev('span').addClass("caret-down");
            }
        }
    }

    // Folder click event for selection
    function Folder_Treeview_click(e) {
        saveType = "";
        $("span").removeClass("treeselected");
        $("li").removeClass("treeselected");
        $(e).addClass("treeselected");
    }

    // get All Folders From Document Library - Rest API
    function getAllFoldersFromLibrary_Treeview(appWebUrl, listName, selectedLibURL, siteName) {
        openWaitDialog();
        var dfdLib = $.Deferred();
        var apiURL = "";
        try {
            if (appWebUrl == "" || appWebUrl == null) {
                apiURL = SPURL + "/_api/web/lists/getbytitle('" + listName + "')/items?$expand=Folder,File&$select=ID,Title,FileLeafRef,Folder/ServerRelativeUrl";
            }
            else if (appWebUrl) {
                apiURL = appWebUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?$expand=Folder,File&$select=ID,Title,FileLeafRef,Folder/ServerRelativeUrl";
            }
            else {
                apiURL = SPURL + "/sites/" + siteName + "/_api/web/lists/getbytitle('" + listName + "')/items?$expand=Folder,File&$select=ID,Title,FileLeafRef,Folder/ServerRelativeUrl";
            }
            $.ajax({
                url: apiURL,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                success: function (data) {
                    getFolderSchema(data.d.results);
                },
                error: function (data) {
                    console.log(JSON.stringify(data));
                    closeWaitDialog();
                    dfdLib.reject(listOfFoldersArray);
                }
            });
            function getFolderSchema(data) {
                if (data.length >= 0) {
                    for (var i = 0; i < data.length; i++) {
                        if (data[i].Folder.ServerRelativeUrl) {
                            listOfFoldersArray.push({ "appWebUrl": appWebUrl, "listName": listName, "selectedLibrarywebURL": selectedLibURL, "folderURL": data[i].Folder.ServerRelativeUrl, "folderName": data[i].FileLeafRef })
                        }
                    }
                }

                dfdLib.resolve(listOfFoldersArray);
                closeWaitDialog();
            }
        }
        catch (error) {
            console.log("getAllFoldersFromLibrary_Treeview: " + error);
            dfdLib.reject(listOfFoldersArray);
            closeWaitDialog();
        }
        return dfdLib.promise();
    };

    // get All Libraries of Sites/subsites - Rest API
    function getAllLibrayFromSite_Treeview(site, parentsiteKey, siteKey, cutmAttr, siteId) {
        openWaitDialog();
        var dfdGET = $.Deferred();
        try {
            var tenantName = SPURL.substr(8, SPURL.length);
            var GraphAPI = "";
            if (GraphAPIToken) {
                if (site == undefined) {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/lists";
                }
                else if (site.split("/sites/")[1] == undefined & (isRoot)) {
                    if (site == "Root") {
                        if (cutmAttr == "rootsubsites") {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + siteId + "/lists";
                        }
                        else {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/lists";
                        }

                    }
                    else {
                        if (cutmAttr == "rootsubsites") {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + siteId + "/lists";
                        }
                        else {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":" + site + ":/lists";
                        }
                    }
                }
                else {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + site + ":/lists";
                }
                $.ajax({
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json");
                    },
                    type: "GET",
                    url: GraphAPI,
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + GraphAPIToken
                    }
                }).done(function (response) {
                    var result = response.value;
                    for (var i = 0; i < result.length; i++) {
                        if (result[i].list.template === "documentLibrary" && !result[i].list.hidden) {
                            var docURL = result[i].webUrl;
                            var webURL = docURL.substring(0, docURL.lastIndexOf("/"));
                            var docLibKey = result[i].name + "_" + i;
                            listOfDocLibsArray.push({
                                "site": site, "appWebUrl": webURL, "siteURL": result[i].webUrl, "displayName": result[i].displayName, "name": result[i].name, "parentsiteKey": parentsiteKey, "siteKey": siteKey, "docLibKey": docLibKey
                            })
                        }
                    }
                    dfdGET.resolve(listOfDocLibsArray);
                    closeWaitDialog();

                }).fail(function (response) {
                    console.log('error:- ' + response.responseText);
                    dfdGET.reject(listOfDocLibsArray);
                    closeWaitDialog();
                });
            }
        } catch (error) {
            console.log("callAjaxGet: " + error);
            dfdGET.reject(listOfDocLibsArray);
            closeWaitDialog();
        }
        return dfdGET.promise();
    }

    // get All SubSites Of SiteCollection - Rest API & Libray of Site Render
    function getAllsubSitesOfSiteCollection_Treeview(siteCollection, siteCollectionURL, isCollection, customAttr, parentsiteKey, level, siteId) {
        openWaitDialog();
        var dfd = $.Deferred();
        var tenantName = SPURL.substr(8, SPURL.length);
        var GraphAPI = "";
        if (GraphAPIToken) {
            if (isCollection) {
                if (siteCollection != undefined) {
                    if (customAttr == "rootsubsites") {
                        GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + siteId + "/sites";
                    }
                    else {
                        GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + siteCollection + ":/sites";
                    }
                }
                else {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/sites/";
                }
            }
            else {
                if (siteCollection.split(":")[1] == undefined) {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + siteCollection + ":/sites";
                }
                else {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/" + siteCollection + "/sites";
                }
            }
            var iSiteId = siteId;
            $.ajax({
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json");
                },
                type: "GET",
                url: GraphAPI,
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + GraphAPIToken
                }
            }).done(function (response) {
                var result = response.value;
                if (result) {
                    var siteId = iSiteId;
                    getAllLibrayFromSite_Treeview(siteCollection, parentsiteKey, siteKey, customAttr, siteId).then(function (libData) {
                        if (libData && libData.length > 0) {
                            openWaitDialog();
                            var liHTML = "";
                            for (var i = 0; i < listOfDocLibsArray.length; i++) {
                                liHTML += " <li id='" + parentsiteKey + "'><span class='caretCustom caret-down treeDocLib' ><div class='type' hidden>documentlibrary</div><div class='appWebUrl' hidden> " + listOfDocLibsArray[i].appWebUrl + "</div>";
                                liHTML += "<div class='listName' hidden> " + listOfDocLibsArray[i].name + "</div><div class='selectedLibURL' hidden> " + listOfDocLibsArray[i].siteURL + "</div>";
                                liHTML += "<div class='docLibKey' hidden> " + listOfDocLibsArray[i].docLibKey + "</div><a href='#'><i class='ms-Icon ms-Icon--FabricDocLibrary' aria-hidden='true'></i><span class='singleLineEllipse'> " + listOfDocLibsArray[i].displayName + "</span></a></span>"; // main site collection shared doc li level - 2 open/close                    
                                liHTML += "<ul class='active nested' id='" + listOfDocLibsArray[i].docLibKey + "'>";
                                liHTML += "</ul>";
                                liHTML += "</li>";
                            }
                            if (parentsiteKey != "") {
                                var skey = parentsiteKey.trim();
                                $("li[id='" + skey + "']").html(liHTML);
                            }
                            listOfDocLibsArray = [];
                            closeWaitDialog();
                        }
                        else {

                            var txt = parentsiteKey.trim();
                            $("div.sitekey:contains(" + txt + ")").parent().removeClass("caretCustom");
                            $("div.sitekey:contains(" + txt + ")").parent().removeClass("caret-down");
                            $("div.sitekey:contains(" + txt + ")").parent().next().addClass("nested");

                        }
                    });

                    var hasNoSubsite = true;
                    if (result.length > 0) {
                        for (var i = 0; i < result.length; i++) {
                            if (result[i].webUrl.indexOf(tenantName) > -1) {
                                hasNoSubsite = false;
                                var rootsubSite = result[i].webUrl.split(tenantName)[1] + ":";
                                var subsitesVar = result[i].webUrl.split("/sites/")[1] == undefined ? rootsubSite : result[i].webUrl.split("/sites/")[1];
                                if (isCollection) {
                                    siteCollection = "Root";
                                }
                                var siteKey = result[i].name + "_" + i;
                                var siteId = result[i].id;
                                listOfSitesArray.push({
                                    "ParentSite": siteCollection, "ParentSiteURL": siteCollectionURL, "hasSubSite": true, "SubSiteDisplayName": result[i].name, "SubSiteName": result[i].name, "SubSiteURL": result[i].webUrl, "IsRoot": isCollection, "parentSiteKey": parentsiteKey, "SiteKey": siteKey, "level": level, "siteId": siteId
                                })
                            }
                        }
                        if (hasNoSubsite) {
                            listOfSitesArray.push({
                                "ParentSite": siteCollection, "ParentSiteURL": siteCollectionURL, "hasSubSite": false, "SubSiteDisplayName": "", "SubSiteName": "", "SubSiteURL": "", "IsRoot": isCollection, "parentSiteKey": parentsiteKey, "SiteKey": "", "level": 0, "siteId": siteId
                            })
                        }
                    }
                    else {
                        listOfSitesArray.push({
                            "ParentSite": siteCollection, "ParentSiteURL": siteCollectionURL, "hasSubSite": false, "SubSiteDisplayName": "", "SubSiteName": "", "SubSiteURL": "", "IsRoot": isCollection, "parentSiteKey": parentsiteKey, "SiteKey": "", "level": 0, "siteId": siteId
                        })
                    }
                    dfd.resolve(listOfSitesArray);
                }
                else {
                    dfd.resolve(listOfSitesArray);
                    closeWaitDialog();
                }
            }).fail(function (response) {
                dfd.reject(listOfSitesArray);
                console.log('error:- ' + response.responseText);
                closeWaitDialog();
            });
        }
        return dfd.promise();
    }

    // get All SiteCollections - Rest API
    function getAllSiteCollection_Treeview(webUrl) {
        openWaitDialog();
        var dfd = $.Deferred();
        try {
            var tempArray = webUrl.split(".");
            var mySitePath = tempArray[0] + "-my." + tempArray[1] + "." + tempArray[2] + "/personal"; //"https://binaryrepublik516-my.sharepoint.com/personal";
            var url = webUrl + "/_api/search/query?querytext='NOT Path:" + mySitePath + "/* contentclass:sts_site'&rowLimit=499&TrimDuplicates=false";
            callAjaxGet(url).done(function (data) {
                var resultsCount = data.d.query.PrimaryQueryResult.RelevantResults.RowCount;
                RowResult = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
                for (var i = 0; i < resultsCount; i++) {
                    var row = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results[i];
                    var siteUrl = row.Cells.results[6].Value;
                    var siteTitle = row.Cells.results[3].Value;
                    var siteKey = siteTitle + "_" + i;
                    if (siteUrl.indexOf("/portals/") < 0) {
                        listOfSiteCollectionsArray.push({ "siteUrl": siteUrl, "siteTitle": siteTitle, "siteKey": siteKey, "siteId": "" })
                    }
                }
                dfd.resolve(listOfSiteCollectionsArray);
            });
        }
        catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };

    // Treeview --> Create document in lib
    function createDocumentInDestLib_Treeview(checkNext) {
        try {
            openWaitDialog();
            $('#alertMessage').css('display', 'none');
            $('.alert-msg').css('display', 'none');
            $('.permissionalert-msg').css('display', 'none');
            var docName = $('#txtNewFileName').val();
            docName = docName.trim();
            var docnameLen = docName.length;
            if (docName == "" || docName == null || docName.match('%') || docName.match('"')
                || docName.match('\'') || docName.match(';') || docName.match('#')) {
                $('#alertMessage').css('display', 'block');
                closeWaitDialog();
                return false;
            } else {
                var selectedSiteRelativeURL = ""
                var webRedirectURL = "";// this is used to create open Document URL
                var tokenUrl = "";
                $('#txtDocumentName').val("");
                var FolderRelativePath = "";
                var chkDocTitle = "";
                var DocumentInternalName = "";
                var DocLibraryUrl = "";
                var DocumentLibraryName = "";
                var FolderName = "";
                var appWebUrl = "";
                var type = "";

                // start here to save doc
                if ($(".treeselected") == null || $(".treeselected").length == 0) {
                    $("#PinnedLocationMsg").html("");
                    pinnedString = "<p>Please select at least one location.</p>";
                    $("#PinnedLocationMsg").append(pinnedString);
                    closeWaitDialog();
                    return;
                }
                else if ($(".treeselected").length == 1) {
                    type = $(".treeselected").find('.type').text().trim();
                    if (saveType != "pinned") {
                        appWebUrl = $(".treeselected").find('.appWebUrl').text().trim();
                        selectedSiteRelativeURL = $(".treeselected").find('.appWebUrl').text().trim();
                        tokenUrl = selectedSiteRelativeURL
                        webRedirectURL = selectedSiteRelativeURL;
                        selectedSiteRelativeURL = selectedSiteRelativeURL ? selectedSiteRelativeURL.replace(SPURL, "") : "";

                        if (type == "folder") {
                            FolderRelativePath = $(".treeselected").find('.folderURL').text().trim();
                            FolderRelativePath = FolderRelativePath.replace(SPURL, "");
                            FolderName = $(".treeselected").find('.folderName').text().trim();
                            DocLibraryUrl = FolderRelativePath;
                            DocumentLibraryName = FolderName;
                        }
                        if (type == "documentlibrary") {
                            DocLibraryUrl = $(".treeselected").find('.selectedLibURL').text().trim();
                            DocLibraryUrl = DocLibraryUrl.replace(SPURL, "");
                            DocumentLibraryName = $(".treeselected").find('.listName').text().trim();
                        }
                    }
                    else {
                        appWebUrl = SPURL + "/" + $(".treeselected").find('.siteurl').text().trim();
                        selectedSiteRelativeURL = $(".treeselected").find('.pinurl').text().trim();
                        tokenUrl = appWebUrl
                        webRedirectURL = appWebUrl;

                        if (type == "folder") {
                            FolderRelativePath = $(".treeselected").find('.pinurl').text().trim();
                            FolderRelativePath = FolderRelativePath.replace(SPURL, "");
                            FolderName = $(".treeselected").find('.pinname').text().trim();
                            DocLibraryUrl = FolderRelativePath;
                            DocumentLibraryName = FolderName;
                        }
                        if (type == "documentlibrary") {
                            DocLibraryUrl = $(".treeselected").find('.pinurl').text().trim();
                            DocLibraryUrl = DocLibraryUrl.replace(SPURL, "");
                            DocumentLibraryName = $(".treeselected").find('.pinname').text().trim();
                        }
                    }
                }
                else {
                    $("#PinnedLocationMsg").html("");
                    pinnedString = "<p>Please select only one location.</p>";
                    $("#PinnedLocationMsg").append(pinnedString);
                    return;
                }
                chkDocTitle = oldfile;
                if (FolderRelativePath === "") {
                    webRedirectURL += "/_api/web/GetFolderByServerRelativeUrl('" + DocLibraryUrl + "/')";
                }
                else {
                    webRedirectURL += "/_api/web/GetFolderByServerRelativeUrl('" + FolderRelativePath + "/')";
                }

                var docName = $('#txtNewFileName').val();
                var folderName = $('#SPDocFolders').val();
                var docnameLen = docName.length;
                var newFileName = docName + docsNodeNewFileExtention;
                var destServerRelURL = destinationServerRelativeUrl;
                if (chkDocTitle !== "") {
                    createDocument(newFileName, webRedirectURL, chkDocTitle, tokenUrl, destServerRelURL).then(function (data) {
                        openEditProperties(data, selectedSiteRelativeURL, tokenUrl).then(function (data) {
                            var documentURL = data.d.ServerRedirectedEmbedUri;
                            var editFormURl = "";
                            documentURL = documentURL.replace("=interactivepreview", "=edit");
                            if (flag == 1) {
                                editFormURl = SPURL + selectedSiteRelativeURL + DocumentInternalName + '/Forms/EditForm.aspx?ID=' + data.d.ID;
                                if (platform == "PC") {
                                    if (FolderRelativePath != null && FolderRelativePath != "") {
                                        getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href='ms-word:ofe|u|" + SPURL + FolderRelativePath + "/" + newFileName + "'> " + newFileName + ". </a></p>\n";
                                    }
                                    else {
                                        getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href='ms-word:ofe|u|" + SPURL + DocLibraryUrl + "/" + newFileName + "'> " + newFileName + ". </a></p>\n";
                                    }

                                }
                                else {
                                    getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href=" + documentURL + "> " + newFileName + ". </a></p>\n";
                                }
                                $("#DocumentUrls").append(getdocumentUrlsString);
                                closeWaitDialog();
                            }
                            else {
                                getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href=" + documentURL + "> " + newFileName + ". </a></p>\n";
                                $("#DocumentUrls").append(getdocumentUrlsString);
                                closeWaitDialog();
                            }
                            console.log("File Saved! " + Filecount);
                            if (Filecount == TotalPages) {
                                $("#third_step").find("input").attr("disabled", 'disabled');
                                currentPage = null, TotalPages = null;
                                len = 0, Filecount = 0, oldfile = ""; destinationServerRelativeUrl = '';
                            }
                        }, function (openEditerFileFail) {
                            console.log(openEditerFileFail);
                        });
                        Filecount = Filecount + 1;
                        if (Filecount < TotalPages) {
                            if (currentTemplateView == "List") {
                                oldfile = $("#listOfTemplate").find("input:checked").eq(Filecount).parent().attr("documentTitle");
                                destinationServerRelativeUrl = $("#listOfTemplate").find("input:checked").eq(Filecount).parent().attr("serverrelativeURL");
                            }
                            else {
                                oldfile = $("#DocTemplatesBoxView").find("input:checked").eq(Filecount).parent().attr("documentTitle");
                                destinationServerRelativeUrl = $("#DocTemplatesBoxView").find("input:checked").eq(Filecount).parent().attr("serverrelativeURL");
                            }
                            $('.Contentss')[0].innerHTML = "<b>Template Name:</b> " + oldfile;
                            $('#page')[0].innerText = (Filecount + 1) + " of " + TotalPages;
                            $('#txtNewFileName').val("");
                            $('#txtNewFileName').attr('value', "");
                            $('#txtNewFileName').removeAttr('value');
                        }
                        if ((Filecount) == TotalPages - 1) {
                            $('#popupnext').css('display', 'none');
                            $('#popupsave').css('display', 'block');
                        }
                        //remove save button on its last click.
                        if ((Filecount) == TotalPages) {
                            $('#popupsave').css('background-color', 'rgb(221, 221, 221)');
                            $('#popupsave').css('pointer-events', 'none');
                            $('#txtNewFileName').val("");
                        }
                        ////// auto pin location ////

                        if (checkNext != false) {
                            if (saveType != "pinned") {
                                var siteurl = tokenUrl.replace(SPURL, "");
                                if (siteurl == "") {
                                    siteurl = "/";
                                }
                                var itemArray = {
                                    '__metadata': {
                                        'type': 'SP.Data.DocsNodePinnedLocationsListItem'
                                    },
                                    "DocumentLibrary": DocumentLibraryName,
                                    "DocumentLibraryURL":
                                    {
                                        '__metadata': { 'type': 'SP.FieldUrlValue' },
                                        'Description': DocLibraryUrl,
                                        'Url': DocLibraryUrl
                                    },
                                    "PinnedType": type,
                                    "SiteURL":
                                    {
                                        '__metadata': { 'type': 'SP.FieldUrlValue' },
                                        'Description': siteurl,
                                        'Url': siteurl
                                    },
                                };
                                if ($("#pinnedcheckbox").prop('checked') == true) {
                                    $("#PinnedLocationMsg").html("");
                                    checkExistingPinned(DocLibraryUrl).then(function (data) {
                                        if (!IsPinnedLocation) {
                                            formSaveNotes(itemArray);
                                        }
                                        else {
                                            pinnedString = "<p>Selected location is already pinned.</p>";
                                            $("#PinnedLocationMsg").append(pinnedString);
                                        }
                                    });
                                }
                            }
                        }
                        /////

                    }, function (errorMsg) {
                        console.log(errorMsg);
                        closeWaitDialog();
                    });
                }
                else {
                    closeWaitDialog();
                }
            }
        }
        catch (error) {
            console.log("createDocumentInDestLib_Treeview : " + error);
            closeWaitDialog();
        }
    }

    // Treeview --> get Pinned Locations
    function getSPPinnedLocations() {
        var dfd = $.Deferred();
        openWaitDialog();
        var liHTML = "";
        var liHTMLAll = "";
        $('#SPPinned').html(liHTML);
        $('#SPPinnedAll').html(liHTMLAll);
        try {
            var UserName = localStorage.getItem('userDisplayName');
            if (SPToken) {
                var siteConfigListUrl = SPURL + "/_api/web/lists/getbytitle('" + sitePinnedLocations + "')/items?$select=ID,DocumentLibrary,DocumentLibraryURL,PinnedType,SiteURL,Author/Title&$expand=Author&$filter=Author/Title eq '" + UserName + "' &$orderby=ID%20desc";
                $.ajax({
                    url: siteConfigListUrl,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                    success: function (data) {
                        if (data.d.results.length > 0) {
                            var arrayResult = data.d.results
                            $('.pinshowmore').hide();
                            liHTML = "<ul class='pin_location'>";
                            for (var i = 0; i < arrayResult.length; i++) {
                                if (i < 3) {
                                    liHTML += "<li class='pinnedselected pinnedselect'>";
                                    liHTML += "       <div class='pindoc'><div hidden class='siteurl'>" + arrayResult[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + arrayResult[i].PinnedType + "</div><div hidden class='pinname'>" + arrayResult[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + arrayResult[i].DocumentLibraryURL.Description.trim() + "</div>";
                                    liHTML += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                    liHTML += "            <h4>" + arrayResult[i].DocumentLibrary + "</h4>";
                                    liHTML += "            <span class='path'>" + arrayResult[i].DocumentLibraryURL.Description.trim() + "</span>";
                                    liHTML += "        </div>";
                                    liHTML += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + arrayResult[i].ID + "</div>";
                                    liHTML += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                    liHTML += "        </a>";
                                    liHTML += "    </li>";
                                }
                                else {
                                    if (liHTMLAll == "") {
                                        liHTMLAll += liHTML;
                                        if ($('#SPPinnedAll').css('display') == 'none') {
                                            $('.pinshowmore').show();
                                        }
                                    }
                                    liHTMLAll += "<li class='pinnedselected pinnedselect'>";
                                    liHTMLAll += "       <div class='pindoc'><div hidden class='siteurl'>" + arrayResult[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + arrayResult[i].PinnedType + "</div><div hidden class='pinname'>" + arrayResult[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + arrayResult[i].DocumentLibraryURL.Description.trim() + "</div>";
                                    liHTMLAll += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                    liHTMLAll += "            <h4>" + arrayResult[i].DocumentLibrary + "</h4>";
                                    liHTMLAll += "            <span class='path'>" + arrayResult[i].DocumentLibraryURL.Description.trim() + "</span>";
                                    liHTMLAll += "        </div>";
                                    liHTMLAll += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + arrayResult[i].ID + "</div>";
                                    liHTMLAll += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                    liHTMLAll += "        </a>";
                                    liHTMLAll += "    </li>";
                                }
                            }
                            liHTML += "</ul>";
                            liHTMLAll += "</ul>";
                            if (arrayResult.length == 3) {
                                $('.pinshowless').hide();
                                $('#SPPinnedAll').css('display', 'none');
                                $('#SPPinned').css('display', 'block');
                            }
                            $('#SPPinned').html(liHTML);
                            $('#SPPinnedAll').html(liHTMLAll);
                            $("#createFile").attr('disabled', 'disabled');
                            $("#createFile").css('background-color', '');
                            $("#createFile").css('cursor', 'default');
                            $(".removepinned").click(function () {
                                removeExistingPinned(this);
                            });

                            $(".pinnedselected").click(function () {
                                pinnedLocations_click(this);
                            });

                        }
                        else {
                            liHTML = "<p>    No pinned locations are found..!! </p>";
                            $('#SPPinned').html(liHTML);
                            $('#SPPinnedAll').html(liHTMLAll);
                        }
                        dfd.resolve(data);
                        closeWaitDialog();
                    },
                    error: function (errordata) {
                        console.log(errordata);
                        dfd.reject();
                        closeWaitDialog();
                    }
                });
            }
        }
        catch (error) {
            console.log("getSPPinnedLocations Ajax Call Error: " + error.responseText);
            dfd.reject();
        }
        return dfd.promise();
    }

    // Treeview --> remove Pinned Locations
    function removeExistingPinned(e) {
        var pinnedId = $(e).find('.pinnedId').text().trim();
        var dfd = $.Deferred();
        try {
            if (SPToken) {
                var siteConfigListUrl = SPURL + "/_api/web/lists/getbytitle('" + sitePinnedLocations + "')/items(" + pinnedId + ")";
                $.ajax({
                    url: siteConfigListUrl,
                    method: "POST",
                    headers: { "Accept": "application/json; odata=verbose", "content-type": "application/json;odata=verbose", "X-RequestDigest": SPToken, "IF-MATCH": "*", 'Authorization': 'Bearer ' + SPToken, "X-HTTP-Method": "DELETE" },
                    success: function (data) {
                        getSPPinnedLocations();
                        //show more --> show more div -->
                        dfd.resolve(data);
                    },
                    error: function (errordata) {
                        console.log("removeExistingPinned Ajax Call Error: " + errordata.responseText);
                        dfd.reject(data);
                    }
                });
            }

        }
        catch (error) {
            console.log("removeExistingPinned: ", JSON.parse(error.responseText).error.message.Value);
            dfd.reject(data);
        }
        return dfd.promise();
    }

    // Treeview --> check exisitng Pinned Locations
    function checkExistingPinned(pinnedURL) {
        var dfd = $.Deferred();
        try {
            IsPinnedLocation = false;
            var UserName = localStorage.getItem('userDisplayName');
            if (SPToken) {
                var siteConfigListUrl = SPURL + "/_api/web/lists/getbytitle('" + sitePinnedLocations + "')/items?$select=ID,DocumentLibrary,DocumentLibraryURL,PinnedType,SiteURL,Author/Title&$expand=Author&$filter=Author/Title eq '" + UserName + "'";
                $.ajax({
                    url: siteConfigListUrl,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                    success: function (data) {
                        if (data.d.results.length > 0) {
                            var arrayResult = data.d.results
                            for (var i = 0; i < arrayResult.length; i++) {
                                if (arrayResult[i].DocumentLibraryURL.Description.trim() == pinnedURL.trim()) {
                                    IsPinnedLocation = true;
                                }
                            }
                        }
                        dfd.resolve(data);
                    },
                    error: function (errordata) {
                        console.log("checkExistingPinned Ajax Call Error: " + errordata.responseText);
                        dfd.reject();
                    }
                });
            }
        }
        catch (error) {
            console.log("checkExistingPinned Ajax Call Error: " + errordata.responseText);
            dfd.reject(data);
        }
        return dfd.promise();
    }

    //Treeview -> get followed sites
    function getMyFollowedSites() {
        var favdef = $.Deferred();
        try {
            var proxyURL = "https://cors-anywhere.herokuapp.com/";
            var favoUrl = proxyURL + SPURL + "/_vti_bin/homeapi.ashx/sites/followed";
            callAjaxGet(favoUrl).done(function (data) {
                favdef.resolve(data.Items);
            });
        }
        catch (err) {
            console.log("getMyFollowedSites: " + err);
        }
        return favdef.promise();
    };

    //Treeview --> save Pinned Location
    function formSaveNotes(listItem) {
        try {
            var tes = [];
            getValues(SPURL + "/").done(function (token) {
                var url = SPURL + "/_api/web/lists/GetByTitle('" + sitePinnedLocations + "')/items";
                tes = { "Accept": "application/json;odata=verbose", "X-RequestDigest": token, "IF-MATCH": "*", 'Authorization': 'Bearer ' + SPToken };
                $.ajax({
                    url: url,
                    type: "POST",
                    contentType: "application/json;odata=verbose",
                    data: JSON.stringify(listItem),
                    headers: tes,
                    success: function (data) {
                    },
                    error: function (error) {
                        console.log(error);
                    }
                });
            });
        }
        catch (error) {
            console.log(error);
        }
    }

    //Treeview --> createDocument
    function createDocument(newFileName, destURl, chkDocTitle, tokenURl, destServerRelURL) {
        var dfd = $.Deferred();
        try {
            copyFile(newFileName, destURl, chkDocTitle, tokenURl, destServerRelURL).done(function (data) {
                dfd.resolve(data);
            });
        } catch (error) {
            console.log("createDocument: " + error);
            dfd.reject(error);
        }
        return dfd.promise();
    }

    //Treeview --> copyFile
    function copyFile(newFileName, destURl, chkDocTitle, tokenURl, destServerRelURL) {
        var dfdCopy = $.Deferred();
        try {
            var sourceSiteUrl = SPURL + templateServerRelURL + "/_api/web/getfilebyserverrelativeurl('" + destServerRelURL + "')/$value";
            var xhr = new XMLHttpRequest();
            xhr.open('GET', sourceSiteUrl, true);
            xhr.setRequestHeader('binaryStringResponseBody', 'true');
            xhr.setRequestHeader('Authorization', 'Bearer ' + SPToken);
            xhr.responseType = 'blob';
            xhr.onload = function (oEvent) {
                var arrayBuffer = xhr.response;
                if (arrayBuffer) {
                    createDocumentinSelectLb(arrayBuffer, newFileName, destURl, tokenURl).done(function (data) {
                        dfdCopy.resolve(data);
                    });
                }
            };
            xhr.onerror = function (error) {
                console.log('There was an error!');
            };
            xhr.send();
        }
        catch (error) {
            console.log('copyFile: ' + error);
            dfdCopy.reject(error);
        }
        return dfdCopy.promise();
    }

    //Treeview --> create Document in Selected Library
    function createDocumentinSelectLb(fdata, newFileName, destURl, tokenURl) {
        var dfdCrtNFile = $.Deferred();
        $('.permissionalert-msg').css('display', 'none');
        try {
            if (SPToken) {
                getValues(tokenURl).done(function (formtoken) {
                    $.ajax({
                        url: destURl + "/Files/Add(url='" + newFileName + "',overwrite=false)",
                        method: "POST",
                        processData: false,
                        headers: {
                            "Accept": "application/json; odata=verbose",
                            "X-RequestDigest": formtoken,
                            'Authorization': 'Bearer ' + SPToken,
                            "content-length": fdata.byteLength
                        },
                        contentType: "application/json;odata=verbose",
                        data: fdata,
                        success: function (data) {
                            $('.alert-msg').css('display', 'none');
                            $('.permissionalert-msg').css('display', 'none');
                            dfdCrtNFile.resolve(data);
                        },
                        error: function (err) {
                            var result = JSON.parse(err.responseText);
                            switch (err.status) {
                                case 423:
                                    $('.alert-msg').css('display', 'block');
                                    closeWaitDialog();
                                    break;
                                case 400:
                                    $('.alert-msg').css('display', 'block');
                                    closeWaitDialog();
                                    break;
                                case 401:
                                    $('.permissionalert-msg').css('display', 'block');
                                    closeWaitDialog();
                                    break;
                                case 403:
                                    $('.permissionalert-msg').css('display', 'block');
                                    closeWaitDialog();
                                    break;
                                case 404:
                                    $('.permissionalert-msg').css('display', 'block');
                                    closeWaitDialog();
                                    break;
                            }
                            console.log("createNewFile: " + err);
                            dfdCrtNFile.reject(err);
                            closeWaitDialog();
                        }
                    });
                });
            }
        } catch (error) {
            console.log("createNewFile: " + error);
        }
        return dfdCrtNFile.promise();
    }

    //Treeview --> copy content in created Document 
    function openEditProperties(data, selectedSiteRelativeURL, tokenUrl) {
        var dfdCrtNFile = $.Deferred();
        openWaitDialog();
        var apiURL = tokenUrl + "/_api/web/getfilebyserverrelativeurl('" + data.d.ServerRelativeUrl + "')/ListItemAllFields";
        $.ajax({
            url: apiURL,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
            success: function (data) {
                dfdCrtNFile.resolve(data);
                closeWaitDialog();
            },
            error: function (data) {
                console.log(JSON.stringify(data));
                dfdCrtNFile.reject(data);
                closeWaitDialog();
            }
        });
        return dfdCrtNFile.promise();
    }

    //Treeview --> get List of Template From Source List
    function getListofTemplateFromSourceList() {

        removeErrorMessage();

        var docsTemplateList = "";
        if (SPToken) {

            var url = SPURL + templateServerRelURL + "/_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Editor/Name,Editor/Title,LinkFilename,ContentTypeId,ContentType/Id,ContentType/Name,FileLeafRef,*&$expand=File,ContentType,Editor/Id&$filter=((substringof('" + docsNodeNewFileExtention + "',FileLeafRef)))&$top=5000";

            $.ajax({
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json; odata=verbose");
                },
                type: "GET",
                url: url,
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + SPToken,
                }
            }).done(function (response) {
                // var docsTemplateList = "";
                var pptDocument = [];

                // var responseData = response.d.results;
                // for (var i = 0; i < responseData.length; i++) {
                //     if (responseData[i].LinkFilename.split('.').pop() == "docx") {
                //         if (responseData[i].LinkFilename.match('%') || responseData[i].LinkFilename.match('"')
                //             || responseData[i].LinkFilename.match('\'') || responseData[i].LinkFilename.match(';') || responseData[i].LinkFilename.match('#')) {
                //             //Do nothing
                //         }
                //         else {
                //             pptDocument.push(responseData[i]);
                //         }
                //     }
                // }
                // filterViewArrayList = pptDocument;
                // filteredData = pptDocument;
                // if (platform == "OfficeOnline") {
                //     getalltemp(pptDocument);
                // }
                // else {
                //     getallClientTemp(pptDocument);
                // }
                // $("#previewbtn").attr('disabled', "disabled");
                // $("#createFile").attr('disabled', 'disabled');
                // $("#nextbtn").attr('disabled', 'disabled');
                // $("#Clearflt").css("display", "none");
                // $('#refreshList').click(function () {
                //     $("#noDataFoundLbl").hide();
                //     if (currentTemplateView == "List") {
                //         $('#listOfTemplate').css('display', 'block');
                //         $('#DocTemplatesBoxView').css('display', 'none');
                //     }
                //     else {
                //         $('#listOfTemplate').css('display', 'none');
                //         $('#DocTemplatesBoxView').css('display', 'block');
                //     }
                //     $('#txtTemplateSearch').val("");
                //     $("#previewbtn").attr('disabled', 'disabled');
                //     $("#previewbtn").css('background', '');
                //     $("#previewbtn").css('cursor', 'default');
                //     $("#nextbtn").attr('disabled', 'disabled');
                //     $("#nextbtn").css('background-color', '');
                //     $("#nextbtn").css('cursor', 'default');
                //     $("#filterUL li").find('div.link').css("background-color", "");
                //     $("#filterUL li").find('div.link').css("color", "");
                //     arryOfColumnAndItem = [];
                //     filterColumns = [];
                //     gboxViewhtml = "";
                //     getDataFromFilter(pptDocument);;
                //     $("#filterUL").css("display", "none");
                //     $("#listOfTemplate").find('input').each(function () {
                //         $("#listOfTemplate").find('input').on("change", handleChange);
                //     });
                //     $("#DocTemplatesBoxView").find('input').each(function () {
                //         $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                //     });
                // });
                // $("#listOfTemplate").find('input').each(function () {
                //     $("#listOfTemplate").find('input').on("change", handleChange);
                // });
                // $("#DocTemplatesBoxView").find('input').each(function () {
                //     $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                // });
                // $("#previewbtn").on("click", function (e) {
                //     showPreview();
                //     _OpenPreviewPane(currentTemplateView);
                // });
            }).fail(function (response) {
                console.log('error:- ' + response.responseText);
                docsTemplateList = "<div class='displayMessage'>No Template Library found in DocsNode Admin Panel</div>";
                var erroeMeg = JSON.parse(response.responseText);
                erroeMeg = erroeMeg["odata.error"].message.value;
                docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel</div>";
                $('#listOfTemplate').html(docsTemplateList); $('#listOfTemplate').html(docsTemplateList);
                $("#nextbtn").attr('disabled', 'disabled');
            });
        }

        //// Added: 28th June ( Amartya )

        let payload = {
            SPOUrl: "https://docsnode.sharepoint.com",
            tenant: "641333af-c280-4e39-8f0c-1a52f0be8dc7",
            FolderPath: "",
            AccountName: "demouser@docsnode.com",
            TenantName: "docsnode",
        }

        $.ajax({
            url: GET_LOCATION_DETAILS_URL,
            beforeSend: function (request) {
                request.setRequestHeader("Accept", "application/json; odata=verbose");
            },
            dataType: "json",
            headers: {
                'Authorization': 'Bearer ' + SPToken,
            },
            type: "POST",
            data: JSON.stringify(payload),

        }).done(function (response) {
            var Files = response.d.Files.results.filter(data => data.Name.lastIndexOf(".doc") > 1);
            var pptDocument = [];

            
            for (var i = 0; i < Files.length; i++) {
                pptDocument.push(Files[i]);
                
            }
            // var responseData = response.d.results;
            // for (var i = 0; i < responseData.length; i++) {
            //     if (responseData[i].LinkFilename.split('.').pop() == "docx") {
            //         if (responseData[i].LinkFilename.match('%') || responseData[i].LinkFilename.match('"')
            //             || responseData[i].LinkFilename.match('\'') || responseData[i].LinkFilename.match(';') || responseData[i].LinkFilename.match('#')) {
            //             //Do nothing
            //         }
            //         else {
            //             pptDocument.push(responseData[i]);
            //         }
            //     }
            // }
            filterViewArrayList = pptDocument;
            filteredData = pptDocument;
            if (platform == "OfficeOnline") {
                getalltemp(pptDocument);
            }
            else {
                getallClientTemp(pptDocument);
            }
            $("#previewbtn").attr('disabled', "disabled");
            $("#createFile").attr('disabled', 'disabled');
            $("#nextbtn").attr('disabled', 'disabled');
            $("#Clearflt").css("display", "none");
            $('#refreshList').click(function () {
                $("#noDataFoundLbl").hide();
                if (currentTemplateView == "List") {
                    $('#listOfTemplate').css('display', 'block');
                    $('#DocTemplatesBoxView').css('display', 'none');
                }
                else {
                    $('#listOfTemplate').css('display', 'none');
                    $('#DocTemplatesBoxView').css('display', 'block');
                }
                $('#txtTemplateSearch').val("");
                $("#previewbtn").attr('disabled', 'disabled');
                $("#previewbtn").css('background', '');
                $("#previewbtn").css('cursor', 'default');
                $("#nextbtn").attr('disabled', 'disabled');
                $("#nextbtn").css('background-color', '');
                $("#nextbtn").css('cursor', 'default');
                $("#filterUL li").find('div.link').css("background-color", "");
                $("#filterUL li").find('div.link').css("color", "");
                arryOfColumnAndItem = [];
                filterColumns = [];
                gboxViewhtml = "";
                getDataFromFilter(pptDocument);;
                $("#filterUL").css("display", "none");
                $("#listOfTemplate").find('input').each(function () {
                    $("#listOfTemplate").find('input').on("change", handleChange);
                });
                $("#DocTemplatesBoxView").find('input').each(function () {
                    $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                });
            });
            $("#listOfTemplate").find('input').each(function () {
                $("#listOfTemplate").find('input').on("change", handleChange);
            });
            $("#DocTemplatesBoxView").find('input').each(function () {
                $("#DocTemplatesBoxView").find('input').on("change", handleChange);
            });
            $("#previewbtn").on("click", function (e) {
                showPreview();
                _OpenPreviewPane(currentTemplateView);
            });
            // console.log("Success: ", response);

        }).fail(function (error) {
            console.error('error:- ' + response.responseText);
            docsTemplateList = "<div class='displayMessage'>No Template Library found in DocsNode Admin Panel</div>";
            var erroeMeg = JSON.parse(response.responseText);
            erroeMeg = erroeMeg["odata.error"].message.value;
            docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel</div>";
            $('#listOfTemplate').html(docsTemplateList); $('#listOfTemplate').html(docsTemplateList);
            $("#nextbtn").attr('disabled', 'disabled');
            // console.error("Error: ", error);
        });

        //// Added: 28th June (Amartya)
    };

    //Clear Filter
    function ClearFilterandRebindList() {
        try {
            arryOfColumnAndItem = [];
            filterColumns = []; //popupsave
            gboxViewhtml = "";
            $('#txtTemplateSearch').val("");
            $("#filterUL").css("display", "none");
            var selectedViewName = "";
            var getselectView = ".selectView";
            selectedViewName = gSelectedView;
            getDocumentsListBasedOnView(selectedViewName);
            $("#Clearflt").css("display", "none");
        } catch (error) {
            console.log("ClearFilterandRebindList: " + error);
        }
    }

    function checkPinnedLocationListExist() {
        try {
            if (SPToken) {
                var siteConfigListUrl = SPURL + "/_api/Web/Lists?$filter=title eq '" + sitePinnedLocations + "'";
                $.ajax({
                    url: siteConfigListUrl,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                    success: function (data) {
                        if (data.d.results[0] == null && data.d.results[0] == undefined) {
                            getValues(SPURL + "/").then(function (token) {
                                $.ajax({
                                    url: SPURL + "/_api/web/lists",
                                    type: "POST",
                                    data: JSON.stringify({
                                        '__metadata': { 'type': 'SP.List' },
                                        'BaseTemplate': 100,
                                        'Title': sitePinnedLocations
                                    }),
                                    headers: {
                                        "accept": "application/json;odata=verbose",
                                        "content-type": "application/json;odata=verbose",
                                        "X-RequestDigest": token,
                                        'Authorization': 'Bearer ' + SPToken

                                    },
                                    success: function (data) {
                                        var DocumentLibraryNameColumn = {
                                            '__metadata': { 'type': 'SP.FieldText' },
                                            'FieldTypeKind': 2,
                                            'Title': 'DocumentLibrary'
                                        };
                                        $.ajax({
                                            url: SPURL + "/_api/lists/getbytitle('" + sitePinnedLocations + "')/fields",
                                            type: "POST",
                                            data: JSON.stringify(DocumentLibraryNameColumn),
                                            headers: {
                                                "accept": "application/json;odata=verbose",
                                                "content-type": "application/json;odata=verbose",
                                                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                                'Authorization': 'Bearer ' + SPToken,
                                            },
                                            success: function (data) {
                                                var DocumentLibraryURLColumn = {
                                                    '__metadata': { 'type': 'SP.FieldUrl' },
                                                    'FieldTypeKind': 11,
                                                    'Title': 'DocumentLibraryURL',
                                                    'DisplayFormat': 1
                                                };
                                                $.ajax({
                                                    url: SPURL + "/_api/lists/getbytitle('" + sitePinnedLocations + "')/fields",
                                                    type: "POST",
                                                    data: JSON.stringify(DocumentLibraryURLColumn),
                                                    headers: {
                                                        "accept": "application/json;odata=verbose",
                                                        "content-type": "application/json;odata=verbose",
                                                        "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                                        'Authorization': 'Bearer ' + SPToken,
                                                    },
                                                    success: function (data) {
                                                        //////////////
                                                        var PinnedTypeColumn = {
                                                            '__metadata': { 'type': 'SP.FieldText' },
                                                            'FieldTypeKind': 2,
                                                            'Title': 'PinnedType'
                                                        };
                                                        $.ajax({
                                                            url: SPURL + "/_api/lists/getbytitle('" + sitePinnedLocations + "')/fields",
                                                            type: "POST",
                                                            data: JSON.stringify(PinnedTypeColumn),
                                                            headers: {
                                                                "accept": "application/json;odata=verbose",
                                                                "content-type": "application/json;odata=verbose",
                                                                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                                                'Authorization': 'Bearer ' + SPToken,
                                                            },
                                                            success: function (data) {
                                                                ////////////////////
                                                                var SiteURLColumn = {
                                                                    '__metadata': { 'type': 'SP.FieldUrl' },
                                                                    'FieldTypeKind': 11,
                                                                    'Title': 'SiteURL',
                                                                    'DisplayFormat': 1
                                                                };
                                                                $.ajax({
                                                                    url: SPURL + "/_api/lists/getbytitle('" + sitePinnedLocations + "')/fields",
                                                                    type: "POST",
                                                                    data: JSON.stringify(SiteURLColumn),
                                                                    headers: {
                                                                        "accept": "application/json;odata=verbose",
                                                                        "content-type": "application/json;odata=verbose",
                                                                        "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                                                        'Authorization': 'Bearer ' + SPToken,
                                                                    },
                                                                    success: function (data) {
                                                                        checkConfigurationLogoListExist();
                                                                    },
                                                                    error: function (error) {
                                                                        console.log("Check Configuration Logo List Exist Error: ", JSON.parse(error.responseText).error.message.Value);
                                                                    }
                                                                });
                                                                //////////////
                                                            },
                                                            error: function (error) {
                                                                console.log("Check DocsNode Pinned Location Error: ", JSON.parse(error.responseText).error.message.Value);
                                                            }
                                                        });
                                                        /////////
                                                    },
                                                    error: function (error) {
                                                        console.log("Check DocsNode Pinned Location Error: ", JSON.parse(error.responseText).error.message.Value);
                                                    }
                                                });
                                            },
                                            error: function (error) {
                                                console.log("Check DocsNode Pinned Location Error: ", JSON.parse(error.responseText).error.message.Value);
                                            }
                                        });
                                    },
                                    error: function (error) {
                                        console.log("Check DocsNode Pinned Location Error: ", JSON.parse(error.responseText).error.message.Value);
                                    }
                                }).fail(function (error) {
                                    console.log("Check DocsNode Pinned Location Error: ", JSON.parse(error.responseText).error.message.Value);

                                });
                            });
                        }
                        else {
                            checkConfigurationLogoListExist();
                        }

                    },
                    error: function (errordata) {
                        console.log(errordata);
                    }
                });
            }
        }
        catch (error) {
            console.log("Check DocsNode Pinned Location Error: ", JSON.parse(error.responseText).error.message.Value);
        }
    }

    function checkConfigurationLogoListExist() {
        try {
            if (SPToken) {
                var ConfigListUrl = SPURL + "/_api/Web/Lists?$filter=title eq '" + ConfigurationListName + "'";
                $.ajax({
                    url: ConfigListUrl,
                    method: "GET",
                    headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                    success: function (data) {
                        if (data.d.results[0] == null && data.d.results[0] == undefined) {
                            getValues(SPURL + "/").then(function (token) {
                                $.ajax({
                                    url: SPURL + "/_api/web/lists",
                                    type: "POST",
                                    data: JSON.stringify({
                                        '__metadata': { 'type': 'SP.List' },
                                        'BaseTemplate': 100,
                                        'Title': ConfigurationListName
                                    }),
                                    headers: {
                                        "accept": "application/json;odata=verbose",
                                        "content-type": "application/json;odata=verbose",
                                        "X-RequestDigest": token,
                                        'Authorization': 'Bearer ' + SPToken

                                    },
                                    success: function (data) {
                                        var AppLogoColumn = {
                                            '__metadata': { 'type': 'SP.FieldUrl' },
                                            'FieldTypeKind': 11,
                                            'Title': 'AppLogo',
                                            'DisplayFormat': 0
                                        };

                                        $.ajax({
                                            url: SPURL + "/_api/lists/getbytitle('" + ConfigurationListName + "')/fields",
                                            type: "POST",
                                            data: JSON.stringify(AppLogoColumn),
                                            headers: {
                                                "accept": "application/json;odata=verbose",
                                                "content-type": "application/json;odata=verbose",
                                                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                                'Authorization': 'Bearer ' + SPToken,
                                            },
                                            success: function (data) {
                                                var IsActiveColumn = {
                                                    '__metadata': { 'type': 'SP.FieldNumber' },
                                                    'FieldTypeKind': 9,
                                                    'Title': 'IsActive'
                                                };
                                                $.ajax({
                                                    url: SPURL + "/_api/lists/getbytitle('" + ConfigurationListName + "')/fields",
                                                    type: "POST",
                                                    data: JSON.stringify(IsActiveColumn),
                                                    headers: {
                                                        "accept": "application/json;odata=verbose",
                                                        "content-type": "application/json;odata=verbose",
                                                        "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                                        'Authorization': 'Bearer ' + SPToken,
                                                    },
                                                    success: function (data) {
                                                    },
                                                    error: function (error) {
                                                        console.log(error);
                                                    }
                                                });

                                            },
                                            error: function (error) {
                                                console.log(error);
                                            }
                                        });
                                    },
                                    error: function (error) {
                                        console.log(error);
                                    }
                                });
                            }).fail(function (error) {
                                console.log(error);

                            });
                        }
                        else {
                        }
                    },
                    error: function (errordata) {
                        console.log("Check Configuration Logo List Error: ", errordata);
                    }
                });

            }
        }
        catch (error) {
            console.log("Check Configuration Logo List Error: ", JSON.parse(error.responseText).error.message.Value);
        }
    }

    function getSelectedSiteId(selectedSite) {
        var dfd = $.Deferred();
        if (GraphAPIToken) {
            var tenantName = SPURL.substr(8, SPURL.length);
            var GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/" + selectedSite + "?$select=id";
            $.ajax({
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json");
                },
                type: "GET",
                url: GraphAPI,
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + GraphAPIToken,
                }
            }).done(function (response) {
                var siteWebID = "";
                if (response) {
                    siteWebID = response.id;
                }
                dfd.resolve(siteWebID);
            }).fail(function (response) {
                dfd.reject(response.responseText);
            });
        }
        return dfd.promise();
    };

    function getDocumentsListBasedOnView(selectedViewName) {
        var self = this;
        try {
            getColumnFieldName(selectedViewName).done(function (data) {
                $("#Clearflt").css("display", "none");
            });
            getListItemsForView(selectedViewName).done(function (data) {
                var items = data.d.results;
                if (items.length != 0) {
                    var filterstr = "((";
                    for (var i = 0; i < items.length; i++) {
                        filterstr += "(Id eq " + items[i].Id + ") or ";
                    }
                    filterstr = filterstr.replace(/or([^or]*)$/, '$1');
                    filterstr += ") and (substringof('" + docsNodeNewFileExtention + "', FileLeafRef))";
                    filterstr += ")";
                    getviewdocument(filterstr).done(function (data) {
                        var pptDocument = [];
                        var response = data.d.results;
                        for (var i = 0; i < response.length; i++) {
                            if (response[i].LinkFilename.split('.').pop() == "docx") {
                                if (response[i].LinkFilename.match('%') || response[i].LinkFilename.match('"')
                                    || response[i].LinkFilename.match('\'') || response[i].LinkFilename.match(';') || response[i].LinkFilename.match('#')) {
                                    //Do nothing
                                }
                                else {
                                    pptDocument.push(response[i]);
                                }
                            }
                        }
                        if (pptDocument == 0) {
                            $("#btndropdown").attr('disabled', 'disabled');
                            $('#refreshList').attr('disabled', 'disabled');
                            $('#txtTemplateSearch').attr('disabled', 'disabled');
                        }
                        else {
                            $("#btndropdown").removeAttr('disabled');
                            $('#refreshList').removeAttr('disabled');
                            $("#txtTemplateSearch").removeAttr('disabled');
                        }
                        filterViewArrayList = pptDocument;
                        filteredData = pptDocument;
                        gboxViewhtml = '';
                        if (platform == "OfficeOnline") {
                            getalltemp(pptDocument);
                        }
                        else {
                            getallClientTemp(pptDocument);
                        }
                        $('#refreshList').show();
                        $("#previewbtn").attr('disabled', 'disabled');
                        $("#previewbtn").css('background', '');
                        $("#previewbtn").css('cursor', 'default');
                        $("#nextbtn").attr('disabled', "disabled");
                        $("#nextbtn").css('background-color', '');
                        $("#nextbtn").css('cursor', 'default');
                    }).fail(
                        function (error) {
                            console.log(JSON.stringify(error));
                        });
                }
                else {
                    $("#btndropdown").attr('disabled', 'disabled');
                    $('#refreshList').hide();
                    filteredData = [];
                    if (platform == "OfficeOnline") {
                        getalltemp(filteredData);
                    }
                    else {
                        getallClientTemp(filteredData);
                    }
                }
            });
        } catch (error) {
            console.log("getDocumentsListBasedOnView: " + error);
        }
    }

    function getListItemsForView(viewTitle) {
        var dfdCrtNFile = $.Deferred();
        try {
            var viewQueryUrl = SPURL + templateServerRelURL + "/_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/Views/getbytitle('" + viewTitle + "')/ViewQuery";
            $.ajax({
                url: viewQueryUrl,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                success: function (data) {
                    var viewQuery = data.d.ViewQuery;
                    getListItems(SPURL, TemplateLibraryDisplayName, viewQuery).done(function (result) {
                        dfdCrtNFile.resolve(result);
                    });
                },
                error: function (data) {
                    console.log(JSON.stringify(data));
                }
            });
        } catch (error) {
            console.log("getListItemsForView: " + error);
        }
        return dfdCrtNFile.promise();
    }

    function getListItems(webUrl, listTitle, queryText) {
        var dfdCrtNFiles = $.Deferred();
        var wURL = webUrl + templateServerRelURL;
        try {
            var viewXml = "<View Scope='Recursive' ><Query>" + queryText + "<QueryOptions><ViewAttributes Scope='Recursive' /></QueryOptions></Query></View>";
            var url = webUrl + templateServerRelURL + "/_api/web/lists(guid'" + listTitle + "')/getitems";
            var queryPayload = {
                'query': {
                    'ViewXml': viewXml
                }
            };
            var query = JSON.stringify(queryPayload.query);

            if (SPToken) {
                getValues(wURL).done(function (formtoken) {
                    $.ajax({
                        url: url + "(query=@v1)?@v1=" + query,
                        method: "POST",
                        headers: {
                            'Authorization': 'Bearer ' + SPToken,
                            "X-RequestDigest": formtoken,
                            "Accept": "application/json; odata=verbose",
                            "content-type": "application/json; odata=verbose"
                        },
                        success: function (data) {
                            dfdCrtNFiles.resolve(data);
                        },
                        error: function (err) {
                            console.log(err);
                        }
                    });
                });
            }

        }
        catch (error) {
            console.log("getListItems: " + error);
        }
        return dfdCrtNFiles.promise();
    }

    function getValues(tokenURl) {
        var dfdReqDig = $.Deferred();
        try {

            $.ajax({
                url: tokenURl + "/_api/contextinfo",
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

        } catch (error) {
            console.log("getValues: " + error);
        }
        return dfdReqDig.promise();
    }
    function getviewdocument(filterstring) {
        var result;
        var dfdgetdoc = $.Deferred();
        var webUrl = SPURL;
        var url = webUrl + templateServerRelURL + "/_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items?$select=ID,Editor/Name,Editor/Title,LinkFilename,ContentTypeId,ContentType/Id,ContentType/Name,*&$expand=File,ContentType,Editor/Id&$filter=";
        url = url + filterstring;
        try {
            callAjaxGet(url).done(function (data) {
                result = data;
                dfdgetdoc.resolve(result);
            });
        } catch (err) {
            dfdgetdoc.reject();
        }
        return dfdgetdoc.promise();
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

    function setHierarchy(ddlID, attributName) {
        $("#" + ddlID).html($('#' + ddlID + ' option').sort(function (x, y) {
            if ($(y).val() !== "0") {
                $(y).text($(y).attr(attributName) + " " + $(y).val());
            }
            return parseInt($(x).attr(attributName).re) < $(y).attr(attributName) ? -1 : 1;
        }));
        $("#" + ddlID).get(0).selectedIndex = 0;
    }

    function getListOfLibraryFromWeb() {
        openWaitDialog();
        removeErrorMessage();
        DestinationWebRelativeUrl = $('#txtSPRelativeURL').val();
        if (DestinationWebRelativeUrl.indexOf('\\') >= 0) {
            showErrorMessage("Please avoid Backslash (\\) in URL");
            closeWaitDialog();
            return false;
        }
        else if (DestinationWebRelativeUrl.indexOf('http') >= 0 || DestinationWebRelativeUrl.indexOf('.com') >= 0 || DestinationWebRelativeUrl.indexOf('.') >= 0) {
            showErrorMessage("Please enter valid Relative Web URL");
            closeWaitDialog();
            return false;
        }
        if (GraphAPIToken) {
            var GraphAPI = "";

            var tenantName = SPURL.substr(8, SPURL.length);
            if (DestinationWebRelativeUrl != null && DestinationWebRelativeUrl.trim() != "") {
                if (DestinationWebRelativeUrl.charAt('0') !== '/') {
                    DestinationWebRelativeUrl = "/" + DestinationWebRelativeUrl
                }
                GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":" + DestinationWebRelativeUrl + ":/lists";
            }
            else {
                GraphAPI = "https://graph.microsoft.com/v1.0/sites/root/lists";
            }
            $.ajax({
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json");
                },
                type: "GET",
                url: GraphAPI,
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + GraphAPIToken,
                }
            }).done(function (response) {
                var result = response.value;
                var listOfLibrary = "<option value='0'>Select</option>";
                if (response) {
                    for (var i = 0; i < result.length; i++) {
                        if (result[i].list.template === "documentLibrary" && !result[i].list.hidden) {
                            listOfLibrary += "<option internalName='" + result[i].name + "' value='" + result[i].displayName + "'>" + result[i].displayName + "</option>";
                        }
                    }
                }
                else {
                    closeWaitDialog();
                }
                $('#SPLibraryList').html(listOfLibrary);
            }).fail(function (response) {
                console.log('error:- ' + response.responseText);
                showErrorMessage("There was some issue. This site doesn't exist or you don't have permission to access this site");
                closeWaitDialog();
            });
        }
    }

    function getallViews() {
        var dfdAllView = $.Deferred();
        try {
            let listAllDocView = [];
            var url = SPURL + templateServerRelURL + "/_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/views";
            $.ajax({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                success: function (data) {
                    $.each(data.d.results, function (index, key) {
                        if (!data.d.results[index].Hidden && data.d.results[index].Title != "") {
                            listAllDocView.push(data.d.results[index].Title);
                        }
                    });
                    viewlist(listAllDocView);
                    dfdAllView.resolve(listAllDocView);
                },
                error: function (data) {
                    console.log(JSON.stringify(data));
                }
            });
            function viewlist(data) {
                var listofviews = "";
                listofviews += "<li id='lstvw' aria-label='GridViewSmall icon'><i title='List' class='ms-Icon ms-Icon--ViewList dsIconTile' aria-hidden='true'></i></i><span>List</span></li>"
                listofviews += "<li id='grdvw' aria-label='GridViewSmall icon' class='selectView'><i class='ms-Icon ms-Icon--GridViewSmall dsIconTile' title='GridViewSmall' aria-hidden='true'></i><span>Tile</span></li>"
                if (data.length > 0) {
                    for (var i = 0; i < data.length; i++) {
                        listofviews += "<li id='" + data[i] + "'>" + data[i] + "</li>";
                    }
                }

                if (currentTemplateView == "Box") {
                    $('#DocTemplatesBoxView').show();
                    $('#listOfTemplate').hide();
                }
                else {
                    $('#listOfTemplate').show();
                    $('#DocTemplatesBoxView').hide();
                }
                $("#ViewUL").html(listofviews);
                $("#ViewUL li").on("click", function () {
                    $("#ViewUL").hide();
                    filterColumns = [];
                    arryOfColumnAndItem = [];
                    if ($("#ViewUL").children().hasClass('selectView')) {
                        $("#ViewUL li").removeClass('selectView');
                    }
                    $(this).addClass('selectView');
                    if (this.innerText != "List" && this.innerText != "Tile") {
                        openWaitDialog();
                        $("#noDataFoundLbl").hide();
                        getDocumentsListBasedOnView(this.innerHTML);
                        gSelectedView = this.innerText;
                        if (currentTemplateView == 'Box') {
                            $("#grdvw").addClass("selectView");
                            $("#lstvw").removeClass("selectView");
                            $('#DocTemplatesBoxView').show();
                            $('#listOfTemplate').hide();
                        }
                        else {
                            $("#grdvw").removeClass("selectView");
                            $("#lstvw").addClass("selectView");
                            $('#listOfTemplate').show();
                            $('#DocTemplatesBoxView').hide();
                        }
                    }
                    else {
                        currentTemplateView = this.innerText;
                        toggleView();
                    }
                });
            }
        } catch (error) {
            console.log("getallViews: " + error);
            dfdAllView.reject(error);
        }

        return dfdAllView.promise();
    }

    function getColumnFieldName(viewName) {
        var ddfieldName = $.Deferred();

        var Url = SPURL + templateServerRelURL + "/_api/Web/Lists(guid'" + TemplateLibraryDisplayName + "')/Views/getbytitle('" + viewName + "')/ViewFields";
        try {
            var columnFieldName;

            $.ajax({
                url: Url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose", 'Authorization': 'Bearer ' + SPToken },
                success: function (data) {
                    columnFieldName = data.d.Items.results;
                    filterColumnList(data.d.Items.results);
                    ddfieldName.resolve(columnFieldName);
                },
                error: function (data) {
                    console.log("getColumnFieldName: " + data);
                }
            });

            function filterColumnList(data) {
                getFieldName(data);
            }
        }
        catch (error) {
            console.log("getColumnFieldName: " + error);
        }
        return ddfieldName.promise();
    }

    function getFieldName(columnfieldname) {
        let items = "";
        try {
            items += "<li class ='clearFields' id ='Clearflt'><i data-icon-name='List' class='ms-Icon ms-Icon--ClearFilter' role='presentation' aria-hidden='true'></i><span>Clear Filter</span></li>";
            for (let i = 0; i < columnfieldname.length; i++) {
                if (columnfieldname[i].indexOf("_x0020_") > -1) {
                    items += "<li class='filterFields' id = " + columnfieldname[i] + "><div class='link'><i id='fasi'  class='ms-Icon ms-Icon--ChevronDown' aria-hidden='true'></i>" + columnfieldname[i].replace(new RegExp('_x0020_', 'g'), ' ') + "</div><ul class='" + columnfieldname[i] + "' id='" + columnfieldname[i] + "'></ul></li>";
                }
                else if (columnfieldname[i] == "Editor") {
                    items += "<li class='filterFields' id = " + columnfieldname[i] + "><div class='link'><i id='fasi' class='ms-Icon ms-Icon--ChevronDown' aria-hidden='true'></i>" + columnfieldname[i].replace(new RegExp('Editor', 'g'), 'Modified By') + "</div><ul class='" + columnfieldname[i] + "' id='" + columnfieldname[i] + "'></ul></li>";
                }
                else if (columnfieldname[i] == "DocIcon") {
                    continue;
                }
                else if (columnfieldname[i] == "LinkFilename") {
                    continue;
                }
                else if (columnfieldname[i] == "ID") {
                    continue;
                }
                else {
                    items += "<li class='filterFields' id = " + columnfieldname[i] + "><div class='link'><i id='fasi' class='ms-Icon ms-Icon--ChevronDown' aria-hidden='true'></i>" + columnfieldname[i] + "</div><ul class='" + columnfieldname[i] + "' id='" + columnfieldname[i] + "'></ul></li>";
                }
            }
            var ul = document.getElementById("filterUL");
            ul.innerHTML = items;
            $("#btndropdown").off("mouseover");
            $("#btndropdown").on("mouseover", function () {
                $("#filterUL").show();
                $("#ViewUL").hide();
                $(".filterFields ul").css('display', 'none');
                $("#filterUL").css({ 'display': 'block' });
                var mousehover = false;
                $("#filterUL").mouseleave(function () { mousehover = true; });
                $("#filterUL").mouseenter(function () { mousehover = false; });
                $(".filterFields .link ").off("click");
                $(".filterFields .link ").on("click", function () {
                    getalldataoffieldcolumn(filterViewArrayList, this);
                });
                $(".input-group").hover(function () {
                    if (mousehover) {
                        $("#filterUL").hide();
                        $('.filterFields').find('ul').removeClass('active');
                    }
                });
                $(".list-item-sec").hover(function () {
                    if (mousehover) {
                        $("#filterUL").hide();
                        $('.filterFields').find('ul').removeClass('active');
                    }
                });
                $('#c-shareddrive').hover(function () {
                    if (mousehover) {
                        $("#filterUL").hide();
                        $('.filterFields').find('ul').removeClass('active');
                    }
                });
                $(".list-item-sec").on('click', function () {
                    if (mousehover) {
                        $("#filterUL").hide();
                    }
                });
            });

            $("#Clearflt").on("click", function () {
                openWaitDialog();
                ClearFilterandRebindList();
            });

        } catch (error) {
            console.log("getAllFields: " + error);
        }
    }

    function getalldataoffieldcolumn(data, prop) {
        var name = $(prop).parent().attr("id");
        var internalName = name;
        var options = [];
        var items = "";
        var optionsItems = [];
        var mousehover = false;

        data.map(function (value, i) {
            if (String(value[internalName]) && options.map(function (record) { return record.value; }).indexOf(value[internalName]) == -1) {
                if (internalName == "Editor") {
                    options.push({
                        "name": internalName,
                        "value": value[internalName].Title
                    });
                }
                else if (internalName == "Modified") {
                    var modifiedDate = _getFormattedDate(value[internalName]);
                    options.push({
                        "name": internalName,
                        "value": modifiedDate
                    });
                }
                else if (internalName == "ContentType") {
                    options.push({
                        "name": internalName,
                        "value": value[internalName].Name
                    });
                }
                else {
                    options.push({
                        "name": internalName,
                        "value": value[internalName]
                    });
                }
            }
        });
        options = removeDuplicates(options, "value");
        for (let i = 0; i < options.length; i++) {
            if (options[i].value != null) {
                optionsItems.push(options[i]);
                if (options[i].value == true) {
                    options[i].value = "Yes";
                    var filter = internalName + ':' + true;
                    items += "<li class='valueItems' filter='" + filter + "' id='true'>" + options[i].value + "</li>";
                }
                else if (options[i].value == false) {
                    options[i].value = "No";
                    var filter = internalName + ':' + false;
                    items += "<li class='valueItems' filter='" + filter + "' id='false'>" + options[i].value + "</li>";
                }
                else {
                    if (!isNaN(options[i].value)) {
                        var filter = internalName + ':' + options[i].value;
                    }
                    else {
                        var filter = internalName + ':' + options[i].value.replace(/\s+/g, '');
                    }
                    items += "<li class='valueItems' filter='" + filter + "' id='" + options[i].value + "'>" + options[i].value + "</li>";
                }
            }
        }
        if (optionsItems.length == 0) {
            items += "<li id='NA' disabled ='disabled' >N/A</li>";
        }
        $("." + name).html(items);
        $('.filterFields').find('ul').slideUp();
        if ($(prop).siblings().hasClass('active')) {
            $(prop).siblings().removeClass('active');
            $('.filterFields').find('ul').removeClass('active');
        }
        else {
            $('.filterFields').find('ul').removeClass('active');
            $(prop).siblings().slideDown().addClass('active');
        }

        buildFilterDropDownUI(filterColumns);

        $(".valueItems").on("click", function () {
            openWaitDialog();
            var flagPresent = false;
            var itemselected = $(this).attr("id");
            if (arryOfColumnAndItem.length > 0) {
                arryOfColumnAndItem.forEach(function (item, index) {
                    return item["value"] == itemselected ? flagPresent = true : null;
                });
            }
            else {
                arryOfColumnAndItem.push({ "key": internalName, "value": itemselected });
                flagPresent = true;
            }
            if (flagPresent == false) {
                arryOfColumnAndItem.push({ "key": internalName, "value": itemselected });
            }
            getDataFromFilter(data);
            $("#filterUL").css('display', 'none');
            $("." + name).css('display', 'none');
            $('#txtTemplateSearch').val("");
            $("#previewbtn").attr('disabled', 'disabled');
            $("#previewbtn").css('background', '');
            $("#previewbtn").css('cursor', 'default');
            $("#nextbtn").attr('disabled', "disabled");
            $("#nextbtn").css('background-color', '');
            $("#nextbtn").css('cursor', 'default');
            closeWaitDialog();

        });
    }

    function removeDuplicates(myArr, prop) {
        return myArr.filter(function (obj, pos, arr) {
            return arr.map(function (mapObj) {
                return mapObj[prop]
            }).indexOf(obj[prop]) === pos;
        });
    }

    // use to apply css to selected filter
    function buildFilterDropDownUI(filterColumns) {
        try {

            for (var i = 0; i < filterColumns.length; i++) {
                var findOuterLI = "li#" + filterColumns[i].key;
                var filterAttr = filterColumns[i].key + ':' + filterColumns[i].value.replace(/\s+/g, '');
                var findinnerLI = $('li[filter="' + filterAttr + '"]');
                $("#filterUL").find(findOuterLI).find('div.link').css("background-color", "#04aba3");
                findinnerLI.css("background-color", "#04aba3");
                findinnerLI.css("color", "#fff");
                $("#filterUL").find(findOuterLI).find('div.link').css("color", "#fff");
            }
            if (filterColumns.length == 0) {
                $("#filterUL li").css("background-color", "");
                $("#filterUL li").css("color", "");
            }
        } catch (error) {
            console.log("buildFilterDropDownUI: " + error);
        }
    }
    function buildFilterCssArray(arr) {
        try {
            if (arr.value.indexOf(":;") > -1) {
                var tempValues = arr.value;
                var tempValuesArr = tempValues.split(":;");
                for (var i = 0; i < tempValuesArr.length; i++) {
                    filterColumns.push({ "key": arr.key, "value": tempValuesArr[i] });
                }
            }
            else {
                filterColumns.push({ "key": arr.key, "value": arr.value });
            }
        } catch (error) {
            console.log("buildFilterCssArray: " + error);
        }
    }
    function getDataFromFilter(data) {
        var self = this;
        var filterByFieldValue = [];
        var results = [];
        var filterCriteria = [];
        var filterString = "";
        var tempString = "";
        arryOfColumnAndItem.map(function (item) { return item.key }).forEach(function (item) {
            if (results.indexOf(item) === -1)
                results.push(item);
        });

        for (var key = 0; key < results.length; key++) {
            for (var j = 0; j < arryOfColumnAndItem.length; j++) {
                if (arryOfColumnAndItem[j].key == results[key]) {
                    filterString += arryOfColumnAndItem[j].value + ':;';
                }
            }
            filterString = filterString.slice(0, -2);
            filterCriteria.push({ "key": results[key], "value": filterString });
            filterString = "";
        }

        Array.prototype.flexFilter = function (info) {
            // Set our variables
            var matchesFilter, matches = [], count;
            matchesFilter = function (item) {
                count = 0;
                for (var n = 0; n < info.length; n++) {
                    if (info[n]["key"] == "DocIcon") {
                        tempString = item.File.Name;
                        var fileNameLen = tempString.length;
                        var lstindex = tempString.lastIndexOf('.');
                        var fileExt = tempString.substr(lstindex + 1, fileNameLen);
                        fileExt = fileExt.toLowerCase();
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(fileExt) > -1) {
                            count++;
                        }
                    }
                    else if (info[n]["key"] == "Editor") {
                        tempString = item.Editor.Title;
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(tempString) > -1) {
                            count++;
                        }
                    }
                    else if (info[n]["key"] == "Modified") {
                        tempString = _getFormattedDate(item[info[n]["key"]]);
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(tempString) > -1) {
                            count++;
                        }
                    }
                    else if (info[n]["key"] == "ContentType") {
                        tempString = item[info[n]["key"]].Name;
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(tempString) > -1) {
                            count++;
                        }
                    }
                    else {
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(item[info[n]["key"]]) > -1) {
                            count++;
                        }
                    }
                }
                // If TRUE, then the current item in the array meets 
                return count == info.length;
            }
            // Loop through each item in the array
            for (var i = 0; i < this.length; i++) {
                // Determine if the current item matches the filter criteria
                if (this[i].ContentType.Name != "Folder") {
                    if (matchesFilter(this[i])) {
                        matches.push(this[i]);
                    }
                }
            }
            // Give us a new array containing the objects matching the filter criteria
            return matches;
        }

        filterByFieldValue = data.flexFilter(filterCriteria);
        if (arryOfColumnAndItem.length > 0) {
            if (platform == "OfficeOnline") {
                getalltemp(filterByFieldValue);
            }
            else {
                getallClientTemp(filterByFieldValue);
            }
            filteredData = filterByFieldValue;
            $("#Clearflt").css("display", "block");
        }
        else {
            if (platform == "OfficeOnline") {
                getalltemp(filterViewArrayList);
            }
            else {
                getallClientTemp(filterViewArrayList);
            }
            filteredData = filterViewArrayList;
            $("#Clearflt").css("display", "none");
        }
        buildFilterDropDownUI(filterColumns);
    }
    function _getFormattedDate(docModifieddate) {
        try {
            var newdate = new Date(docModifieddate);
            return (newdate.getMonth() + 1) + '/' + newdate.getDate() + '/' + newdate.getFullYear();
        } catch (error) {
            console.log("_getFormattedDate: " + error);
        }
    }
    function getalltemp(data) {
        try {
            var pptString = "";
            var Count = 0;
            $('#DocTemplatesBoxView').html("");
            for (var i = 0; i < data.length; i++) {
                var fileNameLen = data[i].Name.length;
                var lstindex = data[i].Name.lastIndexOf('.');
                var fileExt = data[i].Name.substr(lstindex + 1, fileNameLen);
                fileExt = fileExt.toLowerCase();
                if (data[i].LinkFilename.indexOf(docsNodeFilterExtention) != -1 || data[i].LinkFilename.indexOf(docsNodeFilterExtention.toUpperCase()) != -1) {

                    pptString += "<li contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].Modified + "'";
                    pptString += "documentTitle='" + (data[i].Name != null ? data[i].Name : "") + "' docId='" + data[i].ID + "' documentGUID='" + data[i].UniqueId + "'";
                    pptString += "modifiedName='" + data[i].Editor.Title + "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
                    pptString += "ext='.docx'";
                    pptString += ">";
                    pptString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].UniqueId + "></input>";
                    pptString += "<i class='ms-Icon ms-Icon--WordLogo' title='WordLogo' aria-hidden='true'></i>";
                    pptString += "<span> " + data[i].LinkFilename + "</span></li>";
                }
            }
            if (data.length > 0) {
                $('#listOfTemplate').html(pptString);
                if (currentTemplateView == "Box") {
                    $('#DocTemplatesBoxView').css('display', 'block');
                    getalltempBoxView(data, Count);
                }
                else {
                    $('#DocTemplatesBoxView').css('display', 'none');
                    $("#listOfTemplate").find('input').each(function () {
                        $("#listOfTemplate").find('input').on("change", handleChange);
                    });
                    closeWaitDialog();
                }
                $("#btndropdown").removeAttr('disabled');
            }
            else {
                $("#listOfTemplate").html("<p style='color:Red'>No Records Found...!!</p>");
                $("#DocTemplatesBoxView").html("<p style='color:Red'>No Records Found...!!</p>");
                $("#btndropdown").attr('disabled', 'disabled');
                closeWaitDialog();
            }
        }
        catch (error) {
            console.log(error);
        }
    }
    function getalltempBoxView(data, Count) {
        try {
            var boxViewString = "";
            var imageURLs = "";
            if (data.length > 0) {
                if (data[Count].File.LinkingUri != null) {
                    var fileURL = "";
                    var boxThumbnailURL = "";
                    fileURL = data[Count].File.LinkingUri.split('?d')[0];
                    boxThumbnailURL = data[Count].File.LinkingUri.split('DocsNodeAdmin')[0];
                    boxThumbnailURL += "DocsNodeAdmin/_layouts/15/getpreview.ashx?path=" + fileURL;
                    fileURL = fileURL.replace(SPURL, "");
                    var url = SPURL + templateServerRelURL + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
                    if (SPToken) {
                        var xhr = new window.XMLHttpRequest();
                        xhr.open("GET", url, true);
                        xhr.setRequestHeader("Accept", "application/json; odata=verbose");
                        xhr.setRequestHeader("Authorization", "Bearer " + SPToken);
                        //Now set response type
                        xhr.responseType = 'arraybuffer';
                        xhr.addEventListener('load', function () {
                            if (xhr.status === 200) {
                                imageURLs = boxThumbnailURL;
                            }
                            else {
                                imageURLs = "images/icons/doc-96.png";
                            }
                            bindingWebAllTemp(imageURLs, data, Count);
                        })
                        xhr.send();
                    }
                }
            }
        }
        catch (error) {
            console.log(error);
        }
    }
    function bindingWebAllTemp(imageURLs, data, i) {
        var fileNameLen = data[i].File.Name.length;
        var lstindex = data[i].File.Name.lastIndexOf('.');
        var fileExt = data[i].File.Name.substr(lstindex + 1, fileNameLen);
        fileExt = fileExt.toLowerCase();
        if (data[i].LinkFilename.indexOf(docsNodeFilterExtention) != -1 || data[i].LinkFilename.indexOf(docsNodeFilterExtention.toUpperCase()) != -1) {
            boxViewString += "<li thumbnail='' contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].Modified + "'";
            boxViewString += "documentTitle='" + (data[i].File.Name != null ? data[i].File.Name : "") + "' docId='" + data[i].ID + "' documentGUID='" + data[i].File.UniqueId + "'";
            boxViewString += "modifiedName='" + data[i].Editor.Title + "' serverRelativeURL='" + data[i].File.ServerRelativeUrl + "'";
            boxViewString += "ext='.docx'";
            boxViewString += ">";
            boxViewString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].File.UniqueId + "></input><div class='box-img'>";
            boxViewString += "<div>";
            boxViewString += "<img src='" + imageURLs + "'class='docimgouterbox' /></div></div>";
            boxViewString += "<span style='word-break: break-all' title=" + data[i].LinkFilename + "> " + (data[i].LinkFilename.length > 25 ? data[i].LinkFilename.substring(0, 25) + "..." : data[i].LinkFilename) + "</span></li>";
            $('#DocTemplatesBoxView').append(boxViewString);
            boxViewString = "";
            closeWaitDialog();
        }
        $("#DocTemplatesBoxView").find('input').each(function () {
            $("#DocTemplatesBoxView").find('input').on("change", handleChange);
        });
        if (i < data.length - 1) {
            getalltempBoxView(data, ++i);
        }
    }

    function getallClientTemp(data) {
        var pptString = "";
        boxViewString = "";
        var Counter = 0;
        $('#DocTemplatesBoxView').html("");
        if (currentTemplateView == "List") {
            for (var i = 0; i < data.length; i++) {
                var fileNameLen = data[i].File.Name.length;
                var lstindex = data[i].File.Name.lastIndexOf('.');
                var fileExt = data[i].File.Name.substr(lstindex + 1, fileNameLen);
                fileExt = fileExt.toLowerCase();
                if (data[i].LinkFilename.indexOf(docsNodeFilterExtention) != -1 || data[i].LinkFilename.indexOf(docsNodeFilterExtention.toUpperCase()) != -1) {
                    pptString += "<li contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].Modified + "'";
                    pptString += "documentTitle='" + (data[i].File.Name != null ? data[i].File.Name : "") + "' docId='" + data[i].ID + "' documentGUID='" + data[i].File.UniqueId + "'";
                    pptString += "modifiedName='" + data[i].Editor.Title + "' serverRelativeURL='" + data[i].File.ServerRelativeUrl + "'";
                    pptString += "ext='.docx'";
                    pptString += ">";
                    pptString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].File.UniqueId + "></input>";
                    pptString += "<img src= 'images/" + docsNodeListTemplateLogo + "' class='width20'>";
                    pptString += "<span> " + data[i].LinkFilename + "</span></li>";
                }
            }
            if (data.length > 0) {
                $('#listOfTemplate').html(pptString);
                $('#DocTemplatesBoxView').css('display', 'none');
                $("#listOfTemplate").find('input').each(function () {
                    $("#listOfTemplate").find('input').on("change", handleChange);
                });
                closeWaitDialog();
                $("#btndropdown").removeAttr('disabled');
            }
            else {
                $("#btndropdown").attr('disabled', 'disabled');
                $("#listOfTemplate").html("<p style='color:Red'>No Records Found...!!</p>");
                closeWaitDialog();
            }
        }
        else {
            $('#DocTemplatesBoxView').css('display', 'block');
            if (gboxViewhtml != "") {
                if (data != 0) {
                    $("#btndropdown").removeAttr('disabled');
                    $('#DocTemplatesBoxView').html(gboxViewhtml);
                    $("#DocTemplatesBoxView").find('input').each(function () {
                        $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                    });
                    closeWaitDialog();
                }
                else {
                    toDataURL(Counter, data);
                }
            }
            else {
                openWaitDialog();
                toDataURL(Counter, data);
            }
        }
    }
    function toDataURL(Cnt, data) {
        var boxThumbnailURL = "";
        var proxyURL = "https://cors-anywhere.herokuapp.com/";
        if (data.length > 0) {
            if (data[Cnt].File.LinkingUri != null) {
                var fileURL = data[Cnt].File.LinkingUri.split('?d')[0];
                boxThumbnailURL = data[Cnt].File.LinkingUri.split('DocsNodeAdmin')[0];
                boxThumbnailURL += "DocsNodeAdmin/_layouts/15/getpreview.ashx?path=" + fileURL;
            }
            var xhr = new window.XMLHttpRequest();
            xhr.open('GET', proxyURL + boxThumbnailURL, true);
            xhr.setRequestHeader('Authorization', 'Bearer ' + SPToken);
            xhr.responseType = 'blob';
            xhr.onload = function (event) {
                if (event.srcElement.status == 200) {
                    filereader(event).done(function (imgdata) {
                        binding(imgdata, data, Cnt);
                    });
                }
                else {
                    var imageURLs = "images/icons/doc-96.png";
                    binding(imageURLs, data, Cnt);
                }
                function filereader(event) {
                    var dfdImg = $.Deferred();
                    var reader = new FileReader();
                    reader.onloadend = function () {
                        dfdImg.resolve(reader.result);
                    }
                    reader.readAsDataURL(xhr.response);
                    return dfdImg.promise();
                }
            };
            xhr.send();
        }
        else {
            gboxViewhtml = "";
            $("#btndropdown").attr('disabled', 'disabled');
            $("#DocTemplatesBoxView").html("<p style='color:Red'>No Records Found...!!</p>");
            closeWaitDialog();
        }
    }
    function binding(imgdata, data, i) {
        var fileNameLen = data[i].File.Name.length;
        var lstindex = data[i].File.Name.lastIndexOf('.');
        var fileExt = data[i].File.Name.substr(lstindex + 1, fileNameLen);
        fileExt = fileExt.toLowerCase();
        boxViewString += "<li thumbnail='' contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].Modified + "'";
        boxViewString += "documentTitle='" + (data[i].File.Name != null ? data[i].File.Name : "") + "' docId='" + data[i].ID + "' documentGUID='" + data[i].File.UniqueId + "'";
        boxViewString += "modifiedName='" + data[i].Editor.Title + "' serverRelativeURL='" + data[i].File.ServerRelativeUrl + "'";
        boxViewString += "ext='.docx'";
        boxViewString += ">";
        boxViewString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].File.UniqueId + "></input><div class='box-img'>";
        boxViewString += "<div>";
        boxViewString += "<img src='" + imgdata + "'class='docimgouterbox' /></div></div>";
        boxViewString += "<span style='word-break: break-all' title=" + data[i].LinkFilename + "> " + (data[i].LinkFilename.length > 25 ? data[i].LinkFilename.substring(0, 25) + "..." : data[i].LinkFilename) + "</span></li>";
        gboxViewhtml += boxViewString;
        $('#DocTemplatesBoxView').append(boxViewString);
        boxViewString = '';
        closeWaitDialog();
        i++;
        if (i < data.length) {
            toDataURL(i, data);
        }
        else {
            if (data.length > 0) {
                //gboxViewhtml = boxViewString;
                //$('#DocTemplatesBoxView').html(boxViewString);

                $("#DocTemplatesBoxView").find('input').each(function () {
                    $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                });
                closeWaitDialog();
                $("#btndropdown").removeAttr('disabled');
            }
            else {
                $("#btndropdown").attr('disabled', 'disabled');
                $("#DocTemplatesBoxView").html("<p style='color:Red'>No Records Found...!!</p>");
                closeWaitDialog();
            }
        }
    }

    function imagetoDataURL(fileURL) {
        fileURL = fileURL.replace(SPURL, "");
        var url = SPURL + templateServerRelURL + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
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
                    //callback(base64);
                };
                reader.readAsDataURL(blob);
            }
        })
        xhr.send();
    }
    function _OpenPreviewPane(currentTemplateView) {
        try {
            openWaitDialog();
            var self = this;
            var checkedDocument = $('#templateDocs:checked').val();
            var docId = "";
            var docContentTypeName = "";
            var docModifiedDate = "";
            var docTitle = "";
            var previewInfoLi = "";
            var docModifiedName = "";
            var imgSRC = "";
            var previewDestUrl = '';
            $('#previewInfo').html("");
            $('#preview-frame').css('display', 'none');
            $('#preview-frame').attr("src", "");
            if (currentTemplateView == "Box") {
                docId = $("#DocTemplatesBoxView").find("input:checked").parent().attr("docId");
                docContentTypeName = $("#DocTemplatesBoxView").find("input:checked").parent().attr("contentTypeName");
                docModifiedDate = _getFormattedDate($("#DocTemplatesBoxView").find("input:checked").parent().attr("modifiedDate"));
                docTitle = $("#DocTemplatesBoxView").find("input:checked").parent().attr("documentTitle");
                docModifiedName = $("#DocTemplatesBoxView").find("input:checked").parent().attr("modifiedName");
                imgSRC = $("#DocTemplatesBoxView").find("input:checked").next().find("img").attr("src");
            }
            else {
                docId = $("#listOfTemplate").find("input:checked").parent().attr("docId");
                docContentTypeName = $("#listOfTemplate").find("input:checked").parent().attr("contentTypeName");
                docModifiedDate = _getFormattedDate($("#listOfTemplate").find("input:checked").parent().attr("modifiedDate"));
                docTitle = $("#listOfTemplate").find("input:checked").parent().attr("documentTitle");
                docModifiedName = $("#listOfTemplate").find("input:checked").parent().attr("modifiedName");
                previewDestUrl = $("#listOfTemplate").find("input:checked").parent().attr("serverRelativeURL");
            }
            previewInfoLi = "<li><label> <b>Modified: </b>" + docModifiedDate + "</label></li>";
            previewInfoLi += "<li><label> <b>Modified By: </b>" + docModifiedName + "</label></li>";
            previewInfoLi += "<li><label> <b>Content Type: </b>" + docContentTypeName + "</label></li>";
            getMetaDataFromContentType(docContentTypeName, docId).done(function (data) {
                previewInfoLi += data;
                if (currentTemplateView == 'List') {
                    getDocumentPreview(previewDestUrl).done(function (data) {
                        $('#preview-frame').css('display', 'block');
                        $('#preview-frame').attr("src", data.toString());
                        $('#previewInfo').html("");
                        $('#previewInfo').html(previewInfoLi);
                        $(".custmbtn_preview2").removeAttr('disabled');
                        $("#previewInfo li").css("list-style", "none");
                        closeWaitDialog();
                    });
                } else {
                    $('#preview-frame').attr("src", imgSRC);
                    $('#preview-frame').css('display', 'block');
                    $('#previewInfo').html("");
                    $('#previewInfo').html(previewInfoLi);
                    $(".custmbtn_preview2").removeAttr('disabled');
                    $("#previewInfo li").css("list-style", "none");
                    closeWaitDialog();
                }
            });
            $('#closebutton').on('click', function (e) {
                $("#preview_step").css("display", "none");
                $("#myTabContent1").css("display", "block");
            });
        } catch (error) {
            console.log("_OpenPreviewPane: " + error);
        }
    }
    function getMetaDataFromContentType(docContentTypeName, docId) {
        var dfdMetadata = $.Deferred();
        var self = this;
        try {
            var cFields = [];
            var fieldsArray = [];
            var dateColumnArray = [];
            var pplColumnArray = [];
            var lkpColumnArray = [];
            var currColumnArray = [];
            var HyperLinkColumnArray = [];
            var TaxColumnArray = [];
            var expandColumn = "";
            var selectColumns = "";
            var newli = "";
            var url = SPURL + templateServerRelURL + "/_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/contenttypes?$expand=Fields";
            callAjaxGet(url).done(function (data) {
                $.map(data.d.results, function (value, key) {
                    if (value.Name == docContentTypeName) {
                        cFields = value.Fields.results;
                        for (var j = 0; j < cFields.length; j++) {
                            if (cFields[j].Group !== "_Hidden" && cFields[j].Hidden !== true && cFields[j].TypeDisplayName !== "File"
                                && cFields[j].StaticName !== "Modified_x0020_By" && cFields[j].StaticName !== "Created_x0020_By"
                                && cFields[j].StaticName !== '_dlc_DocId' && cFields[j].StaticName !== '_dlc_DocIdUrl' && cFields[j].StaticName !== '_dlc_DocIdPersistId') {
                                var item = {
                                    Title: cFields[j].Title,
                                    StaticName: cFields[j].StaticName,
                                    InternalName: cFields[j].InternalName,
                                    TypeDisplayName: cFields[j].TypeDisplayName,
                                    TypeAsString: cFields[j].TypeAsString,
                                    LookupField: cFields[j].LookupField,
                                    AllowMultipleValues: cFields[j].AllowMultipleValues,
                                    ShowAsPercentage: cFields[j].ShowAsPercentage,
                                    DisplayFormat: cFields[j].DisplayFormat,
                                    Id: cFields[j].Id,
                                    Group: cFields[j].Group
                                };
                                fieldsArray.push(item);
                            }
                        }
                    }
                });
                for (var i = 0; i < fieldsArray.length; i++) {
                    if (fieldsArray[i].TypeDisplayName === "Lookup") {
                        lkpColumnArray.push(fieldsArray[i].InternalName);
                        expandColumn += fieldsArray[i].InternalName + ',';
                        selectColumns += fieldsArray[i].InternalName + "," + fieldsArray[i].InternalName + "/Id," + fieldsArray[i].InternalName + "/" + fieldsArray[i].LookupField + ",";
                    }
                    else if (fieldsArray[i].TypeDisplayName === "Person or Group") {
                        pplColumnArray.push(fieldsArray[i].InternalName);
                        expandColumn += fieldsArray[i].InternalName + ',';
                        selectColumns += fieldsArray[i].InternalName + "," + fieldsArray[i].InternalName + "/Id," + fieldsArray[i].InternalName + "/Title,";
                    }
                    else if (fieldsArray[i].TypeDisplayName === "Managed Metadata") {
                        TaxColumnArray.push(fieldsArray[i].InternalName);
                        expandColumn += 'TaxCatchAll,';
                        selectColumns += fieldsArray[i].InternalName + ",TaxCatchAll/ID,TaxCatchAll/Term,";
                    }
                    else {
                        if (fieldsArray[i].TypeDisplayName === "Date and Time") {
                            dateColumnArray.push(fieldsArray[i].InternalName);
                        }
                        if (fieldsArray[i].TypeDisplayName === "Hyperlink or Picture") {
                            HyperLinkColumnArray.push(fieldsArray[i].InternalName);
                        }
                        if (fieldsArray[i].TypeDisplayName === "Currency") {
                            currColumnArray.push(fieldsArray[i].InternalName);
                        }
                        selectColumns += fieldsArray[i].InternalName + ",";
                    }
                }
                selectColumns = selectColumns.slice(0, -1);
                expandColumn = expandColumn.slice(0, -1);
                var finalUrl = SPURL + templateServerRelURL + "/_api/web/lists(guid'" + TemplateLibraryDisplayName + "')/items(" + docId + ")?$select=" + selectColumns + "&$expand=" + expandColumn;
                callAjaxGet(finalUrl).done(function (data) {
                    var colArrs = selectColumns.split(',');
                    $.each(colArrs, function (index, value) {
                        if (dateColumnArray.indexOf(value) > -1) {
                            var formatedDate = _getFormattedDate(data.d[value]);
                            newli += "<li><label> <b>" + value + ": </b><span>" + (formatedDate != null ? formatedDate : "") + "</span></label></li>";
                        }
                        else if (HyperLinkColumnArray.indexOf(value) > -1) {
                            for (var i = 0; i < HyperLinkColumnArray.length; i++) {
                                newli += "<li><label> <b>" + value + ": </b><span>" + data.d[HyperLinkColumnArray[i]]["Url"] + "</span></label></li>";
                            }
                        }
                        else if (pplColumnArray.indexOf(value) > -1) {
                            newli += "<li><label> <b>" + value + ": </b><span>" + (data.d[value]["Title"] != null ? data.d[value]["Title"] : "") + "</span></label></li>";
                        }
                        else if (lkpColumnArray.indexOf(value) > -1) {
                            newli += "<li><label> <b>" + value + ": </b><span>" + (data.d[value]["Title"] != null ? data.d[value]["Title"] : "") + "</span></label></li>";
                        }
                        else if (currColumnArray.indexOf(value) > -1) {
                            newli += "<li><label> <b>" + value + ": </b><span> $" + data.d[value] + "</span></label></li>";
                        }
                        else if (TaxColumnArray.indexOf(value) > -1) {
                            var term = getTaxonomyValue(data, value);
                            newli += "<li><label> <b>" + value + ": </b><span>" + term + "</span></label></li>";
                        }
                        else {
                            if (data.d[value] != undefined || data.d[value] != null) {
                                newli += "<li><label> <b>" + value + ": </b><span>" + data.d[value] + "</span></label></li>";
                            }
                        }
                    });
                    newli = newli.replace(new RegExp('_x0020_', 'g'), ' ');
                    newli = newli.replace(new RegExp('_x005f_', 'g'), '_');
                    dfdMetadata.resolve(newli);
                });
            });
        } catch (error) {
            console.log("getMetaDataFromContentType: " + error);
            dfdMetadata.reject(error);
        }
        return dfdMetadata.promise();
    }
    function getDocumentPreview(fileServerRelativeUrl) {
        var dfdDocPreview = $.Deferred();
        try {
            var self = this;
            var proxyURL = "https://cors-anywhere.herokuapp.com/";
            var previewURl = SPURL + templateServerRelURL + "_layouts/15/getpreview.ashx?path=" + SPURL +
                fileServerRelativeUrl;
            if (platform == "OfficeOnline") {
                dfdDocPreview.resolve(previewURl);
            }
            else {
                previewToDataURL(proxyURL + previewURl).done(function (dataUrl) {
                    dfdDocPreview.resolve(dataUrl);
                });
            }
        } catch (error) {
            console.log("getDocumentPreview: " + error);
            dfdDocPreview.reject(error);
        }
        return dfdDocPreview.promise();
    }
    function previewToDataURL(imgURL) {
        var dfdImgdef = $.Deferred();
        var xhr = new XMLHttpRequest();
        xhr.open('GET', imgURL, true);
        xhr.setRequestHeader('Authorization', 'Bearer ' + SPToken);
        xhr.responseType = 'blob';
        xhr.onload = function (event) {
            if (xhr.status === 200) {
                filereader(event).done(function (imgdata) {
                    dfdImgdef.resolve(imgdata);
                });
            } else {
                dfdImgdef.resolve("images/icons/doc-96.png");
            }
            function filereader(event) {
                var dfdImg = $.Deferred();
                var reader = new FileReader();
                reader.onloadend = function () {
                    dfdImg.resolve(reader.result);
                }
                reader.readAsDataURL(xhr.response);
                return dfdImg.promise();
            }
        };
        xhr.send();
        return dfdImgdef.promise();
    }
    function getTaxonomyValue(obj, columnName) {
        try {
            // Iterate over the fields in the row of data
            for (var field in obj.d) {
                if (field === columnName) {
                    // ... get the WssId from the field ...
                    var thisId = obj.d[field].WssId;
                    // ... and loop through the TaxCatchAll data to find the matching Term
                    for (var i = 0; i < obj.d.TaxCatchAll.results.length; i++) {
                        if (obj.d.TaxCatchAll.results[i]["ID"] === thisId) {
                            // Augment the fieldName object with the Term value
                            return obj.d.TaxCatchAll.results[i]["Term"];
                        }
                    }
                }
            }
            // No luck, so return null
            return null;
        } catch (error) {

        }
    }

    // Notification
    function showErrorMessage(errorMsgTxt) {
        $searchResultsDiv.addClass('alert-danger');
        $searchResultsDiv.html(errorMsgTxt);
        $searchResultsDiv.css("display", "block");
        setTimeout(function () { $searchResultsDiv.css("display", "none"); }, 5000);
    }
    function showSuccessMessage(errorMsgTxt) {
        $searchResultsDiv.removeClass('alert-danger');
        $searchResultsDiv.addClass('alert-success');
        $searchResultsDiv.html(errorMsgTxt);
        $searchResultsDiv.css("display", "block");
    }
    function removeErrorMessage() {
        $searchResultsDiv.html("");
        $searchResultsDiv.css("display", "none");
    }

    // Loader
    function openWaitDialog() {
        $searchResultsDiv.html("");
        $("#WaitDialog").show();
        $('#mainContent').addClass('blur');
    }
    function closeWaitDialog() {
        $("#WaitDialog").hide();
        $('#mainContent').removeClass('blur');
    }
}