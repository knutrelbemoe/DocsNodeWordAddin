"use strict";
var isRoot = true;
var platform;
var BRDocsNodeJS = window.BRDocsNodeJS || {};
var currentTemplateView = "Box";
var gSelectedView = "All%20Documents";
var filterViewArrayList = [];
var filteredData = "";
var RowResult = "";
var SPToken = "";
var GraphAPIToken = "";
var SPURL = "";
var listOfSiteCollectionsArray = [];
var listOfSitesArray = [];
var listOfDocLibsArray = [];
var listOfFoldersArray = [];

var listOfTeam = [];
var listOfChannel = [];
var listOfChannelFolder = [];

var listOfOneDrive = [];
var listOfOneDriveFolder = [];
let SAVE_LOCATION = {};

let ORG_ROOT_WEB = {};
let ORG_TENANT = {};
let USER_PROP = {};
let TEAM_Listname = "";
let SAVE_LOCATION_TYPE = "";
let TEMPLATE_FIRST_LOAD = "NOT_LOADED";

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
var GET_LOCATION_DETAILS_URL = 'https://docsnode-functions.azurewebsites.net/api/GetUserFolderFiles?code=e3j9/xnIbGuR1CJZwE4tnTavnwSNkI1Ky9DFxZwsaE2y2adBg0VfTQ==';
var GET_DEFAULT_LOCATION_URL = 'https://docsnode-functions.azurewebsites.net/api/GetDefaultView?code=JtcFoDYy8sKVH9UFXxIeG4KzmWEUJU7mrESLN4UmlqVYDmBmUHIaHA==';
var SET_DEFAULT_LOCATION_URL = 'https://docsnode-functions.azurewebsites.net/api/SetDefaultView?code=DbozssXppGlNDbMyhB2wwjrV9NpTjDCGd7BH5Kzdo75HhroqHsYhhQ==';
var GET_TEAM_CHANNEL_TAB_URL = "https://docsnode-functions.azurewebsites.net/api/GetTab?code=AxH/B1aXrKmPhA3fakfSZyqYufHidOXTx3nSrJ2gRAuB/MlA2pVahQ==";
const GET_MY_ORGANIZATION = "https://docsnode-functions.azurewebsites.net/api/GetMyOrganization?";
const CREATE_TEMPLATE_URL = "https://docsnode-functions.azurewebsites.net/api/CreateFileWithTemplate?code=HasN4OWnEJCplE3BsciBhTbzOcaeXnCHdsVj/iUo0cASm4Xo33tM6g==";
const CHANGE_CREATED_BY_URL = "https://docsnodecore-function.azurewebsites.net/api/UpdateCreator?code=Y1p9oneMFGN9Zs1WleUwis8DtbtMF05jc6vPUlwFJuAsiepQrMhd7A==";
const PIN_LOCATION_URL = 'https://docsnode-functions.azurewebsites.net/api/CreatePinItem?code=BBafPh2QNJiyufHn0OHfnWPZly6ho9Ky0A2pYJubWAMHg8KHHxe29g==';
const SAVE_TEMPLATE_IN_ONE_DRIVE = "https://docsnode-functions.azurewebsites.net/api/CreateFileToOneDrive?code=QM0VbZUFnB9ThmgQIyefWYObQ66KvGhre4a0rXDRXU9hzJtZb85tYg==";
const GET_TEMPLATE_FROM_ONE_DRIVE = "https://docsnode-functions.azurewebsites.net/api/GetOndriveCreatedFile?code=4VgYDDUlaw55nEW/R12DUALfAG1T0DBBNIVDZEKgXKO7AaTkhD3EYg==";
var type = ".doc";

var TemplateItemsArray = [];
var CurrentDirectory = "";
var DefaultLocation = "";


var TeamsLocArray = [];
var ChannelLocArray = [];
var OneDriveLocArray = [];
var PinnedLocArray = [];

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
                            docsNode.GetMyOrganiZation().done((value) => {
                                docsNode.GetConfigurations();
                                docsNode.init();
                                var utility = new BRTemplatesJS.Config();
                                $('#createFile').on("click", docsNode.CreateNewTemplate);
                            });

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

    this.GetMyOrganiZation = function () {

        var deferred = $.Deferred();

        $.ajax({
            url: GET_MY_ORGANIZATION,
            method: "GET",
            async: false,
            headers: {
                "Accept": "application/json; odata=verbose",
                'Authorization': 'Bearer ' + GraphAPIToken
            }
        }).then(function (result) {

            let MyOrgProfile = JSON.parse(result);

            let userProfile = MyOrgProfile.myOrgProfile.user;
            let orgRootWeb = MyOrgProfile.myOrgProfile.root;
            let orgTenant = MyOrgProfile.myOrgProfile.organization.value[0];

            setLocalForageItem("UserProfile", JSON.stringify(userProfile)).done(function (values) {
                getLocalForageItem("UserProfile").done(function (values) {
                    USER_PROP = JSON.parse(values);
                    console.log("USER_PROP :" + JSON.stringify(USER_PROP));
                    setLocalForageItem("OrgRootWeb", JSON.stringify(orgRootWeb)).done(function (values) {
                        getLocalForageItem("OrgRootWeb").done(function (values) {
                            ORG_ROOT_WEB = JSON.parse(values);
                            console.log("ORG_ROOT_WEB :" + JSON.stringify(ORG_ROOT_WEB));
                        });

                        setLocalForageItem("OrgTenant", JSON.stringify(orgTenant)).done(function (values) {
                            getLocalForageItem("OrgTenant").done(function (values) {
                                ORG_TENANT = JSON.parse(values);
                                console.log("ORG_TENANT :" + JSON.stringify(ORG_TENANT));
                                deferred.resolve(values);
                            });

                        });
                    });



                });


            });


        }).fail(function (data) {
            console.log(JSON.stringify(data));
        });

        return deferred.promise();
    };

    this.GetConfigurations = function () {
        var webUrl = SPURL;
        var url = webUrl + templateServerRelURL + "/_api/web/lists/getbytitle('DocsNodeConfiguration')/items?$select=ConfigAssestTitle,ConfigSourceList,ConfigSourceListGUID,ConfigSourceListPath";
        $.ajax({
            url: url,
            method: "GET",
            async: false,
            headers: {
                "Accept": "application/json; odata=verbose",
                'Authorization': 'Bearer ' + SPToken
            }
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
            } else {
                currentTemplateView = "Box";
                $("#noDataFoundLbl").hide();
                $('#DocTemplatesBoxView').show();
                $('#listOfTemplate').hide();
                $('#lstvw').removeClass("selectView");
                $('#grdvw').addClass("selectView");
            }
            if (platform == "OfficeOnline") {
                getalltemp(filteredData);
                // getallClientTemp(filteredData);
            } else {
                getallClientTemp(filteredData);
            }

            $("#listOfTemplate").find("input:checked").each(function () {
                if ($(this).attr('class').indexOf("default-location") === -1) {
                    $(this).prop('checked', false)
                }
            });
            $("#DocTemplatesBoxView").find("input:checked").each(function () {
                if ($(this).attr('class').indexOf("default-location") === -1) {
                    $(this).prop('checked', false)
                }
            })
            $("#previewbtn").attr('disabled', 'disabled');
            $("#previewbtn").css('background', '');
            $("#previewbtn").css('color', '');
            $("#previewbtn").css('cursor', 'default');
            $("#nextbtn").attr('disabled', 'disabled');
            $("#nextbtn").css('background-color', '');
            $("#nextbtn").css('color', '');
            $("#nextbtn").css('cursor', 'default');
            $('li:contains("' + gSelectedView + '")').addClass('selectView');

            // Add to bind Method again
            $('#DocTemplatesBoxView li, #listOfTemplate li').each(function (index, item) {
                $(this).unbind('click');
                $(this).bind({
                    click: function () {
                        let folderClicked = $(this).attr('documenttitle');
                        let contenttypename = $(this).attr('contenttypename');
                        if (contenttypename === "Folder") {
                            getListofTemplateFromSourceList(folderClicked);
                        }
                    },
                });

            });

            $('#DocTemplatesBoxView .breadcrumb-span, #listOfTemplate .breadcrumb-span').each(function (index, item) {
                $(this).unbind('click');
                $(this).bind({
                    click: function () {
                        let folderClicked = $(this).attr('path');
                        getListofTemplateFromSourceList(folderClicked, true);
                    },
                });
            });

            $('#DocTemplatesBoxView .box-default-location, #listOfTemplate .list-default-location').each(function (index, item) {
                $(this).unbind('click');
                if (DefaultLocation === CurrentDirectory) {
                    $(this).prop("checked", true);
                }
                $(this).bind({
                    click: function (event) {
                        $(this).prop('checked', event.target.checked);
                        setDefaultLocation(event.target.checked);
                    },
                });
            });

        } catch (e) {
            console.log("toggleView: " + e);
        }
    }

    function setDefaultLocation(checked) {
        openWaitDialog();
        let tempCurrentDirectory = CurrentDirectory.split("/");
        let FolderPath = tempCurrentDirectory.length === 1 ?
            "Home" : tempCurrentDirectory.slice(0, tempCurrentDirectory.length - 1).join("/");
        let payload = {
            SPOUrl: ORG_ROOT_WEB.webUrl,
            tenant: ORG_TENANT.id,
            UserEmail: USER_PROP.userPrincipalName,
            Status: "True",
        }
        if (!checked) {
            payload.FolderPath = "Home";
            DefaultLocation = "";
        } else {
            payload.FolderPath = FolderPath;
            DefaultLocation = CurrentDirectory;
        }

        $.ajax({
            url: SET_DEFAULT_LOCATION_URL,
            beforeSend: function (request) {
                request.setRequestHeader("Accept", "application/json; odata=verbose");
            },
            dataType: "json",
            headers: {
                'Authorization': 'Bearer ' + SPToken,
            },
            type: "POST",
            contentType: "application/json",
            data: JSON.stringify(payload),
            // 
        }).done(function (response) {
            closeWaitDialog();
        }).fail(function (err) {
            closeWaitDialog();
        });
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
        } else {
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
        } else {
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
        Filecount = 0;

        if (DefaultPageFlag == true) {
            $("#myTabContent1").css('display', 'none');
            $(".lib-section").css('display', 'block');
            PreviousPage();
        } else {
            $("#myTabContent1").css('display', 'none');
            $(".lib-section").css('display', 'block');
            PreviousPage();
        }
        if ($("#pinnedcheckbox").prop('checked') == true) {
            getSPPinnedLocations();

            // getPinLocation();
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
        } catch (error) {
            console.log("showSecondScreen: " + error);
        }
    };

    // save template
    this.CreateNewTemplate = function () {
        $('.alert-msg').css('display', 'none');
        $('.permissionalert-msg').css('display', 'none');

        showSavePanel();
        _showCreatePopup();

        getLocalForageItem("SaveLocation").done(function (values) {
            SAVE_LOCATION = JSON.parse(values);
            getLocalForageItem("SaveLocationType").done(function (values) {
                SAVE_LOCATION_TYPE = values;
                if (SAVE_LOCATION_TYPE.toUpperCase() === "TEAMS") {
                    LoadSaveLocationLibrary().done((resp) => {


                    });
                }
                console.log("SaveLocationType Data : " + values);
            });


            console.log("SaveLocation Data : " + values);
        });




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
            } else {
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
            if ($(this).attr('class').indexOf("default-location") === -1) {
                selected.push($(this));
            }

        });
        $("#DocTemplatesBoxView").find("input:checked").each(function () {
            if ($(this).attr('class').indexOf("default-location") === -1) {
                selected.push($(this).parent());
            }
        });

        setLocalForageItem("CurrentTemplateDir", CurrentDirectory).done(function (values) {
            //console.log("Team Data : " + values);
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
        } else if (selected.length > 1) {
            $("#previewbtn").attr('disabled', 'disabled');
            $("#previewbtn").css('background', '');
            $("#previewbtn").css('color', '');
            $("#previewbtn").css('cursor', 'default');
        } else {
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
            $("#ViewUL").mouseleave(function () {
                mousehover = true;
            });
            $("#ViewUL").mouseenter(function () {
                mousehover = false;
            });
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
            } else {
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

        let defaultURLPayload = {
            SPOUrl: ORG_ROOT_WEB.webUrl,
            UserEmail: USER_PROP.userPrincipalName,
            tenant: ORG_TENANT.id
        }
        $.ajax({
            url: GET_DEFAULT_LOCATION_URL,
            dataType: "json",
            type: "POST",
            data: JSON.stringify(defaultURLPayload),
            contentType: "application/json",
            async: false


        }).done(function (response) {

            var defaultFolderPath = "";
            var status = false;
            if (response.d.results.length > 0) {
                defaultFolderPath = response.d.results[0].FolderPath;
                status = response.d.results[0].Active;
            }

            CurrentDirectory = (defaultFolderPath === "Home" || !status) ? "" : defaultFolderPath;
            DefaultLocation = CurrentDirectory + "/";
        }).fail(function (error) {
            console.error(error);
        });

        getListofTemplateFromSourceList();

        //Treeview  --> show more/ show less events
        $(".SPTreeViewMore").on('click', function (e) {
            if ($('.treeshowmore').css('display') == 'block') {
                $('#sharepointLoc').hide();
                $('#sharepointLocAll').show();
                $('.treeshowmore').hide();
                $('.treeshowless').show();

                // get All SiteCollections - Render
                if (listOfSiteCollectionsArray != null && listOfSiteCollectionsArray.length > 0) {
                    getAllSiteCollections_Treeview_Render();
                }
            } else {
                $('#sharepointLocAll').hide();
                $('#sharepointLoc').show();
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
            $(".pin-items").children('.ms-Pivot-content:visible').each(function () {
                var tabSelected = $(this).data("content");
                switch (tabSelected) {
                    case "Team":
                        if ($('.pinshowmore').is(':visible')) {
                            $('#teamItems').hide();
                            $('#teamItemsAll').show();
                            $('.pinshowmore').hide();
                            $('.pinshowless').show();
                        } else {
                            $('#teamItems').show();
                            $('#teamItemsAll').hide();
                            $('.pinshowmore').show();
                            $('.pinshowless').hide();
                        }
                        break;
                    case "SharePoint":
                        if ($('.pinshowmore').is(':visible')) {
                            $('#SPPinned').hide();
                            $('#SPPinnedAll').show();
                            $('.pinshowmore').hide();
                            $('.pinshowless').show();
                        } else {
                            $('#SPPinned').show();
                            $('#SPPinnedAll').hide();
                            $('.pinshowmore').show();
                            $('.pinshowless').hide();
                        }

                        break;
                    case "OneDrive":
                        if ($('.pinshowmore').is(':visible')) {
                            $('#oneDriveitem').hide();
                            $('#oneDriveitemAll').show();
                            $('.pinshowmore').hide();
                            $('.pinshowless').show();
                        } else {
                            $('#oneDriveitem').show();
                            $('#oneDriveitemAll').hide();
                            $('.pinshowmore').show();
                            $('.pinshowless').hide();
                        }
                        break;
                }

            });

            // if ($('.pinshowmore').is(':visible')) {
            //     $('#SPPinned').hide();
            //     $('#SPPinnedAll').show();
            //     $('.pinshowmore').hide();
            //     $('.pinshowless').show();
            // } else {
            //     $('#SPPinned').show();
            //     $('#SPPinnedAll').hide();
            //     $('.pinshowmore').show();
            //     $('.pinshowless').hide();
            // }

            $("span").removeClass("treeselected");
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
            $("li").removeClass("treeselected");
        });

        $("#nextbtn").click(function (e) {
            //checkPinnedLocationListExist();

            // getPinLocation();
            // if (listOfSiteCollectionsArray == null || listOfSiteCollectionsArray.length == 0) {


            // get Pinned Locations
            getSPPinnedLocations();

            // get MS Team treeview
            getAllTeams_Treeview();

            // get Fav SiteCollections
            getFavSites_Treeview_Render();

            //One drive render.
            getOneDrive();

            //}
        });
        $("#nextbtn2").click(function (e) {
            // checkPinnedLocationListExist();
            //if (listOfSiteCollectionsArray == null || listOfSiteCollectionsArray.length == 0) {

            // get Pinned Locations
            getSPPinnedLocations();

            // get MS Team treeview
            getAllTeams_Treeview();

            // get Fav SiteCollections
            getFavSites_Treeview_Render();

            //One drive render.
            getOneDrive();
            //   }
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

    function pinnedlocations_click(e) {
        $("span").removeClass("treeselected");
        $("li").removeClass("treeselected");
        $(e).addClass("treeselected");

        let siteurl = $(e).data("siteurl");
        let locationname = $(e).data("locationname");
        let locationurl = $(e).data("locationurl");
        let pintype = $(e).data("pintype");

        let type = '';

        if (pintype.toLowerCase() === "onedrive") {
            siteurl = "Onedrive";
            type = "OneDrive-PinItem";
        } else if (pintype.toLowerCase() === "teams") {
            type = "Teams-PinItem";
        } else {
            type = "SharePoint-PinItem";
        }


        onSetSaveLocation(locationname, locationurl, siteurl, type, true);

        saveType = "pinned";

        $('#createFile').removeAttr('disabled');
        $("#createFile").css('background-color', '#04aba3');
        $("#createFile").css('cursor', 'pointer');
        $("#createFile").css('color', '#ffffff');
    }

    /////////////// get All Fav SiteCollections - Render///////////////////// Modified on 14.07.2020 by Arijit to bind data in SP tab
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

                    //  $('#SPFavTreeView').html(treeViewHTMLFav); 
                    $('#sharepointLoc').html(treeViewHTMLFav);
                    $("#treeviewUL>li>span.treeSpan").off('click');
                    $("#treeviewUL>li>span.treeSpan").on('click', function () {
                        getAllSites_Treeview_click(this);
                    });
                    var cntHeight = $('.side_body_shadow').height();
                    console.log('cnt', cntHeight);
                    $('.tab_sidemounted_area').css({
                        'width': cntHeight
                    });
                    $('.tab_sidemounted_area').fadeIn(100);

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
        } else {
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
            $('#sharepointLocAll').html(treeViewHTML);

            $("#treeviewUL>li>span.treeSpan").click(function () {
                getAllSites_Treeview_click(this);
            });
        }
    }

    /////////////// get All SubSites Of SiteCollection - render/////////////////////
    function getAllSites_Treeview_click(e) {

        onSetSaveLocation("", "", "", "SharePoint");

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

                            $("ul[id='" + skey + "'] .treeSpan1").off("click");
                            $("ul[id='" + skey + "'] .treeSpan1").on("click", function (event) {
                                getAllSites_Treeview_click($(this));
                            });



                        }
                    }
                    listOfSitesArray = [];

                });
        } else {

            var docLibKey = $(e).find('.sitekey').text().trim();

            if (docLibKey != "") {
                $("ul[id='" + docLibKey + "']").addClass("nested");
                $("ul[id='" + docLibKey + "']").prev('span').addClass("caret-down");
            }
        }
    }

    // get All Folders From Document Library - render
    function getAllFoldersFromLibrary_Treeview_click(e) {

        onSetSaveLocation("", "", "", "SharePoint");

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
            } else {
                $("#createFile").attr('disabled', 'disabled');
                $("#createFile").css('background-color', '');
                $("#createFile").css('cursor', 'default');
            }

            getAllFoldersFromLibrary_Treeview(appWebUrl, displayName, selectedLibURL, siteName).then(function (folData) {

                if (listOfFoldersArray.length > 0) {
                    var liHTML = "";
                    for (var i = 0; i < listOfFoldersArray.length; i++) {
                        var folderName = listOfFoldersArray[i].folderName;
                        var folderURL = listOfFoldersArray[i].folderURL;
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

                        $("ul[id='" + skey + "'] .treeFolder").off("click");
                        $("ul[id='" + skey + "'] .treeFolder").on("click", function (event) {
                            Folder_Treeview_click($(this));
                        });


                    }
                } else {
                    var txt = docLibKey.trim();
                    $("div.docLibKey:contains(" + txt + ")").parent().removeClass("caretCustom");
                    $("div.docLibKey:contains(" + txt + ")").parent().removeClass("caret-down");

                    $("span").removeClass("treeselected");
                    $("li").removeClass("treeselected");
                    saveType = "";
                    $("div.docLibKey:contains(" + txt + ")").parent().addClass("treeselected");
                }
            });
        } else {

            var docLibKey = $(e).find('.docLibKey').text().trim();

            if (docLibKey != "") {
                $("ul[id='" + docLibKey + "']").addClass("nested");
                $("ul[id='" + docLibKey + "']").prev('span').addClass("caret-down");
            }
        }
    }

    // Folder click event for selection
    function Folder_Treeview_click(e) {
        onSetSaveLocation("", "", "", "SharePoint");
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
            } else if (appWebUrl) {
                apiURL = appWebUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?$expand=Folder,File&$select=ID,Title,FileLeafRef,Folder/ServerRelativeUrl";
            } else {
                apiURL = SPURL + "/sites/" + siteName + "/_api/web/lists/getbytitle('" + listName + "')/items?$expand=Folder,File&$select=ID,Title,FileLeafRef,Folder/ServerRelativeUrl";
            }
            $.ajax({
                url: apiURL,
                method: "GET",
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Authorization': 'Bearer ' + SPToken
                },
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
                            listOfFoldersArray.push({
                                "appWebUrl": appWebUrl,
                                "listName": listName,
                                "selectedLibrarywebURL": selectedLibURL,
                                "folderURL": data[i].Folder.ServerRelativeUrl,
                                "folderName": data[i].FileLeafRef
                            })
                        }
                    }
                }

                dfdLib.resolve(listOfFoldersArray);
                closeWaitDialog();
            }
        } catch (error) {
            console.log("getAllFoldersFromLibrary_Treeview: " + error);
            dfdLib.reject(listOfFoldersArray);
            closeWaitDialog();
        }
        return dfdLib.promise();
    }

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
                } else if (site.split("/sites/")[1] == undefined & (isRoot)) {
                    if (site == "Root") {
                        if (cutmAttr == "rootsubsites") {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + siteId + "/lists";
                        } else {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/lists";
                        }

                    } else {
                        if (cutmAttr == "rootsubsites") {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + siteId + "/lists";
                        } else {
                            GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":" + site + ":/lists";
                        }
                    }
                } else {
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
                                "site": site,
                                "appWebUrl": webURL,
                                "siteURL": result[i].webUrl,
                                "displayName": result[i].displayName,
                                "name": result[i].name,
                                "parentsiteKey": parentsiteKey,
                                "siteKey": siteKey,
                                "docLibKey": docLibKey
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
                    } else {
                        GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + siteCollection + ":/sites";
                    }
                } else {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + "/sites/";
                }
            } else {
                if (siteCollection.split(":")[1] == undefined) {
                    GraphAPI = "https://graph.microsoft.com/v1.0/sites/" + tenantName + ":/sites/" + siteCollection + ":/sites";
                } else {
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

                                $("ul[id='" + skey + "'] span.treeDocLib").off("click");
                                $("ul[id='" + skey + "'] span.treeDocLib").on("click", function (event) {
                                    getAllFoldersFromLibrary_Treeview_click($(this));
                                });
                            }
                            listOfDocLibsArray = [];
                            closeWaitDialog();
                        } else {

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
                                    "ParentSite": siteCollection,
                                    "ParentSiteURL": siteCollectionURL,
                                    "hasSubSite": true,
                                    "SubSiteDisplayName": result[i].name,
                                    "SubSiteName": result[i].name,
                                    "SubSiteURL": result[i].webUrl,
                                    "IsRoot": isCollection,
                                    "parentSiteKey": parentsiteKey,
                                    "SiteKey": siteKey,
                                    "level": level,
                                    "siteId": siteId
                                })
                            }
                        }
                        if (hasNoSubsite) {
                            listOfSitesArray.push({
                                "ParentSite": siteCollection,
                                "ParentSiteURL": siteCollectionURL,
                                "hasSubSite": false,
                                "SubSiteDisplayName": "",
                                "SubSiteName": "",
                                "SubSiteURL": "",
                                "IsRoot": isCollection,
                                "parentSiteKey": parentsiteKey,
                                "SiteKey": "",
                                "level": 0,
                                "siteId": siteId
                            })
                        }
                    } else {
                        listOfSitesArray.push({
                            "ParentSite": siteCollection,
                            "ParentSiteURL": siteCollectionURL,
                            "hasSubSite": false,
                            "SubSiteDisplayName": "",
                            "SubSiteName": "",
                            "SubSiteURL": "",
                            "IsRoot": isCollection,
                            "parentSiteKey": parentsiteKey,
                            "SiteKey": "",
                            "level": 0,
                            "siteId": siteId
                        })
                    }
                    dfd.resolve(listOfSitesArray);
                } else {
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
                        listOfSiteCollectionsArray.push({
                            "siteUrl": siteUrl,
                            "siteTitle": siteTitle,
                            "siteKey": siteKey,
                            "siteId": ""
                        })
                    }
                }
                dfd.resolve(listOfSiteCollectionsArray);
            });
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };

    // Treeview --> Create document in lib
    function createDocumentInDestLib_Treeview(checkNext) {
        try {
            openWaitDialog();
            $('#alertMessage,#errorMessage').css('display', 'none');
            $('.alert-msg').css('display', 'none');
            $('.permissionalert-msg').css('display', 'none');
            var docName = $('#txtNewFileName').val();
            docName = docName.trim();
            var docnameLen = docName.length;
            if (docName == "" || docName == null || docName.match('%') || docName.match('"') ||
                docName.match('\'') || docName.match(';') || docName.match('#')) {
                $('#alertMessage').css('display', 'block');
                closeWaitDialog();
                return false;
            } else {
                var selectedSiteRelativeURL = ""
                var webRedirectURL = ""; // this is used to create open Document URL
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
                if (!SAVE_LOCATION.DestFolderRelUrl) {
                    if ($(".treeselected") == null || $(".treeselected").length == 0) {
                        $("#PinnedLocationMsg").html("");
                        pinnedString = "<p>Please select at least one location.</p>";
                        $("#PinnedLocationMsg").append(pinnedString);
                        closeWaitDialog();
                        return;
                    } else if ($(".treeselected").length == 1) {
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
                        } else {
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
                    } else {
                        $("#PinnedLocationMsg").html("");
                        pinnedString = "<p>Please select only one location.</p>";
                        $("#PinnedLocationMsg").append(pinnedString);
                        return;
                    }
                }
                chkDocTitle = oldfile;
                getLocalForageItem("CurrentTemplateDir").done(function (values) {
                    const relativeTemplateDir = values;
                    if (relativeTemplateDir !== "/") {
                        if (relativeTemplateDir.slice(-1) === "/") {
                            chkDocTitle = relativeTemplateDir + oldfile;
                        } else {
                            chkDocTitle = relativeTemplateDir + "/" + oldfile;
                        }

                    }

                    if (FolderRelativePath === "") {
                        webRedirectURL += "/_api/web/GetFolderByServerRelativeUrl('" + DocLibraryUrl + "/')";
                    } else {
                        webRedirectURL += "/_api/web/GetFolderByServerRelativeUrl('" + FolderRelativePath + "/')";
                    }

                    var docName = $('#txtNewFileName').val();
                    var folderName = $('#SPDocFolders').val();
                    var docnameLen = docName.length;
                    var newFileName = docName + docsNodeNewFileExtention;
                    var destServerRelURL = destinationServerRelativeUrl;
                    if (chkDocTitle !== "") {


                        createDocument(newFileName, webRedirectURL, chkDocTitle, tokenUrl, destServerRelURL).then(function (data) {

                                if (SAVE_LOCATION_TYPE.toUpperCase() === "SHAREPOINT" || SAVE_LOCATION_TYPE.toUpperCase() === "SHAREPOINT-PINITEM") {
                                    openEditProperties(data, selectedSiteRelativeURL, tokenUrl).then(function (data) {
                                        var documentURL = data.d.ServerRedirectedEmbedUri;
                                        var editFormURl = "";
                                        documentURL = documentURL.replace("=interactivepreview", "=edit");
                                        if (flag == 1) {
                                            editFormURl = SPURL + selectedSiteRelativeURL + DocumentInternalName + '/Forms/EditForm.aspx?ID=' + data.d.ID;
                                            if (platform == "PC") {
                                                // if (FolderRelativePath != null && FolderRelativePath != "") {
                                                //     getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href='ms-word:ofe|u|" + SPURL + FolderRelativePath + "/" + newFileName + "'> " + newFileName + ". </a></p>\n";
                                                // } else {
                                                //     getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href='ms-word:ofe|u|" + SPURL + DocLibraryUrl + "/" + newFileName + "'> " + newFileName + ". </a></p>\n";
                                                // }
                                                getdocumentUrlsString = `<p class='saved-file-list'>Your Document is saved <strong>${newFileName}</strong>. </p>`;
                                                getdocumentUrlsString += `<p class='saved-file-list'><a class='anchor-button' target='_blank' href=${documentURL}> View in Browser </a>`;
                                                getdocumentUrlsString += `<a class='anchor-button' target='_blank' href='ms-word:ofe|u|${ORG_ROOT_WEB.webUrl}${DocLibraryUrl}/${newFileName}'>View in App</a></p>`


                                            } else {
                                                getdocumentUrlsString = `<p class='saved-file-list'>Your Document is saved <strong>${newFileName}</strong>. </p>`;
                                                getdocumentUrlsString += `<p class='saved-file-list'><a class='anchor-button' target='_blank' href=${documentURL}> View in Browser </a>`;
                                                getdocumentUrlsString += `<a class='anchor-button' target='_blank' href='ms-word:ofe|u|${ORG_ROOT_WEB.webUrl}${DocLibraryUrl}/${newFileName}'>View in App</a></p>`

                                                // getdocumentUrlsString = "<p>Your Document is saved <a class='docUrls' target='_blank' href=" + documentURL + "> " + newFileName + ". </a></p>\n";
                                            }
                                            $("#DocumentUrls").append(getdocumentUrlsString);
                                            closeWaitDialog();
                                        } else {
                                            getdocumentUrlsString = `<p class='saved-file-list'>Your Document is saved <strong>${newFileName}</strong>. </p>`;
                                            getdocumentUrlsString += `<p class='saved-file-list'><a class='anchor-button' target='_blank' href=${documentURL}> View in Browser </a>`;
                                            getdocumentUrlsString += `<a class='anchor-button' target='_blank' href='ms-word:ofe|u|${ORG_ROOT_WEB.webUrl}${DocLibraryUrl}/${newFileName}'>View in App</a></p>`
                                            $("#DocumentUrls").append(getdocumentUrlsString);
                                            closeWaitDialog();
                                        }
                                        console.log("File Saved! " + Filecount);
                                        if (Filecount == TotalPages) {
                                            $("#third_step").find("input").attr("disabled", 'disabled');
                                            currentPage = null, TotalPages = null;
                                            len = 0, Filecount = 0, oldfile = "";
                                            destinationServerRelativeUrl = '';
                                        }
                                    }, function (openEditerFileFail) {
                                        console.log(openEditerFileFail);
                                    });
                                }
                                Filecount = Filecount + 1;
                                if (Filecount < TotalPages) {
                                    if (currentTemplateView == "List") {
                                        oldfile = $("#listOfTemplate").find("input:checked").eq(Filecount).parent().attr("documentTitle");
                                        destinationServerRelativeUrl = $("#listOfTemplate").find("input:checked").eq(Filecount).parent().attr("serverrelativeURL");
                                    } else {
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

                                if (checkNext != false && SAVE_LOCATION_TYPE.toUpperCase().includes("SHAREPOINT")) {
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
                                            "DocumentLibraryURL": {
                                                '__metadata': {
                                                    'type': 'SP.FieldUrlValue'
                                                },
                                                'Description': DocLibraryUrl,
                                                'Url': DocLibraryUrl
                                            },
                                            "PinnedType": type,
                                            "SiteURL": {
                                                '__metadata': {
                                                    'type': 'SP.FieldUrlValue'
                                                },
                                                'Description': siteurl,
                                                'Url': siteurl
                                            },
                                        };
                                        if ($("#pinnedcheckbox").prop('checked') == true) {
                                            $("#PinnedLocationMsg").html("");
                                            checkExistingPinned(DocLibraryUrl).then(function (data) {
                                                if (!IsPinnedLocation) {
                                                    formSaveNotes(itemArray);
                                                } else {
                                                    pinnedString = "<p>Selected location is already pinned.</p>";
                                                    $("#PinnedLocationMsg").append(pinnedString);
                                                }
                                            });
                                        }
                                    }
                                }
                                /////

                            },
                            function (errorMsg) {
                                console.log(errorMsg);
                                closeWaitDialog();
                            });

                    } else {
                        closeWaitDialog();
                    }

                });
            }
        } catch (error) {
            console.log("createDocumentInDestLib_Treeview : " + error);
            closeWaitDialog();
        }
    }

    function GetClientAppLinkforOneDrive(destUrl, destFolder, filename) {

        let fileUrl = "";
        if (destFolder !== "") {
            fileUrl = destUrl.substr(0, destUrl.toLowerCase().indexOf("/_layouts") + 1) + "/documents/" + destFolder + "/" + filename;
        } else {
            fileUrl = destUrl.substr(0, destUrl.toLowerCase().indexOf("/_layouts") + 1) + "/documents/" + filename;
        }

        fileUrl = fileUrl.replace("//", "/").replace("https:/", "https://");


        let clientAppLink = "";

        if (filename.toLowerCase().includes("ppt") || filename.toLowerCase().includes("potx")) {
            clientAppLink = "ms-powerpoint:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("xls") || filename.toLowerCase().includes("xltx")) {
            clientAppLink = "ms-excel:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("doc") || filename.toLowerCase().includes("dotx")) {
            clientAppLink = "ms-word:ofe|u|" + fileUrl;
        }
        return clientAppLink;
    }

    function GetClientAppLinkforTeams(destUrl, destFolder, filename) {

        let fileUrl = "";
        if (destFolder.toLowerCase().includes("sites")) {
            fileUrl = ORG_ROOT_WEB.webUrl + "/" + destFolder + "/" + filename;
        } else {
            fileUrl = destUrl + "/" + destFolder + "/" + filename;
        }

        fileUrl = fileUrl.replace("//", "/").replace("https:/", "https://");


        let clientAppLink = "";

        if (filename.toLowerCase().includes("ppt") || filename.toLowerCase().includes("potx")) {
            clientAppLink = "ms-powerpoint:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("xls") || filename.toLowerCase().includes("xltx")) {
            clientAppLink = "ms-excel:ofe|u|" + fileUrl;
        }
        if (filename.toLowerCase().includes("doc") || filename.toLowerCase().includes("dotx")) {
            clientAppLink = "ms-word:ofe|u|" + fileUrl;
        }
        return clientAppLink;
    }

    function LoadSaveLocationLibrary() {
        var dfd = $.Deferred();
        const eachTeamLibPayload = {
            tenant: ORG_TENANT.id,
            SPOUrl: ORG_ROOT_WEB.webUrl,
            Team: SAVE_LOCATION.DestSite
        };
        $.ajax({
            url: GET_TEAM_LIBRARY_NAMES,
            beforeSend: function (request) {
                request.setRequestHeader("Accept", "application/json; odata=verbose");
            },
            dataType: "json",
            headers: {
                'Authorization': 'Bearer ' + SPToken,
            },
            type: "POST",
            contentType: "application/json",
            data: JSON.stringify(eachTeamLibPayload),
            // 
        }).done((data) => {

            let teamLibInternalNames = [];
            data.d.results.map(libNames => {
                const eachLibInternalNames = {};

                // eachTemplate.isSelected = false;
                if (libNames.Title && (libNames.Title.toLowerCase() === "documents" || libNames.Title.toLowerCase() === "dokumenter")) {
                    eachLibInternalNames.EntityTypeName = libNames.EntityTypeName.toString().replace("_x0020_", " ");
                    eachLibInternalNames.Title = decodeURI(libNames.Title);
                    teamLibInternalNames.push(eachLibInternalNames);
                }


            });

            if (teamLibInternalNames && teamLibInternalNames.length > 0) {
                TEAM_Listname = teamLibInternalNames[0].Title;
            }
            dfd.resolve(TEAM_Listname);
        });
        return dfd.promise();
    }

    function saveTemplateInTeams(currentTemplate, inputDocumentName) {

        var dfd = $.Deferred();
        let documentExtention = currentTemplate.substr(currentTemplate.indexOf("."));
        let inputDocument = {
            name: inputDocumentName.substring(0, inputDocumentName.indexOf(documentExtention))
        };
        let savedDocumentName = inputDocument.name;
        try {
            let payloadCreateFile = {
                // tenant: "docsnode.com",
                tenant: ORG_TENANT.id,
                SPOUrl: ORG_ROOT_WEB.webUrl,
                sourceFileName: currentTemplate,
                DestFolderRelUrl: SAVE_LOCATION.DestFolderRelUrl,
                DestSite: SAVE_LOCATION.DestSite,
                FileName: inputDocument.name
            };

            console.log(payloadCreateFile);

            $.ajax({
                url: CREATE_TEMPLATE_URL,
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json; odata=verbose");
                },
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + SPToken,
                },
                type: "POST",
                contentType: "application/json",
                data: JSON.stringify(payloadCreateFile),
                // 
            }).done((data) => {
                console.log(data.d);
                if (data.error && data.error.code === "-2130575257, Microsoft.SharePoint.SPException") {
                    inputDocument.isValid = false;
                    inputDocument.errorMessage = 'Duplicate file name, file already exists';
                    $('#errorMessage').html(inputDocument.errorMessage).css('display', 'block');
                } else {

                    inputDocument.isValid = false;
                    inputDocument.name = "";
                    inputDocument.errorMessage = "";

                    let documentURL = data.d.ServerRedirectedEmbedUri.replace("&action=interactivepreview", "");
                    let clientAppurl = GetClientAppLinkforTeams(SAVE_LOCATION.DestSite, SAVE_LOCATION.DestFolderRelUrl, inputDocumentName);
                    var getdocumentUrlsString = `<p class='saved-file-list'>Your Document is saved <strong>${inputDocumentName}</strong>. </p>`;
                    getdocumentUrlsString += `<p class='saved-file-list'><a class='anchor-button' target='_blank' href=${documentURL}> View in Browser </a>`;
                    getdocumentUrlsString += `<a class='anchor-button' target='_blank' href='${clientAppurl}'>View in App</a></p>`

                    $("#DocumentUrls").append(getdocumentUrlsString);
                    closeWaitDialog();
                    if (Filecount == TotalPages) {
                        $("#third_step").find("input").attr("disabled", 'disabled');
                        currentPage = null, TotalPages = null;
                        len = 0, Filecount = 0, oldfile = "";

                    }

                    dfd.resolve(data);

                    let payloadChangeCreatedBy = {
                        tenantName: ORG_ROOT_WEB.siteCollection.hostname.split(".")[0],
                        siteUrl: SAVE_LOCATION.DestSite.toString(),
                        listName: TEAM_Listname,
                        // listName: "Documents",
                        itemID: data.d.Id,
                        emailid: USER_PROP.userPrincipalName
                    };

                    // console.log(payloadChangeCreatedBy);
                    $.ajax({
                        url: CHANGE_CREATED_BY_URL,
                        beforeSend: function (request) {
                            request.setRequestHeader("Accept", "application/json; odata=verbose");
                        },
                        dataType: "json",
                        headers: {
                            'Authorization': 'Bearer ' + SPToken,
                        },
                        type: "POST",
                        contentType: "application/json",
                        data: JSON.stringify(payloadChangeCreatedBy),
                        // 
                    }).done((dataChangeCreatedBy) => {
                        // console.log(dataChangeCreatedBy);


                    });
                    var pinlocationChecked = $('#pinnedcheckbox').is(":checked");
                    if (pinlocationChecked && !SAVE_LOCATION.LocationAlreadyPinned) {
                        let payloadPinLocation = {
                            tenant: ORG_TENANT.id,
                            SPOUrl: ORG_ROOT_WEB.webUrl,
                            PinnedType: "Teams",
                            DocumentLibrary: SAVE_LOCATION.DestFolderRelUrl,
                            DocumentLibraryUrl: SAVE_LOCATION.DestSite,
                            SiteUrl: SAVE_LOCATION.SiteUrl

                        };
                        $.ajax({
                            url: PIN_LOCATION_URL,
                            beforeSend: function (request) {
                                request.setRequestHeader("Accept", "application/json; odata=verbose");
                            },
                            dataType: "json",
                            headers: {
                                'Authorization': 'Bearer ' + SPToken,
                            },
                            type: "POST",
                            contentType: "application/json",
                            data: JSON.stringify(payloadPinLocation),
                            // 
                        }).done((dataPinLocation) => {
                            let payloadChangeCreatedByPinLoc = {
                                tenantName: ORG_ROOT_WEB.siteCollection.hostname.split(".")[0],
                                siteUrl: ORG_ROOT_WEB.webUrl,
                                listName: "DocsNodePinnedLocations",
                                itemID: dataPinLocation.d.Id,
                                emailid: USER_PROP.userPrincipalName
                            };
                            $.ajax({
                                url: CHANGE_CREATED_BY_URL,
                                beforeSend: function (request) {
                                    request.setRequestHeader("Accept", "application/json; odata=verbose");
                                },
                                dataType: "json",
                                headers: {
                                    'Authorization': 'Bearer ' + SPToken,
                                },
                                type: "POST",
                                contentType: "application/json",
                                data: JSON.stringify(payloadChangeCreatedByPinLoc),
                                // 
                            }).done((dataChangeCreatedByPinLoc) => {



                            });
                        });

                    }
                }
            });


        } catch (error) {
            console.log('copyFile: ' + error);
            dfd.reject(error);
        }
        return dfd.promise();

    }

    function saveTemplateInOneDrive(currentTemplate, inputDocumentName) {
        var dfd = $.Deferred();
        let documentExtention = currentTemplate.substr(currentTemplate.indexOf("."));
        let inputDocument = {
            name: inputDocumentName.substring(0, inputDocumentName.indexOf(documentExtention))
        };
        let savedDocumentName = inputDocument.name
        try {
            let payloadCreateFile = {
                // tenant: "docsnode.com",
                tenant: ORG_TENANT.id,
                SPOUrl: ORG_ROOT_WEB.webUrl,
                FileName: inputDocument.name,
                sourceFileName: currentTemplate,
                FolderName: SAVE_LOCATION.DestFolderRelUrl,
                userGuidId: USER_PROP.id,
            };

            console.log(payloadCreateFile);


            $.ajax({
                url: SAVE_TEMPLATE_IN_ONE_DRIVE,
                beforeSend: function (request) {
                    request.setRequestHeader("Accept", "application/json; odata=verbose");
                },
                dataType: "json",
                headers: {
                    'Authorization': 'Bearer ' + SPToken,
                },
                type: "POST",
                contentType: "application/json",
                data: JSON.stringify(payloadCreateFile),
                // 
            }).done((data) => {
                console.log(data.d);
                if (data.error && data.error.code === "-2130575257, Microsoft.SharePoint.SPException") {
                    inputDocument.isValid = false;
                    inputDocument.errorMessage = 'Duplicate file name, file already exists';
                    $('#errorMessage').html(inputDocument.errorMessage).css('display', 'block');
                } else {

                    inputDocument.isValid = false;
                    inputDocument.name = "";
                    inputDocument.errorMessage = "";


                    dfd.resolve(data);

                    setTimeout(() => {
                        $.ajax({
                            url: GET_TEMPLATE_FROM_ONE_DRIVE,
                            beforeSend: function (request) {
                                request.setRequestHeader("Accept", "application/json; odata=verbose");
                            },
                            dataType: "json",
                            headers: {
                                'Authorization': 'Bearer ' + SPToken,
                            },
                            type: "POST",
                            contentType: "application/json",
                            data: JSON.stringify(payloadCreateFile),
                            // 
                        }).done((templateData) => {

                            closeWaitDialog();
                            if (Filecount == TotalPages) {
                                $("#third_step").find("input").attr("disabled", 'disabled');
                                currentPage = null, TotalPages = null;
                                len = 0, Filecount = 0, oldfile = "";

                            }
                            let documentURL = templateData.webUrl;
                            let clientAppurl = GetClientAppLinkforOneDrive(templateData.webUrl, SAVE_LOCATION.DestFolderRelUrl, inputDocumentName);
                            var getdocumentUrlsString = `<p class='saved-file-list'>Your Document is saved <strong>${inputDocumentName}</strong>. </p>`;
                            getdocumentUrlsString += `<p class='saved-file-list'><a class='anchor-button' target='_blank' href=${documentURL}> View in Browser </a>`;
                            getdocumentUrlsString += `<a class='anchor-button' target='_blank' href='${clientAppurl}'>View in App</a></p>`

                            $("#DocumentUrls").append(getdocumentUrlsString);

                            var pinlocationChecked = $('#pinnedcheckbox').is(":checked");

                            if (pinlocationChecked && !SAVE_LOCATION.LocationAlreadyPinned) {
                                let payloadPinLocation = {
                                    tenant: ORG_TENANT.id,
                                    SPOUrl: ORG_ROOT_WEB.webUrl,
                                    PinnedType: "Onedrive",
                                    DocumentLibrary: SAVE_LOCATION.DestFolderRelUrl,
                                    DocumentLibraryUrl: SAVE_LOCATION.DestSite,
                                    SiteUrl: SAVE_LOCATION.DestSite

                                };
                                $.ajax({
                                    url: PIN_LOCATION_URL,
                                    beforeSend: function (request) {
                                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                                    },
                                    dataType: "json",
                                    headers: {
                                        'Authorization': 'Bearer ' + SPToken,
                                    },
                                    type: "POST",
                                    contentType: "application/json",
                                    data: JSON.stringify(payloadPinLocation),
                                    // 
                                }).done((dataPinLocation) => {

                                    let payloadChangeCreatedByPinLoc = {
                                        tenantName: ORG_ROOT_WEB.siteCollection.hostname.split(".")[0],
                                        siteUrl: ORG_ROOT_WEB.webUrl,
                                        listName: "DocsNodePinnedLocations",
                                        itemID: dataPinLocation.d.Id,
                                        emailid: USER_PROP.userPrincipalName
                                    };
                                    $.ajax({
                                        url: CHANGE_CREATED_BY_URL,
                                        beforeSend: function (request) {
                                            request.setRequestHeader("Accept", "application/json; odata=verbose");
                                        },
                                        dataType: "json",
                                        headers: {
                                            'Authorization': 'Bearer ' + SPToken,
                                        },
                                        type: "POST",
                                        contentType: "application/json",
                                        data: JSON.stringify(payloadChangeCreatedByPinLoc),
                                        // 
                                    }).done((dataChangeCreatedByPinLoc) => {



                                    });

                                });
                            } else {

                            }
                        });

                    }, 7000);
                }

            });
        } catch (error) {
            console.log('copyFile: ' + error);
            dfd.reject(error);
        }
        return dfd.promise();
    }

    function setOnedriveRootasDestination(ele) {

        let content = $(ele).data('content');

        if (content.toUpperCase() === "ONEDRIVE") {

            getLocalForageItem("OneDrive").done(function (values) {

                listOfOneDrive = JSON.parse(values);

                if (listOfOneDrive && listOfOneDrive.length > 0) {
                    onSetSaveLocation("/", listOfOneDrive[0].DriveUrl, "Onedrive", "OneDrive");


                    $('#createFile').removeAttr('disabled');
                    $("#createFile").css('background-color', '#04aba3');
                    $("#createFile").css('cursor', 'pointer');
                    $("#createFile").css('color', '#ffffff');
                }

            });

        } else {
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
        }
    }


    // Treeview --> get Pinned Locations  

    //Modified this function on 13th july by Arijit for Tab section for pinned location. [Start here]
    function getSPPinnedLocations() {
        // var dfd = $.deferred();
        //  openwaitdialog();
        if (SPToken) {

            let payload = {
                SPOUrl: ORG_ROOT_WEB.webUrl,
                UserName: USER_PROP.displayName,
                tenant: ORG_TENANT.id
            }

            $.ajax({
                    url: GET_PIN_LOCATION,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {


                    var resultArr = response.d.results;

                    var sharepoint = [];
                    var OneDrive = [];
                    var teams = [];
                    resultArr.map(a => {
                        //console.log(a.PinnedType)
                        if (a.PinnedType.toLowerCase() == 'documentlibrary' || a.PinnedType.toLowerCase() == 'folder') {
                            sharepoint.push(a);
                            console.log(sharepoint);
                        } else if (a.PinnedType.toLowerCase() == 'onedrive') {
                            OneDrive.push(a);
                            console.log(OneDrive);
                        } else if (a.PinnedType.toLowerCase() == 'teams') {
                            teams.push(a);
                            console.log(teams);
                        } else {

                        }

                    });
                    //teams sec pinned location data bind
                    var teamHtml = "<ul class='pin_location '>";
                    var teamhtmlAll = "";
                    if (teams.length > 0) {
                        for (var i = 0; i < teams.length; i++) {
                            if (i < 3) {
                                teamHtml += "<li class='pinnedselected pinnedselect' data-siteurl ='" + decodeURI(teams[i].SiteURL.Url) +
                                    "' data-locationname ='" + teams[i].DocumentLibrary +
                                    "' data-locationurl ='" + decodeURI(teams[i].DocumentLibraryURL.Url) +
                                    "' data-pintype ='" + teams[i].PinnedType +
                                    "' >";
                                teamHtml += "       <div class='pindoc'><div hidden class='siteurl'>" + teams[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + teams[i].PinnedType + "</div><div hidden class='pinname'>" + teams[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + teams[i].DocumentLibraryURL.Description.trim() + "</div>";
                                teamHtml += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                teamHtml += "            <h4>" + teams[i].DocumentLibrary + "</h4>";
                                teamHtml += "            <span class='path'>" + teams[i].DocumentLibraryURL.Url.trim() + "</span>";
                                teamHtml += "        </div>";
                                teamHtml += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + teams[i].ID + "</div>";
                                teamHtml += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                teamHtml += "        </a>";
                                teamHtml += "    </li>";
                            } else {
                                if (teamhtmlAll == "") {
                                    teamhtmlAll += teamHtml;
                                    if ($('#teamItemsAll').css('display') == 'none') {
                                        $('.pinshowmore').show();
                                    }
                                }
                                teamhtmlAll += "<li class='pinnedselected pinnedselect' data-siteurl ='" + decodeURI(teams[i].SiteURL.Url) +
                                    "' data-locationname ='" + teams[i].DocumentLibrary +
                                    "' data-locationurl ='" + decodeURI(teams[i].DocumentLibraryURL.Url) +
                                    "' data-pintype ='" + teams[i].PinnedType +
                                    "' >";
                                teamhtmlAll += "       <div class='pindoc'><div hidden class='siteurl'>" + teams[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + teams[i].PinnedType + "</div><div hidden class='pinname'>" + teams[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + teams[i].DocumentLibraryURL.Description.trim() + "</div>";
                                teamhtmlAll += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                teamhtmlAll += "            <h4>" + teams[i].DocumentLibrary + "</h4>";
                                teamhtmlAll += "            <span class='path'>" + teams[i].DocumentLibraryURL.Url.trim() + "</span>";
                                teamhtmlAll += "        </div>";
                                teamhtmlAll += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + teams[i].ID + "</div>";
                                teamhtmlAll += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                teamhtmlAll += "        </a>";
                                teamhtmlAll += "    </li>";
                            }
                        }
                        teamHtml += "</ul>";
                        teamhtmlAll += "</ul>";
                        //if (teams.length == 3) {
                        //$('.pinshowless').hide();
                        $('#teamItemsAll').css('display', 'none');
                        $('#teamItems').css('display', 'block');
                        // }
                        $('#teamItems').html(teamHtml);
                        $('#teamItemsAll').html(teamhtmlAll);

                        $("#createfile").attr('disabled', 'disabled');
                        $("#createfile").css('cursor', 'default');
                        $("#createfile").css('background-color', '');

                    } else {
                        teamHtml = "<p>    no pinned locations are found..!! </p>";
                        $('#teamItems').html(teamHtml);
                        $('#teamItemsAll').html(teamhtmlAll);
                    }


                    // onedrive section Pinnedlocation data bind
                    var onedriveHtml = "<ul class='pin_location'>";
                    var onedriveHtmlAll = "";
                    if (OneDrive.length > 0) {
                        for (var i = 0; i < OneDrive.length; i++) {
                            if (i < 3) {
                                onedriveHtml += "<li class='pinnedselected pinnedselect' data-siteurl ='" + decodeURI(OneDrive[i].SiteURL.Url) +
                                    "' data-locationname ='" + OneDrive[i].DocumentLibrary +
                                    "' data-locationurl ='" + decodeURI(OneDrive[i].DocumentLibraryURL.Url) +
                                    "' data-pintype ='" + OneDrive[i].PinnedType +
                                    "' >";
                                onedriveHtml += "       <div class='pindoc'><div hidden class='siteurl'>" + OneDrive[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + OneDrive[i].PinnedType + "</div><div hidden class='pinname'>" + OneDrive[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + OneDrive[i].DocumentLibraryURL.Description.trim() + "</div>";
                                onedriveHtml += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                onedriveHtml += "            <h4>" + OneDrive[i].DocumentLibrary + "</h4>";
                                onedriveHtml += "            <span class='path'>" + OneDrive[i].DocumentLibraryURL.Url.trim() + "</span>";
                                onedriveHtml += "        </div>";
                                onedriveHtml += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + OneDrive[i].ID + "</div>";
                                onedriveHtml += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                onedriveHtml += "        </a>";
                                onedriveHtml += "    </li>";
                            } else {
                                if (onedriveHtmlAll == "") {
                                    onedriveHtmlAll += onedriveHtml;
                                    if ($('#oneDriveitemAll').css('display') == 'none') {
                                        $('.pinshowmore').show();
                                    }
                                }
                                onedriveHtmlAll += "<li class='pinnedselected pinnedselect' data-siteurl ='" + decodeURI(OneDrive[i].SiteURL.Url) +
                                    "' data-locationname ='" + OneDrive[i].DocumentLibrary +
                                    "' data-locationurl ='" + decodeURI(OneDrive[i].DocumentLibraryURL.Url) +
                                    "' data-pintype ='" + OneDrive[i].PinnedType +
                                    "' >";
                                onedriveHtmlAll += "       <div class='pindoc'><div hidden class='siteurl'>" + OneDrive[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + OneDrive[i].PinnedType + "</div><div hidden class='pinname'>" + OneDrive[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + OneDrive[i].DocumentLibraryURL.Description.trim() + "</div>";
                                onedriveHtmlAll += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                onedriveHtmlAll += "            <h4>" + OneDrive[i].DocumentLibrary + "</h4>";
                                onedriveHtmlAll += "            <span class='path'>" + OneDrive[i].DocumentLibraryURL.Url.trim() + "</span>";
                                onedriveHtmlAll += "        </div>";
                                onedriveHtmlAll += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + OneDrive[i].ID + "</div>";
                                onedriveHtmlAll += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                onedriveHtmlAll += "        </a>";
                                onedriveHtmlAll += "    </li>";
                            }
                        }
                        onedriveHtml += "</ul>";
                        onedriveHtmlAll += "</ul>";
                        //if (OneDrive.length == 3) {
                        //$('.pinshowless').hide();
                        $('#oneDriveitemAll').css('display', 'none');
                        $('#oneDriveitem').css('display', 'block');
                        //}
                        $('#oneDriveitem').html(onedriveHtml);
                        $('#oneDriveitemAll').html(onedriveHtmlAll);

                        $("#createfile").attr('disabled', 'disabled');
                        $("#createfile").css('cursor', 'default');
                        $("#createfile").css('background-color', '');

                    } else {
                        onedriveHtml = "<p>    no pinned locations are found..!! </p>";
                        $('#oneDriveitem').html(onedriveHtml);
                        $('#oneDriveitemAll').html(onedriveHtmlAll);
                    }


                    //SharePoint sec pinned location data bind 


                    var sharepointHtml = "<ul class='pin_location'>";
                    var sharepointHtmlAll = "";
                    if (sharepoint.length > 0) {
                        for (var i = 0; i < sharepoint.length; i++) {
                            if (i < 3) {
                                sharepointHtml += "<li class='pinnedselected pinnedselect' data-siteurl ='" + decodeURI(sharepoint[i].SiteURL.Url) +
                                    "' data-locationname ='" + sharepoint[i].DocumentLibrary +
                                    "' data-locationurl ='" + decodeURI(sharepoint[i].DocumentLibraryURL.Url) +
                                    "' data-pintype ='" + sharepoint[i].PinnedType +
                                    "' >";
                                sharepointHtml += "       <div class='pindoc'><div hidden class='siteurl'>" + sharepoint[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + sharepoint[i].PinnedType + "</div><div hidden class='pinname'>" + sharepoint[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + sharepoint[i].DocumentLibraryURL.Description.trim() + "</div>";
                                sharepointHtml += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                sharepointHtml += "            <h4>" + sharepoint[i].DocumentLibrary + "</h4>";
                                sharepointHtml += "            <span class='path'>" + sharepoint[i].DocumentLibraryURL.Url.trim() + "</span>";
                                sharepointHtml += "        </div>";
                                sharepointHtml += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + sharepoint[i].ID + "</div>";
                                sharepointHtml += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                sharepointHtml += "        </a>";
                                sharepointHtml += "    </li>";
                            } else {
                                if (sharepointHtmlAll == "") {
                                    sharepointHtmlAll += sharepointHtml;
                                    if ($('#sharepointitemAll').css('display') == 'none') {
                                        $('.pinshowmore').show();
                                    }
                                }
                                sharepointHtmlAll += "<li class='pinnedselected pinnedselect' data-siteurl ='" + decodeURI(sharepoint[i].SiteURL.Url) +
                                    "' data-locationname ='" + sharepoint[i].DocumentLibrary +
                                    "' data-locationurl ='" + decodeURI(sharepoint[i].DocumentLibraryURL.Url) +
                                    "' data-pintype ='" + sharepoint[i].PinnedType +
                                    "' >";
                                sharepointHtmlAll += "       <div class='pindoc'><div hidden class='siteurl'>" + sharepoint[i].SiteURL.Description.trim() + "</div><div hidden class='type'>" + sharepoint[i].PinnedType + "</div><div hidden class='pinname'>" + sharepoint[i].DocumentLibrary + "</div><div hidden class='pinurl'>" + sharepoint[i].DocumentLibraryURL.Description.trim() + "</div>";
                                sharepointHtmlAll += "          <i class='ms-Icon ms-Icon--DocLibrary' aria-hidden='true'></i>";
                                sharepointHtmlAll += "            <h4>" + sharepoint[i].DocumentLibrary + "</h4>";
                                sharepointHtmlAll += "            <span class='path'>" + sharepoint[i].DocumentLibraryURL.Url.trim() + "</span>";
                                sharepointHtmlAll += "        </div>";
                                sharepointHtmlAll += "       <a href='#' class='pinicon removepinned' title='Unpin this location'><div hidden class='pinnedId'>" + sharepoint[i].ID + "</div>";
                                sharepointHtmlAll += "           <i class='ms-Icon ms-Icon--Pinned' aria-hidden='true'></i>";
                                sharepointHtmlAll += "        </a>";
                                sharepointHtmlAll += "    </li>";
                            }
                        }
                        sharepointHtml += "</ul>";
                        sharepointHtmlAll += "</ul>";
                        //if (sharepoint.length == 3) {
                        //$('.pinshowless').hide();
                        $('#sharepointitemAll').css('display', 'none');
                        $('#sharepointitem').css('display', 'block');
                        // }
                        $('#sharepointitem').html(sharepointHtml);
                        $('#sharepointitemAll').html(sharepointHtmlAll);

                        $("#createfile").attr('disabled', 'disabled');
                        $("#createfile").css('background-color', '');
                        $("#createfile").css('cursor', 'default');


                    } else {
                        sharepointHtml = "<p>    no pinned locations are found..!! </p>";
                        $('#sharepointitem').html(sharepointHtml);
                        $('#sharepointitemAll').html(sharepointHtmlAll);
                    }

                    $(".removepinned").click(function (evt) {
                        removeexistingpinned($(this), evt);
                    });

                    $(".pinnedselected").click(function () {
                        pinnedlocations_click($(this));
                    });

                    // dfd.resolve(resultArr);
                    //  closewaitdialog();

                })
                .fail(function (response) {
                    console.log(errordata);
                    //  dfd.reject();
                    // closewaitdialog();
                });
        }
    }

    //Modified this function on 13th july by Arijit for Tab section for pinned location. [End here]

    // Treeview --> remove Pinned Locations
    function removeexistingpinned(e, event) {
        event.stopPropagation();
        event.preventDefault();

        var pinnedId = $(e).find('.pinnedId').text().trim();
        var unpinCallback = function (pinnedId) {
            var dfd = $.Deferred();
            try {
                if (SPToken) {
                    var siteConfigListUrl = SPURL + "/_api/web/lists/getbytitle('" + sitePinnedLocations + "')/items(" + pinnedId + ")";
                    $.ajax({
                        url: siteConfigListUrl,
                        method: "POST",
                        headers: {
                            "Accept": "application/json; odata=verbose",
                            "content-type": "application/json;odata=verbose",
                            "X-RequestDigest": SPToken,
                            "IF-MATCH": "*",
                            'Authorization': 'Bearer ' + SPToken,
                            "X-HTTP-Method": "DELETE"
                        },
                        success: function (data) {
                            getSPPinnedLocations();
                            //show more --> show more div -->
                            dfd.resolve(data);

                            $("#dialog").dialog("close");
                        },
                        error: function (errordata) {
                            console.log("removeExistingPinned Ajax Call Error: " + errordata.responseText);
                            dfd.reject(data);
                        }
                    });
                }

            } catch (error) {
                console.log("removeExistingPinned: ", JSON.parse(error.responseText).error.message.Value);
                dfd.reject(data);
            }
            return dfd.promise();
        };

        custom_alert("Sure to Remove the Pin location?", "Remove Pin location", true, unpinCallback, pinnedId);


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
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        'Authorization': 'Bearer ' + SPToken
                    },
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
        } catch (error) {
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
            //  var proxyURL = "";
            var favoUrl = proxyURL + SPURL + "/_vti_bin/homeapi.ashx/sites/followed";
            callAjaxGet(favoUrl).done(function (data) {
                favdef.resolve(data.Items);
            });
        } catch (err) {
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
                tes = {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": token,
                    "IF-MATCH": "*",
                    'Authorization': 'Bearer ' + SPToken
                };
                $.ajax({
                    url: url,
                    type: "POST",
                    contentType: "application/json;odata=verbose",
                    data: JSON.stringify(listItem),
                    headers: tes,
                    success: function (data) {},
                    error: function (error) {
                        console.log(error);
                    }
                });
            });
        } catch (error) {
            console.log(error);
        }
    }

    //Treeview --> createDocument
    function createDocument(newFileName, destURl, chkDocTitle, tokenURl, destServerRelURL) {
        var dfd = $.Deferred();
        try {
            if (SAVE_LOCATION_TYPE.toUpperCase().includes("TEAMS")) {
                saveTemplateInTeams(chkDocTitle, newFileName).done(function (data) {
                    dfd.resolve(data);
                });;
            }
            if (SAVE_LOCATION_TYPE.toUpperCase().includes("ONEDRIVE")) {
                saveTemplateInOneDrive(chkDocTitle, newFileName).done(function (data) {
                    dfd.resolve(data);
                });;
            } else {
                copyFile(newFileName, destURl, chkDocTitle, tokenURl, destServerRelURL).done(function (data) {
                    dfd.resolve(data);
                });
            }
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
        } catch (error) {
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
            headers: {
                "Accept": "application/json; odata=verbose",
                'Authorization': 'Bearer ' + SPToken
            },
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

    // Dialog
    function custom_alert(message, title, isConfirm, callback, id) {
        if (!title)
            title = 'Alert';
        if (!message)
            message = 'No Message to Display.';
        if (!isConfirm) {
            $("#dialog").html(message).dialog({
                modal: true,
                title: title,
                resizable: false,
                // width: 300,
                // height: 150,
                open: function (event, ui) {
                    setTimeout(function () {
                        $("#dialog").dialog("close");
                    }, 4000);
                }
            });
        } else {
            $("#dialog").html(message).dialog({
                modal: true,
                title: title,
                resizable: false,
                // width: 300,
                // height: 150,
                buttons: {
                    "Confirm": function () {
                        callback(id);
                    },
                    "Cancel": function () {
                        $(this).dialog("close");
                    }
                }
            });
        }
        // $('<div></div>').html(message).dialog({
        //     title: title,
        //     resizable: false,
        //     modal: true,
        //     // buttons: {
        //     //     'Ok': function () {
        //     //         $(this).dialog('close');
        //     //     }
        //     // },
        //     open: function (event, ui) {
        //         setTimeout(function () {
        //             $("#dialog").dialog("close");
        //         }, 2000);
        //     }

        // })
    }

    //Treeview --> get List of Template From Source List
    function getListofTemplateFromSourceList(folder, isBreadCrumbClicked) {

        openWaitDialog();
        var OldFolderPath = CurrentDirectory;

        folder = !folder ? // Home 
            "" :
            isBreadCrumbClicked ?
            folder :
            CurrentDirectory + folder;

        CurrentDirectory = !folder ?
            !CurrentDirectory || isBreadCrumbClicked ?
            "" :
            CurrentDirectory + "/" :
            folder + "/";

        removeErrorMessage();

        if (TEMPLATE_FIRST_LOAD === "NOT_LOADED") {
            localStorage.removeItem('TemplateItemsArray');
        }

        var LocalStorageTemplateItemsArray = localStorage.getItem('TemplateItemsArray');
        var items = undefined;
        if (LocalStorageTemplateItemsArray && TEMPLATE_FIRST_LOAD === "LOADED") {

            var parsedTemplateItemsArray = JSON.parse(LocalStorageTemplateItemsArray);
            var FolderDetails = parsedTemplateItemsArray.find((d => d.Directory === CurrentDirectory));
            if (FolderDetails) {
                items = FolderDetails.Items;
            }
        }

        var docsTemplateList = "";

        if (SPToken && !items) {

            //// Added: 28th June ( Amartya ) : Start

            let payload = {
                SPOUrl: ORG_ROOT_WEB.webUrl,
                tenant: ORG_TENANT.id,
                FolderPath: CurrentDirectory,
                AccountName: USER_PROP.userPrincipalName,
                TenantName: ORG_ROOT_WEB.siteCollection.hostname.split(".")[0],
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
                contentType: "application/json",
                data: JSON.stringify(payload),
                // 
            }).done(function (response) {
                TEMPLATE_FIRST_LOAD = "LOADED";
                var TemplateItemsArray = [];
                if (LocalStorageTemplateItemsArray) {
                    TemplateItemsArray = JSON.parse(LocalStorageTemplateItemsArray);
                    TemplateItemsArray.push({
                        "Directory": CurrentDirectory,
                        "Items": response

                    })
                } else {
                    TemplateItemsArray.push({
                        "Directory": CurrentDirectory,
                        "Items": response

                    });

                }
                localStorage.setItem('TemplateItemsArray', JSON.stringify(TemplateItemsArray));

                if (!response.folderAccess) {
                    var Files = response.d.Files.results.filter(data => data.Name.lastIndexOf(".doc") > 1);
                    var Folders = [...response.d.Folders.results.filter(d => d.Name && d.Name.toUpperCase() !== "FORMS")];
                    var pptDocument = [];


                    for (var i = 0; i < Folders.length; i++) {
                        let CurrentFile = Folders[i];
                        CurrentFile.ContentType = {
                            Name: "Folder"
                        }
                        pptDocument.push(CurrentFile);

                    }
                    for (var i = 0; i < Files.length; i++) {
                        let CurrentFile = Files[i];
                        CurrentFile.ContentType = {
                            Name: "Document"
                        }
                        pptDocument.push(CurrentFile);

                    }

                    filterViewArrayList = pptDocument;
                    filteredData = pptDocument;

                    $("#previewbtn").attr('disabled', "disabled");
                    $("#createFile").attr('disabled', 'disabled');
                    $("#nextbtn").attr('disabled', 'disabled');
                    $("#Clearflt").css("display", "none");

                    $("#refreshList").unbind('click');
                    $('#refreshList').on("click", function () {
                        $("#noDataFoundLbl").hide();
                        if (currentTemplateView == "List") {
                            $('#listOfTemplate').css('display', 'block');
                            $('#DocTemplatesBoxView').css('display', 'none');
                        } else {
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
                        getDataFromFilter(filteredData);
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
                    if (platform == "OfficeOnline") {
                        getalltemp(pptDocument);
                        // getallClientTemp(pptDocument);
                    } else {
                        getallClientTemp(pptDocument);
                    }
                } else {
                    CurrentDirectory = OldFolderPath;
                    custom_alert("Access Denied!!", "Access Denied!!");
                }

            }).complete(function () {

                $('#DocTemplatesBoxView li, #listOfTemplate li').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            // Do something on click
                            let folderClicked = $(this).attr('documenttitle');
                            let contenttypename = $(this).attr('contenttypename');
                            if (contenttypename === "Folder") {
                                getListofTemplateFromSourceList(folderClicked);
                            }
                        },
                        mouseenter: function () {
                            // Do something on mouseenter
                        }
                    });
                    // item.off('click').on('click', function (event) {
                    // });

                });

                $('#DocTemplatesBoxView .breadcrumb-span, #listOfTemplate .breadcrumb-span').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            // Do something on click
                            let folderClicked = $(this).attr('path');
                            getListofTemplateFromSourceList(folderClicked, true);
                        },
                        mouseenter: function () {
                            // Do something on mouseenter
                        }
                    });
                    // item.off('click').on('click', function (event) {
                    // });

                });
                $('#DocTemplatesBoxView .box-default-location, #listOfTemplate .list-default-location').each(function (index, item) {
                    $(this).unbind('click');
                    if (DefaultLocation === CurrentDirectory) {
                        $(this).prop("checked", true);
                    }
                    $(this).bind({
                        click: function (event) {
                            $(this).prop('checked', event.target.checked);
                            setDefaultLocation(event.target.checked);
                        },
                    });
                });
                closeWaitDialog();

            }).fail(function (error) {
                console.error('error:- ' + error);
                docsTemplateList = "<div class='displayMessage'>No Template Library found in DocsNode Admin Panel</div>";
                docsTemplateList = "<div class='displayMessage'>" + erroeMeg + "\nNo Template Library found in DocsNode Admin Panel</div>";
                $('#listOfTemplate').html(docsTemplateList);
                $("#nextbtn").attr('disabled', 'disabled');
            });

            //// Added: 28th June (Amartya) : END

        } else {
            if (!items.folderAccess) {
                var Files = items.d.Files.results.filter(data => data.Name.lastIndexOf(".doc") > 1);
                var Folders = [...items.d.Folders.results.filter(d => d.Name && d.Name.toUpperCase() !== "FORMS")];
                var pptDocument = [];


                for (var i = 0; i < Folders.length; i++) {
                    let CurrentFile = Folders[i];
                    CurrentFile.ContentType = {
                        Name: "Folder"
                    }
                    pptDocument.push(CurrentFile);

                }
                for (var i = 0; i < Files.length; i++) {
                    let CurrentFile = Files[i];
                    CurrentFile.ContentType = {
                        Name: "Document"
                    }
                    pptDocument.push(CurrentFile);

                }

                filterViewArrayList = pptDocument;
                filteredData = pptDocument;

                $("#previewbtn").attr('disabled', "disabled");
                $("#createFile").attr('disabled', 'disabled');
                $("#nextbtn").attr('disabled', 'disabled');
                $("#Clearflt").css("display", "none");

                $("#refreshList").unbind('click');
                $('#refreshList').on("click", function () {
                    $("#noDataFoundLbl").hide();
                    if (currentTemplateView == "List") {
                        $('#listOfTemplate').css('display', 'block');
                        $('#DocTemplatesBoxView').css('display', 'none');
                    } else {
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
                    getDataFromFilter(filteredData);
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
                if (platform == "OfficeOnline") {
                    getalltemp(pptDocument);
                    // getallClientTemp(pptDocument);
                } else {
                    getallClientTemp(pptDocument);
                }

                $('#DocTemplatesBoxView li, #listOfTemplate li').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            // Do something on click
                            let folderClicked = $(this).attr('documenttitle');
                            let contenttypename = $(this).attr('contenttypename');
                            if (contenttypename === "Folder") {
                                getListofTemplateFromSourceList(folderClicked);
                            }
                        },
                        mouseenter: function () {
                            // Do something on mouseenter
                        }
                    });
                    // item.off('click').on('click', function (event) {
                    // });

                });

                $('#DocTemplatesBoxView .breadcrumb-span, #listOfTemplate .breadcrumb-span').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            // Do something on click
                            let folderClicked = $(this).attr('path');
                            getListofTemplateFromSourceList(folderClicked, true);
                        },
                        mouseenter: function () {
                            // Do something on mouseenter
                        }
                    });
                    // item.off('click').on('click', function (event) {
                    // });

                });

                $('#DocTemplatesBoxView .box-default-location, #listOfTemplate .list-default-location').each(function (index, item) {
                    $(this).unbind('click');
                    if (DefaultLocation === CurrentDirectory) {
                        $(this).prop("checked", true);
                    }
                    $(this).bind({
                        click: function (event) {
                            $(this).prop('checked', event.target.checked);
                            setDefaultLocation(event.target.checked);
                        },
                    });
                });
                if (platform == "OfficeOnline") {
                    closeWaitDialog();
                }
            } else {
                CurrentDirectory = OldFolderPath;
                custom_alert("Access Denied!!", "Access Denied!!");
                closeWaitDialog();
            }
        }


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
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        'Authorization': 'Bearer ' + SPToken
                    },
                    success: function (data) {
                        if (data.d.results[0] == null && data.d.results[0] == undefined) {
                            getValues(SPURL + "/").then(function (token) {
                                $.ajax({
                                    url: SPURL + "/_api/web/lists",
                                    type: "POST",
                                    data: JSON.stringify({
                                        '__metadata': {
                                            'type': 'SP.List'
                                        },
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
                                            '__metadata': {
                                                'type': 'SP.FieldText'
                                            },
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
                                                    '__metadata': {
                                                        'type': 'SP.FieldUrl'
                                                    },
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
                                                            '__metadata': {
                                                                'type': 'SP.FieldText'
                                                            },
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
                                                                    '__metadata': {
                                                                        'type': 'SP.FieldUrl'
                                                                    },
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
                        } else {
                            checkConfigurationLogoListExist();
                        }

                    },
                    error: function (errordata) {
                        console.log(errordata);
                    }
                });
            }
        } catch (error) {
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
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        'Authorization': 'Bearer ' + SPToken
                    },
                    success: function (data) {
                        if (data.d.results[0] == null && data.d.results[0] == undefined) {
                            getValues(SPURL + "/").then(function (token) {
                                $.ajax({
                                    url: SPURL + "/_api/web/lists",
                                    type: "POST",
                                    data: JSON.stringify({
                                        '__metadata': {
                                            'type': 'SP.List'
                                        },
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
                                            '__metadata': {
                                                'type': 'SP.FieldUrl'
                                            },
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
                                                    '__metadata': {
                                                        'type': 'SP.FieldNumber'
                                                    },
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
                                                    success: function (data) {},
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
                        } else {}
                    },
                    error: function (errordata) {
                        console.log("Check Configuration Logo List Error: ", errordata);
                    }
                });

            }
        } catch (error) {
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
                            if (response[i].Name.split('.').pop() == "docx") {
                                if (response[i].Name.match('%') || response[i].Name.match('"') ||
                                    response[i].Name.match('\'') || response[i].Name.match(';') || response[i].Name.match('#')) {
                                    //Do nothing
                                } else {
                                    pptDocument.push(response[i]);
                                }
                            }
                        }
                        if (pptDocument == 0) {
                            $("#btndropdown").attr('disabled', 'disabled');
                            $('#refreshList').attr('disabled', 'disabled');
                            $('#txtTemplateSearch').attr('disabled', 'disabled');
                        } else {
                            $("#btndropdown").removeAttr('disabled');
                            $('#refreshList').removeAttr('disabled');
                            $("#txtTemplateSearch").removeAttr('disabled');
                        }
                        filterViewArrayList = pptDocument;
                        filteredData = pptDocument;
                        gboxViewhtml = '';
                        if (platform == "OfficeOnline") {
                            getalltemp(pptDocument);
                            // getallClientTemp(pptDocument);
                        } else {
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
                } else {
                    $("#btndropdown").attr('disabled', 'disabled');
                    $('#refreshList').hide();
                    filteredData = [];
                    if (platform == "OfficeOnline") {
                        getalltemp(filteredData);
                        // getallClientTemp(filteredData);
                    } else {
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
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Authorization': 'Bearer ' + SPToken
                },
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

        } catch (error) {
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
                    headers: {
                        "Accept": "application/json; odata=verbose",
                        'Authorization': 'Bearer ' + SPToken
                    },
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
        } else if (DestinationWebRelativeUrl.indexOf('http') >= 0 || DestinationWebRelativeUrl.indexOf('.com') >= 0 || DestinationWebRelativeUrl.indexOf('.') >= 0) {
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
            } else {
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
                } else {
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
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Authorization': 'Bearer ' + SPToken
                },
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

                // Commented to Hide All Items view
                // if (data.length > 0) {
                //     for (var i = 0; i < data.length; i++) {
                //         listofviews += "<li id='" + data[i] + "'>" + data[i] + "</li>";
                //     }
                // }

                if (currentTemplateView == "Box") {
                    $('#DocTemplatesBoxView').show();
                    $('#listOfTemplate').hide();
                } else {
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
                        } else {
                            $("#grdvw").removeClass("selectView");
                            $("#lstvw").addClass("selectView");
                            $('#listOfTemplate').show();
                            $('#DocTemplatesBoxView').hide();
                        }
                    } else {
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
                headers: {
                    "Accept": "application/json; odata=verbose",
                    'Authorization': 'Bearer ' + SPToken
                },
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
        } catch (error) {
            console.log("getColumnFieldName: " + error);
        }
        return ddfieldName.promise();
    }

    function getFieldName(columnfieldname) {
        let items = "";
        try {
            items += "<li disabled class ='clearFields' id ='Clearflt'><i data-icon-name='List' class='ms-Icon ms-Icon--ClearFilter' role='presentation' aria-hidden='true'></i><span>Clear Filter</span></li>";
            for (let i = 0; i < columnfieldname.length; i++) {
                if (columnfieldname[i].indexOf("_x0020_") > -1) {
                    items += "<li class='filterFields' id = " + columnfieldname[i] + "><div class='link'><i id='fasi'  class='ms-Icon ms-Icon--ChevronDown' aria-hidden='true'></i>" + columnfieldname[i].replace(new RegExp('_x0020_', 'g'), ' ') + "</div><ul class='" + columnfieldname[i] + "' id='" + columnfieldname[i] + "'></ul></li>";
                } else if (columnfieldname[i] == "Editor") {
                    items += "<li class='filterFields' id = " + columnfieldname[i] + "><div class='link'><i id='fasi' class='ms-Icon ms-Icon--ChevronDown' aria-hidden='true'></i>" + columnfieldname[i].replace(new RegExp('Editor', 'g'), 'Modified By') + "</div><ul class='" + columnfieldname[i] + "' id='" + columnfieldname[i] + "'></ul></li>";
                } else if (columnfieldname[i] == "DocIcon") {
                    continue;
                } else if (columnfieldname[i] == "LinkFilename") {
                    continue;
                } else if (columnfieldname[i] == "ID") {
                    continue;
                } else {
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
                $("#filterUL").css({
                    'display': 'block'
                });
                var mousehover = false;
                $("#filterUL").mouseleave(function () {
                    mousehover = true;
                });
                $("#filterUL").mouseenter(function () {
                    mousehover = false;
                });
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
                // openWaitDialog();
                // ClearFilterandRebindList();
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
            if (String(value[internalName]) && options.map(function (record) {
                    return record.value;
                }).indexOf(value[internalName]) == -1) {
                if (internalName == "Editor") {
                    options.push({
                        "name": internalName,
                        "value": value[internalName].Title
                    });
                } else if (internalName == "Modified") {
                    var modifiedDate = _getFormattedDate(value[internalName]);
                    options.push({
                        "name": internalName,
                        "value": modifiedDate
                    });
                } else if (internalName == "ContentType") {
                    options.push({
                        "name": internalName,
                        "value": value[internalName].Name
                    });
                } else {
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
                } else if (options[i].value == false) {
                    options[i].value = "No";
                    var filter = internalName + ':' + false;
                    items += "<li class='valueItems' filter='" + filter + "' id='false'>" + options[i].value + "</li>";
                } else {
                    if (!isNaN(options[i].value)) {
                        var filter = internalName + ':' + options[i].value;
                    } else {
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
        } else {
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
            } else {
                arryOfColumnAndItem.push({
                    "key": internalName,
                    "value": itemselected
                });
                flagPresent = true;
            }
            if (flagPresent == false) {
                arryOfColumnAndItem.push({
                    "key": internalName,
                    "value": itemselected
                });
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
                    filterColumns.push({
                        "key": arr.key,
                        "value": tempValuesArr[i]
                    });
                }
            } else {
                filterColumns.push({
                    "key": arr.key,
                    "value": arr.value
                });
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
        arryOfColumnAndItem.map(function (item) {
            return item.key
        }).forEach(function (item) {
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
            filterCriteria.push({
                "key": results[key],
                "value": filterString
            });
            filterString = "";
        }

        Array.prototype.flexFilter = function (info) {
            // Set our variables
            var matchesFilter, matches = [],
                count;
            matchesFilter = function (item) {
                count = 0;
                for (var n = 0; n < info.length; n++) {
                    if (info[n]["key"] == "DocIcon") {
                        tempString = item.Name;
                        var fileNameLen = tempString.length;
                        var lstindex = tempString.lastIndexOf('.');
                        var fileExt = tempString.substr(lstindex + 1, fileNameLen);
                        fileExt = fileExt.toLowerCase();
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(fileExt) > -1) {
                            count++;
                        }
                    } else if (info[n]["key"] == "Editor") {
                        tempString = item.Editor.Title;
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(tempString) > -1) {
                            count++;
                        }
                    } else if (info[n]["key"] == "Modified") {
                        tempString = _getFormattedDate(item[info[n]["key"]]);
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(tempString) > -1) {
                            count++;
                        }
                    } else if (info[n]["key"] == "ContentType") {
                        tempString = item[info[n]["key"]].Name;
                        buildFilterCssArray(info[n]);
                        if (info[n]["value"].indexOf(tempString) > -1) {
                            count++;
                        }
                    } else {
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
                // getallClientTemp(filterByFieldValue);
            } else {
                getallClientTemp(filterByFieldValue);
            }
            filteredData = filterByFieldValue;
            $("#Clearflt").css("display", "block");
        } else {
            if (platform == "OfficeOnline") {
                getalltemp(filterViewArrayList);
                // getallClientTemp(filterViewArrayList);
            } else {
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

    function getalltempBoxView(data, Count) {
        try {
            var imageURLs = "";
            if (data.length > 0) {
                if (data[Count].LinkingUri != null) {
                    var fileURL = "";
                    var boxThumbnailURL = "";
                    fileURL = data[Count].LinkingUri.split('?d')[0];
                    boxThumbnailURL = data[Count].LinkingUri.split('DocsNodeAdmin')[0];
                    boxThumbnailURL += "DocsNodeAdmin/_layouts/15/getpreview.ashx?path=" + fileURL;
                    fileURL = fileURL.replace(SPURL, "");
                    var url = SPURL + templateServerRelURL + "_api/web/GetFileByServerRelativeUrl('" + fileURL + "')/OpenBinaryStream";
                    // if (xhr.status === 200) {
                    imageURLs = boxThumbnailURL;
                    // bindingWebAllTemp(imageURLs, data, Count);
                    // }
                    // else {
                    // imageURLs = "images/icons/Word-Iocn.png";
                    // }
                    // if (SPToken) {
                    //     var xhr = new window.XMLHttpRequest();
                    //     xhr.open("GET", url, true);
                    //     xhr.setRequestHeader("Accept", "application/json; odata=verbose");
                    //     xhr.setRequestHeader("Authorization", "Bearer " + SPToken);
                    //     //Now set response type
                    //     xhr.responseType = 'arraybuffer';
                    //     xhr.addEventListener('load', function () {
                    //     })
                    //     xhr.send();
                    // }
                } else {
                    imageURLs = "./../../images/icons/folder-open.jpg";
                }
                bindingWebAllTemp(imageURLs, data, Count);
            }
        } catch (error) {
            console.log(error);
        }
    }

    // function getallClientTemp(data) {
    //     var pptString = "";
    //     boxViewString = "";
    //     var Counter = 0;
    //     $('#DocTemplatesBoxView').html("");
    //     if (currentTemplateView == "List") {
    //         for (var i = 0; i < data.length; i++) {
    //             var fileNameLen = data[i].Name.length;
    //             var lstindex = data[i].Name.lastIndexOf('.');
    //             var fileExt = data[i].Name.substr(lstindex + 1, fileNameLen);
    //             fileExt = fileExt.toLowerCase();
    //             // if (data[i].Name.indexOf(docsNodeFilterExtention) != -1 || data[i].Name.indexOf(docsNodeFilterExtention.toUpperCase()) != -1) {
    //             pptString += "<li contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].Modified + "'";
    //             pptString += "documentTitle='" + (data[i].Name != null ? data[i].Name : "") + "' docId='" + data[i].ID + "' documentGUID='" + data[i].UniqueId + "'";
    //             // pptString += "modifiedName='" + data[i].Editor.Title + "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
    //             pptString += "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
    //             pptString += "ext='.docx'";
    //             // if (data[i].ContentType.Name === "Folder") {
    //             //     pptString += " onclick=alert('Clicked on Folder')";
    //             // }
    //             pptString += ">";
    //             if (data[i].ContentType.Name === "Document") {
    //                 pptString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].UniqueId + "></input>";
    //             }
    //             pptString += "<img src= 'images/" + docsNodeListTemplateLogo + "' class='width20'>";
    //             pptString += "<span> " + data[i].Name + "</span></li>";
    //             // }
    //         }
    //         if (data.length > 0) {
    //             $('#listOfTemplate').html(pptString);
    //             $('#DocTemplatesBoxView').css('display', 'none');
    //             $("#listOfTemplate").find('input').each(function () {
    //                 $("#listOfTemplate").find('input').on("change", handleChange);
    //             });
    //             closeWaitDialog();
    //             $("#btndropdown").removeAttr('disabled');
    //         } else {
    //             $("#btndropdown").attr('disabled', 'disabled');
    //             $("#listOfTemplate").html("<p style='color:Red'>No Records Found...!!</p>");
    //             closeWaitDialog();
    //         }
    //     } else {
    //         $('#DocTemplatesBoxView').css('display', 'block');
    //         if (gboxViewhtml != "") {
    //             if (data != 0) {
    //                 $("#btndropdown").removeAttr('disabled');
    //                 $('#DocTemplatesBoxView').html(gboxViewhtml);
    //                 $("#DocTemplatesBoxView").find('input').each(function () {
    //                     $("#DocTemplatesBoxView").find('input').on("change", handleChange);
    //                 });
    //                 closeWaitDialog();
    //             } else {
    //                 toDataURL(Counter, data);
    //             }
    //         } else {
    //             openWaitDialog();
    //             toDataURL(Counter, data);
    //         }
    //     }
    // }


    function getalltemp(data) {
        try {
            var pptString = "";
            var Count = 0;
            $('#DocTemplatesBoxView').html("");

            var htmlStrWhenHome = `<div class="breadcrumb-div">
                                    <span path="" class="breadcrumb-span" title="Home">Home</span> 
            </div> `;

            var htmlStrWhenFolder = `<div class="breadcrumb-div"> <span path="" class="breadcrumb-span">
                                    <i class="ms-Icon ms-Icon--HomeSolid" aria-hidden="true"></i>
                                    <i class="ms-Icon ms-Icon--ChevronRightMed" aria-hidden="true"></i>
                                </span>`;

            // 08th July Amartya for Breadcrumb
            if (!CurrentDirectory) {
                $('#DocTemplatesBoxView').html(htmlStrWhenHome);
            } else {
                var DirectoryArr = CurrentDirectory.split("/");

                DirectoryArr.map((dirName, index) => {

                    var path = DirectoryArr.slice(0, index + 1).join("/");

                    htmlStrWhenFolder = index < DirectoryArr.length - 2 ?
                        htmlStrWhenFolder.concat(
                            `<span path = "${path}" class="breadcrumb-span" title = ${dirName}>
                                    <i class="ms-Icon ms-Icon--FabricFolderFill" aria-hidden="true"></i>
                                    <i class="ms-Icon ms-Icon--ChevronRightMed" aria-hidden="true"></i>
                            </span> `) :
                        htmlStrWhenFolder.concat(`<span path = "${path}" class="breadcrumb-span"> ${dirName}</span> `);
                })
                htmlStrWhenFolder = htmlStrWhenFolder.concat(`</div> `);
                //$('#DocTemplatesBoxView').html(`${toggleClassBox} ${htmlStrWhenFolder} `);
                $('#DocTemplatesBoxView').html(` ${htmlStrWhenFolder} `);
            }
            for (var i = 0; i < data.length; i++) {
                var fileNameLen = data[i].Name.length;
                var lstindex = data[i].Name.lastIndexOf('.');
                var fileExt = data[i].Name.substr(lstindex + 1, fileNameLen);
                fileExt = fileExt.toLowerCase();
                // if (data[i].Name.indexOf(docsNodeFilterExtention) != -1 || data[i].Name.indexOf(docsNodeFilterExtention.toUpperCase()) != -1) {

                pptString += "<li contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].TimeLastModified + "'";
                pptString += "documentTitle='" + (data[i].Name != null ? data[i].Name : "") + "' docId='" + data[i].ListItemAllFields.ID + "' documentGUID='" + data[i].UniqueId + "'";
                if (data[i].ContentType.Name === "Document") {
                    pptString += "modifiedName='" + data[i].ListItemAllFields.FieldValuesAsText.Editor + "'";
                }
                pptString += "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
                pptString += "ext='.docx'";
                // if (data[i].ContentType.Name === "Folder") {
                //     pptString += " onclick=alert('Clicked on Folder')";
                // }
                pptString += ">";
                if (data[i].ContentType.Name === "Document") {
                    pptString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].UniqueId + "></input>";
                    pptString += "<i class='ms-Icon ms-Icon--WordLogo' title='WordLogo' aria-hidden='true'></i>";
                } else {
                    pptString += "<i class='ms-Icon ms-Icon--FabricFolderFill list-folder-icon' aria-hidden='true'></i>";

                }
                pptString += "<span> " + data[i].Name + "</span></li>";
                // }
            }
            if (data.length > 0) {
                if (!CurrentDirectory) {
                    $('#listOfTemplate').html(htmlStrWhenHome + pptString);
                } else {
                    // $('#listOfTemplate').html(toggleClassList + htmlStrWhenFolder + pptString);
                    $('#listOfTemplate').html(htmlStrWhenFolder + pptString);
                }
                if (currentTemplateView == "Box") {
                    $('#DocTemplatesBoxView').css('display', 'block');
                    getalltempBoxView(data, Count);
                } else {
                    $('#DocTemplatesBoxView').css('display', 'none');
                    $("#listOfTemplate").find('input').each(function () {
                        $("#listOfTemplate").find('input').on("change", handleChange);
                    });
                    closeWaitDialog();
                }

                $("#btndropdown").removeAttr('disabled');
            } else {
                // $("#listOfTemplate").html(`${toggleClassList} ${htmlStrWhenFolder} <p style='color:Red'>No Records Found...!!</p>`);
                // $("#DocTemplatesBoxView").html(`${toggleClassBox} ${htmlStrWhenFolder} <p style='color:Red'>No Records Found...!!</p>`);
                $("#listOfTemplate").html(`${htmlStrWhenFolder} <p style='color:Red'>No Records Found...!!</p>`);
                $("#DocTemplatesBoxView").html(`${htmlStrWhenFolder} <p style='color:Red'>No Records Found...!!</p>`);
                $("#btndropdown").attr('disabled', 'disabled');

                closeWaitDialog();
            }
        } catch (error) {
            console.log(error);
        }
    }

    function getallClientTemp(data) {
        var pptString = "";
        boxViewString = "";
        gboxViewhtml = "";
        var Counter = 0;
        $('#DocTemplatesBoxView').html("");
        var htmlStrWhenHome = `<div class="breadcrumb-div">
                                    <span path="" class="breadcrumb-span" title="Home">Home</span> 
            </div> `;

        var htmlStrWhenFolder = `<div class="breadcrumb-div"> <span path="" class="breadcrumb-span">
                                    <i class="ms-Icon ms-Icon--HomeSolid" aria-hidden="true"></i>
                                    <i class="ms-Icon ms-Icon--ChevronRightMed" aria-hidden="true"></i>
                                </span>`;
        if (CurrentDirectory) {
            var DirectoryArr = CurrentDirectory.split("/");

            DirectoryArr.map((dirName, index) => {

                var path = DirectoryArr.slice(0, index + 1).join("/");

                htmlStrWhenFolder = index < DirectoryArr.length - 2 ?
                    htmlStrWhenFolder.concat(
                        `<span path = "${path}" class="breadcrumb-span" title = ${dirName}>
                                                        <i class="ms-Icon ms-Icon--FabricFolderFill" aria-hidden="true"></i>
                                                        <i class="ms-Icon ms-Icon--ChevronRightMed" aria-hidden="true"></i>
                                                </span> `) :
                    htmlStrWhenFolder.concat(`<span path = "${path}" class="breadcrumb-span"> ${dirName}</span> `);
            })
            htmlStrWhenFolder = htmlStrWhenFolder.concat(`</div> `);
            //$('#DocTemplatesBoxView').html(`${toggleClassBox} ${htmlStrWhenFolder} `);
            // $('#DocTemplatesBoxView').html(` ${htmlStrWhenFolder} `);
        }

        if (!CurrentDirectory) {
            $('#DocTemplatesBoxView').html(`${htmlStrWhenHome}`);
        } else {
            $('#DocTemplatesBoxView').html(`${htmlStrWhenFolder}`);
        }
        if (currentTemplateView == "List") {
            for (var i = 0; i < data.length; i++) {
                var fileNameLen = data[i].Name.length;
                var lstindex = data[i].Name.lastIndexOf('.');
                var fileExt = data[i].Name.substr(lstindex + 1, fileNameLen);
                fileExt = fileExt.toLowerCase();
                pptString += "<li contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].TimeLastModified + "'";
                pptString += "documentTitle='" + (data[i].Name != null ? data[i].Name : "") + "' docId='" + data[i].ListItemAllFields.ID + "' documentGUID='" + data[i].UniqueId + "'";

                if (data[i].ContentType.Name === "Document") {
                    pptString += "modifiedName='" + data[i].ListItemAllFields.FieldValuesAsText.Editor + "'";
                }

                pptString += "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
                pptString += "ext='.docx'";
                pptString += ">";
                if (data[i].ContentType.Name === "Document") {
                    pptString += "<input id='templateDocs' type='checkbox' class='checkbox' value =" + data[i].UniqueId + "></input>";
                    pptString += "<i class='ms-Icon ms-Icon--WordLogo' title='WordLogo' aria-hidden='true'></i>";
                } else {
                    pptString += "<i class='ms-Icon ms-Icon--FabricFolderFill list-folder-icon' aria-hidden='true'></i>";
                }
                pptString += "<span> " + data[i].Name + "</span></li>";
                // }
            }
            if (data.length > 0) {
                if (!CurrentDirectory) {
                    $('#listOfTemplate').html(htmlStrWhenHome + pptString);
                } else {
                    $('#listOfTemplate').html(htmlStrWhenFolder + pptString);
                }
                $('#DocTemplatesBoxView').css('display', 'none');
                $("#listOfTemplate").find('input').each(function () {
                    $("#listOfTemplate").find('input').on("change", handleChange);
                });
                closeWaitDialog();
                $("#btndropdown").removeAttr('disabled');
            } else {
                $("#btndropdown").attr('disabled', 'disabled');
                $("#listOfTemplate").html(`${htmlStrWhenFolder} <p style='color:Red'>No Records Found...!!</p>`);
                closeWaitDialog();
            }
        } else {
            $('#DocTemplatesBoxView').css('display', 'block');
            if (gboxViewhtml != "") {
                if (data != 0) {
                    $("#btndropdown").removeAttr('disabled');

                    $('#DocTemplatesBoxView').append(`${gboxViewhtml}`);

                    $("#DocTemplatesBoxView").find('input').each(function () {
                        $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                    });
                    closeWaitDialog();
                } else {
                    toDataURL(Counter, data);
                }
            } else {
                openWaitDialog();
                toDataURL(Counter, data);
            }
        }
    }

    function getImageCORSCall(boxThumbnailURL, cnt) {

        var proxyURL = "https://cors-anywhere.herokuapp.com/";
        var dfdReq = $.Deferred();
        try {

            $.ajax({
                url: proxyURL + boxThumbnailURL,
                method: "GET",
                cache: true,
                xhrFields: {
                    responseType: 'blob'
                },
                headers: {
                    'Authorization': 'Bearer ' + SPToken,
                    'cache-control': 'max-age=3600',
                },
                success: function (data) {
                    var reader = new FileReader();
                    reader.onloadend = function () {
                        var imagewithCnt = {
                            imageData: reader.result,
                            cnt: cnt
                        }
                        dfdReq.resolve(imagewithCnt);
                    }
                    reader.readAsDataURL(data);
                },
                error: function (err) {
                    console.log("getValues: " + err);
                    dfdReq.reject(err);
                }
            });

        } catch (error) {
            console.log("getValues: " + error);
        }
        return dfdReq.promise();
    }

    function toDataURL(Cnt, data) {

        if (data.length > 0) {
            data.map(function (item) {

                if (item.ContentType.Name === "Folder") {
                    var imgdata = "./../../images/icons/folder-open.jpg";
                    binding(imgdata, data, Cnt);
                    Cnt += 1;
                }
                if (item.ContentType.Name === "Document") {
                    var fileURL = item.LinkingUri.split('?d')[0];
                    var boxThumbnailURL = item.LinkingUri.split('DocsNodeAdmin')[0];
                    boxThumbnailURL += "DocsNodeAdmin/_layouts/15/getpreview.ashx?path=" + fileURL;

                    getImageCORSCall(boxThumbnailURL, Cnt).done((respItem) => {
                        binding(respItem.imageData, data, respItem.cnt);

                    });
                    Cnt += 1;
                }


            });
        } else {
            gboxViewhtml = "";
            $("#btndropdown").attr('disabled', 'disabled');
            $("#DocTemplatesBoxView").append("<p style='color:Red'>No Records Found...!!</p>");
            closeWaitDialog();
        }

        //  var boxThumbnailURL = "";
        // var proxyURL = "https://cors-anywhere.herokuapp.com/";

        ////  var proxyURL = "";
        //  if (data.length > 0) {

        //      if (data[Cnt].LinkingUri != null) {
        //          var fileURL = data[Cnt].LinkingUri.split('?d')[0];
        //          boxThumbnailURL = data[Cnt].LinkingUri.split('DocsNodeAdmin')[0];
        //          boxThumbnailURL += "DocsNodeAdmin/_layouts/15/getpreview.ashx?path=" + fileURL;
        //      }
        //      var xhr = new window.XMLHttpRequest();
        //      xhr.open('GET', proxyURL + boxThumbnailURL, true);
        //      xhr.setRequestHeader('Authorization', 'Bearer ' + SPToken);
        //      xhr.setRequestHeader('cache-control', 'max-age=3600');
        //      xhr.responseType = 'blob';
        //      xhr.onload = function (event) {
        //          if (event.srcElement.status == 200) {
        //              filereader(event).done(function (imgdata) {
        //                  if (data[Cnt].ContentType.Name === "Folder") {
        //                      imgdata = "./../../images/icons/folder-open.jpg";
        //                  }
        //                  binding(imgdata, data, Cnt);
        //              });
        //          } else {
        //              var imageURLs = "images/icons/Word-Iocn.png";
        //              binding(imageURLs, data, Cnt);
        //          }

        //          function filereader(event) {
        //              var dfdImg = $.Deferred();
        //              var reader = new FileReader();
        //              reader.onloadend = function () {
        //                  dfdImg.resolve(reader.result);
        //              }
        //              reader.readAsDataURL(xhr.response);
        //              return dfdImg.promise();
        //          }
        //      };
        //      xhr.send();
        //  } else {
        //      gboxViewhtml = "";
        //      $("#btndropdown").attr('disabled', 'disabled');
        //      $("#DocTemplatesBoxView").append("<p style='color:Red'>No Records Found...!!</p>");
        //      closeWaitDialog();
        //  }
    }


    function bindingWebAllTemp(imageURLs, data, i) {
        var fileNameLen = data[i].Name.length;
        var lstindex = data[i].Name.lastIndexOf('.');
        var fileExt = data[i].Name.substr(lstindex + 1, fileNameLen);
        fileExt = fileExt.toLowerCase();
        if (data[i].Name.indexOf(docsNodeFilterExtention) != -1 ||
            data[i].Name.indexOf(docsNodeFilterExtention.toUpperCase()) != -1 || data[i].ContentType.Name === "Folder") {
            boxViewString += "<li class='checkboxToggler' thumbnail='' contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].TimeLastModified + "'";
            boxViewString += "documentTitle='" + (data[i].Name != null ? data[i].Name : "") + "' docId='" + data[i].ListItemAllFields.ID + "' documentGUID='" + data[i].UniqueId + "'";
            if (data[i].ContentType.Name === "Document") {
                boxViewString += "modifiedName='" + data[i].ListItemAllFields.FieldValuesAsText.Editor + "'";
            }
            boxViewString += "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
            boxViewString += "ext='.docx'";

            if (data[i].ContentType.Name === "Folder") {
                boxViewString += `foldername = '${data[i].Name}'`;
            }
            boxViewString += ">";
            if (data[i].ContentType.Name === "Document") {
                //boxViewString += "<input id='templateDocs' type='checkbox' class='checkbox' value ='" + data[i].UniqueId + "'></input><div class='box-img'>";
                boxViewString += "<input id='templateDocs" + data[i].UniqueId + "' class='templateDocs'  type='checkbox' class='checkbox' value ='" + data[i].UniqueId + "'></input><label class='clearfix' for='templateDocs" + data[i].UniqueId + "'><span class='styleCheck'><i class='ms-Icon ms-Icon--CompletedSolid' title=''></i ></span><div class='box-img' >";
            }
            boxViewString += "<div>";
            boxViewString += "<img src='" + imageURLs + "'class='docimgouterbox' /></div></div>";
            boxViewString += "<span style='word-break: break-all' title='" + data[i].Name + "'> " + (data[i].Name.length > 25 ? data[i].Name.substring(0, 25) + "..." : data[i].Name) + "</span></li>";
            $('#DocTemplatesBoxView').append(boxViewString);
            boxViewString = "";
            // closeWaitDialog();
        }
        $("#DocTemplatesBoxView").find('input').each(function () {
            $("#DocTemplatesBoxView").find('input').on("change", handleChange);
        });
        if (i < data.length - 1) {
            getalltempBoxView(data, ++i);
        }
    }

    function binding(imgdata, data, i) {
        var fileNameLen = data[i].Name.length;
        var lstindex = data[i].Name.lastIndexOf('.');
        var fileExt = data[i].Name.substr(lstindex + 1, fileNameLen);
        fileExt = fileExt.toLowerCase();

        if (data[i].Name.indexOf(docsNodeFilterExtention) != -1 ||
            data[i].Name.indexOf(docsNodeFilterExtention.toUpperCase()) != -1 || data[i].ContentType.Name === "Folder") {
            boxViewString += "<li class='checkboxToggler' thumbnail='' contentTypeName='" + data[i].ContentType.Name + "' modifiedDate='" + data[i].TimeLastModified + "'";
            boxViewString += "documentTitle='" + (data[i].Name != null ? data[i].Name : "") + "' docId='" + data[i].ListItemAllFields.ID + "' documentGUID='" + data[i].UniqueId + "'";
            // boxViewString += "modifiedName='" + data[i].Editor.Title + "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
            boxViewString += "' serverRelativeURL='" + data[i].ServerRelativeUrl + "'";
            boxViewString += "ext='.docx'";
            if (data[i].ContentType.Name === "Folder") {
                boxViewString += `foldername = '${data[i].Name}'`;
            }
            boxViewString += ">";
            if (data[i].ContentType.Name === "Document") {
                boxViewString += "<input id='templateDocs" + data[i].UniqueId + "' class='templateDocs'  type='checkbox' class='checkbox' value ='" + data[i].UniqueId + "'></input><label class='clearfix' for='templateDocs" + data[i].UniqueId + "'><span class='styleCheck'><i class='ms-Icon ms-Icon--CompletedSolid' title=''></i ></span><div class='box-img' >";
            }
            boxViewString += "<div>";
            boxViewString += "<img src='" + imgdata + "'class='docimgouterbox' /></div></div>";
            boxViewString += "<span style='word-break: break-all' title='" + data[i].Name + "'> " + (data[i].Name.length > 25 ? data[i].Name.substring(0, 25) + "..." : data[i].Name) + "</span></li>";
            gboxViewhtml += boxViewString;
        }
        $('#DocTemplatesBoxView').append(boxViewString);
        boxViewString = '';
        // closeWaitDialog();
        //i++;
        //if (i < data.length) {
        // toDataURL(i, data);
        //}
        if (i >= data.length - 1) {
            if (data.length > 0) {
                //gboxViewhtml = boxViewString;
                //$('#DocTemplatesBoxView').html(boxViewString);

                $("#DocTemplatesBoxView").find('input').each(function () {
                    $("#DocTemplatesBoxView").find('input').off('change');
                    $("#DocTemplatesBoxView").find('input').on("change", handleChange);
                });
                $('#DocTemplatesBoxView li, #listOfTemplate li').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            // Do something on click
                            let folderClicked = $(this).attr('documenttitle');
                            let contenttypename = $(this).attr('contenttypename');
                            if (contenttypename === "Folder") {
                                getListofTemplateFromSourceList(folderClicked);
                            }
                        },
                    });
                });

                $('#DocTemplatesBoxView .breadcrumb-span, #listOfTemplate .breadcrumb-span').each(function (index, item) {
                    $(this).unbind('click');
                    $(this).bind({
                        click: function () {
                            // Do something on click
                            let folderClicked = $(this).attr('path');
                            getListofTemplateFromSourceList(folderClicked, true);
                        },
                    });

                });

                $('#DocTemplatesBoxView .box-default-location, #listOfTemplate .list-default-location').each(function (index, item) {
                    $(this).unbind('click');
                    if (DefaultLocation === CurrentDirectory) {
                        $(this).prop("checked", true);
                    }
                    $(this).bind({
                        click: function (event) {
                            $(this).prop('checked', event.target.checked);
                            setDefaultLocation(event.target.checked);
                        },
                    });
                });
                closeWaitDialog();
                $("#btndropdown").removeAttr('disabled');
            } else {
                $("#btndropdown").attr('disabled', 'disabled');
                $("#DocTemplatesBoxView").append("<p style='color:Red'>No Records Found...!!</p>");
                closeWaitDialog();
            }

            //if (i >= data.length - 1) {
            //    closeWaitDialog();
            //}
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
                var blob = new Blob([sampleBytes], {
                    type: "image/jpeg"
                });
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
            //var checkedDocument = $('#templateDocs:checked').val();
            var checkedDocument = $('.templateDocs:checked').val();
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
                docId = $("#DocTemplatesBoxView").find("input:checked").parents("li.checkboxToggler").attr("docId");
                docContentTypeName = $("#DocTemplatesBoxView").find("input:checked").parents("li.checkboxToggler").attr("contentTypeName");
                docModifiedDate = _getFormattedDate($("#DocTemplatesBoxView").find("input:checked").parents("li.checkboxToggler").attr("modifiedDate"));
                docTitle = $("#DocTemplatesBoxView").find("input:checked").parents("li.checkboxToggler").attr("documentTitle");
                docModifiedName = $("#DocTemplatesBoxView").find("input:checked").parents("li.checkboxToggler").attr("modifiedName");
                imgSRC = $("#DocTemplatesBoxView").find("input:checked").next().find("img").attr("src");
            } else {
                docId = $("#listOfTemplate").find("input:checked").parents("li").attr("docId");
                docContentTypeName = $("#listOfTemplate").find("input:checked").parents("li").attr("contentTypeName");
                docModifiedDate = _getFormattedDate($("#listOfTemplate").find("input:checked").parents("li").attr("modifiedDate"));
                docTitle = $("#listOfTemplate").find("input:checked").parents("li").attr("documentTitle");
                docModifiedName = $("#listOfTemplate").find("input:checked").parents("li").attr("modifiedName");
                previewDestUrl = $("#listOfTemplate").find("input:checked").parents("li").attr("serverRelativeURL");
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
                            if (cFields[j].Group !== "_Hidden" && cFields[j].Hidden !== true && cFields[j].TypeDisplayName !== "File" &&
                                cFields[j].StaticName !== "Modified_x0020_By" && cFields[j].StaticName !== "Created_x0020_By" &&
                                cFields[j].StaticName !== '_dlc_DocId' && cFields[j].StaticName !== '_dlc_DocIdUrl' && cFields[j].StaticName !== '_dlc_DocIdPersistId') {
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
                    } else if (fieldsArray[i].TypeDisplayName === "Person or Group") {
                        pplColumnArray.push(fieldsArray[i].InternalName);
                        expandColumn += fieldsArray[i].InternalName + ',';
                        selectColumns += fieldsArray[i].InternalName + "," + fieldsArray[i].InternalName + "/Id," + fieldsArray[i].InternalName + "/Title,";
                    } else if (fieldsArray[i].TypeDisplayName === "Managed Metadata") {
                        TaxColumnArray.push(fieldsArray[i].InternalName);
                        expandColumn += 'TaxCatchAll,';
                        selectColumns += fieldsArray[i].InternalName + ",TaxCatchAll/ID,TaxCatchAll/Term,";
                    } else {
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
                        } else if (HyperLinkColumnArray.indexOf(value) > -1) {
                            for (var i = 0; i < HyperLinkColumnArray.length; i++) {
                                newli += "<li><label> <b>" + value + ": </b><span>" + data.d[HyperLinkColumnArray[i]]["Url"] + "</span></label></li>";
                            }
                        } else if (pplColumnArray.indexOf(value) > -1) {
                            newli += "<li><label> <b>" + value + ": </b><span>" + (data.d[value]["Title"] != null ? data.d[value]["Title"] : "") + "</span></label></li>";
                        } else if (lkpColumnArray.indexOf(value) > -1) {
                            newli += "<li><label> <b>" + value + ": </b><span>" + (data.d[value]["Title"] != null ? data.d[value]["Title"] : "") + "</span></label></li>";
                        } else if (currColumnArray.indexOf(value) > -1) {
                            newli += "<li><label> <b>" + value + ": </b><span> $" + data.d[value] + "</span></label></li>";
                        } else if (TaxColumnArray.indexOf(value) > -1) {
                            var term = getTaxonomyValue(data, value);
                            newli += "<li><label> <b>" + value + ": </b><span>" + term + "</span></label></li>";
                        } else {
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
            //  var proxyURL = "";
            var previewURl = SPURL + templateServerRelURL + "_layouts/15/getpreview.ashx?path=" + SPURL +
                fileServerRelativeUrl;
            if (platform == "OfficeOnline") {
                dfdDocPreview.resolve(previewURl);
            } else {
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
        xhr.setRequestHeader('cache-control', 'max-age=3600');
        xhr.responseType = 'blob';
        xhr.onload = function (event) {
            if (xhr.status === 200) {
                filereader(event).done(function (imgdata) {
                    dfdImgdef.resolve(imgdata);
                });
            } else {
                dfdImgdef.resolve("images/icons/Word-Iocn.png");
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
        setTimeout(function () {
            $searchResultsDiv.css("display", "none");
        }, 5000);
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





    /////////////// get All Team Location - Render///////////////////// Created on 14.07.2020 by Arijit 
    //function getTeams_Treeview_Render() {
    //    var favData = [];
    //    var treeViewHTMLFav = "";
    //    var hasNoFav = true;

    //    getAllTeams_Treeview().then(function (sdata) {
    //        if (listOfSiteCollectionsArray != null && listOfSiteCollectionsArray.length > 0) {

    //            var sitesCollectionArray = listOfSiteCollectionsArray.filter(function (objFav) {
    //                return objFav.siteUrl
    //            });

    //            getMyFollowedSites().done(function (favResults) {
    //                for (var i = 0; i < sitesCollectionArray.length; i++) {
    //                    for (var j = 0; j < favResults.length; j++) {
    //                        if (sitesCollectionArray[i].siteUrl === favResults[j].Url) {
    //                            favData.push(favResults[j]);
    //                        }
    //                    }
    //                }
    //                treeViewHTMLFav = "<ul id='treeviewUL'>";
    //                for (var i = 0; i < favData.length; i++) {
    //                    hasNoFav = false;
    //                    var siteKey = favData[i].Title + "_" + i;

    //                    treeViewHTMLFav += " <li id='" + favData[i].Title + "'><span class='caretCustom caret-down treeSpan' ><div class='type' hidden>sitecollection</div>"
    //                    treeViewHTMLFav += " <div class='level' hidden> " + 0 + "</div > <div class='sitekey' hidden> " + siteKey + "</div> <div class='siteurl' hidden> " + favData[i].Url + "</div> "
    //                    treeViewHTMLFav += " <div class='sitetitle' hidden> " + favData[i].Title + "</div > <div class='siteId' hidden> " + favData[i].ItemReference.SiteId + "</div>"
    //                    treeViewHTMLFav += " <a href= '#' > <i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i><span class='singleLineEllipse'>" + favData[i].Title + "</span></a ></span > "; // main site collection li level -1 open
    //                    treeViewHTMLFav += " <ul class='active nested' id='" + siteKey + "'>"; // ul level - 2 open
    //                    treeViewHTMLFav += " </ul>"; // Site shared documents
    //                    treeViewHTMLFav += " </li>"; // Site  
    //                }
    //                treeViewHTMLFav += "</ul>"; //treeviewUL

    //                if (hasNoFav) {
    //                    treeViewHTMLFav = "<p> No favorite sites are found..!! </p>";
    //                }

    //                //  $('#SPFavTreeView').html(treeViewHTMLFav); 
    //                $('#sharepointLoc').html(treeViewHTMLFav);
    //                $(".treeSpan").off('click');
    //                $(".treeSpan").on('click', function () {
    //                    getAllSites_Treeview_click(this);
    //                });

    //                closeWaitDialog();
    //            }).fail(function (jqXHR, textStatus) {
    //                closeWaitDialog();
    //            });
    //        }
    //    });
    //}


    // Get All team location for tree view 15.07.2020- Arijit
    function getAllTeams_Treeview() {
        openWaitDialog();
        /// var dfd = $.Deferred();
        try {

            if (SPToken) {
                let payload = {
                    SPOUrl: "https://graph.microsoft.com",
                    UsrGUID: USER_PROP.id,
                    tenant: ORG_TENANT.id,
                }

                $.ajax({
                    url: GET_USER_MSTEAMS,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    //  var teamResultArr = response.d.results;
                    console.log(response.value);

                    var resultsCount = response.value;
                    var firstRow = "<ul id='teamviewUl'>";
                    for (var i = 0; i < resultsCount.length; i++) {
                        //var teamName = resultsCount[i].displayName;
                        //var teamID = resultsCount[i].id;
                        //listOfTeam.push({ "teamName": teamName, "teamID": teamID });
                        var team = {};
                        team.TeamName = resultsCount[i].displayName;
                        team.TeamId = resultsCount[i].id;
                        team.TeamUrl = resultsCount[i]["@odata.editLink"];
                        listOfTeam.push(team);

                        //dfd.resolve(response.value);
                        firstRow += " <li class='parentLi'><span class='caretCustom treeSpan' >"
                        firstRow += " <a href= '#' class='firstLavel' data-teamName='" + resultsCount[i].displayName + "' data-expLevel='0' data-teamId='" + resultsCount[i].id + "'> <i class='ms-Icon ms-Icon--TeamsLogoInverse' aria-hidden='true'></i><span class='singleLineEllipse'>" + resultsCount[i].displayName + "</span></a ></span > ";


                    }
                    firstRow += "</ul>"; //treeviewUL

                    setLocalForageItem("Teams", JSON.stringify(listOfTeam)).done(function (values) {
                        console.log("Team Data : " + values);

                        $('.location-items>ul>li.ms-Pivot-link').off('click');
                        $('.location-items>ul>li.ms-Pivot-link').on('click', function () {
                            setOnedriveRootasDestination($(this));
                        });
                    })

                    $('#teamLoc').html(firstRow);

                    $('#teamviewUl>li>span.treeSpan').on("click", function () {
                        onExpandTeam($(this).find('.firstLavel'));
                    });
                    //dfd.resolve(listOfTeam);
                    closeWaitDialog();
                });
            }
        } catch (err) {
            console.log("getTeam_Treeview: " + err);
            dfd.reject('Error');
            closeWaitDialog();
        }
        //return dfd.promise();
    };


    function getTeamWebURL(teamID) {
        openWaitDialog();
        var dfd = $.Deferred();
        try {

            if (SPToken) {
                let payload = {
                    GroupId: teamID,
                    tenant: ORG_TENANT.id,
                }

                $.ajax({
                    url: GET_TEAMS_WEBURL,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    //var teamURLResult = response.d.results;
                    console.log(response.value)
                    dfd.resolve(response.value);

                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };



    function getLibInternalName(teamweburl) {
        openWaitDialog();
        var dfd = $.Deferred();
        try {

            if (SPToken) {
                let payload = {
                    SPOUrl: ORG_ROOT_WEB.webUrl,
                    Team: teamweburl,
                    tenant: ORG_TENANT.id,
                }

                $.ajax({
                    url: GET_TEAM_LIBRARY_NAMES,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {

                    dfd.resolve(response);
                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };



    function getChannel(teamID, teamurl, libinternalname) {
        openWaitDialog();
        var dfd = $.Deferred();
        try {

            if (SPToken) {
                let payload = {
                    SPOUrl: ORG_ROOT_WEB.webUrl,
                    tenant: ORG_TENANT.id,
                    TeamID: teamID,
                    TeamURL: teamurl,
                    LibInternalName: libinternalname,
                }

                $.ajax({
                    url: GET_MSTEAM_CHANNELS,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    console.log('getChannel success')
                    dfd.resolve(response);

                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };



    function getChannelFolder(payload) {
        //  openWaitDialog();
        var dfd = $.Deferred();
        try {

            if (SPToken) {

                $.ajax({
                    url: GET_CHANNEL_FOLDER,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    console.log('folder', response);
                    dfd.resolve(response);

                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };



    function getTab(payload) {
        // openWaitDialog();
        var dfd = $.Deferred();
        try {

            if (SPToken) {


                $.ajax({
                    url: GET_TEAM_CHANNEL_TAB_URL,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    console.log(response);

                    dfd.resolve(response);

                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };


    function onExpandTeam(e) {

        if ($(e).parents('span.caretCustom').siblings("ul.nested").length > 0) {
            $(e).parents('span.caretCustom').siblings("ul.nested").removeClass("nested");
            $(e).parents('span.caretCustom').addClass("caret-down");
            return;
        } else if ($(e).parents('span.caretCustom').siblings("ul:eq(0)").length === 1) {


            $(e).parents('span.caretCustom').siblings("ul:eq(0)").addClass("nested");
            $(e).parents('span.caret-down').removeClass("caret-down");
            return;
        }



        $("#createFile").attr('disabled', 'disabled');
        $("#createFile").css('background-color', '');
        $("#createFile").css('cursor', 'default');


        var id = $(e).data('teamid');
        var container = $(e).parent().parent('li');
        getTeamWebURL(id).then(function (data) {

            var teamSPOUrl = data;

            getLibInternalName(teamSPOUrl).then(function (results) {

                var eachteam = {};
                results.d.results.map(libNames => {

                    // eachTemplate.isSelected = false;
                    if (libNames.Title && (libNames.Title.toLowerCase() === "documents" || libNames.Title.toLowerCase() === "dokumenter")) {
                        eachteam.EntityTypeName = libNames.EntityTypeName.toString().replace("_x0020_", " ");

                    }


                });

                getLocalForageItem("Teams").done(function (values) {

                    listOfTeam = JSON.parse(values);

                    listOfTeam.map(function (item, inx) {
                        if (item.TeamId === id) {
                            item.TeamSPOUrl = teamSPOUrl;
                            item.TeamShareFolderName = eachteam.EntityTypeName;
                        }

                    });

                    setLocalForageItem("Teams", JSON.stringify(listOfTeam)).done(function (values) {
                        console.log("Team Data : " + values);
                    });

                });

                //console.log('data Received', data)
                getChannel(id, teamSPOUrl, eachteam.EntityTypeName).then(function (channels) {
                    let allMsTeamChannels = [];

                    channels.map(channel => {
                        const eachTeamChannel = {};
                        if (channel) {
                            eachTeamChannel.ChannelDisplayName = channel.displayName;
                            eachTeamChannel.ChannelUrl = channel.webUrl;
                            eachTeamChannel.ChannelId = channel.id;
                            eachTeamChannel.FolderType = channel["@odata.type"];

                            if (channel.membershipType && channel.membershipType === "private") {
                                eachTeamChannel.FolderType = eachTeamChannel.FolderType + ".private";
                            }

                        }

                        allMsTeamChannels.push(eachTeamChannel);
                    });

                    getLocalForageItem("Teams").done(function (values) {

                        listOfTeam = JSON.parse(values);

                        listOfTeam.map(function (item, inx) {
                            if (item.TeamId === id) {
                                item.TeamChannels = allMsTeamChannels;

                            }

                        });

                        setLocalForageItem("Teams", JSON.stringify(listOfTeam)).done(function (values) {
                            console.log("Team Data : " + values);
                        });

                    });

                    var channelHtml = "<ul id='teamFolUl' class='treefolderUl'>";
                    allMsTeamChannels.map((item, index) => {
                        let channelInnerName = item.ChannelUrl && item.ChannelUrl.substring(item.ChannelUrl.lastIndexOf("/") + 1, item.ChannelUrl.lastIndexOf("?"));
                        channelInnerName = channelInnerName && channelInnerName.replace(/\+/gi, "%20");

                        channelHtml += " <li class='parentLi'><span class='caretCustom caret-down treeDocLib' >"
                        channelHtml += "<a href= '#' class='secondLevel' data-teamId='" + id +
                            "' data-teamUrl='" + teamSPOUrl +
                            "' data-channelName='" + item.ChannelDisplayName +
                            "' data-channelId='" + item.ChannelId +
                            "' data-channelurl='" + item.ChannelUrl +
                            "' data-level='" + 1 +
                            "' data-foldertype='" + item.FolderType + "'>" +
                            "<i class='ms-Icon ms-Icon--FabricDocLibrary' aria-hidden='true'></i>" +
                            "<span class='singleLineEllipse'>" + item.ChannelDisplayName + "</span>" +
                            "</a ></span > ";
                    });
                    channelHtml += "</ul>"; //treeviewUL

                    $(container).append(channelHtml);
                    $(container).find('.caretCustom').addClass('caret-down');

                    $('#teamFolUl>li>span.treeDocLib').on("click", function () {
                        onExpandChannel($(this).find('.secondLevel'));
                    });

                    closeWaitDialog();
                });
                //getLibInternalNamegetLibInternalName(data).then(function (intName) {         
                //    getChannel(id, data, intName).then(function (channels) {
                //        console.log('channel', channels) 
                //    });
                //})
            });
        })
    }


    function onSetSaveLocation(DestFolderRelUrl, DestSite, SiteUrl, type, LocationAlreadyPinned) {

        if (type.toUpperCase() !== "SHAREPOINT" || type.toUpperCase() !== "SHAREPOINT-PINITEM") {

            let saveLocationPayload = {

                DestFolderRelUrl: DestFolderRelUrl,
                DestSite: DestSite,
                SiteUrl: SiteUrl,
                LocationAlreadyPinned: LocationAlreadyPinned ? LocationAlreadyPinned : false

            }

            setLocalForageItem("SaveLocation", JSON.stringify(saveLocationPayload)).done(function (values) {
                console.log("SaveLocation Data : " + values);
            });

        }
        setLocalForageItem("SaveLocationType", type).done(function (values) {
            console.log("SaveLocationType Data : " + values);
        });

    }

    function onExpandChannel(channel) {



        type = $(channel).data('foldertype');

        if (type.includes("folder") || type.includes("channel")) {
            $('#createFile').removeAttr('disabled');
            $("#createFile").css('background-color', '#04aba3');
            $("#createFile").css('cursor', 'pointer');
            $("#createFile").css('color', '#ffffff');
        } else {
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
        }

        $("span a").removeClass("treeselected");
        $("li a").removeClass("treeselected");

        $(channel).addClass("treeselected");



        let teamId = $(channel).data('teamid');
        let channelId = $(channel).data('channelid')
        let teamUrl = $(channel).data('teamurl');
        let channelName = $(channel).data('channelname');
        let channelUrl = $(channel).data('channelurl');
        let FolderType = $(channel).data('foldertype');
        let container = $(channel).parent().parent('li');

        let TeamShareFolderName;
        // var folderpath = intName + "/" + channelName
        let channelInnerName = channelUrl && channelUrl.substring(channelUrl.lastIndexOf("/") + 1, channelUrl.lastIndexOf("?"));
        channelInnerName = channelInnerName && channelInnerName.replace(/\+/gi, "%20");



        if (FolderType && FolderType.includes("sharepoint.folder")) {
            if (channelUrl) {

                let relPath = channelUrl && channelUrl;

                let siteUrl = teamUrl;

                onSetSaveLocation(
                    relPath,
                    siteUrl,
                    siteUrl,
                    "Teams"
                );
            }
        } else {

            listOfTeam.map(function (item, inx) {
                if (item.TeamId === teamId) {

                    TeamShareFolderName = item.TeamShareFolderName;

                }

            });

            onSetSaveLocation(
                TeamShareFolderName + "/" + channelInnerName,
                teamUrl,
                teamUrl,
                "Teams"
            );

        }

        if ($(channel).parents('span.caretCustom').siblings("ul.nested").length > 0) {
            $(channel).parents('span.caretCustom').siblings("ul.nested").removeClass("nested");
            $(channel).parents('span.caretCustom').addClass("caret-down");
            return;
        } else if ($(channel).parents('span.caretCustom').siblings("ul:eq(0)").length === 1) {


            $(channel).parents('span.caretCustom').siblings("ul:eq(0)").addClass("nested");
            $(channel).parents('span.caret-down').removeClass("caret-down");
            return;
        }
        openWaitDialog();
        getLocalForageItem("Teams").done(function (values) {

            listOfTeam = JSON.parse(values);

            listOfTeam.map(function (item, inx) {
                if (item.TeamId === teamId) {

                    TeamShareFolderName = item.TeamShareFolderName;

                }

            });
            const payload = {
                SPOUrl: ORG_ROOT_WEB.webUrl,
                tenant: ORG_TENANT.id,
                FolderRelPath: FolderType === "sharepoint.folder" ? channelUrl : TeamShareFolderName + "/" + channelInnerName,
                SiteUrl: teamUrl,
                ChannelID: channelId,
                TeamID: teamId,
                FolderType: FolderType === "sharepoint.folder" ? FolderType : "#microsoft.graph.channel"

            };
            getChannelFolder(payload).then(function (res) {
                let allMsTeamsChannelFolder = [];
                if (!res.error) {

                    let folders = res.d.results;

                    if (folders.length > 0) {
                        folders.map(folder => {
                            const eachFolder = {};
                            // eachTemplate.isSelected = false;

                            eachFolder.FolderName = folder.Name;
                            eachFolder.FolderUrl = folder.__metadata.uri;
                            eachFolder.FolderRelativeUrl = folder.ServerRelativeUrl;
                            eachFolder.FolderLevel = 2;
                            eachFolder.ChildCount = folder.ItemCount;
                            eachFolder.FolderType = payload.FolderType;

                            if (payload.FolderType === "tab.sharepoint.folder" && folder.Name === "Forms") {
                                console.log(eachFolder);
                            } else if (payload.FolderType && payload.FolderType.includes("#microsoft.graph.channel")) {
                                let siteUrl = folder.__metadata.uri && folder.__metadata.uri.substr(0, folder.__metadata.uri.toLowerCase().indexOf("/_api"));
                                if (payload.SiteUrl && siteUrl.toLowerCase() !== payload.SiteUrl.toLowerCase()) {
                                    eachFolder.FolderType = "#microsoft.graph.channel.private";
                                    allMsTeamsChannelFolder.push(eachFolder);
                                } else {
                                    allMsTeamsChannelFolder.push(eachFolder);
                                }
                            } else {
                                allMsTeamsChannelFolder.push(eachFolder);
                            }
                        });
                    }
                }



                const tabPayload = {
                    tenant: ORG_TENANT.id,
                    ChannelID: channelId,
                    TeamID: teamId
                };

                getTab(tabPayload).then(function (res) {
                    // console.log(res);
                    // var file = res.value;
                    if (res.value) {
                        if (res.value.length > 0) {
                            res.value.map(folder => {
                                const eachFolder = {};
                                // eachTemplate.isSelected = false;

                                eachFolder.FolderName = decodeURIComponent(folder.displayName);
                                eachFolder.FolderUrl = folder.configuration.contentUrl;
                                eachFolder.FolderRelativeUrl = String(folder.configuration.contentUrl).substr(String(folder.configuration.contentUrl).toLowerCase().indexOf("/sites"), (String(folder.configuration.contentUrl).length - 1));
                                eachFolder.FolderLevel = 2;
                                eachFolder.FolderType = "tab.sharepoint.folder";
                                allMsTeamsChannelFolder.push(eachFolder);
                            });
                        }
                    }

                    if (allMsTeamsChannelFolder.length > 0) {
                        getLocalForageItem("Teams").done(function (values) {

                            listOfTeam = JSON.parse(values);

                            listOfTeam.map((item, inx) => {
                                if (item.TeamId === teamId) {

                                    item.TeamChannels.map((channel, inx) => {

                                        if (channel.ChannelId === channelId) {
                                            channel.FolderList = allMsTeamsChannelFolder;
                                            channel.ISSelectable = true;
                                            channel.FolderCount = channel.FolderList ? channel.FolderList.length : 0;

                                        }
                                    });

                                }

                            });

                            setLocalForageItem("Teams", JSON.stringify(listOfTeam)).done(function (values) {
                                console.log("Team Data : " + values);
                            });

                        });


                        var folderHtml = "<ul id='folderUl_2' class='scUl treefolderUl'>";
                        allMsTeamsChannelFolder.map((folders, inx) => {
                            folderHtml += " <li class='parentLi'><span class='caretCustom caret-down treeDocLib' >"
                            folderHtml += "<a href= '#' class='nextLevel' data-teamid='" + teamId +
                                "'data-teamurl='" + teamUrl +
                                "' data-teamsharefoldername='" + TeamShareFolderName +
                                "' data-channelname='" + channelName +
                                "' data-channelid='" + channelId +
                                "' data-channelurl='" + channelUrl +
                                "' data-channelfoldertype='" + FolderType +

                                "' data-foldertype='" + folders.FolderType +
                                "' data-foldername='" + folders.FolderName +
                                "' data-folderlevel='" + 2 +
                                "' data-folderurl='" + folders.FolderUrl +
                                "' data-folderrelativeurl='" + folders.FolderRelativeUrl +

                                "'> <i class='ms-Icon ms-Icon--FabricFolderFill' aria-hidden='true'></i>" +
                                "<span class='singleLineEllipse'>" + folders.FolderName + "</span></a ></span > ";
                        });
                        folderHtml += "</ul>"; //treeviewUL
                        $(container).append(folderHtml);
                        $(container).find('.caretCustom').addClass('caret-down');


                        $('#folderUl_2>li>span.treeDocLib').on("click", function () {
                            onExpandFolder($(this).find('.nextLevel'), 2);
                        });
                    }

                    closeWaitDialog();

                });



            });

        });

    }

    function onExpandFolder(ele, level) {



        type = $(ele).data('foldertype');

        if (type.includes("folder") || type.includes("channel")) {
            $('#createFile').removeAttr('disabled');
            $("#createFile").css('background-color', '#04aba3');
            $("#createFile").css('cursor', 'pointer');
            $("#createFile").css('color', '#ffffff');
        } else {
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
        }

        $("span a").removeClass("treeselected");
        $("li a").removeClass("treeselected");

        $(ele).addClass("treeselected");


        let teamId = $(ele).data('teamid');
        let TeamShareFolderName = $(ele).data('teamsharefoldername');

        let channelId = $(ele).data('channelid')
        let channelUrl = $(ele).data('channelurl');
        let channelFolderType = $(ele).data('channelfoldertype');
        let teamUrl = $(ele).data('teamurl');
        let channelName = $(ele).data('channelname');
        let foldertype = $(ele).data('foldertype');
        let FolderLevel = $(ele).data('folderlevel');
        let FolderRelativeUrl = $(ele).data('folderrelativeurl');
        let FolderUrl = $(ele).data('folderurl');
        let FolderName = $(ele).data('foldername');
        //let intName = "Shared Documents"
        //let path = intName + "/" + channelName + '/' + folder
        let container = $(ele).parent().parent('li');


        let channelInnerName = channelUrl && channelUrl.substring(channelUrl.lastIndexOf("/") + 1, channelUrl.lastIndexOf("?")) || '';
        channelInnerName = channelInnerName && channelInnerName.replace(/\+/gi, "%20");

        let folderRelPath = '';

        folderRelPath = FolderRelativeUrl && FolderRelativeUrl.substr(FolderRelativeUrl.indexOf(channelInnerName)) || '';

        let payload = {
            SPOUrl: ORG_ROOT_WEB.webUrl,
            tenant: ORG_TENANT.id,
            FolderRelPath: ((foldertype && foldertype.includes("sharepoint.folder")) || channelFolderType === "sharepoint.folder") ? FolderUrl : TeamShareFolderName + "/" + folderRelPath,
            SiteUrl: teamUrl,
            ChannelID: channelId,
            TeamID: teamId,
            FolderType: (foldertype && foldertype.includes("sharepoint.folder")) ? foldertype : channelFolderType

        };

        if (foldertype && (foldertype.includes("sharepoint.folder") || foldertype.includes("#microsoft.graph.channel.private"))) {
            if (FolderUrl) {
                let splitString = FolderUrl.split("/");
                let teamName = splitString[4] && splitString[4];

                let relPath = FolderRelativeUrl && FolderRelativeUrl;

                let siteUrl = ORG_ROOT_WEB.webUrl + "/sites/" + teamName;

                onSetSaveLocation(relPath, siteUrl, siteUrl, "Teams");

                payload = {
                    SPOUrl: ORG_ROOT_WEB.webUrl,
                    tenant: ORG_TENANT.id,
                    FolderRelPath: FolderRelativeUrl,
                    SiteUrl: siteUrl,
                    ChannelID: channelId,
                    TeamID: teamId,
                    FolderType: foldertype

                };
            }
        } else {
            onSetSaveLocation(
                TeamShareFolderName + "/" + folderRelPath,
                teamUrl,
                teamUrl,
                "Teams"
            );
        }

        if ($(ele).parents('span.caretCustom').siblings("ul.nested").length > 0) {
            $(ele).parents('span.caretCustom').siblings("ul.nested").removeClass("nested");
            $(ele).parents('span.caretCustom').addClass("caret-down");
            return;
        } else if ($(ele).parents('span.caretCustom').siblings("ul:eq(0)").length === 1) {


            $(ele).parents('span.caretCustom').siblings("ul:eq(0)").addClass("nested");
            $(ele).parents('span.caret-down').removeClass("caret-down");
            return;
        }

        openWaitDialog();
        getChannelFolder(payload).then(function (res) {

            if (res.error) {
                closeWaitDialog();
            }
            let folders = res.d.results;
            let allMsTeamsChannelFolder = [];

            if (folders.length > 0) {
                folders.map(folder => {
                    const eachFolder = {};
                    // eachTemplate.isSelected = false;

                    eachFolder.FolderName = folder.Name;
                    eachFolder.FolderUrl = folder.__metadata.uri;
                    eachFolder.FolderRelativeUrl = folder.ServerRelativeUrl;
                    eachFolder.FolderLevel = parseInt(FolderLevel) + 1;
                    eachFolder.ChildCount = folder.ItemCount;
                    eachFolder.FolderType = payload.FolderType;

                    if (payload.FolderType === "tab.sharepoint.folder" && folder.Name === "Forms") {
                        console.log(eachFolder);
                    } else if (payload.FolderType && payload.FolderType.includes("#microsoft.graph.channel")) {
                        let siteUrl = folder.__metadata.uri && folder.__metadata.uri.substr(0, folder.__metadata.uri.toLowerCase().indexOf("/_api"));
                        if (payload.SiteUrl && siteUrl.toLowerCase() !== payload.SiteUrl.toLowerCase()) {
                            eachFolder.FolderType = "#microsoft.graph.channel.private";
                            allMsTeamsChannelFolder.push(eachFolder);
                        } else {
                            allMsTeamsChannelFolder.push(eachFolder);
                        }
                    } else {
                        allMsTeamsChannelFolder.push(eachFolder);
                    }
                });

                getLocalForageItem("Teams").done(function (values) {

                    listOfTeam = JSON.parse(values);

                    listOfTeam.map((item, inx) => {
                        if (item.TeamId === teamId) {

                            item.TeamChannels.map((channel, inx) => {

                                if (channel.ChannelId === channelId) {

                                    channel.FolderList.map((chnfoldr) => {
                                        if (chnfoldr.FolderLevel === level && chnfoldr.FolderRelativeUrl === FolderRelativeUrl) {
                                            chnfoldr.SubFolders = allMsTeamsChannelFolder;
                                            chnfoldr.ISSelectable = true;
                                            chnfoldr.FolderCount = chnfoldr.SubFolders ? chnfoldr.SubFolders.length : 0;

                                        } else {
                                            setSubFolder(channel, FolderRelativeUrl, chnfoldr, allMsTeamsChannelFolder, parseInt(FolderLevel) - 1, 2);
                                        }

                                    });

                                }
                            });

                        }

                    });

                    setLocalForageItem("Teams", JSON.stringify(listOfTeam)).done(function (values) {
                        console.log("Team Data : " + values);
                        listOfTeam = JSON.parse(values);

                        let nxtFolderLevel = parseInt(FolderLevel) + 1;
                        var folderHtml = "<ul id='folderUl_" + nxtFolderLevel + "' class='scUl treefolderUl'>";

                        // listOfTeam.map((item, inx) => {
                        //     if (item.TeamId === teamId) {

                        //         item.TeamChannels.map((channel, inx) => {

                        //             if (channel.ChannelId === channelId) {


                        //                 channel.FolderList.map((chnfoldr) => {
                        //                     if (chnfoldr.FolderLevel === level && chnfoldr.FolderRelativeUrl === FolderRelativeUrl) {

                        allMsTeamsChannelFolder.map((folders, inx) => {
                            folderHtml += " <li class='parentLi'><span class='caretCustom caret-down treeDocLib' >"
                            folderHtml += "<a href= '#' class='nextLevel' data-teamid='" + teamId +
                                "'data-teamurl='" + teamUrl +
                                "' data-teamsharefoldername='" + TeamShareFolderName +
                                "' data-channelname='" + channelName +
                                "' data-channelId='" + channelId +
                                "' data-channelurl='" + channelUrl +
                                "' data-channelfoldertype='" + channelFolderType +

                                "' data-foldertype='" + folders.FolderType +
                                "' data-foldername='" + folders.FolderName +
                                "' data-folderlevel='" + (parseInt(FolderLevel) + 1) +
                                "' data-folderurl='" + folders.FolderUrl +
                                "' data-folderrelativeurl='" + folders.FolderRelativeUrl +

                                "'> <i class='ms-Icon ms-Icon--FabricFolderFill' aria-hidden='true'></i>" +
                                "<span class='singleLineEllipse'>" + folders.FolderName + "</span></a ></span > ";
                        });
                        //                     }
                        //                 });
                        //             }
                        //         });
                        //     }
                        // });

                        folderHtml += "</ul>"; //treeviewUL
                        $(container).append(folderHtml);
                        $(container).find('.caretCustom').addClass('caret-down');


                        $('#folderUl_' + nxtFolderLevel + '>li>span.treeDocLib').on("click", function () {
                            onExpandFolder($(this).find('.nextLevel'), nxtFolderLevel);
                        });
                    });

                });


            }
            closeWaitDialog();

        });

    }

    function setSubFolder(rootchannel, rootfolderRelativeUrl, initialFolderlist, respFolderList, folderLevel, initialLevel) {


        for (let i = initialLevel; i <= folderLevel; i++) {
            if (initialLevel === folderLevel && initialFolderlist.SubFolders) {

                initialFolderlist.SubFolders.map((subfoldr, inx) => {
                    if (subfoldr.FolderRelativeUrl === rootfolderRelativeUrl) {
                        subfoldr.SubFolders = respFolderList;
                    }
                });

                return initialFolderlist;
            } else {

                if (initialFolderlist.SubFolders) {
                    initialFolderlist.SubFolders.map((subfoldr, inx) => {
                        return setSubFolder(rootchannel, rootfolderRelativeUrl, subfoldr, respFolderList, folderLevel, i);
                    });
                }


            }



        }
    }

    function getOneDrive() {
        openWaitDialog();
        //var dfd = $.Deferred();
        try {

            if (SPToken) {
                let payload = {
                    UsrGUID: USER_PROP.id,
                    tenant: ORG_TENANT.id,
                }

                $.ajax({
                    url: GET_USER_ONEDRIVE,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    console.log(response);
                    var resultsCount = response.value;
                    let allOneDrive = [];
                    if (resultsCount) {
                        resultsCount.map(drive => {
                            const eachDriveItem = {};
                            // eachTemplate.isSelected = false;

                            if (drive && drive.folder) {
                                eachDriveItem.DriveName = drive.name;
                                eachDriveItem.DriveId = drive.id;
                                eachDriveItem.DriveUrl = drive.webUrl;
                                eachDriveItem.DriveFolderPath = drive.name;
                                eachDriveItem.DriveLevel = 0;
                                if (drive.parentReference) {
                                    eachDriveItem.ParentDriveId = drive.parentReference.driveId;
                                    eachDriveItem.ParentDrivePath = drive.parentReference.path;
                                }
                                if (drive.folder) {
                                    eachDriveItem.ChildernCount = drive.folder.childCount;
                                }
                                eachDriveItem.ListItemNavigationLink = drive["listItem@odata.navigationLink"];
                                if (drive["#microsoft.graph.createUploadSession"]) {
                                    eachDriveItem.UploadSessionLink = drive["#microsoft.graph.createUploadSession"].target;
                                }
                                allOneDrive.push(eachDriveItem);
                            }


                        });


                        setLocalForageItem("OneDrive", JSON.stringify(allOneDrive)).done(function (values) {
                            console.log("OneDrive Data : " + values);
                        })

                    }



                    var firstRow = "<ul id='teamviewUl' class='driveViewUl'>";
                    resultsCount.map(item => {
                        if (item && item.folder) {
                            firstRow += " <li class='parentLi'><span class='caretCustom treeSpan' >"
                            firstRow += " <a href= '#' class='firstLavelOne' data-onename='" + item.name + "' data-oneid='" + item.id + "' data-onedriveurl= '" + item.webUrl + "' data-drivelevel= '" + 0 + "' > <i class='ms-Icon ms-Icon--OneDriveFolder16' aria-hidden='true'></i><span class='singleLineEllipse'>" + item.name + "</span></a ></span > ";
                        }

                    });
                    firstRow += "</ul>"; //treeviewUL

                    $('#oneDriveLoc').html(firstRow);

                    $('.driveViewUl>li>span.treeSpan').on("click", function () {
                        expandDriveFolder($(this).find('.firstLavelOne'));
                    });

                    //dfd.resolve(response.d.results);
                    closeWaitDialog();
                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            //dfd.reject('Error');
            closeWaitDialog();
        }
        //return dfd.promise();
    };

    function expandDriveFolder(ele) {


        type = "onedrive";

        if (type.includes("onedrive")) {
            $('#createFile').removeAttr('disabled');
            $("#createFile").css('background-color', '#04aba3');
            $("#createFile").css('cursor', 'pointer');
            $("#createFile").css('color', '#ffffff');
        } else {
            $("#createFile").attr('disabled', 'disabled');
            $("#createFile").css('background-color', '');
            $("#createFile").css('cursor', 'default');
        }

        $("span a").removeClass("treeselected");
        $("li a").removeClass("treeselected");

        $(ele).addClass("treeselected");


        var onename = $(ele).data('onename');
        var oneid = $(ele).data('oneid');
        var onedriveurl = $(ele).data('onedriveurl');
        var drivelevel = $(ele).data('drivelevel');

        var container = $(ele).parent().parent('li');


        onSetSaveLocation(
            onename,
            onedriveurl,
            "Onedrive",
            "OneDrive"
        );


        if ($(ele).parents('span.caretCustom').siblings("ul.nested").length > 0) {
            $(ele).parents('span.caretCustom').siblings("ul.nested").removeClass("nested");
            $(ele).parents('span.caretCustom').addClass("caret-down");
            return;
        } else if ($(ele).parents('span.caretCustom').siblings("ul:eq(0)").length === 1) {


            $(ele).parents('span.caretCustom').siblings("ul:eq(0)").addClass("nested");
            $(ele).parents('span.caret-down').removeClass("caret-down");
            return;
        }


        getOneDriveFolder(oneid).then(function (res) {
            console.log(res)
            var folders = res.value;


            const allOneDrive = [];

            folders.map(drive => {
                const eachDriveItem = {};
                // eachTemplate.isSelected = false;

                if (drive && drive.folder) {
                    eachDriveItem.DriveName = drive.name;
                    eachDriveItem.DriveId = drive.id;
                    eachDriveItem.DriveUrl = drive.webUrl;
                    if (drive.parentReference) {
                        eachDriveItem.ParentDriveId = drive.parentReference.driveId;
                        eachDriveItem.ParentDrivePath = drive.parentReference.path;
                    }
                    if (drive.folder) {
                        eachDriveItem.ChildernCount = drive.folder.childCount;
                    }
                    eachDriveItem.ListItemNavigationLink = drive["listItem@odata.navigationLink"];
                    if (drive["#microsoft.graph.createUploadSession"]) {
                        eachDriveItem.UploadSessionLink = drive["#microsoft.graph.createUploadSession"].target;
                    }
                    allOneDrive.push(eachDriveItem);
                }


            });

            getLocalForageItem("OneDrive").done(function (values) {

                listOfOneDrive = JSON.parse(values);

                listOfOneDrive.map(function (item, inx) {
                    if (item.DriveId === oneid) {
                        allOneDrive.map((driveItem) => {
                            driveItem.DriveFolderPath = onename + "/" + driveItem.DriveName;
                        });

                        item.DriveChannels = allOneDrive;
                        item.DriveChannelLoaded = true;
                        item.ISSelectable = true;
                        item.DriveLevel = (drivelevel ? parseInt(drivelevel) : 0) + 1;
                        item.ChildCount = 0;
                        if (allOneDrive.length > 0) {
                            item.ChildCount = allOneDrive.length;
                        }
                    }
                });

                setLocalForageItem("OneDrive", JSON.stringify(listOfOneDrive)).done(function (values) {
                    console.log("OneDrive Data : " + values);
                });

                var folderHtml = "<ul class='scUlNext'>";

                allOneDrive.map(item => {
                    folderHtml += " <li class='parentLi'><span class='caretCustom caret-down treeDocLib' >"
                    folderHtml += " <a href= '#' class='firstLavelOne' data-oneid='" + item.DriveId + "' data-onename='" + item.DriveFolderPath + "' data-onedriveurl= '" + item.DriveUrl + "' data-drivelevel= '" + (drivelevel ? parseInt(drivelevel) : 0) + 1 + "'  > <i class='ms-Icon ms-Icon--FabricFolderFill' aria-hidden='true'></i><span class='singleLineEllipse'>" + item.DriveName + "</span></a ></span > ";
                });
                folderHtml += "</ul>"; //treeviewUL
                $(container).append(folderHtml);
                $(container).find('.caretCustom').addClass('caret-down');

                $('.scUlNext>li>span.treeDocLib').off('click');
                $('.scUlNext>li>span.treeDocLib').on("click", function () {
                    expandDriveFolder($(this).find('.firstLavelOne'));
                });

                closeWaitDialog();

            });




        });
    }


    function getOneDriveFolder(itemID) {
        openWaitDialog();
        var dfd = $.Deferred();
        try {

            if (SPToken) {
                let payload = {
                    ItemID: itemID,
                    UsrGUID: USER_PROP.id,
                    tenant: ORG_TENANT.id,
                }

                $.ajax({
                    url: GET_USER_ONEDRIVE_HIERARCHY,
                    beforeSend: function (request) {
                        request.setRequestHeader("Accept", "application/json; odata=verbose");
                    },
                    dataType: "json",
                    headers: {
                        'Authorization': 'Bearer ' + SPToken,
                    },
                    type: "POST",
                    contentType: "application/json",
                    data: JSON.stringify(payload),
                    // 
                }).done(function (response) {
                    console.log(response);

                    dfd.resolve(response);
                });
            }
        } catch (err) {
            console.log("getAllSiteCollection_Treeview: " + err);
            dfd.reject('Error');
        }
        return dfd.promise();
    };


}