//Importing files and objects creation
import * as React from 'react';
import * as $ from "jquery";
import styles from './DocsNodeAdmin.module.scss';
import { IDocsNodeAdminProps } from './IDocsNodeAdminProps';
import { Selection, SelectionMode, IColumn, ColumnActionsMode } from 'office-ui-fabric-react/lib/DetailsList';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import DatabaseConfiguration from './DatabaseConfiguration';
import { IDetailsListDocumentsState, IDocument } from './IDetailsListDocumentsStates';
import { CommandBarButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { RichText } from '@pnp/spfx-controls-react/lib/RichText';
import constant from './Constant';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Nav } from 'office-ui-fabric-react/lib/Nav';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import CommonUtility from "./CommonUtility";
import { toast, ToastContainer } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.min.css';

//Created objects of other files
const DC: DatabaseConfiguration = new DatabaseConfiguration();
const CU: CommonUtility = new CommonUtility();

//Declared and initialized global variables for See more functionalities 
let _INTERVAL_DELAY = 100;

//SPComponentLoader.loadCss('https://unpkg.com/react-dropdown-tree-select@1.11.3/dist/styles.css');

export default class DocsNodeAdmin extends React.Component<IDocsNodeAdminProps, IDetailsListDocumentsState> {

  //Declaration and initialization of variables
  public _selection: Selection;
  public _SlidesItemArray = [];
  public _ImagesItemArray = [];
  public _TextSnippetItemArray = [];
  public _CategoriesItemArray = [];
  public _PlaceHolderItemArray = [];
  public _CategoriesItemsData = [];
  public _CategoriesDrpDwnItemArray = [];
  public _ParentCategoriesItemArray = [];
  public _ParentCategItemArray = [];
  public _SearchResult = [];
  public _lastIndexWithData: number;
  public _ITEMS_COUNT = 0;
  public _items = [];
  public _lastIntervalId: number;
  public _siteCollitem = [];
  public richText;
  public _listAllSubsite = [];
  public columns: IColumn[];
  public CategoryColumns: IColumn[];
  public PlaceHolderColumn: IColumn[];
  public isRoot = true;
  public listAllSubsite = [];
  public _SiteColectionValue = "";
  public _Subsiteurl = "";
  public _ListName = "";
  public _FieldName = "";

  /**
   *Props initialization. */
  constructor(props) {
    super(props);

    const data = [{
      label: 'Node1',
      value: '0-0',
      key: '0-0',
      children: [{
        label: 'Child Node1',
        value: '0-0-1',
        key: '0-0-1',
      }, {
        label: 'Child Node2',
        value: '0-0-2',
        key: '0-0-2',
      }],
    }, {
      label: 'Node2',
      value: '0-1',
      key: '0-1',
    }];

    //Creating columns for Slide,Image,Text snippet
    this.columns = [
      {
        key: 'column1',
        name: 'File Type',
        className: styles.fileIconCell,
        iconClassName: styles.fileIconHeaderIcon,
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'FileType',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item: IDocument) => {
          return (item.FileType != '' ? <img src={item.FileType} className={styles.fileIconImg} /> : <div style={{ fontSize: '16px' }}><Icon iconName='TextDocument' /></div>);
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'FileName',
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        isSorted: null,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return (this.state.tabValue == constant.TextSnippetDisplayName ? <span title={item.FileName}>{item.FileName}</span> : <Link href={item.LinkURL} target='_blank' title={item.FileName}>{item.FileName}</Link>);
        },
        isPadded: true
      },
      {
        key: 'column3',
        name: constant.CategoryName,
        fieldName: constant.CategoryName,
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span title={item.Category}>{item.Category}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Edit',
        fieldName: 'edit',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editItem(item)} />;
        }
      },
      {
        key: 'column5',
        name: 'Delete',
        fieldName: 'delete',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._showDialog(item)} />;
        }
      }
    ];

    //Creating columns for Category List
    this.CategoryColumns = [
      {
        key: 'column1',
        name: constant.CategoryName,
        fieldName: constant.CategoryName,
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span title={item.Category}>{item.Category}</span>;
        }
      },
      {
        key: 'column2',
        name: 'Parent Category',
        fieldName: constant.CategoryParentId,
        minWidth: 100,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span title={item.ParentCategory}>{item.ParentCategory}</span>;
        },
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Category Type',
        fieldName: constant.CategoryType,
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: true,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span title={item.CategoryType}>{item.CategoryType}</span>;
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Edit',
        fieldName: 'edit',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editItem(item)} />;
        }
      },
      {
        key: 'column5',
        name: 'Delete',
        fieldName: 'delete',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return (item.DeleteEditFlag == false ? null : <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._showDialog(item)} />);
        }
      }
    ];

    //Create columns of Place holder
    this.PlaceHolderColumn = [
      {
        key: 'column1',
        name: constant.PlaceHolderDisplayName,
        fieldName: 'PlaceholderName',
        minWidth: 100,
        maxWidth: 300,
        isRowHeader: true,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span title={item.PlaceholderName}>{item.PlaceholderName}</span>;
        }
      },
      {
        key: 'column2',
        name: 'Edit',
        fieldName: 'edit',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._editItem(item)} />;
        }
      },
      {
        key: 'column2',
        name: 'Delete',
        fieldName: 'delete',
        minWidth: 20,
        maxWidth: 40,
        isRowHeader: true,
        isResizable: true,
        isSorted: false,
        onRender: (item: IDocument) => {
          return <IconButton iconProps={{ iconName: 'Delete' }} onClick={() => this._showDialog(item)} />;
        }
      }
    ];

    //Selection of N numbers of rows
    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails()
        });
      }
    });

    //Initializing state
    this.state = {
      items: [],
      columns: this.columns,
      categoryColumns: this.CategoryColumns,
      placeholderColumns: this.PlaceHolderColumn,
      selectionDetails: '',
      isModalSelection: true,
      isCompactMode: false,
      CategaryDropDown: [],
      isDataLoaded: false,
      tabValue: '',
      defaultKey: '',
      showPanel: false,
      defaultNavSelectdKey: 'key1',
      searchValue: '',
      itemDiscription: '',
      itemTextSnippet: '',
      itemTitle: '',
      itemName: '',
      newPlaceHolder: '',
      addOrEditBtn: false,
      listName: '',
      itemId: null,
      renderKey: '',
      hideDialog: true,
      hideSiteAdminDialog: true,
      deleteItem: [],
      showList: true,
      showcategoryListForm: false,
      showPlaceHolderList: false,
      newCategoryTitle: '',
      radioBtnCategory: '',
      data: data,
      disabledChoiceOne: false,
      disabledChoiceTwo: false,
      disabledChoiceThree: false,
      innerDialogbox: true,
      selectKeyInner: 'key1',
      disabledParentCatgy: false,
      showConfigList: false,
      progressLevel: true,
      sitecoll: [],
      subSiteItem: [],
      listItem: [],
      fieldItem: [],
      SiteColectionValue: '',
      SubSiteCollectionKey: '',
      ListKey: '',
      FieldName: '',
      disabledsubsite: true,
      disableddoclib: true,
      disabledfolder: true,
      configImgLibraryOptons: [],
      configSlideLibraryOption: [],
      configTempLibraryOption: [],
      configTextSnipListOption: [],
      tempLibSelectKey: '',
      slideLibSelectKey: '',
      imgLibSelectKey: '',
      textSniSelectKey: '',
      disbledSaveBtn: false,
      imgLogoURl: '',
      imgLogoID: '',
      slideCheckBoxName: '',
      imgCheckBoxName: '',
      textSniCheckBox: '',
      slideChecked: false,
      imgChecked: false,
      textSniChecked: false,
      FavoriteAndAllSites: 'ChevronUp'
    };

    //Binding methods to get current context inside method    
    this._onSearchText = this._onSearchText.bind(this);
    this._onLoadData = this._onLoadData.bind(this);
    this._refreshCategories = this._refreshCategories.bind(this);
    this._getDocsNodeImageItems = this._getDocsNodeImageItems.bind(this);
    this._getDocsNodeSlidesItems = this._getDocsNodeSlidesItems.bind(this);
    this._getDocsNodeTextSnippetItems = this._getDocsNodeTextSnippetItems.bind(this);
    this._getConfigurationListItems = this._getConfigurationListItems.bind(this);
    this._onFilterByCategoryListDataChange = this._onFilterByCategoryListDataChange.bind(this);
    this._showPanel = this._showPanel.bind(this);
    this._addNewData = this._addNewData.bind(this);
    this._onLinkClickAssets = this._onLinkClickAssets.bind(this);
    this._editItem = this._editItem.bind(this);
    this._deleteItem = this._deleteItem.bind(this);
    this._EmptyList = this._EmptyList.bind(this);
    this._SelectCategory = this._SelectCategory.bind(this);
    this._showDialog = this._showDialog.bind(this);
    this._onChangeCategoryType = this._onChangeCategoryType.bind(this);
    this._getDocsNodeCategoryItems = this._getDocsNodeCategoryItems.bind(this);
    this.refreshCategoryDrpDwn = this.refreshCategoryDrpDwn.bind(this);
    this._showInnerDialog = this._showInnerDialog.bind(this);
    this.LinkClickInner = this.LinkClickInner.bind(this);
    this._OnChangeTemplateLib = this._OnChangeTemplateLib.bind(this);
    this._OnChangeSlideLib = this._OnChangeSlideLib.bind(this);
    this._OnChangeImageLib = this._OnChangeImageLib.bind(this);
    this._OnChangeTextSnippetList = this._OnChangeTextSnippetList.bind(this);
    this._addConfigData = this._addConfigData.bind(this);
    this._onCheckboxChange = this._onCheckboxChange.bind(this);
    this._onTextChange = this._onTextChange.bind(this);

  }

  /**
   *This method is invoked immediately after a component is mounted.
   *This will load data from a remote endpoint. */
  public componentDidMount() {
    try {
      //Check wheather user is admin or not
      CU._checkUserIsSiteAdmin().then(async (response) => {
        if (response) {
          //Check new site collection exist or not
          CU._checkSiteCollectionExist(CU.tenantURL() + CU.siteCollectionPath, this.props.context).then(async (d) => {
            if (d) {
              await CU.getGroupID().then(async (responseResult) => {
                if (responseResult == null) {
                  await CU._createNewSiteGrp();
                }
              });
              //Check Default List and Library
              this._checkDefaultListAndLibrary();
            } else {
              //Create New Team Site
              CU._createNewTeamSite(this.props.context)
                .then(async (data: any) => {
                  if (data) {
                    //Add all user in Read permission group
                    await CU._ensureUser(4).then(() => {
                    }, (error) => {
                      console.log('An error occured while adding user');
                    });
                    await CU._createNewSiteGrp();
                    //Set Time out after creating New Team site                                         
                    //Check Default List and Library
                    this._checkDefaultListAndLibrary();
                  }
                });
            }
          });
        } else {
          this.setState({
            hideSiteAdminDialog: false
          });
        }
      });
    } catch (error) {
      console.log('new siteColection : ' + error);
    }
  }

  /**
   *Check the All Default List and Library are exist. */
  public _checkDefaultListAndLibrary(): void {

    //Get default List and Library Array
    var defaultLstArry = DC._defaultListOrLibraryArray();

    //If not, then this would create  new one
    DC._checkListExistsOrNot(defaultLstArry).then(async (data) => {

      //Check the columns of All Default List created are exist 
      //If not, then this would create columns for that List or Library
      await DC._checkForColumnExistence(defaultLstArry);

      //Insert all default items in configuration list
      await DC._insertConfigData(this.props.context);
      var arry = { name: constant.SlidesDisplayName };

      //By default select Slide section in Manage Assets
      this._onLinkClickAssets('', arry);
      //Getthe existance list and Post new list data and add new column if columns are not exists
      await DC._dynamicListOrLibraryArray();
    });
  }

  /**
   *This function is to get all items from DocsNodeSlide Library. */
  public async _getDocsNodeSlidesItems() {

    var newColumns = [];
    //Await to get all items from DocsNodeSlide Library
    this._SlidesItemArray = await DC._getDocsNodeSlidesName();
    this._SearchResult = this._SlidesItemArray;

    //Empty List
    await this._EmptyList('key1', true, false, false, false);

    //Await to get all Categories from DocsNodeCategory List
    await this._refreshCategories(constant.SlidesDisplayName);

    //Refresh Column header and remove arrow if there on section change
    newColumns = this.refreshColumnHeaderSorting(this.columns);

    //This function is for lazy loading of items in List component
    this._onLoadData(this._SlidesItemArray.length, this._SlidesItemArray);

    //Binding the data
    this.setState({
      columns: newColumns,
      CategaryDropDown: this._CategoriesItemArray,
      defaultKey: '',
      renderKey: '',
      tabValue: constant.SlidesDisplayName,
      listName: constant.DocsNodeSlidesName,
      selectionDetails: 'No items selected',
      hideDialog: true,
      searchValue: '',
      hideSiteAdminDialog: true
    });
  }

  /**
   *This function to get all items from Picture Library. */
  public async _getDocsNodeImageItems() {

    var newColumns = [];
    //Empty List
    await this._EmptyList('key2', true, false, false, false);

    //Await to get all Categories from DocsNodeCategory List
    await this._refreshCategories(constant.ImagesDisplayName);

    //Refresh Column header and remove arrow if there on section change
    newColumns = this.refreshColumnHeaderSorting(this.columns);

    //Calling the api to get all items from DocsNodePicture Library
    this._ImagesItemArray = await DC._getDocsNodePictureName();
    this._SearchResult = this._ImagesItemArray;

    //This function is for lazy loading of items in List component
    this._onLoadData(this._ImagesItemArray.length, this._ImagesItemArray);

    //Binding the data
    this.setState({
      columns: newColumns,
      CategaryDropDown: this._CategoriesItemArray,
      defaultKey: '',
      renderKey: '',
      tabValue: constant.ImagesDisplayName,
      listName: constant.DocsNodePictureName,
      selectionDetails: 'No items selected',
      hideDialog: true,
      searchValue: ''
    });
  }

  /**
   * This function is to get all items from Text Snippet List. */
  public async _getDocsNodeTextSnippetItems() {
    var newColumns = [];
    //Empty List
    await this._EmptyList('key3', true, false, false, false);

    //Await to get all Categories from DocsNodeCategory List
    await this._refreshCategories(constant.TextSnippetDisplayName);

    await this.setState({
      tabValue: constant.TextSnippetDisplayName,
    });

    //Refresh Column header and remove arrow if there on section change
    newColumns = this.refreshColumnHeaderSorting(this.columns);

    //Calling the api to get all items from DocsNodeTextSnippet List
    this._TextSnippetItemArray = await DC._getDocsNodeTextSnippetName();
    this._SearchResult = this._TextSnippetItemArray;

    //This function is for lazy loading of items in List component
    this._onLoadData(this._TextSnippetItemArray.length, this._TextSnippetItemArray);

    //Binding the data
    this.setState({
      columns: newColumns,
      CategaryDropDown: this._CategoriesItemArray,
      defaultKey: '',
      renderKey: '',
      listName: constant.DocsNodeTextName,
      selectionDetails: 'No items selected',
      hideDialog: true,
      searchValue: ''
    });
  }

  /**
   * This function is to get items from Category List. */
  public async _getDocsNodeCategoryItems() {
    var newColumns = [];

    //Empty List
    await this._EmptyList('key5', false, true, false, false);

    //Refresh Column header and remove arrow if there on section change
    newColumns = this.refreshColumnHeaderSorting(this.CategoryColumns);

    //Calling the api to get all items from DocsNodeCategory List
    DC._getDocsNodeCategoriesName('').then((responseData) => {

      this._CategoriesItemsData = responseData.DocsNodeCategoriesItemsData;
      this._SearchResult = responseData.DocsNodeCategoriesItemsData;

      this._ParentCategItemArray = responseData.DocsNodeParentCategoriesArrayItems;
      this._ParentCategoriesItemArray = [{ key: 'Header', text: 'Select Parent Category', itemType: DropdownMenuItemType.Header }].concat(this._ParentCategItemArray);

      //This function is for lazy loading of items in List component
      this._onLoadData(this._CategoriesItemsData.length, this._CategoriesItemsData);

      //Binding the data
      this.setState({
        categoryColumns: newColumns,
        tabValue: constant.CategoryDisplayName,
        renderKey: '',
        defaultKey: '',
        listName: constant.DocsNodeCategoriesName,
        hideDialog: true,
        disabledChoiceOne: false,
        disabledChoiceTwo: false,
        disabledChoiceThree: false,
        searchValue: ''
      });
    });
  }

  /**
   * This function is to get all items from Placeholder List. */
  public async _getDocsNodePlaceHolderItems() {
    var newColumns = [];

    //Empty List
    await this._EmptyList('key4', false, false, false, true);

    //Binding Site Collection dropdown in Panel
    await this._placeHolderAddClick();

    //Refresh Column header and remove arrow if there on section change
    newColumns = this.refreshColumnHeaderSorting(this.PlaceHolderColumn);

    //Calling the api to get all items from DocsNodePlaceHolder List
    DC._getDocsNodePlaceHolderName().then((responseData) => {
      this._PlaceHolderItemArray = responseData;
      this._SearchResult = responseData;

      //This function is for lazy loading of items in List component
      this._onLoadData(this._PlaceHolderItemArray.length, this._PlaceHolderItemArray);

      //Binding the data
      this.setState({
        placeholderColumns: newColumns,
        tabValue: constant.PlaceHolderDisplayName,
        renderKey: '',
        defaultKey: '',
        listName: constant.DocsNodePlaceHolderName,
        hideDialog: true,
        disabledChoiceOne: false,
        disabledChoiceTwo: false,
        disabledChoiceThree: false,
        searchValue: ''
      });
    });
  }

  /**
   * This function is to Render Manage Configuration Section. */
  public async _getConfigurationListItems() {
    var imgLogo = '';
    //Empty List
    await this._EmptyList('key6', false, false, true, false);
    await this.setState({
      progressLevel: false,
      disbledSaveBtn: true,
      tempLibSelectKey: '',
      imgLibSelectKey: '',
      imgLogoURl: imgLogo
    });

    //Get all item from configuration list
    var configListData = await DC._getDocsNodeConfigurationName();

    var templateLibName = '';
    var slideLibName = '';
    var imageLibName = '';
    var textSnippetListname = '';
    var imgLibdata = '';
    var libdata = '';
    if (configListData.length > 0) {
      for (var x = 0; x < configListData.length; x++) {
        var resultData = configListData[x];
        if (resultData.assetTitle == constant.TemplateTitle) {
          templateLibName = resultData.sourceLocation;
        }
        if (resultData.assetTitle == constant.SlideTitle) {
          slideLibName = resultData.sourceLocation;
        }
        if (resultData.assetTitle == constant.ImageTitle) {
          imageLibName = resultData.sourceLocation;
        }
        if (resultData.assetTitle == constant.TextSnippetTitle) {
          textSnippetListname = resultData.sourceLocation;
        }
      }
    }

    var getListUrl = CU.tenantURL() + CU.siteCollectionPath;

    //Get all Document Library from newly created Team site
    libdata = await DC._getAllList(getListUrl, 101);

    //Get all Picture Library from newly created Team site
    imgLibdata = await DC._getAllList(getListUrl, 109);

    //Get ProductLog from Picture library
    imgLogo = await DC._getProductLogo();

    var DDoptions = [];
    var DLoptions = [];
    var DMoptions = [];
    var listOption = [];
    var libOption = [];
    var imgLibOption = [];

    if (libdata.length > 0) {
      //Binding all Document Library
      DLoptions = DC._bindingAllList(libdata);
      if (DLoptions.length > 0) {
        libOption.push({ key: 'Header', text: 'Select Library', itemType: DropdownMenuItemType.Header });
        for (var j = 0; j < DLoptions.length; j++) {
          var data = DLoptions[j];
          if (data.key != 'Form Templates' && data.key != 'Site Assets' && data.key != 'Style Library' && data.key != constant.DocsNodeSlidesName) {
            libOption.push(data);
          }
        }
      } else {
        libOption = [{ key: 'Header', text: 'No Library available', itemType: DropdownMenuItemType.Header }];
      }
    }
    if (imgLibdata.length > 0) {
      //Binding all Picture Library
      DMoptions = DC._bindingAllList(imgLibdata);
      if (DMoptions.length > 0) {
        imgLibOption.push({ key: 'Header', text: 'Select Image Library', itemType: DropdownMenuItemType.Header });
        for (var k = 0; k < DMoptions.length; k++) {
          var dataResult = DMoptions[k];
          if (dataResult.key != constant.DocsNodeProductLogoName) {
            imgLibOption.push(dataResult);
          }
        }
      } else {
        imgLibOption = [{ key: 'Header', text: 'No Image Library available', itemType: DropdownMenuItemType.Header }];
      }
    }

    this.setState({
      listName: constant.DocsNodeConfigurationName,
      tabValue: constant.ConfigurationDisplayName,
      configTempLibraryOption: libOption,
      configImgLibraryOptons: imgLibOption,
      tempLibSelectKey: templateLibName,
      imgLibSelectKey: imageLibName,
      disbledSaveBtn: false,
      progressLevel: true,
      imgLogoURl: imgLogo['EncodeURL'],
      imgLogoID: imgLogo['itemID']
    });
  }

  /**
   * This function is for show OR hide section. 
   * @param key 
   * @param listFlag 
   * @param catFlag 
   * @param ConfigFlag 
   * @param placeholderFalg 
   */
  public _EmptyList(key, listFlag, catFlag, ConfigFlag, placeholderFalg): void {
    this.setState({
      items: [],
      isDataLoaded: false,
      defaultNavSelectdKey: key,
      showList: listFlag,
      showcategoryListForm: catFlag,
      showConfigList: ConfigFlag,
      showPlaceHolderList: placeholderFalg,
    });
  }

  /**
   *Following method is for filter the category list from the displaying items in List Component. */
  public _onFilterByCategoryListDataChange(event): void {
    var result = [];
    var flag = false;
    const { tabValue } = this.state;
    if (tabValue == constant.SlidesDisplayName) { //Slide
      this._SlidesItemArray.map((items) => {
        if (items.Category == event.text) {
          result.push(items);
          flag = true;
        } else if (event.text == constant.AllSlideName) {
          result = this._SlidesItemArray;
          flag = true;
        }
      });
      if (flag == false) {
        result = [];
      }
    } else if (tabValue == constant.ImagesDisplayName) { //Image
      this._ImagesItemArray.map((items) => {
        if (items.Category == event.text) {
          result.push(items);
          flag = true;
        } else if (event.text == constant.AllSlideName) {
          result = this._ImagesItemArray;
          flag = true;
        }
      });
      if (flag == false) {
        result = [];
      }
    } else {
      this._TextSnippetItemArray.map((items) => { //Text Snippet
        if (items.Category == event.text) {
          result.push(items);
          flag = true;
        } else if (event.text == constant.AllSlideName) {
          result = this._TextSnippetItemArray;
          flag = true;
        }
      });
      if (flag == false) {
        result = [];
      }
    }
    this._SearchResult = result;
    this.setState({
      items: result,
      defaultKey: event.key
    });
  }

  /**
   *Open Panel on click Add button. */
  public async _showPanel() {
    await this.setState({
      showPanel: true,
      itemTitle: '',
      itemDiscription: '',
      itemTextSnippet: '',
      addOrEditBtn: false,
      itemName: '',
      newCategoryTitle: '',
      renderKey: '',
      radioBtnCategory: '',
      disabledChoiceOne: false,
      disabledChoiceTwo: false,
      disabledChoiceThree: false,
      disabledParentCatgy: false,
      searchValue: '',
      SiteColectionValue: '',
      SubSiteCollectionKey: '',
      ListKey: '',
      FieldName: '',
      disabledsubsite: true,
      disableddoclib: true,
      disabledfolder: true,
      newPlaceHolder: ''
    });
    if (this.state.tabValue == constant.PlaceHolderDisplayName) {
      this.bindRootSiteCollections(this._siteCollitem, 'AllSiteCollections');
    }
  }

  /**
   * In creating new catagory on dropdown change radio button autofilled.
   * @param event 
   */
  public _SelectCategory(event): void {
    var ChoiceOne = false;
    var ChoiceTwo = false;
    var ChoiceThree = false;
    if (event.CategoryType == constant.ImagesDisplayName) {
      ChoiceOne = true;
      ChoiceTwo = false;
      ChoiceThree = true;
    } else if (event.CategoryType == constant.SlidesDisplayName) {
      ChoiceOne = false;
      ChoiceTwo = true;
      ChoiceThree = true;
    } else {
      ChoiceOne = true;
      ChoiceTwo = true;
      ChoiceThree = false;
    }
    this.setState({
      renderKey: event.key,
      radioBtnCategory: event.CategoryType,
      disabledChoiceOne: ChoiceOne,
      disabledChoiceTwo: ChoiceTwo,
      disabledChoiceThree: ChoiceThree
    });
  }

  /**
   * On click catagory refresh button. */
  public refreshCategoryDrpDwn(): void {
    this.setState({
      renderKey: '',
      radioBtnCategory: '',
      disabledChoiceOne: false,
      disabledChoiceTwo: false,
      disabledChoiceThree: false
    });
  }

  /**
   * Refreshing sorting in whenever click on any manage asset sections.
   * @param newColumns 
   */
  public refreshColumnHeaderSorting(newColumns): any {
    newColumns.forEach((newCol: IColumn) => {
      newCol.isSorted = false;
      newCol.isSortedDescending = false;
    });
    return newColumns;
  }

  /**
   *Await refresh category list data in add new panel on changing  manage asset section. */
  public async _refreshCategories(sectionSelected) {
    var mainArray = await DC._getDocsNodeCategoriesName(sectionSelected);
    this._CategoriesDrpDwnItemArray = mainArray.DocsNodeCategoriesArrayItems;
    this._CategoriesItemArray = [{ key: 0, text: constant.AllSlideName }].concat(this._CategoriesDrpDwnItemArray);
    this._CategoriesDrpDwnItemArray = [{ key: 'Header', text: 'Select Category', itemType: DropdownMenuItemType.Header }].concat(this._CategoriesDrpDwnItemArray);
  }

  /**
   * Rendering dropdown options in panel. */
  private _onRenderOption = (options): JSX.Element => {
    return (
      <div>
        {options.data && (options.data == true) ? (<span>
          <Icon style={{ marginRight: '20px' }} />
          -- {options.text}</span>
        ) : <span style={{ fontWeight: 900 }}>{options.text}</span>}
      </div>
    );
  }

  /**
   * Show selected section in Managed Assest menu. */
  public _onLinkClickAssets = (event, item) => {
    if (item.name == constant.SlidesDisplayName) {
      this._getDocsNodeSlidesItems();
    } else if (item.name == constant.ImagesDisplayName) {
      this._getDocsNodeImageItems();
    } else if (item.name == constant.TextSnippetDisplayName) {
      this._getDocsNodeTextSnippetItems();
    } else if (item.name == constant.CategoryDisplayName) {
      this._getDocsNodeCategoryItems();
    } else if (item.name == constant.PlaceHolderDisplayName) {
      this._getDocsNodePlaceHolderItems();
    } else {
      this._getConfigurationListItems();
    }
  }

  /**
   * Navigation view for catagory (DEMO). */
  public LinkClickInner = (event, item) => {
    this.setState({
      selectKeyInner: item.key
    });
  }

  /**
   *Hide panel. */
  private _hidePanel = (): void => {
    this._ParentCategoriesItemArray = [{ key: 'Header', text: 'Select Parent Category', itemType: DropdownMenuItemType.Header }].concat(this._ParentCategItemArray);
    this.setState({
      showPanel: false,
      addOrEditBtn: false
    });
  }

  /**
   *Input text change  in Title text. */
  public _handleChangeTitle = (event) => {
    this.setState({
      itemTitle: event,
    });
  }

  /**
   *Input text change in Discription text. */
  public _handleChangeDiscript = (event) => {
    this.setState({
      itemDiscription: event,
    });
  }

  /**
   *Input text change  in category text. */
  public _handleChangeCategoryTitle = (event) => {
    this.setState({
      newCategoryTitle: event,
    });
  }

  /**
   *Input text change  in placeholder text. */
  public _handleChangePlaceHolder = (event) => {
    this.setState({
      newPlaceHolder: event,
    });
  }

  /**
   * Input text change in textsnippet text. */
  public _handleChangeTextSnippet = (event) => {
    this.setState({
      itemTextSnippet: event
    });
  }

  private _onTextChange = (newText: string) => {
    return newText;
  }

  /**
   *Show dialog box for validation of deleting item. */
  private _showDialog(item: any) {
    this.setState({
      hideDialog: false,
      deleteItem: item
    });
  }

  /**
   *Close the dialog box. */
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  /**
   * For DEMO purpose. */
  private _showInnerDialog() {
    this.setState({
      innerDialogbox: false
    });
  }

  /**
   * Demo purpose. */
  private _closeInnerDialog = (): void => {
    this.setState({ innerDialogbox: true });
  }

  /**
   *Search the item from Displaying items List Component. */
  private _onSearchText = (text: string): void => {
    var searchArry = [];
    if (this.state.tabValue == constant.CategoryDisplayName) {
      this._SearchResult.map((i) => {
        if ((i.Category != undefined ? i.Category.toLowerCase().indexOf(text.toLowerCase()) > -1 : false) || (i.ParentCategory != undefined ? i.ParentCategory.toLowerCase().indexOf(text.toLowerCase()) > -1 : false) || i.CategoryType.toLowerCase().indexOf(text.toLowerCase()) > -1) {
          searchArry.push(i);
        }
      });

    } else if (this.state.tabValue == constant.PlaceHolderDisplayName) {
      this._SearchResult.map((i) => {
        if ((i.PlaceholderName != undefined ? i.PlaceholderName.toLowerCase().indexOf(text.toLowerCase()) > -1 : false)) {
          searchArry.push(i);
        }
      });
    } else {
      this._SearchResult.map((i) => {
        if ((i.Title != undefined ? i.Title.toLowerCase().indexOf(text.toLowerCase()) > -1 : false) || (i.Category != undefined ? i.Category.toLowerCase().indexOf(text.toLowerCase()) > -1 : false) || i.FileName.toLowerCase().indexOf(text.toLowerCase()) > -1) {
          searchArry.push(i);
        }
      });
    }
    //Binding the search data
    this.setState({
      items: searchArry,
      searchValue: text
    });
  }

  /**
   * On change catagory radio button changes.
   * @param event 
   * @param option 
   */
  private _onChangeCategoryType(event, option: any) {
    this.setState({
      radioBtnCategory: option.key
    });
  }

  /**
   *Following function is use to add new item or edit the existing item. */
  public async _addNewData() {

    const {
      renderKey,
      listName,
      addOrEditBtn,
      itemId,
      radioBtnCategory,
      itemName,
      tabValue,
      itemTitle,
      itemDiscription,
      newPlaceHolder,
      newCategoryTitle,
      SiteColectionValue,
      SubSiteCollectionKey,
      ListKey,
      FieldName
    } = this.state;
    var titleValue = '';
    var discriptionValue = '';
    var textsnippet = '';
    var newFile = '';
    var ParentLevel = 0;
    var sitecollurl = '';
    var subsiteurl = '';
    var listurl = '';
    var listfield = '';
    if (listName == constant.DocsNodeCategoriesName) { //Category
      titleValue = newCategoryTitle;
      titleValue = titleValue != null ? titleValue.trim() : '';
      if (titleValue != '' && radioBtnCategory != '') {
        if (renderKey != '' && renderKey != null) {
          ParentLevel = ParentLevel + 1;
        }
        await this.setState({
          progressLevel: false,
          disbledSaveBtn: true
        });
        //Add and Edit list Item
        DC._updateListItem(titleValue, textsnippet, discriptionValue, renderKey, listName, addOrEditBtn, itemId, radioBtnCategory, ParentLevel).then((dataResult) => {

          //Rebind the data
          this._rebindData();
        });
      } else {
        alert('Fill required fields!!');
      }
    } else if (listName == constant.DocsNodeSlidesName || listName == constant.DocsNodePictureName) { //Slide OR Image
      titleValue = itemTitle;
      discriptionValue = itemDiscription;
      titleValue = titleValue != null ? titleValue.trim() : '';
      discriptionValue = discriptionValue != null ? discriptionValue.trim() : '';
      if (addOrEditBtn) {
        if (titleValue != '') {
          await this.setState({
            progressLevel: false,
            disbledSaveBtn: true
          });
          //Edit library item
          DC._uploadFiles('', titleValue, discriptionValue, renderKey, itemName, this.props.context, listName).then((dataResult) => {

            //Rebind the data
            this._rebindData();
          });
        } else {
          alert('Fill required fields!!');
        }
      } else {
        //Add library Item         
        var counter = 0;
        var fileArray = [];
        if (tabValue == constant.SlidesDisplayName) { //Slide
          newFile = document.getElementById('inputTypeFiles')['files'];
          //Validation for Picture Library only PPT files are uploaded
          if (titleValue != '' && newFile.length > 0) {
            for (var k = 0; k < newFile.length; k++) {
              var newSlideFileData = newFile[k];
              if (newSlideFileData['name'].match('%') || newSlideFileData['name'].match('"') || newSlideFileData['name'].match('\'')) {
                alert(newSlideFileData['name'] + "\nFileName should not contains (Single Quote or Percent) special characters.");
              } else {
                //Validation for Slide Library only pptx files are uploaded          
                if (newSlideFileData != undefined ? newSlideFileData['name'].indexOf('.pptx') > -1 : false) {
                  fileArray.push(newSlideFileData);
                  counter++;
                } else {
                  alert(newSlideFileData['name'] + '\nPlease upload valid presentation file.');
                }
              }
            }
            if (counter > 0) {
              //open progressBar
              await this.setState({
                progressLevel: false,
                disbledSaveBtn: true
              });
              for (var m = 0; m < fileArray.length; m++) {
                //Upload new file
                await DC._uploadFiles(fileArray[m], titleValue, discriptionValue, renderKey, '', this.props.context, listName);
                if (m == fileArray.length - 1) {
                  //Rebind the data
                  this._rebindData();
                }
              }
            }
          } else {
            alert('Fill required fields!!');
          }
        } else { //Image
          newFile = document.getElementById('inputTypeFiles')['files'];
          //Validation for Picture Library only image files are uploaded
          if (itemTitle.length > 0 && newFile.length > 0) {
            for (var j = 0; j < newFile.length; j++) {
              var newFileData = newFile[j];
              if (newFileData['name'].match('%') || newFileData['name'].match('"') || newFileData['name'].match('\'')) {
                alert(newFileData['name'] + "\nImage FileName should not contains (Single Quote or Percent) special characters.");
              } else {
                if (newFileData != undefined ? newFileData['name'].indexOf('.png') > -1 || newFileData['name'].indexOf('.jpeg') > -1 || newFileData['name'].indexOf('.jpg') > -1 || newFileData['name'].indexOf('.gif') > -1 || newFileData['name'].indexOf('.bmp') > -1 : false) {
                  fileArray.push(newFileData);
                  counter++;
                } else {
                  alert(newFileData['name'] + '\nPlease upload image with valid format.');
                }
              }
            }
          } else {
            alert('Fill required fields!!');
          }
          if (counter > 0) {
            //open progressBar
            await this.setState({
              progressLevel: false,
              disbledSaveBtn: true
            });
            for (var i = 0; i < fileArray.length; i++) {
              //Upload new file
              await DC._uploadFiles(fileArray[i], titleValue, discriptionValue, renderKey, '', this.props.context, listName);

              if (i == fileArray.length - 1) {
                //Rebind the data
                this._rebindData();
              }
            }
          }
        }
      }
    } else if (listName == constant.DocsNodePlaceHolderName) { //Placeholder
      titleValue = newPlaceHolder;
      discriptionValue = itemDiscription;
      titleValue = titleValue != null ? titleValue.trim() : '';
      discriptionValue = discriptionValue != null ? discriptionValue.trim() : '';
    
      sitecollurl = this._SiteColectionValue;
      subsiteurl = this._Subsiteurl;
      listurl = this._ListName;
      listfield = this._FieldName;
      if (titleValue != '' && sitecollurl != '' && listurl != '' && listfield != '') {
        //Check Placeholder exist or not
        DC._placeHolderExistOrNot(titleValue).then(async (data) => {
          if (data != 'Exist') {
            await this.setState({
              progressLevel: false,
              disbledSaveBtn: true
            });
            //Add and Edit list Item
            DC._updateListItem(titleValue, sitecollurl, subsiteurl, discriptionValue, listName, addOrEditBtn, itemId, listurl, listfield).then((dataResult) => {

              //Rebind the data
              this._rebindData();
            });
          } else {
            alert('Placeholder name already exist!!');
          }
        });
      } else {
        alert('Fill required fields!!');
      }
    } else { //Text Snippet
      titleValue = itemTitle;
      discriptionValue = itemDiscription;
      titleValue = titleValue != null ? titleValue.trim() : '';
      discriptionValue = discriptionValue != null ? discriptionValue.trim() : '';
      textsnippet = document.getElementsByClassName("ql-editor")[0].innerHTML;
      var validContain = CU.extractContent(textsnippet);
      validContain = validContain.trim();
      textsnippet = validContain != '' ? textsnippet : '';
      if (titleValue != '' && textsnippet != '') {
        await this.setState({
          progressLevel: false,
          disbledSaveBtn: true
        });
        //Add and Edit list Item
        DC._updateListItem(titleValue, textsnippet, discriptionValue, renderKey, listName, addOrEditBtn, itemId, radioBtnCategory, ParentLevel).then((dataResult) => {

          //Rebind the data
          this._rebindData();
        });
      } else {
        alert('Fill required fields!!');
      }
    }
  }

  /**
   *This function is use to get item from list and library which is going to be edited. */
  public _editItem(item: any) {

    const { listName } = this.state;
    var filterArray = [];
    this._ParentCategItemArray.filter((itemFilter) => {
      if (itemFilter.CategoryType === item.CategoryType) {
        filterArray.push(itemFilter);
      }
    });
    this._ParentCategoriesItemArray = [{ key: 'Header', text: 'Select Parent Category', itemType: DropdownMenuItemType.Header }].concat(filterArray);

    //Calling the item to get edited
    DC._getLibraryItemToEdit(item, listName).then((itemData) => {
      var ChoiceOne = false;
      var ChoiceTwo = false;
      var ChoiceThree = false;
      var parentCatgy = false;
      if (itemData[0].CategoryType == constant.ImagesDisplayName) {
        ChoiceOne = true;
        ChoiceTwo = false;
        ChoiceThree = true;
      } else if (itemData[0].CategoryType == constant.SlidesDisplayName) {
        ChoiceOne = false;
        ChoiceTwo = true;
        ChoiceThree = true;
      } else {
        ChoiceOne = true;
        ChoiceTwo = true;
        ChoiceThree = false;
      }
      if (itemData[0].CategoryKey == null) {
        parentCatgy = true;
      } else {
        parentCatgy = false;
      }
      this.setState({
        itemTitle: itemData[0].Title,
        newCategoryTitle: itemData[0].Name,
        itemDiscription: itemData[0].Discription,
        itemTextSnippet: itemData[0].TextSnippet,
        itemName: itemData[0].Name,
        itemId: item.Key,
        radioBtnCategory: itemData[0].CategoryType,
        showPanel: true,
        renderKey: itemData[0].CategoryKey,
        addOrEditBtn: true,
        disabledChoiceOne: ChoiceOne,
        disabledChoiceTwo: ChoiceTwo,
        disabledChoiceThree: ChoiceThree,
        disabledParentCatgy: parentCatgy,
        newPlaceHolder: itemData[0].PlaceHolderName,
        SiteColectionValue: itemData[0].SiteCollUrl,
        SubSiteCollectionKey: itemData[0].SubsiteUrl,
        ListKey: itemData[0].ListUrl,
        FieldName: itemData[0].Listfield
      });
    });
  }

  /**
   *This function is use to delete item from List and Library. */
  public async _deleteItem(item: any) {

    const { listName } = this.state;
    //Close Dialog
    await this._closeDialog();

    //Calling the api to delete item
    DC._deleteListItem(item, listName).then((itemData) => {
      if (itemData == 'success') {
        //Rebind the data
        this._rebindData();
      } else {
        alert(itemData);
      }
    });
  }

  /**
   * This function is use to add new data in Configuration List. */
  public async _addConfigData() {

    const {
      tempLibSelectKey,
      imgLibSelectKey,
      imgLogoID } = this.state;

    var tmpLibName = '';
    var imgLibName = '';
    var newFile = '';
    tmpLibName = tempLibSelectKey;
    imgLibName = imgLibSelectKey;
    newFile = document.getElementById('productLogo')['files'][0];
    if (tmpLibName != '' && imgLibName != '') {
      await this.setState({
        progressLevel: false,
        disbledSaveBtn: true
      });
      if (newFile != undefined) {
        if (newFile['name'].match('%') || newFile['name'].match('"') || newFile['name'].match('\'')) {
          alert(newFile['name'] + "\nImage FileName should not contains (Single Quote or Percent) special characters.");
        } else {
          if (newFile != undefined ? newFile['name'].indexOf('.png') > -1 || newFile['name'].indexOf('.jpeg') > -1 || newFile['name'].indexOf('.jpg') > -1 : false) {
            await this.setState({
              imgLogoURl: ''
            });
            await DC._deleteProductLogo(imgLogoID);
            //Upload custom product logo to site assest list
            await DC._uploadProductLogo(newFile, this.props.context).then((response) => {
            });
          } else {
            alert(newFile['name'] + '\nPlease upload image with valid format.');
          }
        }
      }

      let newConfigArrData = [];
      newConfigArrData.push({ assetTitle: constant.TemplateTitle, sourceLocation: tmpLibName, sourceListURL: '', sourceListGUID: '' });
      newConfigArrData.push({ assetTitle: constant.ImageTitle, sourceLocation: imgLibName, sourceListURL: '', sourceListGUID: '' });

      //Add New items to configuration list
      DC._addNewConfigurationListData(newConfigArrData).then(async (responseData) => {

        //Update the new date from configuration list and get in webpart to access new list and library
        await DC._dynamicListOrLibraryArray();

        await this.setState({
          tempLibSelectKey: '',
          imgLibSelectKey: '',
        });

        this._getConfigurationListItems();

        //Success message
        toast.success('Success', {
          position: "top-right",
          autoClose: 2000,
          hideProgressBar: true,
          closeOnClick: false,
          pauseOnHover: false,
          draggable: false,
        });
      });
    } else {
      alert('Select required fields!!');
    }
  }

  /**
   *This function is use call for rebinding the data after add,edit and delete any edit. */
  public async _rebindData() {
    //Close progressBar
    await this.setState({
      progressLevel: true,
      showPanel: false,
      disbledSaveBtn: false
    });

    if (this.state.tabValue == constant.SlidesDisplayName) {
      this._getDocsNodeSlidesItems();
    } else if (this.state.tabValue == constant.ImagesDisplayName) {
      this._getDocsNodeImageItems();
    } else if (this.state.tabValue == constant.TextSnippetDisplayName) {
      this._getDocsNodeTextSnippetItems();
    } else if (this.state.tabValue == constant.CategoryDisplayName) {
      this._getDocsNodeCategoryItems();
    } else if (this.state.tabValue == constant.PlaceHolderDisplayName) {
      this._getDocsNodePlaceHolderItems();
    } else {
      this._getConfigurationListItems();
    }
  }

  /**
   *Following function is use to load the data  synchronously(Lazy loading). */
  private _loadData = (): void => {

    //Set time inteval to display data
    this._lastIntervalId = setInterval(() => {
      if (this._lastIndexWithData < this.state.items.length) {
        const randomQuantity: number = 10;
        const itemsCopy = this.state.items!.slice(0);
        itemsCopy.splice(
          this._lastIndexWithData,
          randomQuantity,
          ...this._items.slice(this._lastIndexWithData, this._lastIndexWithData + randomQuantity)
        );
        this._lastIndexWithData += randomQuantity;
        this.setState({
          items: itemsCopy
        });
      }
      else {
        clearInterval(this._lastIntervalId);
      }
    }, _INTERVAL_DELAY);

  }

  /**
   *This function is use to divide N number of data in parts to display in List Component. */
  private _onLoadData = (itemLength, itemData): void => {

    this._ITEMS_COUNT = itemLength;
    this._items = itemData;

    let items = [];
    if (this._ITEMS_COUNT > 10) {
      //dividing data ramdomly 
      const randomQuantity: number = 10;
      items = this._items.slice(0, randomQuantity).concat(new Array(this._ITEMS_COUNT - randomQuantity));
      this._lastIndexWithData = randomQuantity;

      //Lazy Loading
      this._loadData();
      this.setState({
        isDataLoaded: true,
        items: items
      });
    } else {
      this.setState({
        isDataLoaded: true,
        items: this._items
      });
    }

  }

  /**
   *This function is use for sorting data on clicking header of the column. */
  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const { columns, items, categoryColumns, tabValue, placeholderColumns } = this.state;
    var newColumns: IColumn[] = [];
    if (tabValue == constant.CategoryDisplayName) {
      newColumns = categoryColumns.slice();
    } else if (tabValue == constant.PlaceHolderDisplayName) {
      newColumns = placeholderColumns.slice();
    } else {
      newColumns = columns.slice();
    }
    const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    //reorganize the data in acending or decending
    const newItems = this._copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    if (tabValue == constant.CategoryDisplayName) {
      this.setState({
        categoryColumns: newColumns,
        items: newItems
      });
    } else if (tabValue == constant.PlaceHolderDisplayName) {
      this.setState({
        placeholderColumns: newColumns,
        items: newItems
      });
    } else {
      this.setState({
        columns: newColumns,
        items: newItems
      });
    }
  }

  /**
   *This function is use to acending or decending the data. */
  public _copyAndSort<T>(items, columnKey: string, isSortedDescending?: boolean) {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a, b) => ((isSortedDescending ? (a[key] != '' ? a[key].toLowerCase() : a[key]) < (b[key] != '' ? b[key].toLowerCase() : b[key]) : (a[key] != '' ? a[key].toLowerCase() : a[key]) > (b[key] != '' ? b[key].toLowerCase() : b[key])) ? 1 : -1));
  }

  /**
   *This function is use for selection N Number of rows. */
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).Title;
      default:
        return `${selectionCount} items selected`;
    }
  }

  /**
   *This function is use to show error message if user didn't input text in Title of Choose file to upload. */
  private _getErrorMessage = (value: string): string => {
    return value.length > 0 ? '' : 'Required *';
  }

  /**
   * On click Placeholder section bind site collection. */
  public _placeHolderAddClick() {
    this._siteCollitem = [];
    DC._getAllSiteCollections().then(data => {
      if (data.length > 0) {
        for (var i = 0; i < data.length; i++) {
          this._siteCollitem.push({ siteUrl: data[i].Path, siteTitle: data[i].Title, siteKey: data[i].Title + "_" + i, ServerRelativeUrl: data[i].Path, siteId: '' });
        }
      }
    });
  }


  public bindRootSiteCollections(RootSiteCollectionData, FavoriteOrNot) {
    var self = this;
    if (RootSiteCollectionData != null && RootSiteCollectionData.length > 0) {
      var treeViewHTML = "<ul id='treeviewUL'>";

      RootSiteCollectionData.map( (sitecollection, index) => {
        var siteUrl = sitecollection.siteUrl;
        var siteTitle = sitecollection.siteTitle;
        var siteKey = sitecollection.siteKey;
        var siteId = sitecollection.siteId;
        var siteTitleSD = siteTitle + "_SD";

        treeViewHTML += " <li id='" + siteTitle + "'><span class='" + styles.caret + " " + styles["caret-down"] + " treeSpan' ><div class='type' hidden>sitecollection</div><div class='level' hidden> " + 0 + "</div><div class='sitekey' hidden> " + siteKey + "</div><div class='siteurl' hidden> " + siteUrl + "</div><div class='sitetitle' hidden> " + siteTitle + "</div> <div class='siteId' hidden> " + siteId + "</div> <a href='#'><i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i>" + siteTitle + "</a></span > "; // main site collection li level -1 open
        treeViewHTML += " <ul class='" + styles.active + " " + styles.nested + "' id='" + siteKey + "'>"; // ul level - 2 open
        treeViewHTML += " </ul>"; // Site List
        treeViewHTML += " </li>"; // Site  
      });
      treeViewHTML += "</ul>"; //treeviewUL
      if (FavoriteOrNot == 'AllSiteCollections') {
        $('#SPTreeView').html(treeViewHTML);
      }
      else {
        $('#SPTreeViewFavoruite').html(treeViewHTML);
      }
      $(".treeSpan").off('click');
      $(".treeSpan").on('click', function () {
        self.clickSpan(this);
        self._getSubsites(this);
      });
    }
  }

  public clickSpan(e) {
    $("span").removeClass(styles.treeselected);
    $(e).addClass(styles.treeselected);
    if (e.children[0].textContent=="field")
    {
      this._FieldName = e.children[3].textContent.trim(); // [3] internal name,  [4] display name
    }
  }

  public getFields(e) {
    var self = this;

    if ($(e).siblings("." + styles.nested).length > 0) {
      var docurl = '';
      var appwebURL = '';
      var selectedLibURL = '';
      var docLibKey = '';
      var tenantURL = '';
      var docName = '';
      var listdisplayName = '';

      $("span").removeClass(styles.treeselected);
      $(e).addClass(styles.treeselected);
      docurl = $(e).find(".appWebUrl").text().trim();
      appwebURL = $(e).find(".appWebUrl").text().trim();
      selectedLibURL = $(e).find('.selectedLibURL').text().trim();
      docLibKey = $(e).find('.docLibKey').text();
      tenantURL = CU.tenantURL();
      docName = $(e).find(".listName").text().trim();
      listdisplayName= $(e).find(".listdisplayName").text().trim();

      if (docurl.toLowerCase() == tenantURL.toLowerCase()) {
        docurl = "/";
      }
      else {
        docurl = docurl.substring(tenantURL.length, docurl.length);
      }

      this._SiteColectionValue = appwebURL;
      this._Subsiteurl = appwebURL.replace(tenantURL, "");
      this._ListName = docName;

      DC._getfields(docurl, listdisplayName,docName).then((fieldItems) => {
        try {
          if (fieldItems.length > 0) {

            var liHTML = "";
            for (var i = 0; i < fieldItems.length; i++) {

              var fieldName = fieldItems[i].internalName.trim();
              var fieldDisplayName = fieldItems[i].text.trim();
              var fieldKey = fieldName + "_" + i;
              var appWebUrl = fieldItems[i].internalName.trim();

              liHTML += " <li id='" + fieldKey + "'><span class='treeSpanList "+styles.fieldCaret+"'><div class='type' hidden>field</div><div class='siteURL' hidden> "+appWebUrl+" </div><div class='folderKey' hidden> " + fieldKey + "</div><div class='fieldName' hidden> " + fieldName + "</div><a href='#'><i class='ms-Icon ms-Icon--ColumnOptions' aria-hidden='true'></i> " + fieldDisplayName + "</a></span>"; // Sub Site 1 - Shared Documents 1 //level - 3 open/close
              liHTML += " <ul class='active nested' id='" + fieldKey + "'>"; // field ul level - open 
              liHTML += " </ul>"; // field - ul - close
              liHTML += " </li>";
            }

            fieldItems = [];

            if (docLibKey != "") {
              var skey = docLibKey.trim();
              $("ul[id='" + skey + "']").html(liHTML);
              $("ul[id='" + skey + "']").removeClass(styles.nested);
              $("ul[id='" + skey + "']").prev('span').removeClass(styles["caret-down"]);

              $(document).on("click", ".treeSpanList", function (event) {
                self.clickSpan(this);
              });

            }
          }
        } catch (error) {
          console.log("_createDropDownOptions: " + error);
        }
      });
    }
    else {

      var docListKey = $(e).find('.docLibKey').text().trim();

      if (docListKey != "") {
        $("ul[id='" + docListKey + "']").addClass(styles.nested);
        $("ul[id='" + docListKey + "']").prev('span').addClass(styles["caret-down"]);
      }
    }
  }
  
  /*** Get all SubSite of site collection ***/
  public _getSubsites(e) {
    var self = this;
    if ($(e).siblings("." + styles.nested).length > 0) {

      $("span").removeClass(styles.treeselected);
      $(e).addClass(styles.treeselected);
      var tenantURL = CU.tenantURL();
      var siteUrl = $(e).find('.siteurl').text().trim();
      var siteTitle = $(e).find('.sitetitle').text().trim();
      var siteKeyRoot = $(e).find('.sitekey').text().trim();
      var siteId = $(e).find('.siteId').text().trim();
      var level = $(e).find('.level').text().trim();
      var tenantName = tenantURL.substr(8, tenantURL.length);
      var selectedSiteURL = siteUrl.split("/sites/")[1];
      var cutmAttr = "";

      if (siteUrl.indexOf("/sites/") < 0 && tenantName != siteUrl.substr(8, tenantURL.length)) {
        selectedSiteURL = siteUrl.substring(siteUrl.lastIndexOf("/") + 1, siteUrl.length);
        cutmAttr = "rootsubsites";
      }

      var rootSiteURL = siteUrl.split(tenantName)[1];
      var selectedSiteURLForLib = siteUrl.split("/sites/")[1] == undefined ? rootSiteURL : selectedSiteURL;
      this.isRoot = siteUrl.indexOf("/sites/") < 0 ? true : false;
      level = (parseInt(level) + 1).toString();
      try {
        let url = siteUrl + "/_api/Web/GetSubwebsFilteredForCurrentUser(nWebTemplateFilter=-1)?$filter=(WebTemplate ne 'APP')";
        CU._getRequest(url).then(async (data) => {
          var result = data.d.results;
          self.listAllSubsite = [];
          if (result && result.length > 0) {
            for (var i = 0; i < result.length; i++) {
              if (result[i].ServerRelativeUrl != 0) {
                var webURL = tenantURL + result[i].ServerRelativeUrl;
                if (webURL.indexOf(tenantName) > -1) {
                  var rootsubSite = webURL.split(tenantName)[1] + ":";
                  var subsitesVar = webURL.split("/sites/")[1] == undefined ? rootsubSite : webURL.split("/sites/")[1];
                  if (self.isRoot) {
                    selectedSiteURL = "Root";
                  }
                  var SiteKeyCurrent = result[i].Title + "_" + i;
                  var siteID = result[i].Id;
                  self.listAllSubsite.push({ "ParentSite": selectedSiteURL, "ParentSiteURL": siteUrl, "hasSubSite": true, "SubSiteDisplayName": result[i].Title, "SubSiteName": result[i].Title, "SubSiteURL": webURL, "IsRoot": self.isRoot, "parentSiteKey": siteKeyRoot, "SiteKey": SiteKeyCurrent, "level": level, "siteId": siteID });
                }
              }
            }
            await self._getList(siteUrl, siteKeyRoot, SiteKeyCurrent).then((responseDocLibrary) => {

              var liHTML = "";
              var parentsiteKey = "";

              if (self.listAllSubsite.length > 0) {

                parentsiteKey = self.listAllSubsite[0].parentSiteKey;

                for (var j = 0; j < self.listAllSubsite.length; j++) {

                  if (self.listAllSubsite[j].hasSubSite) {

                    var siteName = self.listAllSubsite[j].SubSiteName.trim();
                    var siteURL = self.listAllSubsite[j].SubSiteURL.trim();
                    var siteKey = self.listAllSubsite[j].SiteKey;
                    var sitesId = self.listAllSubsite[j].siteId;
                    var ParentSiteURL = self.listAllSubsite[j].ParentSiteURL;
                    if (self.listAllSubsite[j].caretSubsite == false && self.listAllSubsite[j].caretDocumentLibrary == false) {
                      liHTML += " <li id='" + siteName + "'><span class='treeSpanSite " + styles.fieldCaret + "' ><div class='type' hidden>site</div><div class='level' hidden> " + level + "</div><div class='sitekey' hidden> " + siteKey + "</div><div class='siteurl' hidden> " + siteURL + "</div><div class='sitetitle' hidden> " + siteName + "</div><div class='siteId' hidden> " + sitesId + "</div> <a href='#'><i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i>" + siteName + "</a></span > "; // subsite level li - 2 open
                    }
                    else {
                      liHTML += " <li id='" + siteName + "'><span class='" + styles.caret + " " + styles["caret-down"] + " treeSpanSite' ><div class='type' hidden>site</div><div class='level' hidden> " + level + "</div><div class='sitekey' hidden> " + siteKey + "</div><div class='siteurl' hidden> " + siteURL + "</div><div class='sitetitle' hidden> " + siteName + "</div><div class='siteId' hidden> " + sitesId + "</div> <a href='#'><i class='ms-Icon ms-Icon--SharepointLogoInverse' aria-hidden='true'></i>" + siteName + "</a></span > "; // subsite level li - 2 open
                    }

                    liHTML += " <ul class='" + styles.active + " " + styles.nested + "' id='" + siteKey + "'>"; // Sub Site 1 ul level - 3 open  
                    liHTML += " </ul>"; // Sub Site 1 - List - ul level - 3 close
                    liHTML += " </li>"; // Sub Site 1 li   level - 2 close 
                  }
                }
                if (responseDocLibrary) {
                  liHTML = responseDocLibrary + liHTML;
                }
                else {
                  liHTML = liHTML;
                }

                if (parentsiteKey != "") {
                  var skey = parentsiteKey.trim();
                  $("ul[id='" + skey + "']").html(liHTML);
                  $("ul[id='" + skey + "']").removeClass(styles.nested);
                  $("ul[id='" + skey + "']").prev('span').removeClass(styles["caret-down"]);

                  $(document).off("click", ".treeSpanSite");
                  $(document).on("click", ".treeSpanSite", function (event) {
                    self.clickSpan(this);
                    self._getSubsites($(this));
                  });

                  $(document).off("click", ".treeList");
                  $(document).on("click", ".treeList", function (event) {
                    self.clickSpan(this);
                    self.getFields($(this));
                  });

                }
              }
            });
          }
          else {
            await self._getList(siteUrl, siteKeyRoot, SiteKeyCurrent).then((responseDocLibrary) => {
              var liHTML = "";
              var parentsiteKey = "";
              if (responseDocLibrary) {
                liHTML = responseDocLibrary + liHTML;
              }
              else {
                liHTML = liHTML;
              }
              parentsiteKey = siteKeyRoot;
              if (parentsiteKey != "") {
                var skey = parentsiteKey.trim();
                $("ul[id='" + skey + "']").html(liHTML);
                $("ul[id='" + skey + "']").removeClass(styles.nested);
                $("ul[id='" + skey + "']").prev('span').removeClass(styles["caret-down"]);

                $(document).off("click", ".treeSpanSite");
                $(document).on("click", ".treeSpanSite", function (event) {
                  self._getSubsites($(this));
                });

                $(document).off("click", ".treeList");
                $(document).on("click", ".treeList", function (event) {
                  self.getFields($(this));
                });

              }
            });

          }
        });
      }
      catch (error) {
        console.log("getSubsites: " + error);
      }
    }
    else {

      var docLibKey = $(e).find('.sitekey').text().trim();

      if (docLibKey != "") {
        $("ul[id='" + docLibKey + "']").addClass(styles.nested);
        $("ul[id='" + docLibKey + "']").prev('span').addClass(styles["caret-down"]);
        $(e).removeClass(styles.treeselected);
      }
    }
  }

  /**
   *Get all list from site collection. */
  public async _getList(site, parentsiteKey, siteKey) {
    //Get all list call
    return await DC._getAllList(site, 100).then(async doclibdata => {
      var self = this;
      var DDoptions = [];
      var DDOptionTreeView = [];
      try {
        if (doclibdata.length > 0) {
          //Binding all list data
          DDoptions = DC._bindingAllList(doclibdata);
          if (DDoptions.length > 0) {
            for (var i = 0; i < DDoptions.length; i++) {
              var docURL = site + "/" + DDoptions[i].key;
              var webURL = docURL.substring(0, docURL.lastIndexOf("/"));
              var docLibKey = DDoptions[i].key + "_" + i;

              DDOptionTreeView.push({
                "site": site, "appWebUrl": webURL, "siteURL": docURL, "displayName": DDoptions[i].text, "name": DDoptions[i].key, "parentsiteKey": parentsiteKey, "siteKey": siteKey, "docLibKey": docLibKey
              });
            }
          }

          if (DDOptionTreeView) {
            var liHTML = "";
            for (var j = 0; j < DDOptionTreeView.length; j++) {
              liHTML += " <li id='" + DDOptionTreeView[j].docLibKey + "'><span class='" + styles.caret + " " + styles["caret-down"] + " treeList' ><div class='type' hidden>doclib</div><div class='appWebUrl' hidden> " + DDOptionTreeView[j].appWebUrl + "</div><div class='listName' hidden> " + DDOptionTreeView[j].name + "</div><div class='listdisplayName' hidden> " + DDOptionTreeView[j].displayName + "</div><div class='selectedLibURL' hidden> " + DDOptionTreeView[j].siteURL + "</div><div class='docLibKey' hidden> " + DDOptionTreeView[j].docLibKey + "</div><a href='#'><i class='ms-Icon ms-Icon--FabricDocLibrary' aria-hidden='true'></i> " + DDOptionTreeView[j].displayName + "</a></span>"; // main site collection list li level - 2 open/close                        
              liHTML += " <ul class='" + styles.active + " " + styles.nested + "' id='" + DDOptionTreeView[j].docLibKey + "'>";
              liHTML += " </ul>";
              liHTML += " </li>";
            }
            DDOptionTreeView = [];
            return liHTML;
          }
        }
      } catch (error) {
        console.log("_getList: " + error);
      }
    });
  }

  /**
   *Change event of template library dropdown  */
  public _OnChangeTemplateLib(event) {
    this.setState({
      tempLibSelectKey: event.key
    });
  }

  /**
   *Change event of Slide library dropdown  */
  public _OnChangeSlideLib(event) {
    this.setState({
      slideLibSelectKey: event.key
    });
  }

  /**
   *Change event of Image library dropdown  */
  public _OnChangeImageLib(event) {
    this.setState({
      imgLibSelectKey: event.key
    });
  }

  /**
   *Change event of Text Snippet List dropdown  */
  public _OnChangeTextSnippetList(event) {
    this.setState({
      textSniSelectKey: event.key
    });
  }

  private _onCheckboxChange(e: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    var itemName = e.currentTarget.getAttribute('aria-label');
    var slideName = '';
    var imgName = '';
    var txtSniName = '';
    var slideIsChecked;
    var imgIsChecked;
    var textIsChecked;
    if (isChecked) {
      if (itemName == constant.SlidesDisplayName) {
        slideName = itemName;
        slideIsChecked = isChecked;
      } else if (itemName == constant.ImagesDisplayName) {
        imgName = itemName;
        imgIsChecked = isChecked;
      } else {
        txtSniName = itemName;
        textIsChecked = isChecked;
      }
    } else {
      if (itemName == constant.SlidesDisplayName) {
        slideName = '';
        slideIsChecked = isChecked;
      } else if (itemName == constant.ImagesDisplayName) {
        imgName = '';
        imgIsChecked = isChecked;
      } else {
        txtSniName = '';
        textIsChecked = isChecked;
      }
    }
    this.setState({
      slideCheckBoxName: slideName,
      imgCheckBoxName: imgName,
      textSniCheckBox: txtSniName,
      slideChecked: slideIsChecked,
      imgChecked: imgIsChecked,
      textSniChecked: textIsChecked,
    });
  }

  /**
   *This method is for rendering result. */
  public render() {

    //Creating constants for state varaibles
    const {
      columns,
      categoryColumns,
      placeholderColumns,
      items,
      selectionDetails,
      itemTitle,
      searchValue,
      itemDiscription,
      itemTextSnippet,
      isModalSelection,
      isDataLoaded,
      CategaryDropDown,
      defaultKey,
      showPanel,
      defaultNavSelectdKey,
      itemName,
      newPlaceHolder,
      addOrEditBtn,
      renderKey,
      showList,
      showcategoryListForm,
      showPlaceHolderList,
      deleteItem,
      hideDialog,
      hideSiteAdminDialog,
      newCategoryTitle,
      radioBtnCategory,
      data,
      tabValue,
      disabledChoiceOne,
      disabledChoiceTwo,
      disabledChoiceThree,
      disabledParentCatgy,
      showConfigList,
      progressLevel,
      sitecoll,
      subSiteItem,
      listItem,
      fieldItem,
      SiteColectionValue,
      SubSiteCollectionKey,
      ListKey,
      FieldName,
      disabledsubsite,
      disableddoclib,
      disabledfolder,
      configImgLibraryOptons,
      configSlideLibraryOption,
      configTempLibraryOption,
      configTextSnipListOption,
      tempLibSelectKey,
      slideLibSelectKey,
      imgLibSelectKey,
      textSniSelectKey,
      disbledSaveBtn,
      imgLogoURl
    } = this.state;

    //Will return the UI of webpart
    return (<div>
      <div className={styles.topNav}>
        <div className={styles.container}>
          <div className={styles.ProductIcon}>
            <img src={String(require('../images/logo.png'))} />
          </div>
          <div className={styles.productName}>
            <span>DocsNode Templates</span>
          </div>
        </div>
        {<Dialog
          hidden={hideSiteAdminDialog}
          dialogContentProps={{
            type: DialogType.normal,
            subText: "You are not Admin. You don't have access to use DocsNode Template Admin."
          }}
        >
        </Dialog>}
      </div>
      <div className={styles.container}>
        <ToastContainer position={toast.POSITION.TOP_LEFT} />
        <div className={styles.leftNav}>
          <div className={styles.manageAssets}>
            <Nav
              selectedKey={defaultNavSelectdKey}
              onLinkClick={this._onLinkClickAssets}
              groups={[
                {
                  name: 'Manage Configuration',
                  links: [
                    { name: constant.ConfigurationDisplayName, url: '', key: 'key6', iconProps: { iconName: 'List' } },
                  ]
                },
                {
                  name: 'Manage Assets',
                  links: [
                    { name: constant.SlidesDisplayName, key: 'key1', url: '', iconProps: { iconName: 'Boards' } },
                    { name: constant.ImagesDisplayName, key: 'key2', url: '', iconProps: { iconName: 'FileImage' } },
                    { name: constant.TextSnippetDisplayName, key: 'key3', url: '', iconProps: { iconName: 'ContextMenu' } },
                    { name: constant.PlaceHolderDisplayName, key: 'key4', url: '', iconProps: { iconName: 'InsertTextBox' } }
                  ]
                },
                {
                  name: 'Manage Assets Categories',
                  links: [
                    { name: constant.CategoryDisplayName, url: '', key: 'key5', iconProps: { iconName: 'List' } },
                  ]
                }
              ]}
            />
          </div>
        </div>
        <div className={styles.rightContent}>
          {!showConfigList && <div className={styles.ribbonTop}>
            <div className={styles.addNewBtn}>
              <CommandBarButton
                iconProps={{ iconName: 'Add' }}
                text="Add New"
                onClick={this._showPanel}
              />
            </div>
            <div className={styles.searchBox}>
              <SearchBox
                placeholder="Search"
                onEscape={ev => {

                }}
                value={searchValue}
                onChange={newValue => this._onSearchText(newValue)}
              />
            </div>
          </div>}
          {showList && <div>
            <ShimmeredDetailsList
              setKey="items"
              items={items!}
              columns={columns}
              selectionMode={SelectionMode.none}
              enableShimmer={!isDataLoaded}
              listProps={{ renderedWindowsAhead: 0, renderedWindowsBehind: 0 }}
              isHeaderVisible={true}
              selection={this._selection}
              selectionPreservedOnEmptyClick={true}
              enterModalSelectionOnTouch={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            />
          </div>}
          {showcategoryListForm &&
            <div>
              <ShimmeredDetailsList
                setKey="items"
                items={items!}
                columns={categoryColumns}
                selectionMode={SelectionMode.none}
                enableShimmer={!isDataLoaded}
                listProps={{ renderedWindowsAhead: 0, renderedWindowsBehind: 0 }}
                isHeaderVisible={true}
                selection={this._selection}
              />
            </div>}
          {showPlaceHolderList &&
            <div>
              <ShimmeredDetailsList
                setKey="items"
                items={items!}
                columns={placeholderColumns}
                selectionMode={SelectionMode.none}
                enableShimmer={!isDataLoaded}
                listProps={{ renderedWindowsAhead: 0, renderedWindowsBehind: 0 }}
                isHeaderVisible={true}
                selection={this._selection}
              />
            </div>}
          {showConfigList && <div>
            <ProgressIndicator progressHidden={progressLevel} />
            <div className={styles.drpdwnConfigMain}>
              <Dropdown
                placeHolder="Select options"
                label={constant.TemplateTitle + ' :'}
                selectedKey={tempLibSelectKey}
                className={styles.drpdwnConfig}
                options={configTempLibraryOption}
                onChanged={this._OnChangeTemplateLib}
                required
              />
            </div>
            <div className={styles.drpdwnConfigMain}>
              <Dropdown
                placeHolder="Select options"
                label={constant.ImageTitle + ' :'}
                className={styles.drpdwnConfig}
                selectedKey={imgLibSelectKey}
                options={configImgLibraryOptons}
                onChanged={this._OnChangeImageLib}
                required
              />
            </div>
            <div className={styles.drpdwnConfigMain + " " + styles.chooseFileMc}>
              <Label>Product Logo :</Label>
              <div className={styles.imgDivClass}>
                <img src={imgLogoURl} style={{ width: '60px' }}></img>
              </div>
              <TextField
                accept=".png,.jpeg,.jpg"
                type="file"
                id="productLogo"
                value=''
                className={styles.imgText}
              />
            </div>
            <div className={styles.drpdwnConfigMain + " " + styles.mc_saveBtn}>
              <Label></Label>
              <PrimaryButton text='Save' disabled={disbledSaveBtn} onClick={this._addConfigData} />
            </div>
          </div>}
          <Dialog
            hidden={hideDialog}
            onDismiss={this._closeDialog}
            dialogContentProps={{
              type: DialogType.normal,
              title: deleteItem.FileName,
              subText: 'Are you sure you want to delete this item?'
            }}
          >
            <DialogFooter>
              <PrimaryButton className={styles.footerbtn} onClick={() => this._deleteItem(deleteItem)} text="Delete" />
              <DefaultButton className={styles.footerbtn} onClick={this._closeDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        </div>
      </div>
      <div id="mySidenavAdd" className={styles.sidenav}>
        <div className={styles.headerSidenav}>
          <Panel
            isOpen={showPanel}
            onDismiss={this._hidePanel}
            type={PanelType.medium}
            headerText={addOrEditBtn == true ? itemName : 'New Item'}
            onRenderFooterContent={this._onRenderFooterContent}
            isFooterAtBottom={true}
          >
            <ProgressIndicator progressHidden={progressLevel} />
            <div className={styles.bodySidenav}>
              {showList && <div className={styles.NewIteamForm}>
                {this.state.tabValue != constant.TextSnippetDisplayName ? addOrEditBtn == true ? null :
                  <div className={styles.chooseFile}>
                    <Label required>Choose file:</Label>
                    <TextField
                      className={styles.FormInput}
                      accept=".ppt,.pptx,image/*"
                      type="file"
                      id='inputTypeFiles'
                      multiple
                      onGetErrorMessage={this._getErrorMessage}
                      validateOnLoad={false} />
                  </div> : null}
                <div className={styles.FormInput}>
                  <Label required>Title:</Label>
                  <TextField id='titleID' placeholder="Enter value here" underlined
                    value={itemTitle}
                    onChanged={this._handleChangeTitle}
                    onGetErrorMessage={this._getErrorMessage}
                    validateOnLoad={false}
                  />
                </div>
                <div className={styles.FormInput}>
                  <Label>Description:</Label>
                  <TextField
                    id='discription'
                    multiline
                    autoAdjustHeight
                    placeholder="Enter value here"
                    underlined value={itemDiscription}
                    onChanged={this._handleChangeDiscript} />
                </div>
                {tabValue == constant.TextSnippetDisplayName && <div className={styles.FormInput}>
                  <Label required>{constant.TextSnippetDisplayName}:</Label>
                  <RichText value={itemTextSnippet}
                    className={styles.richText}
                    onChange={(text) => this._onTextChange(text)}
                  />
                </div>}
                <div className={styles.FormInput}>
                  <Label>Category:</Label>
                  <Dropdown
                    placeHolder="Select options"
                    selectedKey={renderKey}
                    onRenderOption={this._onRenderOption}
                    options={this._CategoriesDrpDwnItemArray}
                    onChanged={this._SelectCategory}
                    style={{ width: '95%', float: 'left', marginBottom: '30px' }}
                  />
                  <IconButton
                    iconProps={{ iconName: 'Refresh' }}
                    style={{ width: '5%' }}
                    onClick={this.refreshCategoryDrpDwn}
                  />
                </div>
              </div>}
              {showcategoryListForm &&
                <div className={styles.NewIteamForm}>
                  <div className={styles.FormInput}>
                    <Label required >Category:</Label>
                    <TextField id='newCategoryName' placeholder="Enter new category here" underlined
                      value={newCategoryTitle}
                      onChanged={this._handleChangeCategoryTitle}
                      onGetErrorMessage={this._getErrorMessage}
                      validateOnLoad={false}
                    />
                  </div>
                  <div className={styles.FormInput}>
                    <Label>Parent Category (Optional):</Label>
                    <Dropdown
                      placeHolder="Select options"
                      selectedKey={renderKey}
                      options={this._ParentCategoriesItemArray}
                      onChanged={this._SelectCategory}
                      style={{ width: '95%', float: 'left', marginBottom: '30px' }}
                      disabled={disabledParentCatgy}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Refresh' }}
                      style={{ width: '5%' }}
                      onClick={this.refreshCategoryDrpDwn}
                      disabled={disabledParentCatgy}
                    />
                  </div>
                  <div className={styles.FormInput}>
                    <ChoiceGroup
                      selectedKey={radioBtnCategory}
                      options={[
                        {
                          key: constant.SlidesDisplayName,
                          text: constant.SlidesDisplayName,
                          disabled: disabledChoiceOne
                        },
                        {
                          key: constant.ImagesDisplayName,
                          text: constant.ImagesDisplayName,
                          disabled: disabledChoiceTwo
                        },
                        {
                          key: constant.TextSnippetDisplayName,
                          text: constant.TextSnippetDisplayName,
                          disabled: disabledChoiceThree
                        }
                      ]}
                      onChange={this._onChangeCategoryType}
                      label="Category Type:"
                      required={true}
                    />
                  </div>
                </div>}
              {showPlaceHolderList && <div className={styles.placeholder_maindiv}>
                <div className={styles.customCommonDivClass}>
                  <Label required>Placeholder Name:</Label>
                  <TextField id='newPlaceHolder' placeholder="Enter new placeholder" underlined
                    value={newPlaceHolder}
                    onChanged={this._handleChangePlaceHolder}
                  />
                </div>
                <div className={styles.customCommonDivClass}>
                  <Label>Description:</Label>
                  <TextField
                    id='placeholderdiscription'
                    multiline
                    autoAdjustHeight
                    placeholder="Enter value here"
                    underlined value={itemDiscription}
                    onChanged={this._handleChangeDiscript} />
                </div>
                {addOrEditBtn == true ? null : <div>
                  <div className={styles.customCommonDivClass}>
                    <div>
                      <Label>All Locations:</Label>
                    </div>
                    <div className={styles.categoriesTree} id="SPTreeView">
                    </div>
                  </div>
                </div>}
              </div>}
            </div>
          </Panel>
        </div>
      </div>
    </div>
    );
  }

  /**
   *Rendering Save and cancel button at footer in panel bar. */
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton className={styles.footerbtn} disabled={this.state.disbledSaveBtn} onClick={this._addNewData}>Save</PrimaryButton>
        <DefaultButton className={styles.footerbtn} onClick={this._hidePanel}>Cancel</DefaultButton>
      </div>
    );
  }
}
