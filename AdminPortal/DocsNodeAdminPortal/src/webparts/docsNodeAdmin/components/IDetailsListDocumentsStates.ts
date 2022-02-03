/**
 *State for DocsNode Temaplates. */
export interface IDetailsListDocumentsState {
  columns: any;
  categoryColumns: any;
  placeholderColumns: any;
  items: IDocument[];
  selectionDetails: string;
  isModalSelection: boolean;
  isCompactMode: boolean;
  CategaryDropDown: any;
  isDataLoaded: boolean;
  searchValue: string;
  tabValue: string;
  defaultKey: string;
  showPanel: boolean;
  defaultNavSelectdKey: string;
  itemDiscription: any;
  itemTextSnippet: any;
  itemTitle: any;
  itemName: any;
  newPlaceHolder: string;
  addOrEditBtn: boolean;
  listName: string;
  itemId: number;
  renderKey: string;
  hideDialog: boolean;
  deleteItem: any;
  showList: boolean;
  showcategoryListForm: boolean;
  showPlaceHolderList: boolean;
  newCategoryTitle: string;
  radioBtnCategory: string;
  data: any;
  disabledChoiceOne: boolean;
  disabledChoiceTwo: boolean;
  disabledChoiceThree: boolean;
  disabledParentCatgy: boolean;
  innerDialogbox: boolean;
  selectKeyInner: string;
  showConfigList: boolean;
  progressLevel: boolean;
  sitecoll: any;
  subSiteItem: any;
  listItem: any;
  fieldItem: any;
  SiteColectionValue: string;
  SubSiteCollectionKey: string;
  ListKey: string;
  FieldName: string;
  disabledsubsite: boolean;
  disableddoclib: boolean;
  disabledfolder: boolean;
  configImgLibraryOptons: any;
  configSlideLibraryOption: any;
  configTempLibraryOption: any;
  configTextSnipListOption: any;
  tempLibSelectKey: string;
  slideLibSelectKey: string;
  imgLibSelectKey: string;
  textSniSelectKey: string;
  disbledSaveBtn: boolean;
  hideSiteAdminDialog: boolean;
  imgLogoURl: string;
  imgLogoID: any;
  slideCheckBoxName: string;
  imgCheckBoxName: string;
  textSniCheckBox: string;
  slideChecked: boolean;
  imgChecked: boolean;
  textSniChecked: boolean;
  FavoriteAndAllSites?: any;
}

/**
 *State for columns of list. */
export interface IDocument {
  Title: string;
  FileName: string;
  PlaceholderName: string;
  Category: string;
  FileType: any;
  LinkURL: string;
  Key: number;
  CategoryType: string;
  ParentCategory: string;
  DeleteEditFlag: boolean;
}