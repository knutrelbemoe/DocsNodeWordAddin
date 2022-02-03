export default class constant {

  //Evenyone Username 
  public static readonly everyOneLoginName = 'c:0(.s|true';
  public static readonly newSiteGroupName = 'DocsNode Members';

  //Site Configuration
  public static readonly SiteCollectionNameColumn = 'SiteCollectionName';
  public static readonly SiteCollectionUrlColumn = 'SiteCollectionURL';
  public static readonly SubSiteNameColumn = 'SubSiteName';
  public static readonly subsiteUrlColumn = 'SubSiteURL';
  public static readonly DocumentLibraryNameColumn = 'DocumentLibrary';
  public static readonly DocumentLibraryURLColumn = 'DocumentLibraryURL';
  public static readonly FolderNameColumn = 'FolderName';
  public static readonly FolderURLColumn = 'FolderURL';

  //Category
  public static readonly CategoryName = 'Category';
  public static readonly CategoryParentId = 'ParentCategory';
  public static readonly CategoryType = 'CategoryType';
  public static readonly CategoryLevel = 'CategoryLevel';

  //Slide
  public static readonly SlidesCategoryName = 'SlidesCategory';
  public static readonly SlidesDiscriptionName = 'SlidesDiscription';

  //Image
  public static readonly ImageCategoryName = 'ImageCategory';

  //Text Snippet
  public static readonly TextSnippetName = 'TextSnippet';
  public static readonly TextSnippetDiscriptionName = 'TSDiscription';
  public static readonly TextCategoryName = 'TextCategory';

  //Placeholder
  public static readonly PlaceHolderName = 'Placeholder';
  public static readonly PlaceHolderDiscrip = 'PlaceholderDiscription';
  public static readonly SiteCollectionUrlName = 'SiteCollectionUrl';
  public static readonly SubSiteUrlName = 'SubSiteUrl';
  public static readonly ListUrlName = 'ListUrl';
  public static readonly ListFieldName = 'ListField';

  //Configuration
  public static readonly ConfigAssetTitleName = 'ConfigAssestTitle';
  public static readonly ConfigSourceListName = 'ConfigSourceList';
  public static readonly ConfigSourceListPathUrl = 'ConfigSourceListPath';
  public static readonly ConfigSourceListGUID = 'ConfigSourceListGUID';
  public static readonly TemplateTitle = 'Template Library';
  public static readonly ImageTitle = 'Images Library';
  public static readonly SlideTitle = 'Slides Library';
  public static readonly TextSnippetTitle = 'Text Snippet List';
  public static readonly CategoryTitle = 'Category List';

  //Create Custom List amd Library
  public static readonly DocsNodeCategoriesName = 'DocsNodeCategories';
  public static readonly DocsNodeTextName = 'DocsNodeText';
  public static readonly DocsNodeSlidesName = 'DocsNodeSlides';
  public static readonly DocsNodePlaceHolderName = 'DocsNodePlaceHolder';
  public static readonly DocsNodeConfigurationName = 'DocsNodeConfiguration';
  public static DocsNodeTemplatesLibraryName = '';
  public static DocsNodePictureName = '';
  public static readonly DocsNodeProductLogoName = 'DocsNodeProductLogo';
  public static readonly DefaultDocsNodeTemplatesLibraryName = 'DocsNodeTemplatesLibrary';
  public static readonly DefaultDocsNodePictureName = 'DocsNodePicture';
  public static readonly DocsNodeSiteConfiguratonName = 'DocsNodeSiteConfiguration';

  //Manage Assets
  public static readonly SlidesDisplayName = 'Slides';
  public static readonly ImagesDisplayName = 'Images';
  public static readonly TextSnippetDisplayName = 'Text Snippet';
  public static readonly PlaceHolderDisplayName = 'Placeholder';
  public static readonly CategoryDisplayName = 'Categories';
  public static readonly ConfigurationDisplayName = 'Configuration';

  public static readonly AllSlideName = 'All Slides';

  //Default Attributes
  public static readonly NewList = 'List';
  public static readonly NewLibrary = 'Library';
  public static readonly Title = 'Title';

  //Document or Image File Name 
  public static readonly InternalLinkFilename = 'LinkFilename';

  //Create New Team Site
  //public static readonly SiteCollDisplayName = 'DocsNode Admins';
  //public static readonly SiteCollUrlName = 'DocsNodeAdmins123';  
  public static readonly SiteCollDisplayName = 'DocsNode Admin';
  public static readonly SiteCollUrlName = 'DocsNodeAdmin';


  public static postBody = {
    alias: constant.SiteCollUrlName,
    displayName: constant.SiteCollDisplayName,
    isPublic: false,
    optionalParams: {
      Classification: '',
      Description: 'DocsNode Templates Admin Panel',
      Owners: [],
    }
  };
}