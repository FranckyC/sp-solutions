# SharePoint Framework modern search Web Parts

![Version](https://img.shields.io/badge/version-3.1.1-green.svg)

## Summary
This solution allows you to build user friendly SharePoint search experiences using SPFx in the modern interface. The main features include:

- Fully customizable SharePoint search query like the good old Content Search Web Part.
- Can either use a static query or be connected to a search box component using SPFx dynamic data.
- Live templating system with Handlebar to meet your requirements in terms of UI + builtin list and tiles templates. Can alos use template from an external file.
- Search results includings previews for Office documents and Office 365 videos.
- Customizable refiners supporting multilingual values for taxonomy based filters.
- Sortable results (unique field).
- Refiners Web Part.
- Pagination Web Part.
- SharePoint best bets support.
- Search query enhancement with NLP tools (like Microsoft LUIS).

A complete 1 hour turorial video is available on the [official SharePoint Developer blog](https://developer.microsoft.com/en-us/sharepoint/blogs/pnp-webcast-sharepoint-framework-modern-search-web-part/):

<a href="https://www.youtube.com/watch?v=g41nvRVwtds" target="_blank"><img src="http://img.youtube.com/vi/g41nvRVwtds/maxresdefault.jpg"/></a>

<p align="center">
  <img src="./images/react-search-refiners.gif"/>
</p>

This sample includes the following components and service(s):

**Web Part(s)**

Component | Description 
----- | ----- 
Search Box Web Part | Allows users to enter free text/KQL search queries connected to a search results Web Part.
Search Results Web Part | Performs static or dynamic search query with customizable parameters sorting and templating. An associated [blog post](http://thecollaborationcorner.com/2017/10/16/build-dynamic-sharepoint-search-experiences-with-refiners-and-paging-with-spfx-office-ui-fabric-and-pnp-js-library/) is available to give you more details about this Web Part implementation.
Search Refiners | Allows users to configure refiners for a search results Web Part.
Search Pagination | Allows users to configure pagination for a search results Web Part.
Search Navigation | Allows users to configure navigation for a search results Web Part

**Back-end service(s)**

Service | Description 
----- | ----- 
Search Query Enhancer | Sample Azure function to demonstrate the use of Microsoft LUIS and other cognitive services to interpret user intents and enhance the search box query accordingly.

## Used SharePoint Framework Version 
![drop](https://img.shields.io/badge/drop-1.7.1-green.svg)

## Applies to

* [SharePoint Framework](https:/dev.office.com/sharepoint)
* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)

## Solution

Solution| Contributors(s)
--------|---------
ModernSearch | Franck Cornu (aequos) - [@FranckCornu](http://www.twitter.com/FranckCornu)<br/>Tarald Gåsbakk (Puzzlepart AS) - [@taraldgasbakk](https://twitter.com/Taraldgasbakk)<br/>Yannick Reekmans - [@yannickreekmans](https://twitter.com/yannickreekmans)<br/>Mikael Svenson (Puzzlepart AS) - [@mikaelsvenson](http://www.twitter.com/mikaelsvenson)

## Web Parts Configuration

### Search Box Web Part

<p align="center"><img width="300px" src="./images/sb_property_pane.png"/><p>

#### Default Search Query Settings 

Setting | Description 
-------|----
Use a dynamic data source | You can set a default query text coming from am other data source. This case is particularly useful when you want to put a search box Web Part on the front page redirecting to an other page with the same query. Use the query string parameter 'q' from the builtin 'Page Environment' data source.

#### Search box options

Setting | Description 
-------|----
Enable query suggestions | The search box supports query suggestions from SharePoint. Refer to the following [article](https://docs.microsoft.com/en-us/sharepoint/search/manage-query-suggestions) to know how to add query suggestions in your SharePoint tenant (caution: it can take up to 24h for changes to take effect).
Send the query to a new page | Sends the search query text to a new page. On that page, use an other search box Web Part configured with a dynamic data source as the default query. This Web Part uses the URL fragment '#' to pass the entered keywords. Make sure you use this data source property in your targeted components to retrieve the query.
Placeholder text | The placeholder text to display in the search box.

#### Search query enhancement

Setting | Description 
-------|----
Use Natural Language Processing service | Turn this option 'on' if you want to enhance the query text with NLP services like LUIS. In the _'Service Url'_ field, enter the URL of the Azure Function endpoint. Refer the instructions in the `'/functions/README.md'` file to set up the service. In this sample, only relevant detected keywords are returned as q new query using LUIS. Enabling debug mode will show you relevant information about the entered query.

---

### Search Results Web Part

<table>
  <tr>
    <td>
      <p align="center"><img src="./images/property_pane.png"/><p>
    </td>
    <td>
      <p align="center"><img src="./images/property_pane2.png"/><p>
    </td>
    <td>
      <p align="center"><img src="./images/property_pane3.png"/><p>
    </td>
  </tr>
<table>

#### Search Query Configuration 

Setting | Description 
-------|----
Search query keywords | Here you choose to use a static search query or a query coming from a data source. It is recommended to use the associated Web Part coming with this sample. The search query is in KQL format so you can use search query variables (See this [post](http://www.techmikael.com/2015/07/sharepoint-rest-do-support-query.html) to know which ones are allowed). You can only plug one source to this Web Part.

<p align="center"><img width="300px" src="./images/wp_connection.png"/><p>

#### Search Settings

Setting | Description 
-------|----
Query template | The search query template in KQL format. You can use search variables here (like Path:{Site}).
Result Source Identifier | The GUID of a SharePoint result source.
Initial sort order | The initial search results sort order. You can use mutliple properties here.
Sortable fields | The search managed properties to use for sorting. Make sure these are sortable. With SharePoint Online, you have to reuse the default ones to do so (RefinableStringXX etc.). The order is the same as they will appear in the sort panel. You can also provide your own custom labels using the following format RefinableString01:"You custom filter label",RefinableString02:"You custom filter label",... If no sortable fields are provided, the 'Sort' button will not be visible.
Connect to a search refiners Web Part | If enable, select the search refiners Web Part to use on the current page to get selected filters. It is a 2 ways connection so don't forget to connect the targeted search refiners to the search results Web Part as well. 
Enable Query Rules | Enable the query rules if applies. Turn this options  'on' to display your SharePoint Promoted results (links only).
Selected properties | The search managed properties to retrieve. You can use these properties then in your Handlebar template with the syntax (`item.property_name`).
Number of items to retrieve per page | Quite explicit. The paging behavior is done directly by the search API (See the *SearchDataProvider.ts* file), not by the code on post-render.

#### Styling Options

Setting | Description 
-------|----
Web Part Title | Shows a title for this Web Part. Set blank if you don't want a title.
Show blank if no result | Shows nothing if there is no result
Show result count | Shows the result count and entered keywords  
Connect to a search pagination Web Part	 | If enable, select the search pagination Web Part to use on the current page to get selected page. It is a 2 ways connection so don't forget to connect the targeted search pagination to the search results Web Part as well. 
Result Layouts options | Choose the template to use to display search results. Some layouts are defined by default (List and Tiles) but you can create your own either by clinkg on the **"Custom"** tile, or **"Edit template"** from an existing chosen template. In custom mode, you can set an external template. It has to be in the same SharePoint tenant. Behind the scenes, the Office UI Fabric core CSS components are used in a isolated way. Custom code templates will also automaticly be displayed here upon registration.
Result types | Allows you to set a custom template at item level according to a specific condition (ex: FileType equals 'pdf').
Handlebars Helpers | Load [handlebar helpers](https://github.com/helpers/handlebars-helpers) to use in your template. Disable this option will make Web Part loading faster if you don't need them.

##### Miscellaneous: Taxonomy values dynamic translation

The search results Web Part supports automatic translation for taxonomy based filters and result metadata according to current site language. To get it work, you must map a new refinable managed property associated with *ows_taxId_<your_column_name>* crawled property and turn on the *'Localization Enabled'* toggle in the search results property pane:

<p align="center">
  <img src="./images/managed-property.png"/>
</p>

<p align="center">
  <img src="./images/enable_localization.png"/>
</p>

---

#### Templates with Handlebars

This Web Part allows you change customize the way you display your search results. The templating feature comes directly from the original [react-content-query-webpart](https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/react-content-query-webpart) so thanks to @spplante!

<p align="center">
  <img width="500px" src="./images/edit_template.png"/>
</p>

##### Available tokens

Setting | Description 
-------|----
`{{showResultsCount}}` | Boolean flag corresponding to the associated in the property pane.
`{{totalRows}}` | The result count.
`{{maxResultsCount}}` | The number of results configured to retrieve in the web part.
`{{actualResultsCount}}` | The actual number of results retreived.
`{{keywords}}` | The search query.
`{{getSummary HitHighlightedSummary}}` | Format the *HitHighlightedSummary* property with recognized words in bold.
`{{getDate <date_managed_property> "<format>}}"` | Format the date with moment.ts according to the current language.
`{{getPreviewSrc item}}` | Determine the image thumbnail URL if applicable.
`{{getUrl item}}` | Get the item URL. For a document, it means the URL to the Office Online instance or the direct URL (to download it).
`{{getUrlField managed_propertyOWSURLH "URL/Title"}}` | Return the URL or Title part of a URL field managed property.
`{{getCountMessage totalRows <?keywords>}}` | Display a friendly message displaying the result and the entered keywords.
`{{<search_managed_property_name>}}` | Any valid search managed property returned in the results set. These are typically managed properties set in the *"Selected properties"* setting in the property pane. You don't need to prefix them with `item.` if you are in the "each" loop.
`{{webUrl}}` | The current web relative url. Use `{{../webUrl}}` inside a loop.
`{{siteUrl}}` | The current site relative url. Use `{{../siteUrl}}` inside a loop.
`{{getUniqueCount items "property"}}` | Get the unique count of a property over the result set (or another array)
`{{getUniqueCount array}}` | Get the unique count of objects in an array. Example: [1,1,1,2,2,4] would return `3`.

Also the [Handlebars helpers](https://github.com/helpers/handlebars-helpers) (188 helpers) are also available. You can also define your own in the *BaseTemplateService.ts* file. See [helper-moment](https://github.com/helpers/helper-moment) for date samples using moment.

##### Use result types

The result types feature is a convenient way to split your templates according to results characteristics instead of making a huge central template with multiple conditions. They can be defined in 'inline' mode or using an external file. You can use the sorting option to determine to order of evaluation for each condition. 

<p align="center">
  <img width="5700px" src="./images/result_types.png"/>
</p>

The following operators are supported:
- Equals
- Contains
- StartsWith
- Greater Or Equal
- Less Or Equal
- Less than
- Greater than
- Is not null

To use it in your main template, just follow this pattern. This block is not mandatory.

```
{{#> resultTypes}}
  {{!-- The block below will be used as default item template if no result types matched --}}
  <div class="template_result">
      <!-- Your default template markup -->
  </div>
  
{{/resultTypes}}
```

Handlebars [partials](https://handlebarsjs.com/partials.html) are used behind the scenes and conditions are built dynamically using a recursive if/else structure.

#### Custom code renderers
You may also define  your own renderers, which most often should be SPFx application customizers. These should use the resultservice to register themselves as renderers, and will upon registration be available as a rendering choice in the "Result Layouts" Section.

More information about custom code renderers may be found in a [seperate sample](../react-search-refiners-renderer), which showcases such a renderer.

#### Out of the box query variables

The following out of the box [query variables](https://docs.microsoft.com/en-us/sharepoint/technical-reference/query-variables) are supported/tested:

##### Site and site collection properties
|**Query variable**|**Definition**|
|:-----|:-----|
|{Site} or {Site.URL}  <br/> |URL of the site from where the query was issued. For example, this value can be used to query content of the managed property Path.  <br/> |
|{Site.ID}  <br/> |GUID of site from where the query was issued. This value corresponds to the value of the managed property SiteID.  <br/> |
|{Site.LCID}  <br/> |Numeric value of the locale as specified by the Regional Settings in the Site Settings on the Site from where the query was issued.  <br/> |
|{Site.Locale}  <br/> |Language of the Site from where the query was issued in ll-cc format — for example, en-us.  <br/> |
|{Site.\<property\>}  <br/> |Any property from the property bag of the site (SPWeb) from where the query was issued, including custom properties.  <br/> |
|{SiteCollection} or {SiteCollection.URL}  <br/> |URL of site collection from where the query was issued. For example, this value can be used to query content of the managed property Path.  <br/> |
|{SiteCollection.ID}  <br/> |GUID of site collection from where the query was issued.  <br/> |
|{SiteCollection.LCID}  <br/> |Numeric value of the locale as specified by the Regional Settings in the Site Settings on the Site Collection from where the query was issued.  <br/> |
|{SiteCollection.Locale}  <br/> |Language of the Site Collection from where the query was issued in ll-cc format — for example, en-us.  <br/> |
|{SiteCollection.\<property\>}  <br/> |Any property from the property bag of the root site (SPWeb) in the site collection (SPSite) from where the query was issued, including custom properties.  <br/> |

##### URL token and request properties
|**Query variable**|**Definition**|
|:-----|:-----|
|{URLToken.\<integer\>}  <br/> |A value from the URL of a page. The integer represents the position of the value in the URL as counted from right to left. For example, for the page http://www.contoso/audio/mp3/1010101, the query variable {URLToken.1} will query for the last value in the URL, 1010101. The query variable {URLToken.3} will query for the third last property in the URL, audio. You can query for values up to the ninth last position in a URL.  <br/> |
|{Request.\<PropertyName\>}  <br/> |A value from the current http request - for example, {Request.Url}.  <br/> |

##### User properties
|**Query variable**|**Definition**|
|:-----|:-----|
|{User} or {User.Name}  <br/> |Display name of the user who issued the query. For example, this value can be used to query content of the managed property Author.  <br/> |
|{User.Email}  <br/> |Email address of the user who issued the query. For example, this value can be used to query content of the managed property WorkEmail.  <br/> |
|{User.PreferredContentLanguage}  <br/> |Language as specified as Preferred Content Language in the profile of the user who issued the query.  <br/> |
|{User.PreferredDisplayLanguage}  <br/> |Language as specified as Preferred Display Language in the profile of the user who issued the query.  <br/> |
|{User.\<property\>}  <br/> |Any property from the user profile of the user who issued the query — for example, SPS-Interests, including custom properties.  <br/> |
   
##### Other properties
|**Query variable**|**Definition**|
|:-----|:-----|
|{Today+/- \<integer value for number of days\>}  <br/> |A date calculated by adding/subtracting the specified number of days to/from the date when the query is issued. Date format is YYYY-MM-DD. For example, this value can be used to query content of the managed property LastModifiedTime.  <br/> |
|{SearchBoxQuery} or {searchTerms} <br/> |The query value entered into a search box on a page.  <br/> |
|{CurrentDisplayLanguage}  <br/> |The current display language based on MUI in ll-cc format.  <br/> |
|{CurrentDisplayLCID}  <br/> |Numeric value of the current display language based on MUI in ll-cc format.  <br/> |

#### Custom query variables

The following custom query variables are supported:

* {Page.&lt;column&gt;} - where column is the internal name of the column.
   * When used with taxonomy columns, use `{Page.Column.Label}` or `{Page.Column.TermID}`
* {CurrentYear} - Todays's date four digits, 2018
* {CurrentMonth} - Today's month, 1-12
* {CurrentDate} - Today's date, 1-31
   * Sample query template: {searchTerms} Write>2010-{CurrentMonth}-{CurrentDate}

#### Best bets

This WP supports SharePoint best bets via SharePoint query rules:

<p align="center">
  <img width="500px" src="./images/query_rules.png"/>
</p>

<p align="center">
  <img width="500px" src="./images/best_bets.png"/>
</p>

#### Elements previews

Previews are available, **only for the list view**, for Office documents and Office 365 videos (not Microsoft Stream). The embed URL is directly taken from the `ServerRedirectedEmbedURL` managed property retrieved from the search results. 

<p align="center">
  <img width="500px" src="./images/result_preview.png"/>
</p>

The WebPart must have the following selected properties in the configuration to get the preview feature work (they are set by default): 
- ServerRedirectedPreviewURL
- ServerRedirectedURL
- contentclass
- ServerRedirectedEmbedURL
- DefaultEncodingURL

This preview is displayed as an _iframe_ when the user clicks on the corresponding preview image. DOM manipulations occur to add the _iframe_ container dynamically aside with the _<img/>_ container.

### Search Refiners

<p align="center"><img width="300px" src="./images/search-refiners-property-pane.png"/><p>

#### Refiner Options

Setting | Description 
-------|----
Refiners | The search managed properties to use as refiners. Make sure these are refinable. With SharePoint Online, you have to reuse the default ones to do so (RefinableStringXX etc.). The order is the same as they will appear in the refnement panel. You can also provide your own custom labels using the following format RefinableString01:"You custom filter label",RefinableString02:"You custom filter label",... This Web Part supports dynamic translation of taxonomy based refiners with few additional configurations in the connected search results Web Part (see above).
Connect to search results Web Part | The search results Web Part to use on the current page to get filters.

#### Styling Options

Setting | Description 
-------|----
Web Part Title | Shows a title for this Web Part. Set blank if you don't want a title.
Show blank if no result | Shows nothing if there is no filter
Filters layout | Choose the template to use to display filters results.

### Search Pagination

<p align="center"><img width="300px" src="./images/search-pagination-property-pane.png"/><p>

#### Pagination Options

Setting | Description 
-------|----
Connect to search results Web Part | The search results Web Part to use on the current page to get page.

### Search Navigation Webpart
This addition is a pivot-navigation link list that may be used to emulate the navigation in the out-of-the-box search center. You can use this webpart on several different pages and set up a searchbox and refiners webpart to handle different types of result. The webpart will automaticly detect if it is on the right page from the set of navigation nodes, and display it in the preferred color. 
<img src="./images/NavNodesMain.gif">

#### Connect to a source
In order to make the search go between the pages where you have set up searches, use this functionality to connect to a search box on the current site. The query from this searchbox will be passed as a url parameter 'q' when navigating to the next page. This is done in the same manner as when you are connecting the search box webpart to the search refiners webpart.

#### Setup Navigation nodes
This step is used to set up the navigation nodes, and needs to be done on all pages it should be displayed on to work correctly.
<img src="./images/ChangeNode.gif">

#### Setup colors
The you may either use the sitecollections theme colors, or use a custom color you choose in the settings pane.
<img src="./images/ChangeColor.gif">

## Features
This Web Part illustrates the following concepts on top of the SharePoint Framework:

- Build an user friendly search experience on the top of the SharePoint search REST API with paging, refiners and query suggestions using the *@pnp* JavaScript library.
- Use [Handlebars](https://handlebarsjs.com/) to create templates for search results according to your requirements like the good old display templates. 
- Using the SPFx [dynamic data feature](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/dynamic-data) to connect Web Parts.
- Using SharePoint taxonomy using JSOM in SPFx (filter translations)
- Integrate the [@pnp/spfx-property-controls](https://github.com/SharePoint/sp-dev-fx-property-controls) in your solution (*PlaceHolder* control).
- Integrate multiple Office UI Fabric components (DocumentCard, Panel, GroupedList, ...) to fit with the native Office 365 theme.
- Use the React container component approach inspiring by the [react-todo-basic sample](https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/react-todo-basic).

<img src="https://telemetry.sharepointpnp.com/sp-dev-solutions/ModernSearch/react-search-refiners" />

## Version history

Version|Date|Comments
-------|----|--------
1.0 | Oct 17, 2017 | Initial release
1.1 | Jan 03, 2018 | Improvements and updating to SPFx drop 1.4
1.2 | Feb 12, 2018 | Added a search box Web Part + Added a "Result Source Id" and "Enable Query Rules" parameters.
1.3 | Apr1, 2018 | Added the result count + entered keywords option
1.4 | May 10, 2018 | <ul><li>Added the query suggestions feature to the search box Web Part</li><li>Added the automatic translation for taxonomy filter values according to the current site locale.</li> <li>Added the option in the search box Web Part to send the query to an other page</ul>
1.5 | Jul 2, 2018 | <ul><li>Added a templating feature for search results with Handlebars inspired by the [react-content-query-webpart](https://github.com/SharePoint/sp-dev-fx-webparts/tree/master/samples/react-content-query-webpart) sample.</li><li>Upgraded to 1.5.1-plusbeta to use the new SPFx dynamic data feature instead of event aggregator for Web Parts communication.</li> <li>Code refactoring and reorganization.</ul>
2.0.0.5 | Sept 18, 2018 | <ul><li>Upgraded to 1.6.0-plusbeta.</li><li>Added dynamic loading of parts needed in edit mode to reduce web part footprint.</li><li>Added configuration to sort.</li><li>Added option to set web part title.</li><li>Added result count tokens.</li><li>Added toggle to load/use handlebars helpers/moment.</li></ul>
2.1.0.0 | Oct 14, 2018 | <ul><li>Bug fixes ([#641](https://github.com/SharePoint/sp-dev-fx-webparts/issues/641),[#642](https://github.com/SharePoint/sp-dev-fx-webparts/issues/642))</li><li>Added document and Office 365 videos previews for the list template.</li><li>Added SharePoint best bets support.</li></ul>
2.1.1.0 | Oct 30, 2018 | <ul><li>Bug fix for editing custom template.</li><li>Bug fix for dynamic loading of video helper library.</li><li>Added support for Page context query variables.</li><li>Added `getUniqueCount` helper function.</li></ul>
2.1.2.0 | Nov 9, 2018 | <ul><li>Bug fix for IE11.</li><li>Added date query variables.</li><li>Added support for both result source id and query template.</li><li>Added `getUniqueCount` helper function.</li></ul>
2.2.0.0 | Nov 11, 2018 | <ul><li>Upgraded to SPFx 1.7.0</li><li>Added a TypeScript Azure Function to demonstrate NLP processing on search query</li><li>Removed extension data source. Now we use the default SPFx 'Page Environment' data source.</li></ul>
2.2.0.1 | Dec 3, 2018 | <ul><li>Remove switch for handlebar helpers, and instead load helpers if used in the template.</li></ul>
2.3.0.0 | Dec 13, 2018 | <ul><li>Upgraded to @pnp/controls 1.13.0</li><li>Added a result types features</li><li>Fix bug regarding dynamic data source connection</li></ul>
2.4.0.0 | Jan 03, 2019 | Added custom code renderer support.
2.4.0.1 | Jan 07, 2019 | Added backwards compability for older sort configurations, and old empty refiner configurations
2.5.0.1 | Feb 11, 2019 | Downgrade @microsoft/sp-office-ui-fabric-core to v1.6.0 to fix theming
2.5.1.0 | Mar 05, 2019 | Added Search Navigation webpart to the package.
3.0.0.0 | Mar 10, 2019 | Created a dedicated refiners Web Part to connect with a search results like to the classic SharePoint search experience.
3.0.1.0 | Mar 14, 2019 | Created a dedicated pagination Web Part to connect with a search results like to the classic SharePoint search experience. Improved Refienrs WP CSS.
3.0.2.0 | Mar 14, 2019 | Fixed regressions with the paging experience
3.0.3.0 | Mar 16, 2019 | Fixed display of custom renderers, in edit mode
3.0.4.0 | Mar 21, 2019 | Fixed loading of Handlebars helpers when having multiple search parts on a page
3.0.5.0 | Mar 26, 2019 | Fixed recreating SearchService on each render
3.1.0.0 | Mar 30, 2019 | Added date range and multi value refiner templates

## Important notice on upgrading the solution from pre v2.2.0.0
**Due to code restucturing we have hit an edge case which impacts upgrades from previous versions. To solve the issue go to `https://<tenant>.sharepoint.com/sites/<appcatalog>/Lists/ComponentManifests` and remove the entries for SearchBox and Search Results, and then upload the .sppkg for the new release.**

**Next you need to loop over all sites which have the web parts installed, and upgrade the App on those sites. Now the web parts should work on new and existing sites. You may use the PnP command `Update-PnPApp` to update the application.**

**If you have deployed the solution as a tenant wide extension, this should not impact you.**

## Important notice on upgrading the solution from pre v3.0.0.0

**Because this version introduces a new standalone search refiners Web Part, you will have to reconfigure all previous refiners from the search results Web Part to this new Web Part. You can get the previous layout for filters by selecting the 'Panel' layout in the new Web Part property pane.**

## Disclaimer
**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository

### SPFx 
- Go to the [spfx](./spfx) directory
- In the command line run:
  - `npm install`
  - `gulp serve`

### Azure Function (Not mandatory)

- Go to the [functions](./functions) directory
- Follow the README.md file instructions
- Set the correct service URL in the Search Box Web Part