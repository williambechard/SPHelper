
# SPHelper
Javascript library to assist with common/useful SharePoint and REST List functions

> Written with [StackEdit](https://stackedit.io/) 
> - Auhtor: William Bechard
> - Contact: ury2ok2000@gmail.com | [william.bechard@us.af.mil](william.bechard@us.af.mil)
> 
## Requirements

 - JQuery

## How To Use
In SharePoint you will want to store this .js file in a document library. Then you just reference it like any other JavaScript library reference:

    <script type="text/javascript"  src="\sites\mySite\SiteAssets\Scripts\SPHelper.js"></script>
   
# Notes

 > - for **siteURL**, if the List is on the same site as your script, rather than using a URL you can use `_spPageContextInfo.webAbsoluteUrl`
 > - Additionally the functions will pull all data from a list, it uses a loop to get past the usual **SharePoint item limit**

# Use
Each function returns a promise. As such the syntax for using any of the functions will be very similar.

    var myPromise = functionName();
    myPromise.done(
         function(data){
	         // data is the return result of the promise
	         // depending on the function it can be an
	         // array of objets or just one, so you may
	         // need to use console.log to determine the 
	         // result format. In general it will be data.d.results
	         console.log(data);
         }
     );

# Functions

### Get All List Items:
    SPHGetListItems(ListName, additionalParameters, siteUrl)
#### Use: 
         var allListItems = SPHGetListItems('list Name', '', siteURL);
         allListItems.done(function(data){
              console.log(data); //data.d.results might work as well 
         }).fail(
             function(e){ 
                  alert('Failed Getting List Items!' + e.responseText);
             }
           );
**additionalParameters** can be anything you want to use in the AJAX call. Commonly this will begin with `?$select=`   and then the field names you want from the list. This is very useful for only **selecting** the fields you want, **filtering**, and **expanding** fields (like the people picker field type). Further options can be found at: https://social.technet.microsoft.com/wiki/contents/articles/31995.sharepoint-2013-working-with-rest-api-using-jquery-ajax.aspx

### Update List Item

    SPHUpdateItemByID(ListTitle, id, itemProperties, siteURL, contentType)
> - id: the ID of the item you want to update
> - contentType: use the function SPHGetItemType('list name')

### Create List Item

    SPHCreateListItem(listName, itemProperties, siteURL, contentType)
> the main difference in this function is itemProperties. This needs to be an object and should be setup like the below example:

    var itemProperties = {};
    itemProperties['Property Name']="some Value";
> also there is a function UpdateFormDigest() in this call that is used so that if the user is Idle before this function is called it will re-ask for credentials so that the list can create the item.

### Get List Item

    SPHGetListItemByID(ListTitle, id, siteUrl)
> no different than the Get All List Items function, except it requires the ID of the item you want back (and of course therefore only returns 1 item). 

### Delete List

    SPHDeleteList(ListTitle, siteURL)

### Delete List Item

    SPHDeleteListItem(ListTitle, id, siteUrl)
    
  > Same as Delete List, except it requires the ID of the item you want Deleted

### Get User

    SPHGetUser(UserID)
> - Requires the ID of a user (as a string). If you are after the logged in user, you can simply pass `_spPageContextInfo.userId.toString()`
> - With the return value of this promise you can extrapolate a lot of info about the user, including **Title, First Name, Last Name, email** etc.

### What Groups a Logged in User is a Part Of

    IsCurrentUserMemberOfGroup(grpNames)
>  - **grpNames**: array of strings of all the groups you want to check
   against

> - the Return  of this promise is either **true** (the member is a part of one of the group names that were passed in the array), or **false**. 


### Get All List Columns (Fields)

    SPHGetListColumns(ListName, siteUrl)
> Does not show Hidden or Read-Only fields

### Get All Lists Under a Site

    SPHGetAllSiteLists(siteUrl)
 
 ### Open SP Modal
 

    openSPModal(URL, modalTitle, modalWidth, modalHeight, refreshOnSave)
   
   > **Note** this is different from the rest of the functions in that it doesnt return a promise. It is simply called like a normal function.
   > 
   > Its purpose is to open a modal window to a SharePoint List (**NewForm/ EditForm/DispForm**)
   > - **URL**: The full URL of the list ASPX that you want to open in the modal (...NewForm.aspx, ...DispForm.aspx?ID=, ...EditForm.aspx?ID=)
   > - **modalTitle**: What we want the modal to show at the top Title bar
   > - **modalWidth**: width of modal window
   > - **modalHeight**: height of modal window
   > - **refreshOnSave**: **true** or **false** (if true page Refreshes after save is clicked)
