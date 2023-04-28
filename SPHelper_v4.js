
/*
Author: MSgt William Joseph Bechard
Date: 13 Feb 2020
Email: william.bechard@us.af.mil

Requirements:
    Jquery must be loaded on the site

Note: if the script you are using is on the same site as the list, rather than filling out the siteURL variable
      you can instead use _spPageContextInfo.webAbsoluteUrl
*/

function SPHDeleteList_v2(ListTitle, siteURL) {
    var deferred=$.Deferred(); 
    var clientContext = new SP.ClientContext(siteURL);
    var oList = clientContext.get_web().get_lists().getByTitle(ListTitle);


    for(var i = 0; i<= 100; i++){
        //assuming the IDs are available, else need some validation  
        var oListItem = oList.getItemById(i);  
        oListItem.deleteObject();
    }

    clientContext.executeQueryAsync(
        function(){
            deferred.resolve("Done");
        },
        function(sender, args){
            deferred.reject(sender, args.get_message() + '\n' + args.get_stackTrace());
        }
    );

    return deferred.promise();
}
function SPHAllLists(siteURL){
    var deferred = $.Deferred();
    var ctx = new SP.ClientContext(siteURL);
    var web = ctx.get_web();
    var lists = web.get_lists();
    ctx.load(lists); 
  
    ctx.executeQueryAsync(
      function(){deferred.resolve(lists)},
      function(sender, args){deferred.reject(sender, args.get_message());}
    );
    return deferred.promise();
  }
 
function SPHAddFieldToList(listName, f, siteURL) {
    var deferred = $.Deferred(); //create the promise
    var clientContext = new SP.ClientContext(siteURL);
    var oList = clientContext.get_web().get_lists().getByTitle(listName);
    var fldCollection = oList.get_fields();
    console.log('Attempting to add Fields to the %s List', listName);
    // Loop through each field
    $.each(f, function(i,v){
        var Default = '';
        var Choices = [];
        var strField = '<Field ';
        for(const property in v){
          if(property !='Default' && property!='Choice') strField+= property + '="' + v[property] + '" '
          else if(property=='Choice') Choices = v[property];
          else if(property=='Default') Default = '<Default>'+ v[property] + '</Default>';
        }

        strField+='>';

        //add Default (if it exists) and close out the Field tag
        //choices?
        if(Choices.length >0){
            strField+='<CHOICES>';
            Choices.map(function(c){ strField+='<CHOISE>' + c + '</CHOISE>'});
            strField+='</CHOICES>'
        }
        
        strField+=Default+'</Field>';
        

        console.log('The Field XML to add: ', strField);

        var field = fldCollection.addFieldAsXml(strField, true,SP.AddFieldOptions.defaultValue);

        field.update();   
    });

    clientContext.executeQueryAsync(
          function(){deferred.resolve()},
          function(sender, args){deferred.reject(args.get_message() + '\n' + args.get_stackTrace());}
    );  
    return deferred.promise(); //return the promise
}
 

/** Info
 * Summary. Deletes ENTIRE list using AJAX
 * Description. Deletes a list. Be careful with this!!!
 *  @param {string} ListTitle Title of the SP List to delete
 *  @param {string} site URL where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * Use:
 * var deleteList = SPHDeleteList('NameOfListToDelete', siteURL);
 * 
 * deleteList.done(function(data){
 *      //success
 * }).fail(function(e){ alert('Failed Delete List! ' + e.responseText);});
 */
function SPHDeleteList(ListTitle, siteURL){
    var fullUrl = siteURL +"/_api/web/lists/GetByTitle('"+ListTitle+"')/items";
    var deferred=$.Deferred();  
    $.ajax({
        url: fullUrl,
        type: "POST",
        headers: {
            "accept": "application/json;odata=verbose",
            "X-Http-Method":"DELETE",
            "content-type": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "IF-MATCH": "*"
        },
        success: function(data){
            deferred.resolve(data);
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}

function SPHListGUID(listName, siteURL){
    var deferred = $.Deferred();
    var context = new SP.ClientContext(siteURL);
    var web = context.get_web();
    list = web.get_lists().getByTitle(listName);
    context.load(list, 'Id');
    context.executeQueryAsync(
        function(){deferred.resolve(list.get_id())},
        function(sender, args){deferred.reject(sender, args.get_message());}
    );

    return deferred.promise();
}

function createSPList(listName, siteURL) {
    var deferred = $.Deferred();
    var clientContext = new SP.ClientContext(siteURL);
    var oWebsite = clientContext.get_web();
    
    var listCreationInfo = new SP.ListCreationInformation();
    listCreationInfo.set_title(listName);
    listCreationInfo.set_templateType(SP.ListTemplateType.genericList);

    var oList = oWebsite.get_lists().add(listCreationInfo);

    clientContext.load(oList);

    clientContext.executeQueryAsync(
        function(){deferred.resolve(oList.get_id())},
        function(sender, args){deferred.reject(sender, args.get_message());}
    );

    return deferred.promise();
}

/** Info
 * Summary. Deletes only 1 item from list based on the item's ID
 * Description. With the passed List Title and the item ID, uses AJAX call
 *         to delete just that item from the list
 *  @param {string} ListTitle the target name of the SP List
 *  @param {int} id the item id that is to be deleted
 *  @param {string} site URL where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * Use:
 * var itemDelete = SPHDeleteListItem('ListTitle', 'StringID of item to delete', 'siteURL');
 * 
 * itemDelete.success(function(data){
 *      //item successfully deleted
 * }).fail(function(e){ alert('Failed Deleting Item! ' + e.responseText);});
 */
function SPHDeleteListItem(ListTitle, id, siteUrl){
    var fullUrl = siteUrl +"/_api/web/lists/GetByTitle('"+ListTitle+"')/items(" + id.toString() + ")";
    var deferred=$.Deferred();  
    $.ajax({
        url: fullUrl,
        type: "DELETE",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val(),
            "IF-MATCH": "*"
        },
        success: function(data){
            deferred.resolve(data);
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}

/** Function Gets SP User ID of the logged in user

 * How To Use
 * var user = SPHGetUserID(); 
 *     user.done(function(data){
 *          console.log(data.d.results[0]);
 *     }).fail(function(e){ alert('Failed Getting User! ' + e.responseText);}); 
 */
function SPHGetUserID(){
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;
    var deferred=$.Deferred();  
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;
    var fullUrl = siteUrl +"/_api/web/CurrentUser";
    
    return $.ajax({
        url: fullUrl,  
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
        },
        success: function(data){

            deferred.resolve(data);
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}

/** Function returns the lookup result of a SP User ID
 *  
 * @param {string} UserID : The string representation of the SP User ID 
 * How To Use
 * var user = SPHGetUser('112312'); //the string number needs to be of a userId in the sharepoint system
 * user.done(function(data){
 *      console.log(data.d.results[0]);
 * }).fail(function(e){ alert('Failed Getting User! ' + e.responseText);}); 
 * or
 * var user = SPHGetUser(_spPageContextInfo.userId.toString());  //this will use the logged in user id
 * user.done(function(data){
 *      console.log(data.d.results[0]);
 * }).fail(function(e){ alert('Failed Getting User! ' + e.responseText);});
 */
function SPHGetUser(UserID)
{
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;
    var deferred=$.Deferred();  
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;
    var fullUrl = siteUrl +"/_api/Web/SiteUserInfoList/Items?$filter=Id eq " + UserID ;
    
    return $.ajax({
        url: fullUrl,  
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
        },
        success: function(data){

            deferred.resolve(data);
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}


/** Info
 * Summary. Gets all items of a SP List
 * Description. Returns data from the SP List. Also fills SPHListItems with that data.
 *        Data can be retrieved by user with SPHReturnItems()
 * @param {string} ListName The name of the SP List that is targeted
 * @param {string} additionalParameters Additional text for the Ajax call after /items 
 *                 will likely begin with ?$select etc or can be an empty string for returning allItems
 * @param {string} site URL where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * Use:  
 * var allListItems = SPHGetListItems('ListName', '', 'siteUrl');  
 *     for advanced pulling from the list if desired but not necessary
 * allListItems.done(function(data){
 *      console.log(data.d.results);
 * }).fail(function(e){ alert('Failed Getting List Items! ' + e.responseText);});
 */
function SPHGetListItems(ListName, additionalParameters, siteUrl)
{ 
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var fullUrl = siteUrl + "/_api/web/lists/GetByTitle('" + ListName + "')/items" + additionalParameters;
    var deferred=$.Deferred();
    var nonext = false;
    var results = [];
    getItems();

    function getItems(){ 
        $.ajax({
            url: fullUrl,
            type: "GET",
            headers: {
                "accept": "application/json;odata=verbose",
                "content-type": "application/json;odata=verbose",
            },
            success: function(data){
                results = results.concat(data.d.results);
                //this allows us to get over the SP item limit of 100 items returned
                if (data.d.__next) {
                    fullUrl = data.d.__next;
                    getItems();
                }else{
                    deferred.resolve(results);
                }
            },
            error: function(data){
                errorFunction(data);
                deferred.reject(data);
            },
            complete: function(data){
            }
        });
    }
    return deferred;
}

//Helper function to resolve the Item type used when updating list items.
//   Note:this will fail if listName isnt simple no space no special char like _
function SPHGetItemType(name) {
    return "SP.Data." + name.charAt(0).toUpperCase() + name.slice(1) + "ListItem";
}

/** Info
 * Summary. Grabs an item from a SP List if the id is found
 * Description. Uses specifically crafted AJAX call to return an item
 *        (to SPHListItems) if the id is found
 * @param {string} ListName : Name of SP List
 * @param {int} id : ID of the item that will be updated
 * @param {object} itemProperties : object of Column/Field name and its value
 *    ex  itemProperties["Title"] = "Title of this item"
 *        itemProperties["Status"] = "Done"  ...etc
 *  @param {string} site URL where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * @param {string} contentType: Unless you need to set to something special you should use the helper function
 *                              SPHGetItemType("name of list") to pass the result to this function as
 *                              it will easily give you the contentType for a given list
 * Use:
 * var updItem = SPHUpdateItemByID('ListName', 'string id to update', itemProperties, 'siteUrl', SPHGetItemType('ListName'));
 * 
 * updItem.done(function(data){
 *       //success!!!
 * }).fail(function(e){ alert('Failed Updating Item by ID! ' + e.responseText);});
 * 
 *   see below for how to structure itemProperties
 * //https://sharepoint.stackexchange.com/questions/105380/adding-new-list-item-using-rest
/*
//specify item properties
var itemProperties = {'Title':'Order task','Description': 'New task'};
*/
function SPHUpdateItemByID(ListTitle, id, itemProperties, siteURL, contentType)
{  
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var siteUrl = siteURL;
    var fullUrl = siteUrl +"/_api/web/lists/GetByTitle('"+ListTitle+"')/items(" + id.toString() + ")";
    var deferred=$.Deferred(); 
    var itemPayload = {
        '__metadata': {'type': contentType}
    };

    for(var prop in itemProperties){
        itemPayload[prop] = itemProperties[prop];
    }
    
    console.log(JSON.stringify(itemPayload));
    $.ajax({
        url: fullUrl,
        type: "POST",
        data: JSON.stringify(itemPayload),  
        headers: {
            "Accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
            "X-RequestDigest" : $("#__REQUESTDIGEST").val(),
            "X-HTTP-Method": "MERGE",
            "If-Match": "*"
        },
        success: function(data){
            deferred.resolve(data);
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}

//https://sharepoint.stackexchange.com/questions/105380/adding-new-list-item-using-rest

/** Info
 * Summary. Grabs an item from a SP List if the id is found
 * Description. Uses specifically crafted AJAX call to return an item
 *        (to SPHListItems) if the id is found
 * @param {string} ListName : Name of SP List
 * @param {int} id : ID of the item that will be updated
 * @param {object} itemProperties : object of Column/Field name and its value
 *    ex  itemProperties["Title"] = "Title of this item"
 *        itemProperties["Status"] = "Done"  ...etc
 * @param {string} site URL where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * @param {string} contentType: Unless you need to set to something special you should use the helper function
 *                              SPHGetItemType("name of list") to pass the result to this function as
 *                              it will easily give you the contentType for a given list
 * Use:
 * //specify item properties
 *   var itemProperties = {};
 *   itemProperties["Name of Property"] = value of property
 *     ie: itemProperties["Title"] = "itemTitle"
 *         itemProperties["Status"] = "Complete"
 * var updItem = SPHCreateListItem('ListName', itemProperties, 'siteURL', SPHGetItemType('ListName'));
 * 
 * updItem.done(function(data){
 *      //success
 * }).fail(function(e){ alert('Failed creating List Item! ' + e.responseText);});
 *     
 * //https://sharepoint.stackexchange.com/questions/105380/adding-new-list-item-using-rest
**/
function SPHCreateListItem(listName, itemProperties, siteURL, contentType) {
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    itemProperties["__metadata"] = { "type": contentType };
    var deferred=$.Deferred(); 
    $.ajax({
        url: siteURL + "/_api/web/lists/getbytitle('" + listName + "')/items",
        type: "POST",
        contentType: "application/json;odata=verbose",
        data: JSON.stringify(itemProperties),
        headers: {
            "Accept": "application/json;odata=verbose",
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
        },
        success: function (data) {
            deferred.resolve(data);
        },
        error: function (data) {
            errorFunction(data);
            deferred.reject(data);
        }
    });

    return deferred;
}

/** Info
 * Summary. Grabs an item from a SP List if the id is found
 * Description. Uses specifically crafted AJAX call to return an item
 *        (to SPHListItems) if the id is found
 * @param {string} ListName 
 * @param {int} id
 * @param {string} site URL where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * Use:
 * var targetItem = SPHGetListItemByID('ListName', 'id as string');
 * 
 * targetItem.done(function(data){
 *     console.log('item =', data.d.results);
 * }.fail(function(e){ alert('Failed Getting List Item by ID! ' + e.responseText);});
 */
function SPHGetListItemByID(ListTitle, id, siteUrl)
{   
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var fullUrl = siteUrl +"/_api/web/lists/GetByTitle('"+ListTitle+"')/items(" + id.toString() + ")";
    var deferred=$.Deferred();  
    $.ajax({
        url: fullUrl,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
        },
        success: function(data){

            deferred.resolve(data);
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}

/** Info
 * Summary. Gets all items of a SP List
 * Description. Returns data from the SP List. Also fills SPHListItems with that data.
 *        Data can be retrieved by user with SPHReturnItems()
 * @param {string} ListName The name of the SP List that is targeted
 * @param {string} siteURL: web address where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * Use:
 * var allColumnsInList = SPHGetListColumns('List Name', 'siteAddress');
 * allColumnsInList.done(function(data){
 *     console.log(data.d.results);
 * }).fail(function(e){ alert('Failed Getting List Columns! ' + e.responseText);});
 */
function SPHGetListColumns(ListName, siteUrl)
{    
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var results = [];
    var fullUrl = siteUrl + "/_api/web/lists/GetByTitle('" + ListName + "')/Fields?$filterfilter=Hidden eq false and ReadOnlyField eq false";
    var deferred=$.Deferred();  
    $.ajax({
        url: fullUrl,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
        },
        success: function(data){
            console.log('sphgetlistcolumns ', data.d.results);
            results = results.concat(data.d.results);
                //this allows us to get over the SP item limit of 100 items returned
                if (data.d.__next) {
                    fullUrl = data.d.__next;
                    getItems();
                }else{
                    deferred.resolve(results);
                }
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}


/** Info
 * Summary. Gets all lists at a site
 * Description. Returns all lists from a sharepoint site.
 * @param {string} siteURL: web address where list exists ex: if my AllItems.aspx address is
 *                   https://.../quality_assurance/Lists/CIMB/AllItems.aspx then
 *                   siteURL = "https://.../quality_assurance"
 * Use:
 * var allSites = SPHGetAllSiteLists('siteAddress');
 * allSites.done(function(data){
 *     console.log(data.d.results);
 * }).fail(function(e){ alert('Failed Getting all Lists from the site! ' + e.responseText);});
 */
function SPHGetAllSiteLists(siteUrl)
{    
    UpdateFormDigest(_spPageContextInfo.webServerRelativeUrl, _spFormDigestRefreshInterval);
    var fullUrl = siteUrl + "/_api/web/lists";
    var deferred=$.Deferred();  
    $.ajax({
        url: fullUrl,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
            "content-type": "application/json;odata=verbose",
        },
        success: function(data){
            results = results.concat(data.d.results);
                //this allows us to get over the SP item limit of 100 items returned
                if (data.d.__next) {
                    fullUrl = data.d.__next;
                    getItems();
                }else{
                    deferred.resolve(results);
                }
        },
        error: function(data){
            errorFunction(data);
            deferred.reject(data);
        }  
    });
    return deferred;
}

/** Info
 * Summary. Displays error info to the console
 * Description. data from the AJAX call is given and the error info is displayed to the console
 * @param {ajax} data AJAX return data
 */
//Error handler
function errorFunction(data)
{
    console.error('error noted in SPHelper_v3.js');
    console.error(data);
    console.error(JSON.stringify(data.responseJSON));
    console.error(data.responseJSON.error.message.value);
    console.error(JSON.stringify(data.responseJSON));
}


/*
Open Modal for SharePoint List Forms (DispForm.aspx, NewForm.aspx etc)
Note: refreshOnSave is if you want the main page to refresh when user clicks on save of the List Form
      true=Page Refreshes   false=Page does not Refresh
*/
function openSPModal(URL, modalTitle, modalWidth, modalHeight, refreshOnSave) {
    var options = {
      url: URL,
      allowMaximize: false,
      showClose: true,
      width: modalWidth,
      height: modalHeight,
      title: modalTitle,
      dialogReturnValueCallback: function dialogReturnValueCallback(dialogResult) {
        if (dialogResult != SP.UI.DialogResult.cancel) {
          //When Save is clicked
          if(refreshOnSave) SP.UI.ModalDialog.RefreshPage(dialogResult); //reload the page
        }
      }
    };
    SP.UI.ModalDialog.showModalDialog(options);
}
