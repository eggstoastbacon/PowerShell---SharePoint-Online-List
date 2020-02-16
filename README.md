# PowerShell---SharePoint-Online-List
Fetch list items from a SharePoint online list using a login cookie instead of an api key. 
This has to make a few passes on the pages as these is a list view limit on the api. 
We get around this limit by iterating through the pages and using $top in the query. 
This requires an account with read access to the list.
This script could be used to create list backups or feed data into reporting systems that do not have a SharePoint online plugin. 
Add your own filters, information of some of the filters here: https://social.technet.microsoft.com/wiki/contents/articles/35796.sharepoint-2013-using-rest-api-for-selecting-filtering-sorting-and-pagination-in-sharepoint-list.aspx
The cookie retrieve function is not my work, unfortunately the author is not in the comments of the function and it's been awhile since I obtained it. Many thanks to whoever that was.
