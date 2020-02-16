# PowerShell---SharePoint-Online-List
Fetch list items from a SharePoint online list using a login cookie instead of an api key. 
This has to make a few passes on the pages as these is a list view limit on the api. 
We get around this limit by iterating 75 items at a time in $pages variable. This requires an account with read access to the list.
This script could be used to create list backups or feed data into reporting systems that do not have a SharePoint online plugin. 

This script can be quite fast.. if your connection is good you can parse over 5000 list items in minutes.

The cookie retrieve function is not my work, unfortuanately the author is not in the comments of the function and it's been awhile since I obtained it. Many thanks to whoever that was.
