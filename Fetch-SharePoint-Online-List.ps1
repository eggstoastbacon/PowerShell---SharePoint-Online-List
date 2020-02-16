# SCRIPT BY EGGSTOASTBACON :: https://github.com/eggstoastbacon/

#This script requires SPOnline-Get-Cookie.ps1, change the path to it below.
#Force TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$userName = "someaccount@microsoft.com"
# Get encrypted password for account
$securedPassword = Get-Content "D:\your\enc\password-file.enc" | ConvertTo-SecureString
#Root of the site your list is in
$urlBase = "https://yourdomain.sharepoint.com/sites/yoursite"
# Add %20 instead of spaces.
$spList = "Your%20List"

#Need to decrypt the password
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedPassword)
$decryptedPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

#Function to fetch the cookie, requires decrypted password
$spCookie = . D:\Scripts\Functions\SPOnline-Get-Cookie.ps1 -url "$urlBase" -format "XML" -username $username -password $decryptedPassword
clear-variable decryptedPassword

#Clean up the cookie
$spCookie = $spCookie.replace("</SPOIDCRL>", "")
$spCookie = $spCookie.replace("<SPOIDCRL>", "")

$credential = New-Object System.Management.Automation.PSCredential ($username, $securedPassword)

# adjust this and add more +75 depending on the size of your list, eg. , 450, 525. 
# You could probably dynamically create variables until you recieved no more list items using while
$data = @()
$page = 0
$count = 0
while (($count -eq 0) -or ($count -ne $countTracker)) {
    $page = $page + 20
    $countTracker = $count
    # Add your own filters, default is page by page until it cannot find anymore items.
    # https://social.technet.microsoft.com/wiki/contents/articles/35796.sharepoint-2013-using-rest-api-for-selecting-filtering-sorting-and-pagination-in-sharepoint-list.aspx
    [System.Uri]$uri = "$urlBase/_api/web/lists/GetByTitle('$spList')/items?`$skiptoken=Paged=TRUE%26p_ID=$page&`$top=20 "
    
    $contentType = "application/json" 
    $method = [Microsoft.PowerShell.Commands.WebRequestMethod]::Get 
    $body = '' 

    $cookie = New-Object System.Net.Cookie
    $cookie.Name = "SPOIDCRL" # Add the name of the cookie
    $cookie.Value = "$spCookie" # Add the value of the cookie
    $cookie.Domain = $uri.DnsSafeHost

    $webSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $webSession.Cookies.Add($cookie)

    $headers = @{ 
        "accept" = "application/json;odata=verbose" 
    } 

    #Plop the PARAMS
    $props = @{
        Uri         = $uri.AbsoluteUri
        Headers     = $headers
        Credential  = $credential
        ContentType = $contentType
        Method      = $method
        WebSession  = $webSession
    }

    $spRESTResults = Invoke-RestMethod @props
    $spRESTResultsCorrected = $spRESTResults -creplace '"Id":', '"Fake-Id":' 
    try {
        $spResults = $spRESTResultsCorrected | ConvertFrom-Json 
    }
    catch { }
    $spListItems = $spResults.d.results
   
    #Store results avoiding duplicates
    foreach ($spListItem in $spListItems | `
            Where-Object { $_.ID -notin $data.ID -and $spListItem -notlike $NULL } ) { 
            
        $data += $spListItem
        $count = $data.count
        #Displays the ID it's added, 
        $spListItem.ID
        #This is where we output data, take $splistitem and add it to a SQL database or output it to a CSV, up to you.             
    }
}
