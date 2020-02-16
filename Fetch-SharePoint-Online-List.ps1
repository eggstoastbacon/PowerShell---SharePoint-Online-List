#This script requires SPOnline-Get-Cookie.Ps1
$data = @()
$userName = "name@yourmicrosoftaccount.com"
# Get encrypted password for account
$securedPassword = Get-Content "\path\to\enc\yourencryptedpassword file.enc" | ConvertTo-SecureString
#Root of the site your list is in
$urlBase = "https://yourdomain.sharepoint.com/yoursite"
# Add %20 instead of spaces.
$spList = "Your%20List"

#Need to decrypt the password
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securedPassword)
$decryptedPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

#Function to fetch the cookie, requires decrypted password
$spCookie = . D:\Path\To\SPOnline-Get-Cookie.ps1 -url "$urlBase" -format "XML" -username $username -password $decryptedPassword
clear-variable decryptedPassword

#Clean up the cookie
$spCookie = $spCookie.replace("</SPOIDCRL>", "")
$spCookie = $spCookie.replace("<SPOIDCRL>", "")

$credential = New-Object System.Management.Automation.PSCredential ($username, $securedPassword)

$pages = (0, 75, 150, 225, 300, 375)

foreach ($page in $pages) {

    [System.Uri]$uri = "$urlBase/_api/web/lists/GetByTitle('$spList)/items?%24skiptoken=Paged%3dTRUE%26p_ID%3d$page&$TOP=1000" # Add the Uri
    $contentType = "application/json" # Add the content type
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

    # Splat the parameters
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
    $spResults = $spRESTResultsCorrected | ConvertFrom-Json 
    $spListItems = $spResults.d.results 

    #Store results avoiding deuplicates and empties "NULL"
    foreach ($spListItem in $spListItems) { 
    if($splistitem -notin $data -and $data -notlike $NULL){
    $data += $splistitem
    }
    }
    }
