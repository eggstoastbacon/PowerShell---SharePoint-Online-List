<#
    .Synopsis 
        Retrieve SPOIDCR cookie for SharePoint Online.
    .Description
        Authenticates to the sts and retrieves the SPOIDCR cookie for SharePoint Online.
        Will use the custom IDP if one has been setup.
        Optionally, can use integrated credentials (when integrated is set to true) with ADFS using the windowsmixed endpoint.
        Results are formattable as XML, JSON, KEYVALUE, and by line.
        
        Makes global variables avaiable at the end of the run.
        $spoidcrl contains the SPOIDCRL cookie

    .Example 
        The following returns the SPOIDCRL cookie value provided a username and password.

        PS> .\spoidcrl.ps1 -url https://contoso.sharepoint.com -username user@contoso.com -password ABCDEFG
    .Example 
        The following returns the SPOIDCRL cookie value using integrated windows credentials. Applies only to ADFS.

        PS> .\spoidcrl.ps1 -url https://contoso.sharepoint.com/sites/site1 -integrated

	.Example 
        The following saves the SPOIDCRL cookie value using integrated windows credentials. Applies only to ADFS.

        PS> .\spoidcrl.ps1 -url https://contoso.sharepoint.com/sites/site1 -integrated -format "XML" | Out-File "c:\temp\spoidcr.txt"

    .PARAMETER url 
        Tenant url (e.g. contoso.sharepoint.com)
    .PARAMETER username
        The username to login with. (e.g. user@contoso.com or user@contoso.onmicrosoft.com)		
    .PARAMETER password
      The password to login with.
    .PARAMETER integrated
      Whether to use integrated credentials (user running PowerShell) instead of explicit credentials.
      Needs to be supported by ADFS.
    .PARAMETER format
      How to format the output. Options include: XML, JSON, KEYVALUE

#>
[CmdletBinding()]
Param(
[Parameter(Mandatory=$true)]
[string]$url,
[Parameter(Mandatory=$false)]
[string]$username,
[Parameter(Mandatory=$false)]
[string]$password,
[Parameter(Mandatory=$false)]
[switch]$integrated = $false,
[Parameter(Mandatory=$false)]
[string]$format
)

$statusText = New-Object System.Text.StringBuilder

function log($info)
{
    if([string]::IsNullOrEmpty($info))
    {
        $info = ""
    }

    [void]$statusText.AppendLine($info)
}

try
{
    if (![uri]::IsWellFormedUriString($url, [UriKind]::Absolute))
    {
        throw "Parameter 'url' is not a valid URI."
    }
    else
    {
        $uri = [uri]::new($url)
        $tenant = $uri.Authority
    }

    if ($tenant.EndsWith("sharepoint.com", [System.StringComparison]::OrdinalIgnoreCase))
    {
        $msoDomain = "sharepoint.com"
    }
    else
    {
        $msoDomain = $tenant
    }

    if ($integrated.ToBool())
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices") | out-null
        [System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") | out-null
        $username = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.UserPrincipalName
    }
    elseif ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($password))
    {
        $credential = Get-Credential -UserName $username -Message "Enter credentials"
        $username = $credential.UserName
        $password = $credential.GetNetworkCredential().Password
    }

    $contextInfoUrl = $url.TrimEnd('/') + "/_api/contextinfo"
    $getRealmUrl = "https://login.microsoftonline.com/GetUserRealm.srf"
    $realm = "urn:federation:MicrosoftOnline"
    $msoStsAuthUrl = "https://login.microsoftonline.com/rst2.srf"
    $idcrlEndpoint = "https://$tenant/_vti_bin/idcrl.svc/"
    $username = [System.Security.SecurityElement]::Escape($username)
    $password = [System.Security.SecurityElement]::Escape($password)

    # Custom STS integrated authentication envelope format index info
    # 0: message id - unique guid
    # 1: custom STS auth url
    # 2: realm
    $customStsSamlIntegratedRequestFormat = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><s:Envelope xmlns:s=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:a=`"http://www.w3.org/2005/08/addressing`"><s:Header><a:Action s:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action><a:MessageID>urn:uuid:{0}</a:MessageID><a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo><a:To s:mustUnderstand=`"1`">{1}</a:To></s:Header><s:Body><t:RequestSecurityToken xmlns:t=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><wsp:AppliesTo xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`"><wsa:EndpointReference xmlns:wsa=`"http://www.w3.org/2005/08/addressing`"><wsa:Address>{2}</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType><t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType></t:RequestSecurityToken></s:Body></s:Envelope>";


    # custom STS envelope format index info
    # {0}: ADFS url, such as https://corp.sts.contoso.com/adfs/services/trust/2005/usernamemixed, its value comes from the response in GetUserRealm request.
    # {1}: MessageId, it could be an arbitrary guid
    # {2}: UserLogin, such as someone@contoso.com
    # {3}: Password
    # {4}: Created datetime in UTC, such as 2012-11-16T23:24:52Z
    # {5}: Expires datetime in UTC, such as 2012-11-16T23:34:52Z
    # {6}: tokenIssuerUri, such as urn:federation:MicrosoftOnline, or urn:federation:MicrosoftOnline-int
    $customStsSamlRequestFormat = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><s:Envelope xmlns:s=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:wsse=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd`" xmlns:saml=`"urn:oasis:names:tc:SAML:1.0:assertion`" xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`" xmlns:wsu=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd`" xmlns:wsa=`"http://www.w3.org/2005/08/addressing`" xmlns:wssc=`"http://schemas.xmlsoap.org/ws/2005/02/sc`" xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><s:Header><wsa:Action s:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action><wsa:To s:mustUnderstand=`"1`">{0}</wsa:To><wsa:MessageID>{1}</wsa:MessageID><ps:AuthInfo xmlns:ps=`"http://schemas.microsoft.com/Passport/SoapServices/PPCRL`" Id=`"PPAuthInfo`"><ps:HostingApp>Managed IDCRL</ps:HostingApp><ps:BinaryVersion>6</ps:BinaryVersion><ps:UIVersion>1</ps:UIVersion><ps:Cookies></ps:Cookies><ps:RequestParams>AQAAAAIAAABsYwQAAAAxMDMz</ps:RequestParams></ps:AuthInfo><wsse:Security><wsse:UsernameToken wsu:Id=`"user`"><wsse:Username>{2}</wsse:Username><wsse:Password>{3}</wsse:Password></wsse:UsernameToken><wsu:Timestamp Id=`"Timestamp`"><wsu:Created>{4}</wsu:Created><wsu:Expires>{5}</wsu:Expires></wsu:Timestamp></wsse:Security></s:Header><s:Body><wst:RequestSecurityToken Id=`"RST0`"><wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType><wsp:AppliesTo><wsa:EndpointReference>  <wsa:Address>{6}</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><wst:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</wst:KeyType></wst:RequestSecurityToken></s:Body></s:Envelope>"

    # mso envelope format index info (Used for custom STS + MSO authentication)
    # 0: custom STS assertion
    # 1: mso endpoint
    $msoSamlRequestFormat = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><S:Envelope xmlns:S=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:wsse=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd`" xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`" xmlns:wsu=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd`" xmlns:wsa=`"http://www.w3.org/2005/08/addressing`" xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><S:Header><wsa:Action S:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action><wsa:To S:mustUnderstand=`"1`">https://login.microsoftonline.com/rst2.srf</wsa:To><ps:AuthInfo xmlns:ps=`"http://schemas.microsoft.com/LiveID/SoapServices/v1`" Id=`"PPAuthInfo`"><ps:BinaryVersion>5</ps:BinaryVersion><ps:HostingApp>Managed IDCRL</ps:HostingApp></ps:AuthInfo><wsse:Security>{0}</wsse:Security></S:Header><S:Body><wst:RequestSecurityToken xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`" Id=`"RST0`"><wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType><wsp:AppliesTo><wsa:EndpointReference><wsa:Address>{1}</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><wsp:PolicyReference URI=`"MBI`"></wsp:PolicyReference></wst:RequestSecurityToken></S:Body></S:Envelope>"

    # mso envelope format index info (Used for MSO-only authentication)
    # 0: mso endpoint
    # 1: username
    # 2: password
    $msoSamlRequestFormat2 = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><S:Envelope xmlns:S=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:wsse=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd`" xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`" xmlns:wsu=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd`" xmlns:wsa=`"http://www.w3.org/2005/08/addressing`" xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><S:Header><wsa:Action S:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action><wsa:To S:mustUnderstand=`"1`">{0}</wsa:To><ps:AuthInfo xmlns:ps=`"http://schemas.microsoft.com/LiveID/SoapServices/v1`" Id=`"PPAuthInfo`"><ps:BinaryVersion>5</ps:BinaryVersion><ps:HostingApp>Managed IDCRL</ps:HostingApp></ps:AuthInfo><wsse:Security><wsse:UsernameToken wsu:Id=`"user`"><wsse:Username>{1}</wsse:Username><wsse:Password>{2}</wsse:Password></wsse:UsernameToken></wsse:Security></S:Header><S:Body><wst:RequestSecurityToken xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`" Id=`"RST0`"><wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType><wsp:AppliesTo><wsa:EndpointReference><wsa:Address>sharepoint.com</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><wsp:PolicyReference URI=`"MBI`"></wsp:PolicyReference></wst:RequestSecurityToken></S:Body></S:Envelope>"


    function Invoke-HttpPost($endpoint, $body, $headers, $session)
    {
        log
        log "Invoke-HttpPost"
        log "url: $endpoint"
        log "post body: $body"

        $params = @{}
        $params.Headers = $headers
        $params.uri = $endpoint
        $params.Body = $body
        $params.Method = "POST"
        $params.WebSession = $session

        $response = Invoke-WebRequest @params -ContentType "application/soap+xml; charset=utf-8" -UseDefaultCredentials -UserAgent ([string]::Empty)
        $content = $response.Content

        return $content
    }

    # Get saml Assertion value from the custom STS
    function Get-AssertionCustomSts($customStsAuthUrl)
    {
        log
        log "Get-AssertionCustomSts"

        $messageId = [guid]::NewGuid()
        $created = [datetime]::UtcNow.ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
        $expires = [datetime]::UtcNow.AddMinutes(10).ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)

        if ($integrated.ToBool())
        {
            log "integrated"

            $customStsAuthUrl = $customStsAuthUrl.ToLowerInvariant().Replace("/usernamemixed","/windowstransport")
            log $customStsAuthUrl

            $requestSecurityToken = [string]::Format($customStsSamlIntegratedRequestFormat, $messageId, $customStsAuthUrl, $realm)
            log $requestSecurityToken

        }
        else
        {
            log "not integrated"

            $requestSecurityToken = [string]::Format($customStsSamlRequestFormat, $customStsAuthUrl, $messageId, $username, $password, $created, $expires, $realm)
            log $requestSecurityToken

        }

        [xml]$customStsXml = Invoke-HttpPost $customStsAuthUrl $requestSecurityToken

        return $customStsXml.Envelope.Body.RequestSecurityTokenResponse.RequestedSecurityToken.Assertion.OuterXml
    }

    function Get-BinarySecurityToken($customStsAssertion, $msoSamlRequestFormatTemp)
    {
        log
        log "Get-BinarySecurityToken"

        if ([string]::IsNullOrWhiteSpace($customStsAssertion))
        {
            log "using username and password"            
            $msoPostEnvelope = [string]::Format($msoSamlRequestFormatTemp, $msoDomain, $username, $password)
        }
        else
        {
            log "using custom sts assertion"                        
            $msoPostEnvelope = [string]::Format($msoSamlRequestFormatTemp, $customStsAssertion, $msoDomain)
        }

        $msoContent = Invoke-HttpPost $msoStsAuthUrl $msoPostEnvelope
    
        # Get binary security token using regex instead of [xml]
        # Using regex to workaround PowerShell [xml] bug where hidden characters cause failure
        [regex]$regex = "BinarySecurityToken Id=.*>([^<]+)<"
        $match = $regex.Match($msoContent).Groups[1]

        return $match.Value
    }

    function Get-SPOIDCRLCookie($msoBinarySecurityToken)
    {
        log
        log "Get-SPOIDCRLCookie"
        log 
        log "BinarySecurityToken: $msoBinarySecurityToken"

        $binarySecurityTokenHeader = [string]::Format("BPOSIDCRL {0}", $msoBinarySecurityToken)
        $params = @{uri=$idcrlEndpoint
                    Method="GET"
                    Headers = @{}
                   }
        $params.Headers["Authorization"] = $binarySecurityTokenHeader
        $params.Headers["X-IDCRL_ACCEPTED"] = "t"

        $resonse = Invoke-WebRequest @params -UserAgent ([string]::Empty)
        $cookie = $resonse.BaseResponse.Cookies["SPOIDCRL"]

        return $cookie
    }

    # Retrieve the configured STS Auth Url (ADFS, PING, etc.)
    function Get-UserRealmUrl($getRealmUrl, $username)
    {
        log
        log "Get-UserRealmUrl"
        log "url: $getRealmUrl"
        log "username: $username"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $body = "login=$username&xml=1"
        $response = Invoke-WebRequest -Uri $getRealmUrl -Method POST -Body $body -UserAgent ([string]::Empty)
    
        return ([xml]$response.Content).RealmInfo.STSAuthURL
    }

    [System.Net.ServicePointManager]::Expect100Continue = $true

    #1 Get custom STS auth url
    $customStsAuthUrl = Get-UserRealmUrl $getRealmUrl $username

    if ($customStsAuthUrl -eq $null)
    {
        #2 Get binary security token from the MSO STS by passing the SAML <Assertion> xml
        $customStsAssertion = $null
        $msoBinarySecurityToken = Get-BinarySecurityToken $customStsAssertion $msoSamlRequestFormat2
    }
    else
    {
        #2 Get SAML <Assertion> xml from custom STS
        $customStsAssertion = Get-AssertionCustomSts $customStsAuthUrl

        #3 Get binary security token from the MSO STS by passing the SAML <Assertion> xml
        $msoBinarySecurityToken = Get-BinarySecurityToken $customStsAssertion $msoSamlRequestFormat
    }

    #3/4 Get SPOIDRCL cookie from SharePoint site by passing the binary security token
    #  Save cookie and reuse with multiple requests
    $idcrl = $null
    $idcrl = Get-SPOIDCRLCookie $msoBinarySecurityToken
    
    if ([string]::IsNullOrEmpty($format))
    {
        $format = [string]::Empty
    }
    else
    {
        $format = $format.Trim().ToUpperInvariant()
    }

    $Global:spoidcrl = $idcrl
        
    if ($format -eq "XML")
    {
        Write-Output ([string]::Format("<SPOIDCRL>{0}</SPOIDCRL>", $idcrl.Value))
    }
    elseif ($format -eq "JSON")
    {
        Write-Output ([string]::Format("{{`"SPOIDCRL`":`"{0}`"}}", $idcrl.Value))
    }
    elseif ($format.StartsWith("KEYVALUE") -or $format.StartsWith("NAMEVALUE"))
    {
        Write-Output ("SPOIDCRL:" + $idcrl.Value)
    }
    else
    {
        Write-Output $idcrl.Value
    }

}
catch
{
    log $error[0]
    "ERROR:" + $statusText.ToString()
}
