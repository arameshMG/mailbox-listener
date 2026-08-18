$port = if ($env:PORT) { $env:PORT } else { "5050" }
$sharedSecret = $env:API_SHARED_SECRET
$appId = $env:EXO_APP_ID
$clientSecret = $env:EXO_CLIENT_SECRET
$tenantId = $env:EXO_TENANT_ID
$organization = $env:EXO_ORGANIZATION

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$port/")
$listener.Start()
Write-Host "Listening on port $port ..."

while ($true) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    try {
        $providedSecret = $request.Headers["X-Api-Key"]

        if ($request.Url.AbsolutePath -eq "/health") {
            $responseText = '{"status":"ok"}'
            $response.StatusCode = 200
        }
        elseif ($providedSecret -ne $sharedSecret) {
            Write-Host "Rejected request: bad or missing API key"
            $responseText = '{"status":"error","message":"Unauthorized"}'
            $response.StatusCode = 401
        }
        else {
            $reader = New-Object System.IO.StreamReader($request.InputStream)
            $rawBody = $reader.ReadToEnd()
            Write-Host "RAW BODY: $rawBody"

            $body = $rawBody | ConvertFrom-Json
            $upn = $body.userPrincipalName

            if ([string]::IsNullOrWhiteSpace($upn)) {
                throw "userPrincipalName was empty or missing in request body"
            }

            Write-Host "Converting mailbox for: $upn"

            # Get an access token using the client secret 
            $tokenBody = @{
                client_id     = $appId
                client_secret = $clientSecret
                scope         = "https://outlook.office365.com/.default"
                grant_type    = "client_credentials"
            }
            $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody

            Connect-ExchangeOnline -AccessToken $tokenResponse.access_token -Organization $organization -ShowBanner:$false

            Set-Mailbox -Identity $upn -Type Shared

            Disconnect-ExchangeOnline -Confirm:$false

            Write-Host "Successfully converted: $upn"
            $responseText = '{"status":"converted","userPrincipalName":"' + $upn + '"}'
            $response.StatusCode = 200
        }
    }
    catch {
        $errorMsg = $_.Exception.Message -replace '"', '\"'
        Write-Host "ERROR: $errorMsg"
        $responseText = '{"status":"error","message":"' + $errorMsg + '"}'
        $response.StatusCode = 500
    }

    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseText)
    $response.ContentType = "application/json"
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}
