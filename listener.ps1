# Mailbox conversion listener
# Runs inside a Docker container on Render.
# Listens for POST requests and converts the specified mailbox to Shared.

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
        elseif ($request.Url.AbsolutePath -eq "/add-to-distribution-list") {
            $reader = New-Object System.IO.StreamReader($request.InputStream)
            $rawBody = $reader.ReadToEnd()
            Write-Host "RAW BODY: $rawBody"

            $body = $rawBody | ConvertFrom-Json
            $upn = $body.userPrincipalName
            $dlIdentity = $body.distributionListIdentity

            if ([string]::IsNullOrWhiteSpace($upn)) {
                throw "userPrincipalName was empty or missing in request body"
            }
            if ([string]::IsNullOrWhiteSpace($dlIdentity)) {
                throw "distributionListIdentity was empty or missing in request body"
            }

            Write-Host "Adding $upn to distribution list $dlIdentity"

            $tokenBody = @{
                client_id     = $appId
                client_secret = $clientSecret
                scope         = "https://outlook.office365.com/.default"
                grant_type    = "client_credentials"
            }
            $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody

            Connect-ExchangeOnline -AccessToken $tokenResponse.access_token -Organization $organization -ShowBanner:$false

            try {
                Add-DistributionGroupMember -Identity $dlIdentity -Member $upn -ErrorAction Stop
                Write-Host "Successfully added $upn to $dlIdentity"

                $responseText = '{"status":"added","userPrincipalName":"' + $upn + '","distributionListIdentity":"' + $dlIdentity + '"}'
                $response.StatusCode = 200
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false
            }
        }
        elseif ($request.Url.AbsolutePath -eq "/remove-from-distribution-list") {
            $reader = New-Object System.IO.StreamReader($request.InputStream)
            $rawBody = $reader.ReadToEnd()
            Write-Host "RAW BODY: $rawBody"

            $body = $rawBody | ConvertFrom-Json
            $upn = $body.userPrincipalName

            if ([string]::IsNullOrWhiteSpace($upn)) {
                throw "userPrincipalName was empty or missing in request body"
            }

            Write-Host "Finding and removing $upn from all distribution lists"

            $tokenBody = @{
                client_id     = $appId
                client_secret = $clientSecret
                scope         = "https://outlook.office365.com/.default"
                grant_type    = "client_credentials"
            }
            $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody

            Connect-ExchangeOnline -AccessToken $tokenResponse.access_token -Organization $organization -ShowBanner:$false

            try {
                $allDLs = Get-DistributionGroup -ResultSize Unlimited
                $removedFrom = @()
                $failedOn = @()

                foreach ($dl in $allDLs) {
                    $isMember = Get-DistributionGroupMember -Identity $dl.Identity -ResultSize Unlimited | Where-Object { $_.PrimarySmtpAddress -eq $upn } | Select-Object -First 1

                    if ($isMember) {
                        try {
                            Remove-DistributionGroupMember -Identity $dl.Identity -Member $upn -Confirm:$false -ErrorAction Stop
                            Write-Host "Removed $upn from $($dl.PrimarySmtpAddress)"
                            $removedFrom += $dl.PrimarySmtpAddress
                        }
                        catch {
                            Write-Host "Failed to remove $upn from $($dl.PrimarySmtpAddress): $($_.Exception.Message)"
                            $failedOn += $dl.PrimarySmtpAddress
                        }
                    }
                    $isMember = $null
                    [System.GC]::Collect()
                }

                $removedJson = ($removedFrom | ForEach-Object { '"' + $_ + '"' }) -join ','
                $failedJson = ($failedOn | ForEach-Object { '"' + $_ + '"' }) -join ','

                $responseText = '{"status":"completed","userPrincipalName":"' + $upn + '","removedFrom":[' + $removedJson + '],"failedOn":[' + $failedJson + ']}'
                $response.StatusCode = 200
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false
            }
        }
        else {
            $reader = New-Object System.IO.StreamReader($request.InputStream)
            $rawBody = $reader.ReadToEnd()
            Write-Host "RAW BODY: $rawBody"

            $body = $rawBody | ConvertFrom-Json
            $upn = $body.userPrincipalName
            $grantAccessUpn = $body.grantAccessUpn   # optional — UPN of whoever should get mailbox access

            if ([string]::IsNullOrWhiteSpace($upn)) {
                throw "userPrincipalName was empty or missing in request body"
            }

            Write-Host "Converting mailbox for: $upn"

            # Get an access token using the client secret (same pattern as Graph token calls)
            $tokenBody = @{
                client_id     = $appId
                client_secret = $clientSecret
                scope         = "https://outlook.office365.com/.default"
                grant_type    = "client_credentials"
            }
            $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody

            Connect-ExchangeOnline -AccessToken $tokenResponse.access_token -Organization $organization -ShowBanner:$false

            try {
                Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                Write-Host "Successfully converted: $upn"

                $permissionGranted = $false
                if (-not [string]::IsNullOrWhiteSpace($grantAccessUpn)) {
                    Write-Host "Granting Full Access + Send As to: $grantAccessUpn"
                    Add-MailboxPermission -Identity $upn -User $grantAccessUpn -AccessRights FullAccess -InheritanceType All -AutoMapping $true -Confirm:$false -ErrorAction Stop | Out-Null
                    Add-RecipientPermission -Identity $upn -Trustee $grantAccessUpn -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                    $permissionGranted = $true
                    Write-Host "Successfully granted access to: $grantAccessUpn"
                }

                $responseText = '{"status":"converted","userPrincipalName":"' + $upn + '","permissionGranted":' + $permissionGranted.ToString().ToLower() + '}'
                $response.StatusCode = 200
            }
            finally {
                Disconnect-ExchangeOnline -Confirm:$false
            }
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
