param(
    [Parameter(Mandatory = $true)]
    [string]$TenantDomain, # e.g. tenant.console.ves.volterra.io

    [Parameter(Mandatory = $true)]
    [string]$ApiToken, # e.g. zQCKkBwWU9/k4ep7BAg=

    [ValidateSet("Table", "CSV", "Text")]
    [string]$OutputFormat = "Table"
)

# Static API path
$ApiPath = "/api/web/custom/namespaces/system/user_roles"
$FullUrl = "https://$TenantDomain$ApiPath"

# Load Microsoft Graph module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph

# Connect to Microsoft Graph if not already connected
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "User.Read.All"
}

# Query XC API
try {
    $response = Invoke-RestMethod -Uri $FullUrl -Headers @{
        "Authorization" = "APIToken $ApiToken"
    } -Method Get
}
catch {
    Write-Error "Failed to query XC API at $FullUrl"
    exit 1
}

if (-not $response.items) {
    Write-Error "No 'items' array found in the response."
    exit 1
}

# Process user emails
$results = @()

$total = $response.items.Count
$counter = 0

foreach ($item in $response.items) {
    $counter++
    $email = $item.email

    Write-Progress -Activity "Verifying Users" `
        -Status "Checking $email" `
        -PercentComplete (($counter / $total) * 100)

    if ([string]::IsNullOrWhiteSpace($email)) { continue }

    try {
        $user = Get-MgUser -UserId $email -ErrorAction Stop
        $exists = "Yes"
        $enabled = if ($user.accountEnabled) { "Yes" } else { "No" }
    }
    catch {
        $exists = "No"
        $enabled = "N/A"
    }

    $results += [PSCustomObject]@{
        Email           = $email
        ExistsInAzureAD = $exists
        Enabled         = $enabled
    }
}


# Output formatting
switch ($OutputFormat) {
    "Table" {
        $results | Sort-Object Email | Format-Table -AutoSize
    }
    "CSV" {
        $csvPath = "XCUserCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $results | Sort-Object Email | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Results written to $csvPath"
    }
    "Text" {
        foreach ($entry in $results | Sort-Object Email) {
            Write-Host "$($entry.Email): $($entry.ExistsInAzureAD)"
        }
    }
}
