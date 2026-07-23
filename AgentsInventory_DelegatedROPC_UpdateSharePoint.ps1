# ==========================================
# CONFIGURATION
# ==========================================

$tenantId    = ""
$clientId    = ""
$username    = ""
$password    =''

$scope = "https://api.powerplatform.com/.default"
$tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

# ==========================================
# 1. BUILD MULTIPART FORM-DATA REQUEST
# ==========================================

Add-Type -AssemblyName System.Net.Http

$httpClient = New-Object System.Net.Http.HttpClient
$multipartContent = New-Object System.Net.Http.MultipartFormDataContent

function Add-FormField {
    param ($name, $value)

    $stringContent = New-Object System.Net.Http.StringContent($value)
    $stringContent.Headers.ContentDisposition = 
        [System.Net.Http.Headers.ContentDispositionHeaderValue]::new("form-data")

    $stringContent.Headers.ContentDisposition.Name = "`"$name`""

    $multipartContent.Add($stringContent)
}

Add-FormField -name "client_id"  -value $clientId
Add-FormField -name "scope"      -value $scope
Add-FormField -name "username"   -value $username
Add-FormField -name "password"   -value $password
Add-FormField -name "grant_type" -value "password"

Write-Host "Requesting delegated token..."

$response = $httpClient.PostAsync($tokenUri, $multipartContent).Result
$responseContent = $response.Content.ReadAsStringAsync().Result | ConvertFrom-Json

$accessToken = $responseContent.access_token

if (-not $accessToken) {
    throw "Failed to acquire token"
}

Write-Host "Token acquired"

# ==========================================
# HEADERS (keep yours as-is)
# ==========================================
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Content-Type", "application/json")
$headers.Add("Authorization", "Bearer $accessToken")

# ==========================================
# PAGINATION VARIABLES
# ==========================================
$top = 5
$skip = 0          # increment by $top after each batch
$allAgents = @()
$hasMore = $true

# ==========================================
# LOOP
# ==========================================
while ($hasMore) {

    Write-Host "Fetching batch... Skip: $skip"

    # IMPORTANT: inject skip dynamically
    $body = @"
{
  `"TableName`": `"PowerPlatformResources`",
  `"Options`": {
    `"Top`": $top,
    `"Skip`": $skip
  },
  `"Clauses`": [
    {
      `"`$type`": `"where`",
      `"FieldName`": `"type`",
      `"Operator`": `"==`",
      `"Values`": [
        `"'microsoft.copilotstudio/agents'`"
      ]
    },
    {
      `"`$type`": `"project`",
      `"FieldList`": [
        `"id`",
        `"name`",
        `"type`",
        `"tenantId`",
        `"kind`",
        `"location`",
        `"resourceGroup`",
        `"subscriptionId`",
        `"managedBy`",
        `"sku`",
        `"plan`",
        `"tags`",
        `"identity`",
        `"zones`",
        `"extendedLocation`",
        `"properties`"
      ]
    }
  ]
}
"@

    try {
        $response = Invoke-RestMethod `
            -Uri "https://api.powerplatform.com/resourcequery/resources/query?api-version=2024-10-01" `
            -Method POST `
            -Headers $headers `
            -Body $body

        # Append results
        $batchCount = 0
        if ($response.data) {
            $allAgents += $response.data
            $batchCount = @($response.data).Count
        }

        # Handle pagination: advance Skip by $top each batch.
        # Stop when a batch returns fewer than $top records (last page).
        if ($batchCount -lt $top) {
            $hasMore = $false
        }
        else {
            $skip += $top
        }

    } catch {
        Write-Host "Error occurred:"
        Write-Host $_
        Start-Sleep -Seconds 5
    }
}

# ==========================================
# OUTPUT
# ==========================================
Write-Host "Total Agents:" $allAgents.Count

$agents = $allAgents | ForEach-Object {

    $p = $_.properties

    [PSCustomObject]@{
        # Top-level
        Id                = $_.id
        Name              = $_.name
        Type              = $_.type
        TenantId          = $_.tenantId
        Location          = $_.location

        # Core properties
        DisplayName       = $p.displayName
        EnvironmentId     = $p.environmentId
        CreatedAt         = $p.createdAt
        CreatedBy         = $p.createdBy
        LastPublishedAt   = $p.lastPublishedAt
        OwnerId           = $p.ownerId
        SchemaName        = $p.schemaName
        CreatedIn         = $p.createdIn

        # Behavior / config
        Model             = $p.model
        Orchestration     = $p.orchestration
        Authentication    = $p.authentication
        IsManaged         = $p.isManaged
        IsCLIAgent        = $p.isCLIAgent
        IsQuarantined     = $p.isQuarantined
        IsWebSearchEnabled= $p.isWebSearchEnabledForKnowledge

        # Counts
        TopicsCount       = $p.componentsCounts.topics
        ToolsCount        = $p.componentsCounts.tools
        KnowledgeCount    = $p.componentsCounts.knowledge
        ConnectedAgents   = $p.componentsCounts.connectedAgents

        # Sharing
        ViewerUsers       = $p.sharedWithViewers.userCount
        ViewerGroups      = $p.sharedWithViewers.groupCount
        EditorUsers       = $p.sharedWithEditors.userCount
        EditorGroups      = $p.sharedWithEditors.groupCount

        # Advanced
        BlueprintId       = $p.entraAgentBlueprintId
        EntraAgentId      = $p.entraAgentId
        TitleId           = $p.titleId

        # Arrays (stringified for CSV)
        Channels          = ($p.channels -join ";")
        Connectors        = ($p.powerPlatformConnectors -join ";")
        Flows             = ($p.flows -join ";")
        Triggers          = ($p.triggers -join ";")
    }
}

$agents | Export-Csv "AllAgents.csv" -NoTypeInformation
$agents

# Keep only agents created in 'Copilot Studio Lite'
#$agents = @($agents | Where-Object { $_.CreatedIn -eq 'Copilot Studio Lite' })
Write-Host "Agents created in 'Copilot Studio Lite': $($agents.Count)" -ForegroundColor Yellow

Connect-PnPOnline -Url "https://hrtbeat.sharepoint.com/sites/Internal" -Interactive -ClientId "a5809a03-0669-4c83-87e6-f7b44c1ce843"

$listName = "AgentsInventory"
$total  = @($agents).Count
Write-Host "Importing $total agents into SharePoint list '$listName'..." -ForegroundColor Yellow

# Safely parse a date string, returning $null when empty/invalid
function ConvertTo-DateOrNull {
    param($value)
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    try { return [datetime]::Parse($value) } catch { return $null }
}

# ==========================================
# RESOLVE CreatedBy / OwnerId GUIDs -> UPN
# (Microsoft Graph, batched + cached for scale)
# ==========================================

# App registration used for Microsoft Graph (client credentials flow)
$graphClientId     = ""
$graphClientSecret = ""
$tenantId = ""

# Acquire a Graph token via client credentials (app-only)
$graphTokenBody = @{
    client_id     = $graphClientId
    client_secret = $graphClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
}
$graphToken = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method POST -Body $graphTokenBody `
        -ContentType "application/x-www-form-urlencoded").access_token

if (-not $graphToken) {
    throw "Failed to acquire Microsoft Graph token for UPN resolution"
}

$graphHeaders = @{
    Authorization  = "Bearer $graphToken"
    "Content-Type" = "application/json"
}

# Collect unique, valid GUIDs from both CreatedBy and OwnerId
$guidRegex = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$uniqueIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($a in $agents) {
    foreach ($id in @($a.CreatedBy, $a.OwnerId)) {
        if ($id -and $id -match $guidRegex) { [void]$uniqueIds.Add($id) }
    }
}

Write-Host "Resolving $($uniqueIds.Count) unique user GUID(s) via Microsoft Graph..." -ForegroundColor Yellow

# Resolve in batches of 1000 using directoryObjects/getByIds (case-insensitive cache)
$upnCache = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
$idList = @($uniqueIds)
for ($start = 0; $start -lt $idList.Count; $start += 1000) {
    $end   = [Math]::Min($start + 999, $idList.Count - 1)
    $chunk = @($idList[$start..$end])
    $payload = @{ ids = $chunk; types = @("user") } | ConvertTo-Json -Depth 3
    try {
        $result = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/directoryObjects/getByIds" `
            -Method POST -Headers $graphHeaders -Body $payload
        foreach ($u in $result.value) {
            if ($u.id -and $u.userPrincipalName) { $upnCache[$u.id] = $u.userPrincipalName }
        }
    } catch {
        Write-Warning "Graph getByIds failed for batch starting at $start : $($_.Exception.Message)"
    }
}

Write-Host "Resolved $($upnCache.Count) UPN(s)." -ForegroundColor Green

# Batch size controls how many items are sent per server round-trip
$batchSize = 100
$batch = New-PnPBatch

$i = 0
foreach ($agent in $agents) {
    $i++

    # Resolve people-picker UPNs from the Graph cache (blank if unresolved)
    $createdByUpn = $null
    if ($agent.CreatedBy -and $upnCache.ContainsKey($agent.CreatedBy)) { $createdByUpn = $upnCache[$agent.CreatedBy] }
    $ownerUpn = $null
    if ($agent.OwnerId -and $upnCache.ContainsKey($agent.OwnerId)) { $ownerUpn = $upnCache[$agent.OwnerId] }

    $values = @{
        "Title"          = $agent.DisplayName
        "AgentId"        = $agent.Name
        "Type"           = $agent.Type
        "TenantId"       = $agent.TenantId
        "EnvironmentId"  = $agent.EnvironmentId
        "SchemaName"     = $agent.SchemaName
        "BlueprintId"    = $agent.BlueprintId
        "EntraAgentId"   = $agent.EntraAgentId
        "Location"       = $agent.Location
        "CreatedIn"      = $agent.CreatedIn
        "Model"          = $agent.Model
        "Orchestration"  = $agent.Orchestration
        "Authentication" = $agent.Authentication
        "Flows"          = $agent.Flows
        "Triggers"       = $agent.Triggers
        "TitleId"        = $agent.TitleId
        "CreatedBy"      = $createdByUpn
        "OwnerId"        = $ownerUpn
    }

    # Numbers: blank or 0 => 0
    $numberFields = @(
        "TopicsCount", "ToolsCount", "KnowledgeCount", "ConnectedAgents",
        "ViewerUsers", "ViewerGroups", "EditorUsers", "EditorGroups"
    )
    foreach ($nf in $numberFields) {
        $nv = "$($agent.$nf)".Trim()
        if ([string]::IsNullOrWhiteSpace($nv)) {
            $values[$nf] = 0
        }
        else {
            $values[$nf] = [int]$nv
        }
    }

    # Booleans: blank or false => No ($false), anything else => Yes ($true)
    $boolFields = @("IsManaged", "IsCLIAgent", "IsQuarantined", "IsWebSearchEnabled")
    foreach ($bf in $boolFields) {
        $bv = "$($agent.$bf)".Trim()
        $values[$bf] = -not ([string]::IsNullOrWhiteSpace($bv) -or $bv -eq "False")
    }

    # Dates
    $createdAt = ConvertTo-DateOrNull $agent.CreatedAt
    if ($createdAt) { $values["CreatedAt"] = $createdAt }
    $lastPublishedAt = ConvertTo-DateOrNull $agent.LastPublishedAt
    if ($lastPublishedAt) { $values["LastPublishedAt"] = $lastPublishedAt }

    # Multi-choice (semicolon separated in the CSV)
    if (-not [string]::IsNullOrWhiteSpace($agent.Channels)) {
        $values["Channels"] = $agent.Channels -split ';'
    }
    if (-not [string]::IsNullOrWhiteSpace($agent.Connectors)) {
        $values["Connectors"] = $agent.Connectors -split ';'
    }

    # Drop empty values so SharePoint keeps them blank
    $clean = @{}
    foreach ($key in $values.Keys) {
        if ($null -ne $values[$key] -and $values[$key] -ne "") {
            $clean[$key] = $values[$key]
        }
    }

    # Queue the item into the current batch (sent in bulk, not one-by-one)
    Add-PnPListItem -List $listName -Values $clean -Batch $batch | Out-Null

    # Flush the batch once it reaches the batch size
    if (($i % $batchSize) -eq 0) {
        try {
            Invoke-PnPBatch -Batch $batch
            Write-Host "[$i/$total] Uploaded batch (up to $batchSize items)" -ForegroundColor Cyan
        } catch {
            Write-Warning "[$i/$total] Batch upload failed: $($_.Exception.Message)"
        }
        $batch = New-PnPBatch
    }
}

# Flush any remaining items in the final partial batch
if ($batch.RequestCount -gt 0) {
    try {
        Invoke-PnPBatch -Batch $batch
        Write-Host "Uploaded final batch" -ForegroundColor Cyan
    } catch {
        Write-Warning "Final batch upload failed: $($_.Exception.Message)"
    }
}

Write-Host "✅ Bulk upload complete. Processed $total agents." -ForegroundColor Green
