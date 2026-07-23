# Connect to SharePoint
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/Internal" -Interactive -ClientId ""

$listName = "AgentsInventory"

# Create List
New-PnPList -Title $listName -Template GenericList -OnQuickLaunch

# ---------------------------
# TEXT FIELDS (Create first)
# ---------------------------
Add-PnPField -List $listName -DisplayName "AgentId" -InternalName "AgentId" -Type Text
Add-PnPField -List $listName -DisplayName "Type" -InternalName "Type" -Type Text
Add-PnPField -List $listName -DisplayName "TenantId" -InternalName "TenantId" -Type Text
Add-PnPField -List $listName -DisplayName "EnvironmentId" -InternalName "EnvironmentId" -Type Text
Add-PnPField -List $listName -DisplayName "SchemaName" -InternalName "SchemaName" -Type Text
Add-PnPField -List $listName -DisplayName "BlueprintId" -InternalName "BlueprintId" -Type Text
Add-PnPField -List $listName -DisplayName "EntraAgentId" -InternalName "EntraAgentId" -Type Text
Add-PnPField -List $listName -DisplayName "TitleId" -InternalName "TitleId" -Type Text  
Add-PnPField -List $listName -DisplayName "AgentCreatedBy" -InternalName "AgentCreatedBy" -Type Text
Add-PnPField -List $listName -DisplayName "AgentOwner" -InternalName "AgentOwner" -Type Text

# ---------------------------
# APPLY MAX LENGTH (FIX)
# ---------------------------
function Set-MaxLength {
    param ($list, $fieldName, $length)

    $field = Get-PnPField -List $list -Identity $fieldName
    $xml = $field.SchemaXml -replace 'MaxLength="\d*"', ""  # clean existing
    $xml = $xml -replace 'Type="Text"', "Type=`"Text`" MaxLength=`"$length`""

    Set-PnPField -List $list -Identity $fieldName -Values @{SchemaXml = $xml}
}

Set-MaxLength -list $listName -fieldName "AgentId" -length 36
Set-MaxLength -list $listName -fieldName "TenantId" -length 36
Set-MaxLength -list $listName -fieldName "EnvironmentId" -length 36
Set-MaxLength -list $listName -fieldName "BlueprintId" -length 36
Set-MaxLength -list $listName -fieldName "EntraAgentId" -length 36

# ---------------------------
# DATE
# ---------------------------
Add-PnPField -List $listName -DisplayName "CreatedAt" -InternalName "CreatedAt" -Type DateTime
Add-PnPField -List $listName -DisplayName "LastPublishedAt" -InternalName "LastPublishedAt" -Type DateTime 
Add-PnPField -List $listName -DisplayName "LastCertifiedDate" -InternalName "LastCertifiedDate" -Type DateTime


# ---------------------------
# PEOPLE
# ---------------------------
Add-PnPField -List $listName -DisplayName "AgentCreatedByPkr" -InternalName "AgentCreatedByPkr" -Type User
Add-PnPField -List $listName -DisplayName "AgentOwnerPkr" -InternalName "AgentOwnerPkr" -Type User

# ---------------------------
# NUMBERS
# ---------------------------
$numberFields = @(
"TopicsCount","ToolsCount","KnowledgeCount",
"ConnectedAgents","ViewerUsers","ViewerGroups",
"EditorUsers","EditorGroups"
)

foreach ($f in $numberFields) {
    Add-PnPField -List $listName -DisplayName $f -InternalName $f -Type Number
}

# ---------------------------
# MULTILINE TEXT
# ---------------------------
Add-PnPField -List $listName -DisplayName "Flows" -InternalName "Flows" -Type Note
Add-PnPField -List $listName -DisplayName "Triggers" -InternalName "Triggers" -Type Note

# ---------------------------
# CHOICE (Single) - FillInChoice enabled so real agent data is accepted
# ---------------------------
Add-PnPField -List $listName -DisplayName "Location" -InternalName "Location" -Type Choice

Add-PnPField -List $listName -DisplayName "CreatedIn" -InternalName "CreatedIn" -Type Choice `
    -Choices @("Copilot Studio","Copilot Studio Lite")

Add-PnPField -List $listName -DisplayName "Model" -InternalName "Model" -Type Choice

Add-PnPField -List $listName -DisplayName "Orchestration" -InternalName "Orchestration" -Type Choice `
    -Choices @("Generative","Classic")

Add-PnPField -List $listName -DisplayName "Authentication" -InternalName "Authentication" -Type Choice `
    -Choices @()

Add-PnPField -List $listName -DisplayName "AgentClassification" -InternalName "AgentClassification" -Type Choice `
    -Choices @("L1","L2")

Add-PnPField -List $listName -DisplayName "CertificationStatus" -InternalName "CertificationStatus" -Type Choice `
    -Choices @("Certified","Not Certified","L1 Assistive","L2 Augumented","Deleted")

# ---------------------------
# YES/NO (Better as Boolean)
# ---------------------------
$boolFields = @("IsManaged","IsCLIAgent","IsQuarantined","IsWebSearchEnabled")

foreach ($f in $boolFields) {
    Add-PnPField -List $listName -DisplayName $f -InternalName $f -Type Boolean
}

# ---------------------------
# MULTI-CHOICE
# ---------------------------
Add-PnPField -List $listName -DisplayName "Channels" -InternalName "Channels" -Type MultiChoice `
    -Choices @("Teams","Microsoft 365 Copilot","Direct Line Channels")

Add-PnPField -List $listName -DisplayName "Connectors" -InternalName "Connectors" -Type MultiChoice `
    -Choices @()

    $fieldsWithFillIn = @(
    "Location",
    "CreatedIn",
    "Model",
    "Orchestration",
    "Authentication",
    "Channels",
    "Connectors"
)

foreach ($field in $fieldsWithFillIn) {
    Set-PnPField -List $listName -Identity $field -Values @{FillInChoice = $true}
}

Write-Host "✅ Choice & MultiChoice fields created with Fill-In enabled!" -ForegroundColor Green

# Get default view
$view = Get-PnPView -List $listName | Where-Object { $_.DefaultView -eq $true }

# Set desired fields in order
Set-PnPView -List $listName -Identity $view.Id -Fields @(
    "LinkTitle",      # Title column (required internal name)
    "AgentId",
    "CreatedAt",
    "AgentCreatedBy",
    "AgentOwner",
    "AgentCreatedByPkr",
    "AgentOwnerPkr",
    "LastPublishedAt",
    "ViewerUsers",
    "ViewerGroups",
    "CertificationStatus",
    "AgentClassification",
    "LastCertifiedDate"
)

Write-Host "✅ List created successfully with correct schema!" -ForegroundColor Green



