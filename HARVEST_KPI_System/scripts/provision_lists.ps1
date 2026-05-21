<#
.SYNOPSIS
    Provisions the six HARVEST KPI SharePoint Lists in one shot.

.DESCRIPTION
    Creates Food Entrepreneurs, Institution, Community Orgs,
    Investors-Funders, Farmers, and Environmental Lists with all
    columns, choices, defaults, and views defined in
    schemas/sharepoint_lists.md.

    Idempotent: safe to re-run. Lists that already exist are skipped.
    Columns that already exist on a List are skipped individually.

.PREREQUISITES
    1. PowerShell 7+ (PowerShell ISE on Windows works; on Mac, install
       PowerShell from <https://github.com/PowerShell/PowerShell>).
    2. PnP.PowerShell module:
           Install-Module -Name PnP.PowerShell -Scope CurrentUser
    3. You must have at least Edit permission on the target SharePoint
       site. No app registration is needed — the script uses interactive
       authentication.

    If your RWJBH IT environment blocks the PnP.PowerShell module, this
    script can't run there. In that case, follow the manual setup steps
    in schemas/sharepoint_lists.md. Both paths produce the same Lists.

.PARAMETER SiteUrl
    Full URL of the HARVEST SharePoint site, e.g.
    https://rwjbh.sharepoint.com/sites/HARVEST

.PARAMETER WhatIf
    Show what would be created without actually creating anything.

.EXAMPLE
    .\provision_lists.ps1 -SiteUrl https://rwjbh.sharepoint.com/sites/HARVEST

.EXAMPLE
    .\provision_lists.ps1 -SiteUrl https://rwjbh.sharepoint.com/sites/HARVEST -WhatIf

.NOTES
    Column-name encoding: SharePoint Lists don't allow $ or % in internal
    column names. The script creates columns with safe internal names
    (Sales_Period_USD, Earned_Revenue_Pct, etc.). The display names are
    set to the human-readable form (Sales Period $, Earned Revenue %).
    The Power Query M code in docs/data_dictionary.md renames these
    on the way into the workbook.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

$ErrorActionPreference = "Stop"

# -----------------------------------------------------------------------------
# Connect
# -----------------------------------------------------------------------------

Write-Host ""
Write-Host "HARVEST KPI — SharePoint Lists Provisioning" -ForegroundColor Green
Write-Host "Site: $SiteUrl"
Write-Host ""

try {
    Import-Module PnP.PowerShell -ErrorAction Stop
} catch {
    Write-Host "PnP.PowerShell module not installed." -ForegroundColor Red
    Write-Host "Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
    Write-Host "If your tenant blocks it, follow schemas/sharepoint_lists.md (manual setup)."
    exit 1
}

Connect-PnPOnline -Url $SiteUrl -Interactive | Out-Null

# -----------------------------------------------------------------------------
# Common columns — added to every List
# -----------------------------------------------------------------------------

$OfsaChoices = @("Availability", "Access", "Utilization", "Stability", "Agency", "Sustainability")
$SectorChoices = @("Production", "Aggregation", "Processing", "Distribution", "Retail-Consumption", "Recovery-Waste")
$PartnerChoices = @("Healthcare", "Agriculture", "Education", "Government", "Philanthropy", "Community-Based Org", "Private Sector", "Food Enterprise")
$LeadChoices = @("HARVEST Lead", "UAC Lead", "Joint")

$CommonColumns = @(
    @{ Name = "Submitted_By"; Type = "Text"; Required = $true; Display = "Submitted By" },
    @{ Name = "Submitted_Date"; Type = "DateTime"; Required = $true; Display = "Submitted Date" },
    @{ Name = "Reporting_Period_Start"; Type = "DateTime"; Required = $true; Display = "Reporting Period Start" },
    @{ Name = "Reporting_Period_End"; Type = "DateTime"; Required = $true; Display = "Reporting Period End" },
    @{ Name = "OFSA_Dimension"; Type = "MultiChoice"; Required = $true; Display = "OFSA Dimension"; Choices = $OfsaChoices },
    @{ Name = "Food_System_Sector"; Type = "Choice"; Required = $false; Display = "Food System Sector"; Choices = $SectorChoices },
    @{ Name = "Partner_Sector"; Type = "Choice"; Required = $false; Display = "Partner Sector"; Choices = $PartnerChoices },
    @{ Name = "Lead_Role"; Type = "Choice"; Required = $true; Display = "Lead Role"; Choices = $LeadChoices; Default = "HARVEST Lead" },
    @{ Name = "Notes"; Type = "Note"; Required = $false; Display = "Notes" }
)

# -----------------------------------------------------------------------------
# Per-List column definitions
# -----------------------------------------------------------------------------

$Lists = @(
    @{
        Name = "Food Entrepreneurs List"
        Description = "Member businesses, conversion funnel, jobs, growth, business support."
        Columns = @(
            @{ Name = "Business_Name"; Type = "Text"; Display = "Business Name" }
            @{ Name = "Stage"; Type = "Choice"; Display = "Stage"; Required = $true; Default = "Engaged"
               Choices = @("Engaged", "Onboarding", "Active Member", "Inactive", "Graduated", "Operational Snapshot") }
            @{ Name = "Date_Engaged"; Type = "DateTime"; Display = "Date Engaged" }
            @{ Name = "Date_Converted"; Type = "DateTime"; Display = "Date Converted" }
            @{ Name = "Jobs_Created_FTE"; Type = "Number"; Display = "Jobs Created FTE" }
            @{ Name = "Jobs_Created_PTE"; Type = "Number"; Display = "Jobs Created PTE" }
            @{ Name = "Sales_Period_USD"; Type = "Currency"; Display = "Sales Period `$" }
            @{ Name = "Sales_YoY_Change_Pct"; Type = "Number"; Display = "Sales YoY Change %" }
            @{ Name = "Sources_Local_Produce"; Type = "Boolean"; Display = "Sources Local Produce" }
            @{ Name = "Wellness_Coaching_Sessions"; Type = "Number"; Display = "Wellness Coaching Sessions" }
            @{ Name = "Demographic_Tags"; Type = "MultiChoice"; Display = "Demographic Tags"
               Choices = @("Woman-owned", "BIPOC-owned", "Immigrant-owned", "Veteran-owned", "LGBTQ+-owned", "First-time entrepreneur") }
            @{ Name = "New_Members"; Type = "Number"; Display = "New Members (weekly count)" }
            @{ Name = "ServSafe_Passed"; Type = "Number"; Display = "ServSafe Passed" }
            @{ Name = "ServSafe_Attempted"; Type = "Number"; Display = "ServSafe Attempted" }
        )
        Views = @(
            @{ Name = "Active Members"; Query = "<Where><Eq><FieldRef Name='Stage'/><Value Type='Choice'>Active Member</Value></Eq></Where>" }
            @{ Name = "Current Quarter"; Query = "<Where><Geq><FieldRef Name='Submitted_Date'/><Value Type='DateTime'><Today OffsetDays='-90'/></Value></Geq></Where>" }
        )
    },
    @{
        Name = "Institution List"
        Description = "Institutional purchasing, menu placements, cross-institution connections, TA partnerships."
        Columns = @(
            @{ Name = "Institution_Name"; Type = "Text"; Display = "Institution Name"; Required = $true }
            @{ Name = "Institution_Type"; Type = "Choice"; Display = "Institution Type"
               Choices = @("Healthcare", "Education", "Government", "Corporate", "Other") }
            @{ Name = "Activity_Type"; Type = "Choice"; Display = "Activity Type"; Required = $true
               Choices = @("Purchase Order", "Menu Placement", "TA Partnership", "Cross-Institution Connection", "Health Outcome Measure") }
            @{ Name = "PO_Value_USD"; Type = "Currency"; Display = "PO Value `$" }
            @{ Name = "Items_On_Menu"; Type = "Number"; Display = "Items On Menu" }
            @{ Name = "Connection_Notes"; Type = "Note"; Display = "Connection Notes" }
            @{ Name = "Health_Outcome_Metric"; Type = "Text"; Display = "Health Outcome Metric" }
            @{ Name = "Health_Outcome_Value"; Type = "Currency"; Display = "Health Outcome Value" }
        )
        Views = @(
            @{ Name = "Current Quarter"; Query = "<Where><Geq><FieldRef Name='Submitted_Date'/><Value Type='DateTime'><Today OffsetDays='-90'/></Value></Geq></Where>" }
        )
    },
    @{
        Name = "Community Orgs List"
        Description = "Community events, tours, attendance, surveys, food recovery, volunteer hours."
        Columns = @(
            @{ Name = "Event_Name"; Type = "Text"; Display = "Event Name"; Required = $true }
            @{ Name = "Event_Type"; Type = "Choice"; Display = "Event Type"; Required = $true
               Choices = @("Teaching Kitchen", "RD Workshop", "Tour", "Partner-Hosted Event", "HARVEST-Hosted Event", "Pop-Up", "Other") }
            @{ Name = "Event_Date"; Type = "DateTime"; Display = "Event Date"; Required = $true }
            @{ Name = "Attendance_Count"; Type = "Number"; Display = "Attendance Count" }
            @{ Name = "Volunteer_Hours"; Type = "Number"; Display = "Volunteer Hours" }
            @{ Name = "Orgs_Engaged"; Type = "Note"; Display = "Orgs Engaged (one per line)" }
            @{ Name = "Sector_Count"; Type = "Number"; Display = "Sector Count" }
            @{ Name = "Multilingual_Cultural"; Type = "Boolean"; Display = "Multilingual / Cultural" }
            @{ Name = "Language_Tradition"; Type = "Text"; Display = "Language / Tradition" }
            @{ Name = "Survey_Satisfaction_Avg"; Type = "Number"; Display = "Survey — Satisfaction Avg" }
            @{ Name = "Survey_Learning_Avg"; Type = "Number"; Display = "Survey — Learning Avg" }
            @{ Name = "Food_Recovery_lbs"; Type = "Number"; Display = "Food Recovery (lbs)" }
            @{ Name = "Food_Recovery_Recipient"; Type = "Text"; Display = "Food Recovery Recipient" }
            @{ Name = "Direct_Service_Provider_Onsite"; Type = "Boolean"; Display = "Direct Service Provider On-site" }
        )
        Views = @(
            @{ Name = "Current Quarter"; Query = "<Where><Geq><FieldRef Name='Event_Date'/><Value Type='DateTime'><Today OffsetDays='-90'/></Value></Geq></Where>" }
            @{ Name = "Multilingual / Cultural"; Query = "<Where><Eq><FieldRef Name='Multilingual_Cultural'/><Value Type='Boolean'>1</Value></Eq></Where>" }
        )
    },
    @{
        Name = "Investors-Funders List"
        Description = "Grants, funder relationships, leveraged capital, governance (JSC/CAB)."
        Columns = @(
            @{ Name = "Funder_Name"; Type = "Text"; Display = "Funder Name"; Required = $true }
            @{ Name = "Activity_Type"; Type = "Choice"; Display = "Activity Type"; Required = $true
               Choices = @("Grant", "Partnership", "Sponsorship", "Leveraged Capital", "SROI Calculation", "Earned Revenue Snapshot", "JSC Meeting", "CAB Meeting", "CAB Member Added", "Professional Development") }
            @{ Name = "Amount_USD"; Type = "Currency"; Display = "Amount `$" }
            @{ Name = "Period_Start"; Type = "DateTime"; Display = "Period Start" }
            @{ Name = "Period_End"; Type = "DateTime"; Display = "Period End" }
            @{ Name = "Earned_Revenue_Pct"; Type = "Number"; Display = "Earned Revenue %" }
            @{ Name = "Leveraged_Capital_Source"; Type = "Text"; Display = "Leveraged Capital Source" }
            @{ Name = "SROI_Ratio"; Type = "Number"; Display = "SROI Ratio" }
            @{ Name = "Status"; Type = "Choice"; Display = "Status"; Default = "Active"
               Choices = @("Active", "Pending", "Closed") }
        )
        Views = @(
            @{ Name = "Active Grants"
               Query = "<Where><And><Eq><FieldRef Name='Activity_Type'/><Value Type='Choice'>Grant</Value></Eq><Eq><FieldRef Name='Status'/><Value Type='Choice'>Active</Value></Eq></And></Where>" }
        )
    },
    @{
        Name = "Farmers List"
        Description = "Farm relationships, produce moved, distribution, demographics. EXTERNAL SHARING ON."
        Columns = @(
            @{ Name = "Farm_Name"; Type = "Text"; Display = "Farm Name"; Required = $true }
            @{ Name = "Farm_Location"; Type = "Text"; Display = "Farm Location" }
            @{ Name = "Farm_Contact"; Type = "Text"; Display = "Farm Contact" }
            @{ Name = "Produce_Types"; Type = "Note"; Display = "Produce Types (one per line)" }
            @{ Name = "Produce_lbs"; Type = "Number"; Display = "Produce (lbs)" }
            @{ Name = "Produce_Value_USD"; Type = "Currency"; Display = "Produce Value `$" }
            @{ Name = "Origin_Region"; Type = "Text"; Display = "Origin Region" }
            @{ Name = "Distribution_Destinations"; Type = "Note"; Display = "Distribution Destinations" }
            @{ Name = "Value_Added_Product"; Type = "Boolean"; Display = "Value-Added Product" }
            @{ Name = "Value_Added_Description"; Type = "Text"; Display = "Value-Added Description" }
            @{ Name = "Demographic_Race_Ethnicity"; Type = "MultiChoice"; Display = "Demographic — Race/Ethnicity"
               Choices = @("Black/African American", "Hispanic/Latino", "Asian", "White", "Native American/Indigenous", "Pacific Islander", "Multiracial", "Prefer not to say") }
            @{ Name = "Demographic_Gender"; Type = "Choice"; Display = "Demographic — Gender"
               Choices = @("Woman", "Man", "Non-binary", "Prefer not to say") }
            @{ Name = "Demographic_Age_Band"; Type = "Choice"; Display = "Demographic — Age Band"
               Choices = @("Under 25", "25–34", "35–44", "45–54", "55–64", "65+", "Prefer not to say") }
        )
        Views = @(
            @{ Name = "Current Month"
               Query = "<Where><Geq><FieldRef Name='Reporting_Period_Start'/><Value Type='DateTime'><Today OffsetDays='-30'/></Value></Geq></Where>" }
            @{ Name = "Value-Added Participants"
               Query = "<Where><Eq><FieldRef Name='Value_Added_Product'/><Value Type='Boolean'>1</Value></Eq></Where>" }
        )
    },
    @{
        Name = "Environmental List"
        Description = "Food waste diversion, composting, food miles, food captured for demos."
        Columns = @(
            @{ Name = "Activity_Type"; Type = "Choice"; Display = "Activity Type"; Required = $true
               Choices = @("Food Waste Diverted", "Composted", "Food Captured for Demo", "Food Miles Snapshot") }
            @{ Name = "Activity_Date"; Type = "DateTime"; Display = "Activity Date"; Required = $true }
            @{ Name = "lbs_Value"; Type = "Number"; Display = "lbs Value" }
            @{ Name = "Dollar_Value"; Type = "Currency"; Display = "Dollar Value" }
            @{ Name = "Disposition"; Type = "Choice"; Display = "Disposition"
               Choices = @("Composted On-Site", "Composted Off-Site", "Donated", "Used in Demo", "Other") }
            @{ Name = "Food_Miles_Avg"; Type = "Number"; Display = "Food Miles Avg" }
            @{ Name = "Source_or_Recipient"; Type = "Text"; Display = "Source or Recipient" }
        )
        Views = @(
            @{ Name = "Current Quarter"
               Query = "<Where><Geq><FieldRef Name='Activity_Date'/><Value Type='DateTime'><Today OffsetDays='-90'/></Value></Geq></Where>" }
        )
    }
)

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

function Add-ChoiceField {
    param($ListTitle, $Name, $Display, $Choices, $Required, $Default, $Multi)
    $existing = Get-PnPField -List $ListTitle -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    - $Name (exists, skipped)"
        return
    }
    $fieldType = if ($Multi) { "MultiChoice" } else { "Choice" }
    $xmlChoices = ($Choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) -join ""
    $defaultXml = if ($Default) { "<Default>$Default</Default>" } else { "" }
    $reqAttr = if ($Required) { "TRUE" } else { "FALSE" }
    $schema = @"
<Field Type='$fieldType' Name='$Name' DisplayName='$Display' Required='$reqAttr' Format='Dropdown'>
  $defaultXml
  <CHOICES>$xmlChoices</CHOICES>
</Field>
"@
    if ($PSCmdlet.ShouldProcess("$ListTitle / $Name", "Create $fieldType field")) {
        Add-PnPFieldFromXml -List $ListTitle -FieldXml $schema | Out-Null
        Write-Host "    + $Name ($fieldType)"
    }
}

function Add-SimpleField {
    param($ListTitle, $Name, $Display, $Type, $Required, $Default)
    $existing = Get-PnPField -List $ListTitle -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "    - $Name (exists, skipped)"
        return
    }
    $params = @{
        List        = $ListTitle
        InternalName = $Name
        DisplayName = $Display
        Type        = $Type
    }
    if ($Required) { $params["Required"] = $true }
    if ($PSCmdlet.ShouldProcess("$ListTitle / $Name", "Create $Type field")) {
        Add-PnPField @params | Out-Null
        # Set default value for Boolean default of false (PnP default)
        if ($Default) {
            Set-PnPField -List $ListTitle -Identity $Name -Values @{ DefaultValue = $Default } | Out-Null
        }
        Write-Host "    + $Name ($Type)"
    }
}

function Add-Column {
    param($ListTitle, $Col)
    $name = $Col.Name
    $display = if ($Col.Display) { $Col.Display } else { $Col.Name }
    $type = $Col.Type
    $required = $Col.Required -eq $true
    $default = $Col.Default

    switch ($type) {
        "Choice" {
            Add-ChoiceField -ListTitle $ListTitle -Name $name -Display $display -Choices $Col.Choices -Required $required -Default $default -Multi $false
        }
        "MultiChoice" {
            Add-ChoiceField -ListTitle $ListTitle -Name $name -Display $display -Choices $Col.Choices -Required $required -Default $default -Multi $true
        }
        default {
            Add-SimpleField -ListTitle $ListTitle -Name $name -Display $display -Type $type -Required $required -Default $default
        }
    }
}

# -----------------------------------------------------------------------------
# Create or update each List
# -----------------------------------------------------------------------------

$summary = @()

foreach ($listSpec in $Lists) {
    $title = $listSpec.Name
    Write-Host ""
    Write-Host "List: $title" -ForegroundColor Cyan

    $existing = Get-PnPList -Identity $title -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  (List exists — adding missing columns only.)"
        $action = "Updated"
    } else {
        if ($PSCmdlet.ShouldProcess($title, "Create SharePoint List")) {
            New-PnPList -Title $title -Template GenericList -OnQuickLaunch | Out-Null
            Set-PnPList -Identity $title -Description $listSpec.Description -EnableAttachments $true | Out-Null
            Write-Host "  + List created."
            $action = "Created"
        }
    }

    Write-Host "  Common columns:"
    foreach ($col in $CommonColumns) {
        Add-Column -ListTitle $title -Col $col
    }

    Write-Host "  Per-list columns:"
    foreach ($col in $listSpec.Columns) {
        Add-Column -ListTitle $title -Col $col
    }

    if ($listSpec.Views) {
        Write-Host "  Views:"
        foreach ($view in $listSpec.Views) {
            $existingView = Get-PnPView -List $title -Identity $view.Name -ErrorAction SilentlyContinue
            if ($existingView) {
                Write-Host "    - $($view.Name) (exists, skipped)"
                continue
            }
            $fields = @("Title", "Submitted_By", "Submitted_Date", "Reporting_Period_Start", "Reporting_Period_End", "Lead_Role")
            if ($PSCmdlet.ShouldProcess("$title / $($view.Name)", "Create View")) {
                Add-PnPView -List $title -Title $view.Name -Fields $fields -Query $view.Query | Out-Null
                Write-Host "    + $($view.Name)"
            }
        }
    }

    $summary += [PSCustomObject]@{
        List = $title
        Action = $action
        ColumnsRequested = $CommonColumns.Count + $listSpec.Columns.Count
    }
}

# -----------------------------------------------------------------------------
# Summary
# -----------------------------------------------------------------------------

Write-Host ""
Write-Host "Summary" -ForegroundColor Green
$summary | Format-Table -AutoSize

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "  1. For the Farmers List, enable external sharing per the steps in"
Write-Host "     schemas/sharepoint_lists.md (External sharing setup section)."
Write-Host "  2. Build the five Microsoft Forms per schemas/microsoft_forms.md."
Write-Host "  3. Build the Power Automate flows per schemas/power_automate_flows.md."
Write-Host "  4. Open HARVEST_KPI_Master.xlsx, add Power Query connections using"
Write-Host "     the M code in docs/data_dictionary.md, set the HarvestSite"
Write-Host "     parameter to: $SiteUrl"
Write-Host ""

Disconnect-PnPOnline
