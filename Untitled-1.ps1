<#
.SYNOPSIS
  CIS Azure Foundations Benchmark 1.3.0 (Azure Government) compliance pull via Azure Policy.

.DESCRIPTION
  - Connects to AzureUSGovernment
  - Enumerates subscriptions
  - Locates built-in CIS Azure Foundations Benchmark 1.3.0 (Azure Government) initiative definition
  - Finds policy assignments using that initiative
  - Pulls policy state summary + optional detailed noncompliant records
  - Exports CSVs

.NOTES
  This reports what Azure Policy can evaluate. Some CIS items require Entra ID / manual validation.
#>

param(
  [switch]$IncludeNonCompliantDetails,
  [int]$NonCompliantDetailMaxPerSub = 500,
  [string]$OutDir = ".\cis_gov_output"
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

function Ensure-AzModules {
  $needed = @("Az.Accounts","Az.Resources","Az.PolicyInsights")
  foreach ($m in $needed) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
      throw "Missing module [$m]. Install with: Install-Module $m -Scope CurrentUser"
    }
  }
}

function Get-CisGovInitiativeDefinition {
  # Microsoft documents the Gov-specific CIS 1.3.0 initiative mapping. :contentReference[oaicite:3]{index=3}
  # We locate the built-in policy set definition by display name.
  $defs = Get-AzPolicySetDefinition -Builtin
  $match = $defs | Where-Object {
    $_.Properties.DisplayName -match "CIS Microsoft Azure Foundations Benchmark 1\.3\.0" -and
    $_.Properties.DisplayName -match "Government|Gov|Azure Government|Regulatory Compliance"
  } | Select-Object -First 1

  if (-not $match) {
    # Fallback: match just on CIS + 1.3.0 in display name
    $match = $defs | Where-Object {
      $_.Properties.DisplayName -match "CIS Microsoft Azure Foundations Benchmark 1\.3\.0"
    } | Select-Object -First 1
  }

  return $match
}

function Get-CisAssignmentsForSubscription {
  param(
    [Parameter(Mandatory)] [string]$SubscriptionId,
    [Parameter(Mandatory)] [string]$PolicySetDefinitionId
  )

  Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

  $assignments = Get-AzPolicyAssignment -Scope "/subscriptions/$SubscriptionId" -ErrorAction SilentlyContinue
  $cisAssignments = $assignments | Where-Object {
    $_.Properties.PolicyDefinitionId -eq $PolicySetDefinitionId
  }

  return $cisAssignments
}

function Get-AssignmentComplianceSummary {
  param(
    [Parameter(Mandatory)] [string]$AssignmentId,
    [Parameter(Mandatory)] [string]$SubscriptionId
  )

  # Summary counts by state for the assignment (compliant/noncompliant/unknown/exempt)
  $summary = Get-AzPolicyStateSummary -SubscriptionId $SubscriptionId -PolicyAssignmentId $AssignmentId

  # The cmdlet returns nested structures; normalize the key counts
  $results = [ordered]@{
    SubscriptionId = $SubscriptionId
    PolicyAssignmentId = $AssignmentId
    TimestampUtc = (Get-Date).ToUniversalTime().ToString("s") + "Z"
    Compliant = $null
    NonCompliant = $null
    Unknown = $null
    NotStarted = $null
    Exempt = $null
    Conflicting = $null
    Error = $null
  }

  try {
    # summary.PolicyAssignments is usually an array; pick the row for this assignment
    $pa = $summary.PolicyAssignments | Where-Object { $_.PolicyAssignmentId -eq $AssignmentId } | Select-Object -First 1
    if ($pa) {
      foreach ($s in $pa.Results) {
        switch ($s.PolicyState) {
          "Compliant"    { $results.Compliant    = $s.ResultsCount }
          "NonCompliant" { $results.NonCompliant = $s.ResultsCount }
          "Unknown"      { $results.Unknown      = $s.ResultsCount }
          "NotStarted"   { $results.NotStarted   = $s.ResultsCount }
          "Exempt"       { $results.Exempt       = $s.ResultsCount }
          "Conflicting"  { $results.Conflicting  = $s.ResultsCount }
          "Error"        { $results.Error        = $s.ResultsCount }
        }
      }
    }
  } catch {
    # Keep the raw error but don't stop the entire run
    $results.Error = $_.Exception.Message
  }

  [pscustomobject]$results
}

function Get-NonCompliantDetails {
  param(
    [Parameter(Mandatory)] [string]$SubscriptionId,
    [Parameter(Mandatory)] [string]$AssignmentId,
    [int]$Top = 500
  )

  # Pull noncompliant policy states for the assignment
  # Note: this can be large; keep a sane cap.
  $states = Get-AzPolicyState `
    -SubscriptionId $SubscriptionId `
    -Filter "PolicyAssignmentId eq '$AssignmentId' and ComplianceState eq 'NonCompliant'" `
    -Top $Top

  $states | Select-Object `
    @{n="SubscriptionId";e={$SubscriptionId}},
    @{n="PolicyAssignmentId";e={$AssignmentId}},
    Timestamp,
    ResourceId,
    ResourceType,
    PolicyDefinitionName,
    PolicyDefinitionId,
    PolicySetDefinitionName,
    PolicySetDefinitionId,
    PolicyAssignmentName,
    PolicyAssignmentId,
    ComplianceState
}

# ------------------- Main -------------------

Ensure-AzModules
Connect-Gov

$cisInitiative = Get-CisGovInitiativeDefinition
if (-not $cisInitiative) {
  throw "Could not find the built-in CIS Microsoft Azure Foundations Benchmark 1.3.0 initiative in this tenant (AzureUSGovernment)."
}

$cisPolicySetId = $cisInitiative.PolicySetDefinitionId

$subs = Get-AzSubscription
$allSummaries = @()
$allMissingAssignments = @()
$allNonCompliant = @()

foreach ($sub in $subs) {
  $subId = $sub.Id

  $cisAssignments = Get-CisAssignmentsForSubscription -SubscriptionId $subId -PolicySetDefinitionId $cisPolicySetId

  if (-not $cisAssignments -or $cisAssignments.Count -eq 0) {
    $allMissingAssignments += [pscustomobject]@{
      SubscriptionId = $subId
      SubscriptionName = $sub.Name
      Note = "No CIS 1.3.0 (Azure Government) initiative assignment found at subscription scope."
    }
    continue
  }

  foreach ($a in $cisAssignments) {
    $summary = Get-AssignmentComplianceSummary -AssignmentId $a.PolicyAssignmentId -SubscriptionId $subId
    $allSummaries += $summary

    if ($IncludeNonCompliantDetails.IsPresent) {
      $details = Get-NonCompliantDetails -SubscriptionId $subId -AssignmentId $a.PolicyAssignmentId -Top $NonCompliantDetailMaxPerSub
      $allNonCompliant += $details
    }
  }
}

$summaryPath = Join-Path $OutDir "cis_gov_policy_compliance_summary.csv"
$missingPath = Join-Path $OutDir "cis_gov_missing_assignments.csv"
$detailPath  = Join-Path $OutDir "cis_gov_noncompliant_details.csv"

$allSummaries | Export-Csv -NoTypeInformation -Path $summaryPath
$allMissingAssignments | Export-Csv -NoTypeInformation -Path $missingPath

if ($IncludeNonCompliantDetails.IsPresent -and $allNonCompliant.Count -gt 0) {
  $allNonCompliant | Export-Csv -NoTypeInformation -Path $detailPath
}

Write-Host "Done."
Write-Host "Summary: $summaryPath"
Write-Host "Missing assignments: $missingPath"
if ($IncludeNonCompliantDetails.IsPresent) { Write-Host "Noncompliant details: $detailPath" }

Write-Host ""
Write-Host "Initiative used:"
Write-Host "  DisplayName: $($cisInitiative.Properties.DisplayName)"
Write-Host "  PolicySetDefinitionId: $cisPolicySetId"

