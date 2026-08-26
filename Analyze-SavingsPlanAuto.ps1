# Auto savings-plan analysis - runs end to end from Config.ps1 alone (no BU mapping, no PowerPoint).
# For every subscription it: reads Azure's Compute Savings Plan recommendation (the same data the
# Advisor / Cost blade surfaces), grosses the recommended baseline up to LIST price using the
# negotiated discount, prices the covered/steady portion at Azure's own recommended savings-plan cost
# (benefitCost, which already reflects the negotiated savings-plan pricing), leaves the uncovered/burst
# portion at full list (the on-demand discount is forfeited on a plan), then verifies whether that
# beats staying on the current negotiated rate. -> SavingsPlanAutoAnalysis.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

function Get-SpRecommendation($subId, $term) {
    $p = "/subscriptions/$subId/providers/Microsoft.CostManagement/benefitRecommendations?api-version=2024-08-01&`$filter=properties/lookBackPeriod eq 'Last30Days' and properties/term eq '$term' and properties/scope eq 'Single'"
    try {
        $resp = Invoke-AzRestMethod -Path $p -Method GET -ErrorAction Stop
        if ($resp.StatusCode -ne 200) { return $null }
        return ($resp.Content | ConvertFrom-Json).value |
            Where-Object { $_.properties.armSkuName -eq 'Compute_Savings_Plan' } | Select-Object -First 1
    } catch { return $null }
}

# Steady-state coverage share from a recommendation. 2024-08 API drops coveragePercentage; derive from cost fields.
function Get-CoverageShare($d, $costNoBenefit) {
    $names = $d.PSObject.Properties.Name
    $cov = if (($names -contains 'coveragePercentage') -and [double]$d.coveragePercentage) { [double]$d.coveragePercentage / 100.0 }
           elseif ($costNoBenefit -gt 0) { 1 - ((([double]$d.totalCost) - ([double]$d.benefitCost)) / $costNoBenefit) }
           else { 0.0 }
    return [math]::Max(0.0, [math]::Min(1.0, $cov))
}

# Covered-portion savings-plan cost. Prefer Azure's own recommended benefitCost (already reflects the
# negotiated savings-plan pricing for this usage); fall back to the published-rate model only if missing.
function Get-CoveredSpCost($d, $coverage, $listMonthly, $fallbackSp, $spd) {
    $bc = [double]$d.benefitCost
    if ($bc -gt 0) { return $bc }
    return $coverage * $listMonthly * (1 - $fallbackSp) * (1 - $spd)
}

$subs = Get-AzSubscription -TenantId $tenant -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name
Write-Host ("Analyzing {0} enabled subscriptions in tenant {1} ..." -f $subs.Count, $tenant) -ForegroundColor Cyan

$results = foreach ($s in $subs) {
    $rec3 = Get-SpRecommendation $s.Id 'P3Y'
    if (-not $rec3) { continue }   # no compute savings-plan baseline for this sub

    $d3 = $rec3.properties.recommendationDetails
    $currentMonthly = [double]$rec3.properties.costWithoutBenefit   # today's spend at the negotiated (already-discounted) rate
    if ($currentMonthly -le 0) { continue }

    # Coverage (steady-state share) from the 3-year recommendation.
    $coverage = Get-CoverageShare $d3 $currentMonthly

    # Step 1: turn the recommended baseline into LIST price by removing the negotiated discount.
    $listMonthly = $currentMonthly / (1 - $retailDiscount)

    # Step 2/3: covered/steady portion is priced at Azure's own recommended savings-plan cost
    # (benefitCost already reflects the negotiated savings-plan pricing); uncovered/burst pays full list.
    # $standardSp*/$spd are used only as a fallback when a term has no usable benefitCost.
    $covered3 = Get-CoveredSpCost $d3 $coverage $listMonthly $standardSp3 $spd
    $threeYearMonthly = $covered3 + ((1 - $coverage) * $listMonthly)

    $rec1 = Get-SpRecommendation $s.Id 'P1Y'
    if ($rec1) {
        $d1       = $rec1.properties.recommendationDetails
        $curr1    = [double]$rec1.properties.costWithoutBenefit
        $cov1     = Get-CoverageShare $d1 $curr1
        $list1    = if ($curr1 -gt 0) { $curr1 / (1 - $retailDiscount) } else { $listMonthly }
        $covered1 = Get-CoveredSpCost $d1 $cov1 $list1 $standardSp1 $spd
        $oneYearMonthly = $covered1 + ((1 - $cov1) * $list1)
    } else {
        $covered1 = $coverage * $listMonthly * (1 - $standardSp1) * (1 - $spd)
        $oneYearMonthly = $covered1 + ((1 - $coverage) * $listMonthly)
    }

    # Effective discount off list on the covered portion (live; already includes the deal) - for transparency.
    $coveredList3    = $coverage * $listMonthly
    $covered3OffList = if ($coveredList3 -gt 0) { 1 - ($covered3 / $coveredList3) } else { 0 }

    # Step 4: verify against the current negotiated (40%) spend.
    $annual3 = ($currentMonthly - $threeYearMonthly) * 12
    $annual1 = ($currentMonthly - $oneYearMonthly)   * 12
    $verdict = if ($annual3 -gt 0) { '3-year Savings Plan' } elseif ($annual1 -gt 0) { '1-year Savings Plan' } else { 'Keep current negotiated rate' }

    [pscustomobject]@{
        Subscription       = $s.Name
        SubscriptionId     = $s.Id
        CurrentNegMonthly  = [math]::Round($currentMonthly, 0)
        ListMonthly        = [math]::Round($listMonthly, 0)
        SteadyCoveragePct  = [math]::Round($coverage * 100, 0)
        Covered3OffListPct = [math]::Round($covered3OffList * 100, 1)
        OneYearSpMonthly   = [math]::Round($oneYearMonthly, 0)
        ThreeYearSpMonthly = [math]::Round($threeYearMonthly, 0)
        Commit3PerHour     = [math]::Round([double]$d3.commitmentAmount, 2)
        AnnualSavings1yr   = [math]::Round($annual1, 0)
        AnnualSavings3yr   = [math]::Round($annual3, 0)
        Recommendation     = $verdict
    }
}

$results = @($results | Sort-Object AnnualSavings3yr -Descending)
$results | Export-Csv (Join-Path $folder 'SavingsPlanAutoAnalysis.csv') -NoTypeInformation

# Console rollup.
$wins = @($results | Where-Object { $_.Recommendation -like '*Savings Plan*' })
$totalNeg   = ($results | Measure-Object CurrentNegMonthly  -Sum).Sum
$totalThree = ($results | Measure-Object ThreeYearSpMonthly -Sum).Sum
$estAnnual  = ($results | Measure-Object AnnualSavings3yr   -Sum).Sum

Write-Host ""
Write-Host ("Model: current negotiated rate = {0:P0} off list; covered portion priced at Azure's own recommended savings-plan cost (benefitCost), uncovered at full list." -f $retailDiscount) -ForegroundColor DarkGray
$results | Format-Table Subscription,
    @{n='Current(neg)/mo';e={'{0:N0}' -f $_.CurrentNegMonthly}},
    @{n='List/mo';e={'{0:N0}' -f $_.ListMonthly}},
    @{n='Steady%';e={'{0:N0}%' -f $_.SteadyCoveragePct}},
    @{n='3yr SP/mo';e={'{0:N0}' -f $_.ThreeYearSpMonthly}},
    @{n='3yr save/yr';e={'{0:N0}' -f $_.AnnualSavings3yr}},
    Recommendation -AutoSize

Write-Host ("Subscriptions analyzed: {0}   Savings plan wins: {1}" -f $results.Count, $wins.Count) -ForegroundColor Green
Write-Host ("Estate current negotiated: {0:C0}/mo   optimal 3-yr blend: {1:C0}/mo   estimated annual savings: {2:C0}" -f $totalNeg, $totalThree, $estAnnual) -ForegroundColor Green
Write-Host "Saved SavingsPlanAutoAnalysis.csv" -ForegroundColor Green
