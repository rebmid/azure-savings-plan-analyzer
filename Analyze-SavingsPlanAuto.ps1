# Auto savings-plan analysis - runs end to end from Config.ps1 alone (no BU mapping, no PowerPoint).
# For every subscription it: reads Azure's Compute Savings Plan recommendation (the same data the
# Advisor / Cost blade surfaces), grosses the recommended baseline up to LIST price using the
# negotiated discount, prices the covered/steady portion at the published savings-plan rate PLUS
# the 32.5% deal discount (both stacked off list), leaves the uncovered/burst portion at full list
# (the on-demand discount is forfeited on a plan), then verifies whether that beats staying on the
# current negotiated rate. -> SavingsPlanAutoAnalysis.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

# Covered-portion price factors off list: savings-plan rate AND the 32.5% deal discount, multiplicative.
$sp1Factor = (1 - $standardSp1) * (1 - $spd)   # 1-year
$sp3Factor = (1 - $standardSp3) * (1 - $spd)   # 3-year

function Get-SpRecommendation($subId, $term) {
    $p = "/subscriptions/$subId/providers/Microsoft.CostManagement/benefitRecommendations?api-version=2024-08-01&`$filter=properties/lookBackPeriod eq 'Last30Days' and properties/term eq '$term' and properties/scope eq 'Single'"
    try {
        $resp = Invoke-AzRestMethod -Path $p -Method GET -ErrorAction Stop
        if ($resp.StatusCode -ne 200) { return $null }
        return ($resp.Content | ConvertFrom-Json).value |
            Where-Object { $_.properties.armSkuName -eq 'Compute_Savings_Plan' } | Select-Object -First 1
    } catch { return $null }
}

$subs = Get-AzSubscription -TenantId $tenant -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name
Write-Host ("Analyzing {0} enabled subscriptions in tenant {1} ..." -f $subs.Count, $tenant) -ForegroundColor Cyan

$results = foreach ($s in $subs) {
    $rec3 = Get-SpRecommendation $s.Id 'P3Y'
    if (-not $rec3) { continue }   # no compute savings-plan baseline for this sub

    $d3 = $rec3.properties.recommendationDetails
    $currentMonthly = [double]$rec3.properties.costWithoutBenefit   # today's spend at the negotiated (already-discounted) rate
    if ($currentMonthly -le 0) { continue }

    # Coverage (steady-state share). 2024-08 API drops coveragePercentage; derive from cost fields, fall back if present.
    $names3 = $d3.PSObject.Properties.Name
    $coverage = if (($names3 -contains 'coveragePercentage') -and [double]$d3.coveragePercentage) { [double]$d3.coveragePercentage / 100.0 }
                else { $bc = [double]$d3.benefitCost; $tc = [double]$d3.totalCost; 1 - (($tc - $bc) / $currentMonthly) }
    $coverage = [math]::Max(0.0, [math]::Min(1.0, $coverage))

    # Step 1: turn the recommended baseline into LIST price by removing the negotiated discount.
    $listMonthly = $currentMonthly / (1 - $retailDiscount)

    # Step 2/3: price the plan off list - covered portion gets SP rate + 32.5%; uncovered pays full list.
    $threeYearMonthly = $listMonthly * (($coverage * $sp3Factor) + (1 - $coverage))
    $oneYearMonthly   = $listMonthly * (($coverage * $sp1Factor) + (1 - $coverage))

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
        OneYearSpMonthly   = [math]::Round($oneYearMonthly, 0)
        ThreeYearSpMonthly = [math]::Round($threeYearMonthly, 0)
        Commit3PerHour     = [math]::Round(($coverage * $listMonthly * $sp3Factor) / 730, 2)
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
Write-Host ("Model: current negotiated rate = {0:P0} off list; savings plan priced off list at the published SP rate + {1:P1} deal discount (both stacked)." -f $retailDiscount, $spd) -ForegroundColor DarkGray
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
