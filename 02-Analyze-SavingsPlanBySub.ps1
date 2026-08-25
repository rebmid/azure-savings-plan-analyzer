# Step 2: 3-year Compute Savings Plan recommendation (30-day look-back) per subscription. -> SavingsPlanBySub.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

$targets = Import-Csv (Join-Path $folder 'PerSubCompute.csv') | Where-Object { [double]$_.EligibleCompute -ge 150 }
$results = foreach ($t in $targets) {
    $sub = $t.Id
    $p = "/subscriptions/$sub/providers/Microsoft.CostManagement/benefitRecommendations?api-version=2024-08-01&`$filter=properties/lookBackPeriod eq 'Last30Days' and properties/term eq 'P3Y' and properties/scope eq 'Single'"
    try {
        $resp = Invoke-AzRestMethod -Path $p -Method GET -ErrorAction Stop
        if ($resp.StatusCode -ne 200) { continue }
        $rec = ($resp.Content | ConvertFrom-Json).value | Where-Object { $_.properties.armSkuName -eq 'Compute_Savings_Plan' } | Select-Object -First 1
        if (-not $rec) { continue }
        $d = $rec.properties.recommendationDetails
        $costNoBenefit = [double]$rec.properties.costWithoutBenefit
        $benefitCost   = [double]$d.benefitCost
        $totalCost     = [double]$d.totalCost
        # 2024-08 API drops coveragePercentage/savingsAmount; derive from cost fields, fall back to legacy fields if present.
        $names = $d.PSObject.Properties.Name
        $monthlySavings = if (($names -contains 'savingsAmount') -and [double]$d.savingsAmount) { [double]$d.savingsAmount } else { $costNoBenefit - $totalCost }
        $coverage = if (($names -contains 'coveragePercentage') -and [double]$d.coveragePercentage) { [double]$d.coveragePercentage }
                    elseif ($costNoBenefit -gt 0) { 100 * (1 - (($totalCost - $benefitCost) / $costNoBenefit)) } else { 0 }
        $coverage = [math]::Max(0.0, [math]::Min(100.0, $coverage))
        $savingsPct = if ($costNoBenefit -gt 0) { 100 * $monthlySavings / $costNoBenefit } else { 0 }
        [pscustomobject]@{
            BU             = Get-BU $t.Name
            Subscription   = $t.Name
            MonthlyCompute = [math]::Round($costNoBenefit,0)
            CommitPerHour  = [math]::Round([double]$d.commitmentAmount,2)
            Coverage       = [math]::Round($coverage,0)
            MonthlySavings = [math]::Round($monthlySavings,0)
            AnnualSavings  = [math]::Round($monthlySavings * 12,0)
            SavingsPct     = [math]::Round($savingsPct,0)
        }
    } catch {}
}
$results = $results | Sort-Object AnnualSavings -Descending
$results | Export-Csv (Join-Path $folder 'SavingsPlanBySub.csv') -NoTypeInformation
Write-Host ("Saved SavingsPlanBySub.csv - {0} subscriptions with a 3-year recommendation" -f @($results).Count) -ForegroundColor Green
