# Step 3: 1-year recommendation + current spend, merged with the 3-year data. -> SavingsPlan_1yr_3yr.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

$subsCsv = Import-Csv (Join-Path $folder 'PerSubCompute.csv') | Where-Object { [double]$_.EligibleCompute -ge 150 }
$threeYr = Import-Csv (Join-Path $folder 'SavingsPlanBySub.csv')
$out = foreach ($row in $subsCsv) {
    $sub = $row.Id
    $p = "/subscriptions/$sub/providers/Microsoft.CostManagement/benefitRecommendations?api-version=2024-08-01&`$filter=properties/lookBackPeriod eq 'Last30Days' and properties/term eq 'P1Y' and properties/scope eq 'Single'"
    try {
        $v = ((Invoke-AzRestMethod -Path $p -Method GET).Content | ConvertFrom-Json).value | Where-Object { $_.properties.armSkuName -eq 'Compute_Savings_Plan' } | Select-Object -First 1
        $oneYrAnnual = 0; $monthlyCost = 0
        if ($v) {
            $d = $v.properties.recommendationDetails
            $monthlyCost = [math]::Round([double]$v.properties.costWithoutBenefit,0)
            # 2024-08 API drops savingsAmount; derive from cost fields, fall back to legacy field if present.
            $names = $d.PSObject.Properties.Name
            $mSav = if (($names -contains 'savingsAmount') -and [double]$d.savingsAmount) { [double]$d.savingsAmount } else { [double]$v.properties.costWithoutBenefit - [double]$d.totalCost }
            $oneYrAnnual = [math]::Round($mSav * 12,0)
        }
        $t3 = $threeYr | Where-Object { $_.Subscription -eq $row.Name }
        [pscustomobject]@{ Subscription=$row.Name; MonthlyCost=$monthlyCost; Commit3=[double]$t3.CommitPerHour; Save1=$oneYrAnnual; Save3=[double]$t3.AnnualSavings }
    } catch {}
}
$out | Export-Csv (Join-Path $folder 'SavingsPlan_1yr_3yr.csv') -NoTypeInformation
Write-Host ("Saved SavingsPlan_1yr_3yr.csv - {0} rows" -f @($out).Count) -ForegroundColor Green
