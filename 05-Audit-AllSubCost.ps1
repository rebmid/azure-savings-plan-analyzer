# Step 5: cost across EVERY readable subscription (last 30 days), for reconciliation. -> AllSubCostAudit.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

$subs = Get-AzSubscription -TenantId $tenant -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name
Write-Host ("Enumerated {0} readable subscriptions." -f $subs.Count) -ForegroundColor Cyan
$from = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd'); $to = (Get-Date).ToString('yyyy-MM-dd')

$rows = foreach ($s in $subs) {
    $id = $s.Id; $name = $s.Name
    $body = @{ type='ActualCost'; timeframe='Custom'; timePeriod=@{ from=$from; to=$to };
               dataset=@{ granularity='None'; aggregation=@{ totalCost=@{ name='Cost'; function='Sum' } }; grouping=@(@{ type='Dimension'; name='ServiceName' }) } } | ConvertTo-Json -Depth 8
    $vm=0.0; $comp=0.0; $tot=0.0; $st=0
    try {
        $resp = Invoke-AzRestMethod -Path "/subscriptions/$id/providers/Microsoft.CostManagement/query?api-version=2023-03-01" -Method POST -Payload $body; $st = $resp.StatusCode
        if ($st -eq 200) {
            foreach ($r in ($resp.Content | ConvertFrom-Json).properties.rows) {
                $c=[double]$r[0]; $svc=[string]$r[1]; $tot+=$c
                if ($svc -like 'Virtual Machines*'){ $vm+=$c }
                if ($svc -match 'Virtual Machines|Azure App Service|Functions|Container|Kubernetes|Virtual Machine Scale'){ $comp+=$c }
            }
        }
    } catch { $st = -1 }
    [pscustomobject]@{ Name=$name; Sub=$id; VM=[math]::Round($vm); Compute=[math]::Round($comp); Total=[math]::Round($tot); InDeck=$false; Status=$st }
}
$rows = $rows | Sort-Object Compute -Descending
$rows | Export-Csv (Join-Path $folder 'AllSubCostAudit.csv') -NoTypeInformation
Write-Host ("Saved AllSubCostAudit.csv - {0} subscriptions" -f @($rows).Count) -ForegroundColor Green
