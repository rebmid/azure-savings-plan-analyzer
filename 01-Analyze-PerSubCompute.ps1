# Step 1: per-subscription savings-plan-eligible compute spend for the period.  -> PerSubCompute.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

$eligible = @('Virtual Machines','Azure App Service','Container Instances','Azure Container Apps','Azure Functions','Cloud Services')
$subs = Get-AzSubscription -TenantId $tenant -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Enabled' }
$body = @{ type='ActualCost'; timeframe='Custom'; timePeriod=@{ from=$periodFrom; to=$periodTo };
           dataset=@{ granularity='None'; aggregation=@{ Cost=@{ name='Cost'; function='Sum' } }; grouping=@(@{ type='Dimension'; name='ServiceName' }) } } | ConvertTo-Json -Depth 8

$rows = foreach ($s in $subs) {
    try {
        $resp = Invoke-AzRestMethod -Path "/subscriptions/$($s.Id)/providers/Microsoft.CostManagement/query?api-version=2023-11-01" -Method POST -Payload $body -ErrorAction Stop
        if ($resp.StatusCode -ne 200) { continue }
        $d = ($resp.Content | ConvertFrom-Json).properties
        $iCost = [array]::IndexOf($d.columns.name,'Cost'); $iSvc = [array]::IndexOf($d.columns.name,'ServiceName')
        $vm=0.0; $app=0.0; $cont=0.0
        foreach ($r in $d.rows) {
            $c=[double]$r[$iCost]; $svc=[string]$r[$iSvc]
            if ($svc -eq 'Virtual Machines'){ $vm+=$c }
            elseif ($svc -eq 'Azure App Service'){ $app+=$c }
            elseif ($eligible -contains $svc){ $cont+=$c }
        }
        $tot = $vm+$app+$cont
        if ($tot -gt 0) { [pscustomobject]@{ Name=$s.Name; Id=$s.Id; VM=[math]::Round($vm,0); AppSvc=[math]::Round($app,0); OtherCompute=[math]::Round($cont,0); EligibleCompute=[math]::Round($tot,0) } }
    } catch {}
}
$rows = $rows | Sort-Object EligibleCompute -Descending
$rows | Export-Csv (Join-Path $folder 'PerSubCompute.csv') -NoTypeInformation
Write-Host ("Saved PerSubCompute.csv - {0} subscriptions with eligible compute" -f @($rows).Count) -ForegroundColor Green
$rows | Format-Table Name, @{n='Eligible/mo'; e={ '{0:N0}' -f $_.EligibleCompute }} -AutoSize
