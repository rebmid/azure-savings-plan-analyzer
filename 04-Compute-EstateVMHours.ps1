# Step 4: VM instance-hours per subscription for the period. -> VMHoursBySub.csv
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')

$subs = Import-Csv (Join-Path $folder 'PerSubCompute.csv') | Where-Object { [double]$_.VM -gt 0 }
$body = @{ type='ActualCost'; timeframe='Custom'; timePeriod=@{ from=$periodFrom; to=$periodTo };
           dataset=@{ granularity='None'; aggregation=([ordered]@{ Hours=@{ name='UsageQuantity'; function='Sum' } });
                      filter=@{ dimensions=@{ name='MeterCategory'; operator='In'; values=@('Virtual Machines') } } } } | ConvertTo-Json -Depth 12

$rows = foreach ($s in $subs) {
    try {
        $r = (Invoke-AzRestMethod -Path "/subscriptions/$($s.Id)/providers/Microsoft.CostManagement/query?api-version=2023-11-01" -Method POST -Payload $body).Content | ConvertFrom-Json
        $h = 0.0; if ($r.properties.rows.Count -gt 0) { $h = [double]$r.properties.rows[0][0] }
        if ($h -gt 0) { [pscustomobject]@{ BU=(Get-BU $s.Name); Sub=$s.Name; Hours=[math]::Round($h,0); PerDay=[math]::Round($h/30,0); AvgVMs=[math]::Round($h/720,0) } }
    } catch {}
}
$rows = $rows | Sort-Object Hours -Descending
$rows | Export-Csv (Join-Path $folder 'VMHoursBySub.csv') -NoTypeInformation
Write-Host ("Saved VMHoursBySub.csv - {0} subscriptions" -f @($rows).Count) -ForegroundColor Green
