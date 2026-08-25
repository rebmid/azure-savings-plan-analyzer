# Step 7: build the deck - slide 2 summary + one detail slide per business unit.
# Business units, colors and slide format are generic; all inputs come from Config.ps1.
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')
$deck = Join-Path $folder $deckName

if ($businessUnits.Keys.Count -eq 0) {
    Write-Host "Config.ps1 `$businessUnits is empty. Fill it in (see README STEP 3) before running 07." -ForegroundColor Red
    return
}
if (-not (Test-Path $deck)) {
    Write-Host "Deck not found: $deck  - run 06-Add-MethodologySlide.ps1 first." -ForegroundColor Red
    return
}

$audit = Import-Csv (Join-Path $folder 'AllSubCostAudit.csv')
$savedRecommendations = Import-Csv (Join-Path $folder 'SavingsPlan_1yr_3yr.csv')
$vmHours = Import-Csv (Join-Path $folder 'VMHoursBySub.csv')
$bySub = Import-Csv (Join-Path $folder 'SavingsPlanBySub.csv')

function New-Lookup($rows, $property) {
    $lookup = @{}
    foreach ($row in $rows) { $lookup[[string]$row.$property] = $row }
    return $lookup
}

# Optional prior-access baseline (flags newly-readable subscriptions).
$previousByName = @{}
if ($priorSubsCsv -and (Test-Path (Join-Path $folder $priorSubsCsv))) {
    $previousByName = New-Lookup (Import-Csv (Join-Path $folder $priorSubsCsv)) 'SUBSCRIPTION NAME'
}
$auditByName = New-Lookup $audit 'Name'
$recommendationByName = New-Lookup $savedRecommendations 'Subscription'
$hoursByName = New-Lookup $vmHours 'Sub'
$coverageOverride = @{}; foreach ($row in $bySub) { $coverageOverride[[string]$row.Subscription] = [double]$row.Coverage / 100.0 }

$sp1Factor = (1 - $standardSp1) * (1 - $spd)
$sp3Factor = (1 - $standardSp3) * (1 - $spd)

$detail = foreach ($businessUnit in $businessUnits.Keys) {
    foreach ($name in $businessUnits[$businessUnit]) {
        $auditRow = $auditByName[$name]
        $recommendation = $recommendationByName[$name]
        $previouslyVisible = $previousByName.ContainsKey($name)
        $currentlyReadable = $null -ne $auditRow -and [string]$auditRow.Status -eq '200'
        $hasModel = $null -ne $recommendation -and [double]$recommendation.MonthlyCost -gt 0
        $currentMonthly = if ($hasModel) { [double]$recommendation.MonthlyCost } elseif ($auditRow) { [double]$auditRow.Compute } else { 0.0 }
        $coverage = 0.0; $oneYearMonthly = $null; $threeYearMonthly = $null; $annualSavings = $null
        $publicRetail = $null; $commitOne = $null; $commitThree = $null
        $recommendationText = 'No eligible compute baseline'

        if ($hasModel) {
            if ($coverageOverride.ContainsKey($name)) { $coverage = $coverageOverride[$name] }
            else { $coverage = (([double]$recommendation.Save3 / 12) + ([double]$recommendation.Commit3 * 730)) / $currentMonthly }
            $coverage = [math]::Max(0.0, [math]::Min(1.0, $coverage))
            $publicRetail = $currentMonthly / (1 - $retailDiscount)
            $oneYearMonthly = $publicRetail * (($coverage * $sp1Factor) + (1 - $coverage))
            $threeYearMonthly = $publicRetail * (($coverage * $sp3Factor) + (1 - $coverage))
            $annualSavings = ($currentMonthly - $threeYearMonthly) * 12
            $commitOne = ($coverage * $publicRetail * $sp1Factor) / 730
            $commitThree = ($coverage * $publicRetail * $sp3Factor) / 730
            $recommendationText = if ($annualSavings -gt 0) { '3-year Savings Plan' } else { 'Keep current 40%' }
        }

        [pscustomobject]@{
            BusinessUnit = $businessUnit
            Subscription = $name
            SubscriptionId = if ($auditRow) { $auditRow.Sub } elseif ($previousByName[$name]) { $previousByName[$name].'SUBSCRIPTION ID' } else { '' }
            PreviouslyVisible = $previouslyVisible
            CurrentlyReadable = $currentlyReadable
            NewlyVisible = $currentlyReadable -and -not $previouslyVisible
            HasSavingsBaseline = $hasModel
            VMHoursPerMonth = if ($hoursByName[$name]) { [double]$hoursByName[$name].Hours } else { 0.0 }
            SteadyCoveragePct = if ($hasModel) { [math]::Round($coverage * 100, 0) } else { $null }
            Current40Monthly = [math]::Round($currentMonthly, 0)
            OneYearSpMonthly = if ($hasModel) { [math]::Round($oneYearMonthly, 0) } else { $null }
            ThreeYearSpMonthly = if ($hasModel) { [math]::Round($threeYearMonthly, 0) } else { $null }
            AnnualSavingsVs40 = if ($hasModel) { [math]::Round($annualSavings, 0) } else { $null }
            PublicPaygMonthly = if ($hasModel) { [math]::Round($publicRetail, 0) } else { $null }
            CommitOne = if ($hasModel) { [math]::Round($commitOne, 2) } else { $null }
            CommitThree = if ($hasModel) { [math]::Round($commitThree, 2) } else { $null }
            Recommendation = $recommendationText
        }
    }
}

$summary = foreach ($businessUnit in $businessUnits.Keys) {
    $rows = @($detail | Where-Object BusinessUnit -eq $businessUnit)
    $modeled = @($rows | Where-Object HasSavingsBaseline)
    [pscustomobject]@{
        BusinessUnit = $businessUnit
        MappedSubscriptions = $rows.Count
        CurrentlyReadable = @($rows | Where-Object CurrentlyReadable).Count
        NewlyVisible = @($rows | Where-Object NewlyVisible).Count
        SavingsBaselines = $modeled.Count
        VMHoursPerMonth = [math]::Round(($modeled | Measure-Object VMHoursPerMonth -Sum).Sum, 0)
        Current40Monthly = [math]::Round(($modeled | Measure-Object Current40Monthly -Sum).Sum, 0)
        OneYearSpMonthly = [math]::Round(($modeled | Measure-Object OneYearSpMonthly -Sum).Sum, 0)
        ThreeYearSpMonthly = [math]::Round(($modeled | Measure-Object ThreeYearSpMonthly -Sum).Sum, 0)
        AnnualSavingsVs40 = [math]::Round(($modeled | Measure-Object AnnualSavingsVs40 -Sum).Sum, 0)
    }
}

$detail | Export-Csv (Join-Path $folder 'BusinessUnitSavingsDetail.csv') -NoTypeInformation
$summary | Export-Csv (Join-Path $folder 'BusinessUnitSavingsBreakout.csv') -NoTypeInformation
$detail | Select-Object BusinessUnit,Subscription,SubscriptionId,PreviouslyVisible,CurrentlyReadable,NewlyVisible,HasSavingsBaseline |
    Export-Csv (Join-Path $folder 'BusinessUnitAccessReconciliation.csv') -NoTypeInformation

$navy = 31+56*256+100*65536
$blue = 46+117*256+181*65536
$green = 33+115*256+70*65536
$white = 16777215
$dark = 38+38*256+38*65536
$gray = 242+242*256+242*65536

# Business-unit colors assigned by position, so any BU names work.
$palette = @(
    @{ detail = 214+226*256+240*65536; header = 79+129*256+189*65536;  darkHeader = $true  },  # blue
    @{ detail = 255+228*256+196*65536; header = 244+177*256+131*65536; darkHeader = $false },  # orange
    @{ detail = 214+235*256+205*65536; header = 169+208*256+142*65536; darkHeader = $false },  # green
    @{ detail = 250+219*256+205*65536; header = 244+150*256+120*65536; darkHeader = $false },  # coral
    @{ detail = 221+229*256+240*65536; header = 110+130*256+170*65536; darkHeader = $true  },  # slate
    @{ detail = 235+225*256+245*65536; header = 150+120*256+175*65536; darkHeader = $true  }   # purple
)
$buDetail = @{}; $buHeader = @{}; $buDarkHeader = @{}
$bi = 0
foreach ($bu in $businessUnits.Keys) { $p = $palette[$bi % $palette.Count]; $buDetail[$bu]=$p.detail; $buHeader[$bu]=$p.header; $buDarkHeader[$bu]=$p.darkHeader; $bi++ }

$amberFill = 252+228*256+214*65536
$amberText = 156+40*256+30*65536
$maxDetail = 14
$HDR = @('Subscription name | Business unit','VM-hrs /mo','Steady %','40% now  $/mo & $/hr','Public Azure PAYG /mo','1-yr SP +32.5% /mo','3-yr SP +32.5% /mo','Commit $/hr  1yr / 3yr','Recommend','Annual savings (3-yr plan)')
function F($n){ '${0:N0}' -f [double]$n }
function N($n){ '{0:N0}' -f [double]$n }
function PC($n){ '{0:N0}%' -f (100*[double]$n) }
function F2($n){ '${0:N2}' -f [double]$n }
function Imp($d){ if([double]$d -ge 0){ (F $d)+'/yr' } else { '('+(F ([math]::Abs([double]$d)))+' more)/yr' } }
function Rate40($mo){ if([double]$mo -gt 0){ (F $mo)+"`r"+(F2 ([double]$mo/730))+'/hr' } else { '-' } }
$msoTrue = -1; $msoFalse = 0; $ppLayoutBlank = 12; $horizontalText = 1

$powerPoint = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $powerPoint.Presentations.Open($deck, $false, $false, $false)
    $width = $presentation.PageSetup.SlideWidth
    $height = $presentation.PageSetup.SlideHeight

    function Add-Text($slide, $left, $top, $boxWidth, $boxHeight, $text, $size, $bold, $color, $alignment) {
        $box = $slide.Shapes.AddTextbox($horizontalText, [int]$left, [int]$top, [int]$boxWidth, [int]$boxHeight)
        $box.TextFrame.WordWrap = $msoTrue
        $range = $box.TextFrame.TextRange
        $range.Text = [string]$text; $range.Font.Name = 'Segoe UI'; $range.Font.Size = [single]$size
        $range.Font.Bold = if ($bold) { $msoTrue } else { $msoFalse }; $range.Font.Color.RGB = $color
        if ($alignment) { $range.ParagraphFormat.Alignment = $alignment }
        return $box
    }
    function Add-Rect($slide, $left, $top, $boxWidth, $boxHeight, $fill) {
        $shape = $slide.Shapes.AddShape(1, [int]$left, [int]$top, [int]$boxWidth, [int]$boxHeight)
        $shape.Fill.ForeColor.RGB = $fill; $shape.Line.Visible = $msoFalse; return $shape
    }
    function Add-Header($slide, $title, $kicker) {
        Add-Rect $slide 0 0 $width ([int]($height * 0.155)) $navy | Out-Null
        Add-Rect $slide 0 ([int]($height * 0.155)) $width 4 $blue | Out-Null
        Add-Text $slide ([int]($width * 0.035)) ([int]($height * 0.028)) ([int]($width * 0.93)) ([int]($height * 0.06)) $kicker 12 $true $white 1 | Out-Null
        Add-Text $slide ([int]($width * 0.035)) ([int]($height * 0.072)) ([int]($width * 0.93)) ([int]($height * 0.075)) $title 19 $true $white 1 | Out-Null
    }
    function Render-Table($slide, $rc, $rk) {
        $nRows = $rc.Count
        $tw = [int]($width * 0.975); $tx = [int]($width * 0.0125); $ty = [int]($height * 0.17); $th = [int]($height * 0.64)
        $tbl = $slide.Shapes.AddTable($nRows, 10, $tx, $ty, $tw, $th).Table
        $fr = 0.215,0.072,0.058,0.088,0.085,0.09,0.09,0.10,0.086,0.116
        for ($c = 1; $c -le 10; $c++) { $tbl.Columns[$c].Width = [int]($tw * $fr[$c-1]) }
        $fontSize = if ($nRows -gt 14) { 6 } else { 8 }
        for ($i = 0; $i -lt $nRows; $i++) {
            $rr = $i + 1; $cells = $rc[$i]; $k = $rk[$i]
            for ($c = 1; $c -le 10; $c++) {
                $cell = $tbl.Cell($rr, $c); $tr = $cell.Shape.TextFrame.TextRange
                $tr.Text = [string]$cells[$c-1]; $tr.Font.Name = 'Segoe UI'; $tr.Font.Size = [single]$fontSize
                $cell.Shape.TextFrame.MarginTop = 1; $cell.Shape.TextFrame.MarginBottom = 1
                if ($c -ge 2) { $tr.ParagraphFormat.Alignment = 3 } else { $tr.ParagraphFormat.Alignment = 1 }
                if ($k -like 'buheader|*') {
                    $bu = $k.Substring(9); $cell.Shape.Fill.ForeColor.RGB = $buHeader[$bu]; $tr.Font.Bold = $msoTrue; $tr.Font.Color.RGB = $(if ($buDarkHeader[$bu]) { $white } else { $dark })
                }
                elseif ($k -like 'budetail|*') {
                    $bu = $k.Substring(9); $cell.Shape.Fill.ForeColor.RGB = $buDetail[$bu]; $tr.Font.Color.RGB = $dark
                }
                else {
                    switch ($k) {
                        'hdr'   { $cell.Shape.Fill.ForeColor.RGB = $navy; $tr.Font.Bold = $msoTrue; $tr.Font.Color.RGB = $white }
                        'amber' { $cell.Shape.Fill.ForeColor.RGB = $amberFill; $tr.Font.Color.RGB = $amberText }
                        'total' { $cell.Shape.Fill.ForeColor.RGB = $green; $tr.Font.Bold = $msoTrue; $tr.Font.Color.RGB = $white }
                    }
                }
            }
        }
    }
    function Aggregate($modeled) {
        $payg = ($modeled | Measure-Object PublicPaygMonthly -Sum).Sum
        $covW = if ($payg -gt 0) { (($modeled | ForEach-Object { ($_.SteadyCoveragePct/100.0) * $_.PublicPaygMonthly } | Measure-Object -Sum).Sum) / $payg } else { 0 }
        [pscustomobject]@{
            VMh = ($modeled | Measure-Object VMHoursPerMonth -Sum).Sum
            Cov = $covW
            A = ($modeled | Measure-Object Current40Monthly -Sum).Sum
            Payg = $payg
            One = ($modeled | Measure-Object OneYearSpMonthly -Sum).Sum
            Three = ($modeled | Measure-Object ThreeYearSpMonthly -Sum).Sum
            C1 = ($modeled | Measure-Object CommitOne -Sum).Sum
            C3 = ($modeled | Measure-Object CommitThree -Sum).Sum
            Annual = ($modeled | Measure-Object AnnualSavingsVs40 -Sum).Sum
        }
    }
    function Add-Footnote($slide, $text) {
        $fy = [int]($height * 0.83)
        Add-Rect $slide ([int]($width*0.0175)) $fy ([int]($width*0.965)) ([int]($height*0.15)) $gray | Out-Null
        Add-Rect $slide ([int]($width*0.0175)) $fy 6 ([int]($height*0.15)) $green | Out-Null
        Add-Text $slide ([int]($width*0.03)) ([int]($fy+$height*0.006)) ([int]($width*0.94)) ([int]($height*0.14)) $text 10 $false $dark 1 | Out-Null
    }
    function Render-Summary($slide) {
        Add-Header $slide 'Where to Use the Savings Plan vs. the 40% - Summary by Business Unit' 'WHERE THE SAVINGS PLAN WINS vs. THE 40%  |  SAVINGS-PLAN RATE + 32.5% OFF PAYG  |  ESTATE ROLLUP'
        $rc = New-Object System.Collections.ArrayList; $rk = New-Object System.Collections.ArrayList
        [void]$rc.Add($HDR); [void]$rk.Add('hdr')
        foreach ($bu in $businessUnits.Keys) {
            $all = @($detail | Where-Object BusinessUnit -eq $bu)
            $modeled = @($all | Where-Object HasSavingsBaseline)
            if (-not $modeled) { continue }
            $g = Aggregate $modeled
            [void]$rc.Add(@(("$bu  ($($all.Count) subscriptions)"),(N $g.VMh),(PC $g.Cov),(Rate40 $g.A),(F $g.Payg),(F $g.One),(F $g.Three),((F2 $g.C1)+' / '+(F2 $g.C3)),'Savings Plan',(Imp $g.Annual)))
            [void]$rk.Add("buheader|$bu")
        }
        $t = Aggregate @($detail | Where-Object HasSavingsBaseline)
        [void]$rc.Add(@('ESTATE TOTAL - optimal blend (3-year, all business units)',(N $t.VMh),(PC $t.Cov),(Rate40 $t.A),(F $t.Payg),(F $t.One),(F $t.Three),((F2 $t.C1)+' / '+(F2 $t.C3)),'Savings Plan',(Imp $t.Annual)))
        [void]$rk.Add('total')
        Render-Table $slide $rc $rk
        Add-Footnote $slide ("Business-unit rollup of the columns shown per subscription on the following slides. '40% now' is today's spend under the existing negotiated deal. 'Public Azure PAYG' is the full list rate. Under a savings plan the negotiated on-demand discount is forfeited and replaced by 32.5% off PAYG PLUS the savings-plan rate (both stacked off PAYG). Totals cover the modeled subscriptions; access-only subscriptions are itemized on the per-unit detail slides. Estate optimal blend saves " + (Imp $t.Annual) + " on a 3-year term.")
    }
    function Render-BUDetail($slide, $bu, $chunk, $part, $parts) {
        $suffix = if ($parts -gt 1) { "  (part $part of $parts)" } else { '' }
        Add-Header $slide ("Where to Use the Savings Plan vs. the 40% - $bu$suffix") 'WHERE THE SAVINGS PLAN WINS vs. THE 40%  |  SAVINGS-PLAN RATE + 32.5% OFF PAYG  |  ALL SUBSCRIPTIONS'
        $all = @($detail | Where-Object BusinessUnit -eq $bu)
        $modeled = @($all | Where-Object HasSavingsBaseline)
        $rc = New-Object System.Collections.ArrayList; $rk = New-Object System.Collections.ArrayList
        [void]$rc.Add($HDR); [void]$rk.Add('hdr')
        $g = Aggregate $modeled
        [void]$rc.Add(@(("$bu  ($($all.Count) subscriptions)"),(N $g.VMh),(PC $g.Cov),(Rate40 $g.A),(F $g.Payg),(F $g.One),(F $g.Three),((F2 $g.C1)+' / '+(F2 $g.C3)),'Savings Plan',(Imp $g.Annual)))
        [void]$rk.Add("buheader|$bu")
        foreach ($x in $chunk) {
            $tag = if ($x.NewlyVisible) { ' *' } else { '' }
            $name = '   ' + $x.Subscription + ' | ' + $x.BusinessUnit + $tag
            $kind = "budetail|$bu"
            if ($x.HasSavingsBaseline) {
                $rec = if ([double]$x.AnnualSavingsVs40 -gt 0) { 'Savings Plan' } else { 'Keep 40%' }
                if ([double]$x.AnnualSavingsVs40 -le 0) { $kind = 'amber' }
                [void]$rc.Add(@($name,(N $x.VMHoursPerMonth),(PC ($x.SteadyCoveragePct/100.0)),(Rate40 $x.Current40Monthly),(F $x.PublicPaygMonthly),(F $x.OneYearSpMonthly),(F $x.ThreeYearSpMonthly),((F2 $x.CommitOne)+' / '+(F2 $x.CommitThree)),$rec,(Imp $x.AnnualSavingsVs40)))
            } else {
                $vmh = if ([double]$x.VMHoursPerMonth -gt 0) { N $x.VMHoursPerMonth } else { '-' }
                [void]$rc.Add(@($name,$vmh,'n/a',(Rate40 $x.Current40Monthly),'n/a','n/a','n/a','n/a','No baseline','n/a'))
            }
            [void]$rk.Add($kind)
        }
        Render-Table $slide $rc $rk
        Add-Footnote $slide ("Pricing: under a savings plan the negotiated on-demand discount is forfeited and replaced by 32.5% off PAYG PLUS the savings-plan rate (both off the public PAYG list). The colored row is the business-unit total (subscriptions with a savings baseline). 'n/a' rows are readable subscriptions with no savings-plan baseline yet (billing, identity, sandbox or idle), kept here so every subscription reconciles against Azure. '*' marks a subscription newly readable since the prior baseline.")
    }

    for ($index = $presentation.Slides.Count; $index -ge 2; $index--) { $presentation.Slides[$index].Delete() }

    $summarySlide = $presentation.Slides.Add(2, $ppLayoutBlank)
    Render-Summary $summarySlide

    $slideIndex = 3
    foreach ($businessUnit in $businessUnits.Keys) {
        $buRows = @($detail | Where-Object BusinessUnit -eq $businessUnit) | Sort-Object @{e={[int][bool]$_.HasSavingsBaseline};d=$true}, @{e={[double]$_.Current40Monthly};d=$true}
        if ($buRows.Count -eq 0) { continue }
        $chunks = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $buRows.Count; $i += $maxDetail) { [void]$chunks.Add(@($buRows[$i..([math]::Min($i+$maxDetail-1,$buRows.Count-1))])) }
        for ($c = 0; $c -lt $chunks.Count; $c++) {
            $slide = $presentation.Slides.Add($slideIndex, $ppLayoutBlank)
            Render-BUDetail $slide $businessUnit $chunks[$c] ($c+1) $chunks.Count
            $slide.Export((Join-Path $folder ("bu-detail-{0}-{1}.png" -f ($businessUnit -replace '[^A-Za-z0-9]+','-').Trim('-'), ($c+1))), 'PNG', 1600, 900)
            $slideIndex++
        }
    }

    $presentation.Save()
    $summarySlide.Export((Join-Path $folder 'bu-breakout-summary.png'), 'PNG', 1600, 900)
    $presentation.Close()
}
finally {
    $powerPoint.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
}

$summary | Format-Table -AutoSize
Write-Host "Built deck: $deck" -ForegroundColor Green
