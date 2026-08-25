# Step 6: create the deck (if missing) and add the methodology slide (slide 1).
$folder = $PSScriptRoot
. (Join-Path $folder 'Config.ps1')
$dst = Join-Path $folder $deckName

$NAVY=31+56*256+100*65536; $BLUE=46+117*256+181*65536; $WHITE=16777215; $DARK=38+38*256+38*65536
$LIGHTA=234+240*256+248*65536; $LIGHTB=232+244*256+234*65536; $STRIP=219+229*256+241*65536
$FOOT=242+242*256+242*65536
$ordH=1; $msoTrue=-1; $msoFalse=0; $anchorMid=3

$stepsA = @(
 @('1. Start with real usage.','Azure uses your actual on-demand cost from eligible compute over the last 30 days (also 7 and 60), including your negotiated discounts.'),
 @('2. Price every hour.','It applies the 1-year and 3-year savings-plan rates to each hour''s eligible usage. Each hour''s ideal spend becomes a candidate commitment.'),
 @('3. Simulate hundreds of options.','For each candidate: total cost = (commitment x 24 x days) + pay-as-you-go above the commitment, versus paying all on-demand.'),
 @('4. Keep only what saves.','Candidates that net save are ranked; the greatest 1-year and 3-year savings is highlighted.'),
 @('5. Stay conservative.','It also re-runs on just the last 3 days and takes the LOWER commitment, so a usage drop cannot cause overcommitment.')
)
$stepsB = @(
 @('It is a rate for one clock hour.','Sized to the compute that runs every hour (your steady base). Small steady workloads make small hourly numbers, so fractions of a dollar are normal.'),
 @('Small example (3-year):','A workload that runs steadily for about $200 a month is roughly $0.27 an hour on-demand. Spread over the ~730 hours in a month, a 3-year plan is roughly half off, so you commit to about $0.14 for every hour of the term.'),
 @('You pay it every hour of the term.','8,760 hours/year. Usage above it is pay-as-you-go; unused commitment each hour does not roll over.')
)

$pp=New-Object -ComObject PowerPoint.Application
try{
    if(Test-Path $dst){ $pres=$pp.Presentations.Open($dst,$false,$false,$false); $isNew=$false }
    else { $pres=$pp.Presentations.Add($msoFalse); try{ $pres.PageSetup.SlideSize=15 }catch{}; $isNew=$true }
    $W=$pres.PageSetup.SlideWidth; $H=$pres.PageSetup.SlideHeight
    $mine=$null
    foreach($sl in $pres.Slides){ foreach($sp in $sl.Shapes){ if($sp.HasTextFrame -and $sp.TextFrame.TextRange.Text -like '*How the Savings Plan Commitment*'){ $mine=$sl } } }
    if(-not $mine){ $mine=$pres.Slides.Add($pres.Slides.Count+1,12) }
    $slide=$mine
    $arr=@(); foreach($sp in $slide.Shapes){ $arr+=$sp }; foreach($sp in $arr){ $sp.Delete() }

    function Add-Text($l,$t,$w,$h,$text,$size,$bold,$color,$align){ $tb=$slide.Shapes.AddTextbox($ordH,[int]$l,[int]$t,[int]$w,[int]$h); $tr=$tb.TextFrame.TextRange; $tb.TextFrame.WordWrap=$msoTrue; $tr.Text=[string]$text; $tr.Font.Size=[single]$size; $tr.Font.Bold=$(if($bold){$msoTrue}else{$msoFalse}); $tr.Font.Color.RGB=$color; $tr.Font.Name='Segoe UI'; if($align){$tr.ParagraphFormat.Alignment=$align}; return $tb }
    function Add-Rect($l,$t,$w,$h,$fill){ $rr=$slide.Shapes.AddShape(1,[int]$l,[int]$t,[int]$w,[int]$h); $rr.Fill.ForeColor.RGB=$fill; $rr.Line.Visible=$msoFalse; return $rr }
    function Add-Steps($l,$t,$w,$h,$steps,$sz){
        $tb=$slide.Shapes.AddTextbox($ordH,[int]$l,[int]$t,[int]$w,[int]$h); $tb.TextFrame.WordWrap=$msoTrue; $tb.TextFrame.MarginTop=2; $tb.TextFrame.MarginBottom=2
        $tr=$tb.TextFrame.TextRange; $first=$true
        foreach($s in $steps){
            if($first){ $tr.Text=$s[0]; $lead=$tr; $first=$false } else { $lead=$tr.InsertAfter("`r"+$s[0]) }
            $lead.Font.Bold=$msoTrue; $lead.Font.Size=[single]$sz; $lead.Font.Color.RGB=$NAVY; $lead.Font.Name='Segoe UI'
            $det=$tr.InsertAfter("  "+$s[1]); $det.Font.Bold=$msoFalse; $det.Font.Size=[single]$sz; $det.Font.Color.RGB=$DARK; $det.Font.Name='Segoe UI'
        }
        $tr.ParagraphFormat.SpaceAfter=7
        return $tb
    }

    # Header
    Add-Rect 0 0 $W ([int]($H*0.105)) $NAVY | Out-Null; Add-Rect 0 ([int]($H*0.105)) $W 4 $BLUE | Out-Null
    Add-Text ([int]($W*0.02)) ([int]($H*0.008)) ([int]($W*0.96)) ([int]($H*0.04)) 'SAVINGS PLAN METHODOLOGY   |   SOURCE: MICROSOFT LEARN (COST MANAGEMENT)' 8 $true $WHITE 1 | Out-Null
    Add-Text ([int]($W*0.02)) ([int]($H*0.045)) ([int]($W*0.96)) ([int]($H*0.05)) 'How the Savings Plan Commitment ($/hr) Is Calculated' 14 $true $WHITE 1 | Out-Null

    # Formula strip
    Add-Rect ([int]($W*0.0175)) ([int]($H*0.122)) ([int]($W*0.965)) ([int]($H*0.072)) $STRIP | Out-Null
    $fs=Add-Text ([int]($W*0.03)) ([int]($H*0.128)) ([int]($W*0.94)) ([int]($H*0.062)) '' 11 $false $DARK 1
    $ftr=$fs.TextFrame.TextRange; $ftr.Text='Commitment ($/hr) = the discounted cost of the compute that runs every hour (your steady base). '; $ftr.Font.Bold=$msoTrue; $ftr.Font.Size=[single]11
    $ftr2=$ftr.InsertAfter('You pay that rate 24x7 for the whole term; anything above it is pay-as-you-go; unused hours do not roll over.'); $ftr2.Font.Bold=$msoFalse; $ftr2.Font.Size=[single]11; $ftr2.Font.Color.RGB=$DARK; $ftr2.Font.Name='Segoe UI'

    # Two panels
    $py=[int]($H*0.205); $ph=[int]($H*0.55); $bar=[int]($H*0.048)
    $pAx=[int]($W*0.0175); $pw=[int]($W*0.478); $pBx=[int]($W*0.505)
    # Panel A
    Add-Rect $pAx $py $pw $ph $LIGHTA | Out-Null; Add-Rect $pAx $py $pw $bar $NAVY | Out-Null
    $tA=Add-Text ($pAx+14) $py ($pw-28) $bar 'How Azure calculates the commitment' 12 $true $WHITE 1; $tA.TextFrame.VerticalAnchor=$anchorMid
    Add-Steps ($pAx+16) ($py+$bar+6) ($pw-32) ($ph-$bar-12) $stepsA 10 | Out-Null
    # Panel B
    Add-Rect $pBx $py $pw $ph $LIGHTB | Out-Null; Add-Rect $pBx $py $pw $bar $NAVY | Out-Null
    $tB=Add-Text ($pBx+14) $py ($pw-28) $bar 'Why it is often a fraction of a dollar' 12 $true $WHITE 1; $tB.TextFrame.VerticalAnchor=$anchorMid
    Add-Steps ($pBx+16) ($py+$bar+6) ($pw-32) ($ph-$bar-12) $stepsB 9.5 | Out-Null

    # Footnote
    $fy=[int]($H*0.775)
    Add-Rect ([int]($W*0.0175)) $fy ([int]($W*0.965)) ([int]($H*0.215)) $FOOT | Out-Null
    Add-Rect ([int]($W*0.0175)) $fy ([int]($W*0.004)) ([int]($H*0.215)) $BLUE | Out-Null
    $l1='Source: Microsoft Learn, "Savings plan recommendations" and "How a savings plan discount is applied." Azure generates the recommendation from your real usage at your negotiated (already-discounted) rates, refreshes it several times a day, and sizes it to your steady base.'
    $l2='Each hour, the benefit applies to your highest-discount eligible usage first until the commitment is used up; usage beyond it is pay-as-you-go. The commitment depends on the term (1-year or 3-year discount), not on how large a number you pick.'
    $ft=Add-Text ([int]($W*0.03)) ([int]($fy+$H*0.01)) ([int]($W*0.94)) ([int]($H*0.19)) '' 9.5 $false $DARK 1
    $tr=$ft.TextFrame.TextRange; $tr.Text=$l1; $tr.Font.Bold=$msoTrue; $tr.Font.Size=[single]9.5
    $p2=$tr.InsertAfter("`r"+$l2); $p2.Font.Bold=$msoFalse; $p2.Font.Size=[single]9.5; $p2.Font.Color.RGB=$DARK

    if($isNew){ $pres.SaveAs($dst) } else { $pres.Save() }
    $slide.Export((Join-Path $folder 'method.png'),'PNG',1900,1069)
    Write-Host ("Added methodology slide at index {0} of {1}" -f $slide.SlideIndex, $pres.Slides.Count) -ForegroundColor Green
    $pres.Close()
}
finally{ $pp.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp)|Out-Null }
