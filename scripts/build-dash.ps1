# Builds dash.json from the published agg_lrm CSV.
# Runs on PowerShell Core (Linux runner) as well as Windows.
#
# dash.json is an ARCHIVE: it keeps every month it has ever published. The source sheet
# only holds a rolling window (it dropped August on 1 Oct 2026), so each run takes what
# the source has and carries every other month forward from the published file. A month
# disappearing from the archive is treated as a failure, not a refresh.
#
#   -Extra <paths>  older dash.json files to recover months from (one-off backfills)
#
# Writes dash.json in place. Exits non-zero WITHOUT writing if validation fails.

param(
    [string]$Repo = '.',
    [string[]]$Extra = @(),
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
function Say($m)  { Write-Host $m }
function Fail($m) { Write-Host "::error::$m" }

$LRM_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTL0Tj_UigCmUlZXFC0BrRUTe1gAqJ4nTnDrNVK1R3-Wu-8Xe1YtUyHyLERC5L8ktlkYZI6wPK5t_ac/pub?gid=0&single=true&output=csv'
$MONTH_TOLERANCE = 2    # percent - a closed month may drift as late rows land
$DAY_TOLERANCE   = 12   # percent - any ONE settled day moving more than this is structural

$dashPath = Join-Path $Repo 'dash.json'
$csv      = Join-Path ([IO.Path]::GetTempPath()) 'lrm.csv'

if (-not (Test-Path $dashPath)) { Fail "dash.json not found at $dashPath"; exit 2 }
$live = Get-Content $dashPath -Raw | ConvertFrom-Json

# ---------- 1 fetch ----------
Say "1/5  fetching agg_lrm ..."
Invoke-WebRequest -Uri $LRM_URL -UseBasicParsing -OutFile $csv
$sr = New-Object IO.StreamReader($csv)
$hdrLine = $sr.ReadLine(); $sr.Close()
$expect = 'Action_Date,CITY,Source_Class_final,Source_Sub_Class_final,bill_qualified,score,lead_delivered_to_lrm'
if ($hdrLine.Trim() -ne $expect) { Fail "Unexpected header, nothing rebuilt. Got: $hdrLine"; exit 3 }
Say "     header OK  ($((Get-Item $csv).Length) bytes)"

# ---------- 2 aggregate the source ----------
Say "2/5  aggregating ..."
$cities=@($live.cities)
$core=@('Digital','Referral','BTL','SolarPro')
$channels=@($core) + 'Others'
$coreSet=@{}; foreach ($c in $core) { $coreSet[$c]=1 }
$buckets=@('0-2','2-5','5-7','7-9','9-10','No Score')
$cityIdx=@{}; for($i=0;$i -lt $cities.Count;$i++){$cityIdx[$cities[$i]]=$i}
$chIdx=@{};   for($i=0;$i -lt $channels.Count;$i++){$chIdx[$channels[$i]]=$i}
$bkIdx=@{};   for($i=0;$i -lt $buckets.Count;$i++){$bkIdx[$buckets[$i]]=$i}
$ncr=@{'Delhi'=1;'Noida'=1;'Ghaziabad'=1;'Gurgaon'=1;'Faridabad'=1}

$subs=@{}; $subList=New-Object System.Collections.Generic.List[string]
function SubIdx($name) {
    if (-not $script:subs.ContainsKey($name)) { $script:subs[$name]=$script:subList.Count; $script:subList.Add($name) }
    return $script:subs[$name]
}
$srcDaily=@{}    # 'yyyy-MM-dd|city|ch|sub' -> leads, bql, lrm
$srcBucket=@{}   # 'yyyy-MM-dd|city|ch|sub|bk' -> bql, score*bql

$sr = New-Object IO.StreamReader($csv)
$cols=$sr.ReadLine().Split(','); $map=@{}
for($i=0;$i -lt $cols.Count;$i++){$map[$cols[$i].Trim()]=$i}
$iD=$map['Action_Date']; $iC=$map['CITY']; $iS=$map['Source_Class_final']
$iSS=$map['Source_Sub_Class_final']; $iB=$map['bill_qualified']
$iSc=$map['score']; $iL=$map['lead_delivered_to_lrm']
while ($null -ne ($line=$sr.ReadLine())) {
    if ($line.Trim() -eq '') { continue }
    $f=$line.Split(','); if ($f.Count -lt 7) { continue }
    $ds=$f[$iD].Trim(); if ($ds -eq '') { continue }
    try { $d=[datetime]::ParseExact($ds,'M/d/yyyy',$null) } catch { continue }
    $ch=$f[$iS].Trim(); $sub=$f[$iSS].Trim(); if ($sub -eq '') { $sub='(none)' }
    if ($ch -eq 'Referral' -and $sub -eq 'BTL') { $ch='BTL' }
    if (-not $coreSet.ContainsKey($ch)) { if ($ch -eq '') { continue }; $sub=$ch; $ch='Others' }
    $city=$f[$iC].Trim(); if ($ncr.ContainsKey($city)) { $city='Delhi NCR' }
    if (-not $cityIdx.ContainsKey($city)) { continue }
    $bq=0.0; [void][double]::TryParse($f[$iB],[ref]$bq)
    $lr=0.0; [void][double]::TryParse($f[$iL],[ref]$lr)
    $k="$($d.ToString('yyyy-MM-dd'))|$($cityIdx[$city])|$($chIdx[$ch])|$(SubIdx $sub)"
    if (-not $srcDaily.ContainsKey($k)) { $srcDaily[$k]=@(0,0.0,0.0) }
    $a=$srcDaily[$k]; $a[0]++; $a[1]+=$bq; $a[2]+=$lr
    $sc=$f[$iSc].Trim(); $scv=0.0
    if ($sc -eq '') { $bk='No Score' }
    elseif (-not [double]::TryParse($sc,[ref]$scv)) { $bk='No Score'; $scv=0.0 }
    elseif ($scv -le 2) { $bk='0-2' } elseif ($scv -le 5) { $bk='2-5' }
    elseif ($scv -le 7) { $bk='5-7' } elseif ($scv -le 9) { $bk='7-9' } else { $bk='9-10' }
    $bkk="$k|$($bkIdx[$bk])"
    if (-not $srcBucket.ContainsKey($bkk)) { $srcBucket[$bkk]=@(0.0,0.0) }
    $b=$srcBucket[$bkk]; $b[0]+=$bq; $b[1]+=($scv*$bq)
}
$sr.Close()

function PanByDate($dailyTable) {
    $o=@{}
    foreach ($kv in $dailyTable.GetEnumerator()) {
        $dt=$kv.Key.Split('|')[0]
        if (-not $o.ContainsKey($dt)) { $o[$dt]=0.0 }
        $o[$dt]+=$kv.Value[1]
    }
    return $o
}

# ---------- 3 cut the newest month at its last complete day ----------
$srcPan=PanByDate $srcDaily
$sorted=@($srcPan.Keys | Sort-Object)
if ($sorted.Count -eq 0) { Fail "source has no dated rows - aborting."; exit 3 }
$cut=$sorted[-1]
while ($srcPan[$cut] -eq 0) {
    if ($sorted.Count -le 1) { Fail "every day is zero - aborting."; exit 3 }
    $sorted=$sorted[0..($sorted.Count-2)]; $cut=$sorted[-1]
}
$ref=@($sorted | Select-Object -Last 8 | Select-Object -SkipLast 1 | ForEach-Object { $srcPan[$_] } | Sort-Object)
$med = if ($ref.Count) { $ref[[int][math]::Floor($ref.Count/2)] } else { 0 }
if ($med -gt 0 -and $srcPan[$cut] -lt ($med*0.6)) {
    Say "     last day $cut = $([int]$srcPan[$cut]) vs 7-day median $([int]$med) - partial, cutting it"
    $sorted=$sorted[0..($sorted.Count-2)]; $cut=$sorted[-1]
}
$newestKey=([datetime]$cut).ToString('yyyy-MM')
Say "3/5  newest complete day $cut"
if ([datetime]$cut -lt (Get-Date).Date.AddDays(-1)) {
    Say "     NOTE source has nothing newer than $cut ($([int]((Get-Date).Date - [datetime]$cut).TotalDays) days back) - publishing it anyway"
}

# ---------- 3b archive: source months + every month carried from published/extra files ----------
function DaysOf($dateList, $mk) { @($dateList | Where-Object { $_ -like "$mk-*" }).Count }
$srcDays=@{}
foreach ($dt in $sorted) { if ($srcPan[$dt] -gt 0) { $mk=$dt.Substring(0,7); if (-not $srcDays.ContainsKey($mk)) { $srcDays[$mk]=0 }; $srcDays[$mk]++ } }

# archive candidates, newest file first: published, then extras
$files=@(@{name='published'; j=$live})
foreach ($e in $Extra) { $files += @{name=(Split-Path $e -Leaf); j=(Get-Content $e -Raw | ConvertFrom-Json)} }

$keep=@{}; $keepB=@{}; $plan=[ordered]@{}; $taken=@{}
# source months first
foreach ($mk in $srcDays.Keys) {
    $best=$null; $bestDays=0
    foreach ($fl in $files) { $n=DaysOf $fl.j.dates $mk; if ($n -gt $bestDays) { $best=$fl; $bestDays=$n } }
    if ($mk -eq $newestKey -or $srcDays[$mk] -ge $bestDays) { $taken[$mk]='source' }
}
foreach ($kv in $srcDaily.GetEnumerator())  { $dt=$kv.Key.Split('|')[0]; if ($dt -gt $cut) { continue }; if ($taken[$dt.Substring(0,7)] -eq 'source') { $keep[$kv.Key]=$kv.Value } }
foreach ($kv in $srcBucket.GetEnumerator()) { $dt=$kv.Key.Split('|')[0]; if ($dt -gt $cut) { continue }; if ($taken[$dt.Substring(0,7)] -eq 'source') { $keepB[$kv.Key]=$kv.Value } }

# then every other month, from whichever file holds the most days of it
$allMonths=@{}
foreach ($fl in $files) { foreach ($dt in $fl.j.dates) { $allMonths[$dt.Substring(0,7)]=1 } }
foreach ($mk in ($allMonths.Keys | Sort-Object)) {
    if ($taken.ContainsKey($mk)) { continue }
    $best=$null; $bestDays=0
    foreach ($fl in $files) { $n=DaysOf $fl.j.dates $mk; if ($n -gt $bestDays) { $best=$fl; $bestDays=$n } }
    if (-not $best) { continue }
    $taken[$mk]=$best.name
    $j=$best.j; $jCh=@($j.channels); $jSubs=@($j.subs)
    $mon=[int]$mk.Substring(5,2); $ym=[int]($mk.Substring(0,4)+$mk.Substring(5,2))
    foreach ($r in $j.daily) {
        $dt=$j.dates[$r[0]]; if ($dt -notlike "$mk-*") { continue }
        $chX=$chIdx[$jCh[$r[2]]]; if ($null -eq $chX) { continue }
        $keep["$dt|$($r[1])|$chX|$(SubIdx $jSubs[$r[3]])"]=@([int]$r[4],[double]$r[5],[double]$r[6])
    }
    foreach ($r in $j.bucket) {
        $m0=[int]$r[0]
        if (-not (($m0 -lt 100 -and $m0 -eq $mon) -or $m0 -eq $ym)) { continue }   # old files store month only
        $chX=$chIdx[$jCh[$r[3]]]; if ($null -eq $chX) { continue }
        $dt='{0}-{1:00}' -f $mk, [int]$r[1]
        $keepB["$dt|$($r[2])|$chX|$(SubIdx $jSubs[$r[4]])|$($r[5])"]=@([double]$r[6],[double]$r[7])
    }
}
Say "     months: $((($taken.GetEnumerator() | Sort-Object Name) | ForEach-Object { "$($_.Name) <- $($_.Value)" }) -join ',  ')"

# plans (MOP), keyed by month. Older files carry one plan as mopRate (+ planMonth).
foreach ($fl in ($files | Select-Object -Skip 0)) {
    $j=$fl.j
    if ($j.PSObject.Properties.Name -contains 'plans' -and $j.plans) {
        foreach ($p in $j.plans.PSObject.Properties) { if (-not $plan.Contains($p.Name)) { $plan[$p.Name]=$p.Value } }
    } elseif ($j.PSObject.Properties.Name -contains 'mopRate' -and $j.mopRate) {
        $pm = if ($j.PSObject.Properties.Name -contains 'planMonth' -and $j.planMonth) { $j.planMonth } else { '2026-09' }
        if (-not $plan.Contains($pm)) { $plan[$pm]=$j.mopRate }
    }
}

# ---------- 4 assemble + validate ----------
$pan=PanByDate $keep
$dates=@($pan.Keys | Sort-Object)
$dIdx=@{}; for($i=0;$i -lt $dates.Count;$i++){$dIdx[$dates[$i]]=$i}
$dailyArr=New-Object System.Collections.Generic.List[object]
foreach ($kv in $keep.GetEnumerator()) {
    $p=$kv.Key.Split('|'); $v=$kv.Value
    $dailyArr.Add(@($dIdx[$p[0]],[int]$p[1],[int]$p[2],[int]$p[3],[int]$v[0],[int]$v[1],[int]$v[2]))
}
$bucketArr=New-Object System.Collections.Generic.List[object]
foreach ($kv in $keepB.GetEnumerator()) {
    $p=$kv.Key.Split('|'); $v=$kv.Value
    if ($v[0] -eq 0) { continue }
    $ym=[int]($p[0].Substring(0,4)+$p[0].Substring(5,2))     # bucket month stored as yyyymm
    $bucketArr.Add(@($ym,[int]$p[0].Substring(8,2),[int]$p[1],[int]$p[2],[int]$p[3],[int]$p[4],[int]$v[0],[math]::Round($v[1],2)))
}

$liveDay=PanByDate (& { $o=@{}; foreach ($r in $live.daily) { $o["$($live.dates[$r[0]])|$($r[1])|$($r[2])|$($r[3])"]=@(0,[double]$r[5],0) }; $o })
$liveMonths=@($live.dates | ForEach-Object { $_.Substring(0,7) } | Sort-Object -Unique)
$newMonths =@($dates | ForEach-Object { $_.Substring(0,7) } | Sort-Object -Unique)

Say "4/5  validate:  months $($newMonths -join ', ')"
foreach ($mk in $liveMonths) {
    if ($newMonths -notcontains $mk) { Fail "$mk is in the published file but would be dropped. The archive must never lose a month. Nothing written."; exit 4 }
}
foreach ($mk in $liveMonths) {
    if ($mk -eq $newestKey) { continue }                       # newest month is still filling
    $a=0.0; $b=0.0
    foreach ($dt in $dates)        { if ($dt -like "$mk-*") { $a+=$pan[$dt] } }
    foreach ($dt in $liveDay.Keys) { if ($dt -like "$mk-*") { $b+=$liveDay[$dt] } }
    if ($b -le 0) { continue }
    $drift=100*($a-$b)/$b
    if ([math]::Abs($drift) -gt $MONTH_TOLERANCE) { Fail "$mk moved $([math]::Round($drift,2))% against the published file, beyond $MONTH_TOLERANCE%. Nothing written."; exit 4 }
    Say "     $mk  $([int]$a)  (published $([int]$b), drift $([math]::Round($drift,2))%)"
}
$settledBefore=(Get-Date).Date.AddDays(-3).ToString('yyyy-MM-dd')
$worstDay=$null; $worstPct=0
foreach ($dt in $dates) {
    if ($dt -ge $settledBefore) { continue }
    if (-not $liveDay.ContainsKey($dt) -or $liveDay[$dt] -le 0) { continue }
    $pp=100*($pan[$dt]-$liveDay[$dt])/$liveDay[$dt]
    if ([math]::Abs($pp) -gt [math]::Abs($worstPct)) { $worstPct=$pp; $worstDay=$dt }
}
if ($worstDay -and [math]::Abs($worstPct) -gt $DAY_TOLERANCE) { Fail "$worstDay moved $([math]::Round($worstPct,1))% against the published file, beyond $DAY_TOLERANCE%. Nothing written."; exit 4 }
if ($worstDay) { Say "     largest settled-day move: $worstDay $([math]::Round($worstPct,1))% - within $DAY_TOLERANCE%" }
Say "     checks hold"

# ---------- 5 write ----------
$out=[ordered]@{
    lastDate=$dates[-1]; dates=$dates; cities=$cities; channels=$channels
    subs=$subList.ToArray(); buckets=$buckets
    daily=$dailyArr.ToArray(); bucket=$bucketArr.ToArray()
    plans=$plan
}
Say "     MOP loaded for: $(@($plan.Keys) -join ', ')"
if ($DryRun) { Say "5/5  dry run - dash.json not written"; exit 0 }
[IO.File]::WriteAllText($dashPath, ($out|ConvertTo-Json -Depth 8 -Compress), (New-Object Text.UTF8Encoding $false))
Say "5/5  wrote dash.json - $($newMonths -join ', '), data through $($dates[-1])"
$last=$dates[-1]; $prev=$dates[-2]
Say ""
Say "PAN BQL $last = $([int]$pan[$last])   prior day $prev = $([int]$pan[$prev])"
$mtd=0.0; foreach ($dt in $dates) { if ($dt -like "$newestKey-*") { $mtd+=$pan[$dt] } }
Say "$newestKey MTD through ${last}: $([int]$mtd)"
