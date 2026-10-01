# Builds dash.json from the published agg_lrm CSV.
# Runs on PowerShell Core (Linux runner) as well as Windows.
#
# Month-agnostic: the current month is the month of the newest complete day in the
# source; the comparison month is the one before it. The source sheet only holds a
# rolling current + previous month, so when it drops the comparison month (as it did
# on 1 Oct 2026, losing August) that month is carried forward from the published
# dash.json instead.
#
# Writes dash.json in place. Exits non-zero WITHOUT writing if validation fails.

param(
    [string]$Repo = '.',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
function Say($m)  { Write-Host $m }
function Fail($m) { Write-Host "::error::$m" }

$LRM_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTL0Tj_UigCmUlZXFC0BrRUTe1gAqJ4nTnDrNVK1R3-Wu-8Xe1YtUyHyLERC5L8ktlkYZI6wPK5t_ac/pub?gid=0&single=true&output=csv'
$PREV_TOLERANCE = 2     # percent - a closed month may drift up as late rows land
$DAY_TOLERANCE  = 12    # percent - any ONE settled day moving more than this is structural

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
if ($hdrLine.Trim() -ne $expect) {
    Fail "Unexpected header, nothing rebuilt. Got: $hdrLine"
    exit 3
}
Say "     header OK  ($((Get-Item $csv).Length) bytes)"

# ---------- 2 aggregate every month in the source ----------
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

$daily=@{}    # 'yyyy-MM-dd|city|ch|sub' -> leads, bql, lrm
$bucket=@{}   # 'yyyy-MM-dd|city|ch|sub|bk' -> bql, score*bql
$subs=@{}; $subList=New-Object System.Collections.Generic.List[string]
function SubIdx($name) {
    if (-not $script:subs.ContainsKey($name)) { $script:subs[$name]=$script:subList.Count; $script:subList.Add($name) }
    return $script:subs[$name]
}

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
    $ciX=$cityIdx[$city]; $chX=$chIdx[$ch]; $sbX=SubIdx $sub

    $dk=$d.ToString('yyyy-MM-dd'); $k="$dk|$ciX|$chX|$sbX"
    if (-not $daily.ContainsKey($k)) { $daily[$k]=@(0,0.0,0.0) }
    $a=$daily[$k]; $a[0]++; $a[1]+=$bq; $a[2]+=$lr

    $sc=$f[$iSc].Trim(); $scv=0.0
    if ($sc -eq '') { $bk='No Score' }
    elseif (-not [double]::TryParse($sc,[ref]$scv)) { $bk='No Score'; $scv=0.0 }
    elseif ($scv -le 2) { $bk='0-2' } elseif ($scv -le 5) { $bk='2-5' }
    elseif ($scv -le 7) { $bk='5-7' } elseif ($scv -le 9) { $bk='7-9' } else { $bk='9-10' }
    $bkk="$k|$($bkIdx[$bk])"
    if (-not $bucket.ContainsKey($bkk)) { $bucket[$bkk]=@(0.0,0.0) }
    $b=$bucket[$bkk]; $b[0]+=$bq; $b[1]+=($scv*$bq)
}
$sr.Close()

# ---------- 3 find the last complete day -> current and comparison month ----------
function PanByDate($dailyTable) {
    $o=@{}
    foreach ($kv in $dailyTable.GetEnumerator()) {
        $dt=$kv.Key.Split('|')[0]
        if (-not $o.ContainsKey($dt)) { $o[$dt]=0.0 }
        $o[$dt]+=$kv.Value[1]
    }
    return $o
}
$srcPan = PanByDate $daily
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
$cutD   = [datetime]$cut
$curKey = $cutD.ToString('yyyy-MM')
$prvKey = $cutD.AddMonths(-1).ToString('yyyy-MM')
Say "3/5  series cut at $cut   current month $curKey, compared against $prvKey"
if ($cutD -lt (Get-Date).Date.AddDays(-1)) {
    $age=[int]((Get-Date).Date - $cutD).TotalDays
    Say "     NOTE source has nothing newer than $cut ($age days back) - publishing it anyway"
}

# ---------- 3b keep current + comparison month; carry comparison month if the source dropped it ----------
$prvDays = [datetime]::DaysInMonth([int]$prvKey.Substring(0,4), [int]$prvKey.Substring(5,2))
$srcPrvDays = @($sorted | Where-Object { $_ -like "$prvKey-*" -and $srcPan[$_] -gt 0 }).Count
$carryPrv = $srcPrvDays -lt ($prvDays - 2)

$keep=@{}
foreach ($kv in $daily.GetEnumerator()) {
    $dt=$kv.Key.Split('|')[0]
    if ($dt -gt $cut) { continue }
    if ($dt -like "$curKey-*" -or (-not $carryPrv -and $dt -like "$prvKey-*")) { $keep[$kv.Key]=$kv.Value }
}
$keepB=@{}
foreach ($kv in $bucket.GetEnumerator()) {
    $dt=$kv.Key.Split('|')[0]
    if ($dt -gt $cut) { continue }
    if ($dt -like "$curKey-*" -or (-not $carryPrv -and $dt -like "$prvKey-*")) { $keepB[$kv.Key]=$kv.Value }
}

if ($carryPrv) {
    $livePrv = @($live.dates | Where-Object { $_ -like "$prvKey-*" }).Count
    if ($livePrv -lt ($prvDays - 2)) {
        Fail "Source holds only $srcPrvDays days of $prvKey and the published file holds $livePrv - no complete comparison month anywhere. Nothing written."
        exit 4
    }
    Say "     source has $srcPrvDays days of $prvKey - carrying $prvKey forward from the published file ($livePrv days)"
    $liveCh = @($live.channels); $liveSubs = @($live.subs)
    foreach ($r in $live.daily) {
        $dt=$live.dates[$r[0]]; if ($dt -notlike "$prvKey-*") { continue }
        $chX=$chIdx[$liveCh[$r[2]]]; if ($null -eq $chX) { continue }
        $sbX=SubIdx $liveSubs[$r[3]]
        $keep["$dt|$($r[1])|$chX|$sbX"]=@([int]$r[4],[double]$r[5],[double]$r[6])
    }
    $prvMon=[int]$prvKey.Substring(5,2)
    foreach ($r in $live.bucket) {
        if ([int]$r[0] -ne $prvMon) { continue }
        $chX=$chIdx[$liveCh[$r[3]]]; if ($null -eq $chX) { continue }
        $sbX=SubIdx $liveSubs[$r[4]]
        $dt='{0}-{1:00}' -f $prvKey, [int]$r[1]
        $keepB["$dt|$($r[2])|$chX|$sbX|$($r[5])"]=@([double]$r[6],[double]$r[7])
    }
}

# ---------- 4 assemble + validate ----------
$pan  = PanByDate $keep
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
    $bucketArr.Add(@([int]$p[0].Substring(5,2),[int]$p[0].Substring(8,2),[int]$p[1],[int]$p[2],[int]$p[3],[int]$p[4],[int]$v[0],[math]::Round($v[1],2)))
}

$liveDay=@{}
foreach ($r in $live.daily) {
    $dt=$live.dates[$r[0]]
    if (-not $liveDay.ContainsKey($dt)) { $liveDay[$dt]=0.0 }
    $liveDay[$dt]+=$r[5]
}

# closed comparison month must not move much against what is published
$newPrv=0.0; $oldPrv=0.0
foreach ($dt in $dates)        { if ($dt -like "$prvKey-*") { $newPrv+=$pan[$dt] } }
foreach ($dt in $liveDay.Keys) { if ($dt -like "$prvKey-*") { $oldPrv+=$liveDay[$dt] } }
Say "4/5  validate:  $prvKey total = $([int]$newPrv)  (published $([int]$oldPrv))"
if ($oldPrv -gt 0) {
    $drift=100*($newPrv-$oldPrv)/$oldPrv
    if ([math]::Abs($drift) -gt $PREV_TOLERANCE) {
        Fail "$prvKey moved $([math]::Round($drift,2))% against the published file, beyond $PREV_TOLERANCE%. Nothing written."
        exit 4
    }
    Say "     $prvKey drift $([math]::Round($drift,2))% - within $PREV_TOLERANCE%"
}

$settledBefore=(Get-Date).Date.AddDays(-3).ToString('yyyy-MM-dd')
$worstDay=$null; $worstPct=0
foreach ($dt in $dates) {
    if ($dt -ge $settledBefore) { continue }
    if (-not $liveDay.ContainsKey($dt) -or $liveDay[$dt] -le 0) { continue }
    $p=100*($pan[$dt]-$liveDay[$dt])/$liveDay[$dt]
    if ([math]::Abs($p) -gt [math]::Abs($worstPct)) { $worstPct=$p; $worstDay=$dt }
}
if ($worstDay -and [math]::Abs($worstPct) -gt $DAY_TOLERANCE) {
    Fail "$worstDay moved $([math]::Round($worstPct,1))% against the published file, beyond $DAY_TOLERANCE%. Nothing written."
    exit 4
}
if ($worstDay) { Say "     largest settled-day move: $worstDay $([math]::Round($worstPct,1))% - within $DAY_TOLERANCE%" }
Say "     checks hold"

# ---------- 5 write ----------
# Plan arrays are carried as-is. planMonth records which month they belong to, so the
# page can hide MOP/AOP comparisons instead of showing last month's targets as this month's.
$planMonth = if ($live.PSObject.Properties.Name -contains 'planMonth') { $live.planMonth } else { '2026-09' }
$out=[ordered]@{
    lastDate=$dates[-1]; dates=$dates; cities=$cities; channels=$channels
    subs=$subList.ToArray(); buckets=$buckets
    daily=$dailyArr.ToArray(); bucket=$bucketArr.ToArray()
    aop=$live.aop; mopRate=$live.mopRate; planMonth=$planMonth
}
if ($planMonth -ne $curKey) { Say "     NOTE plan (MOP/AOP) is for $planMonth, not $curKey - page will hide plan comparisons until it is updated" }
if ($DryRun) { Say "5/5  dry run - dash.json not written"; exit 0 }
[IO.File]::WriteAllText($dashPath, ($out|ConvertTo-Json -Depth 6 -Compress), (New-Object Text.UTF8Encoding $false))
Say "5/5  wrote dash.json - data through $($dates[-1])"

$last=$dates[-1]; $prev=$dates[-2]
Say ""
Say "PAN BQL $last = $([int]$pan[$last])   prior day $prev = $([int]$pan[$prev])"
$mtd=0.0; foreach ($dt in $dates) { if ($dt -like "$curKey-*") { $mtd+=$pan[$dt] } }
Say "$curKey MTD through ${last}: $([int]$mtd)"
