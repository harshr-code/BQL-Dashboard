# Builds dash.json from the published agg_lrm CSV.
# Runs on PowerShell Core (Linux runner) as well as Windows.
# Reads the repo's current dash.json to carry MOP/AOP forward and to validate against.
# Writes dash.json in place. Exits non-zero WITHOUT writing if validation fails.

param(
    [string]$Repo = '.',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
function Say($m)  { Write-Host $m }
function Fail($m) { Write-Host "::error::$m" }

$LRM_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTL0Tj_UigCmUlZXFC0BrRUTe1gAqJ4nTnDrNVK1R3-Wu-8Xe1YtUyHyLERC5L8ktlkYZI6wPK5t_ac/pub?gid=0&single=true&output=csv'
$AUG_1_20_EXPECTED = 39359
$AUG_TOLERANCE     = 2
$DAY_TOLERANCE     = 12

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

# ---------- 2 aggregate ----------
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

$daily=@{}; $bucket=@{}; $subs=@{}; $subList=New-Object System.Collections.Generic.List[string]
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
    if ($d.Year -ne 2026 -or ($d.Month -ne 8 -and $d.Month -ne 9)) { continue }

    $ch=$f[$iS].Trim(); $sub=$f[$iSS].Trim(); if ($sub -eq '') { $sub='(none)' }
    if ($ch -eq 'Referral' -and $sub -eq 'BTL') { $ch='BTL' }
    if (-not $coreSet.ContainsKey($ch)) { if ($ch -eq '') { continue }; $sub=$ch; $ch='Others' }
    $city=$f[$iC].Trim(); if ($ncr.ContainsKey($city)) { $city='Delhi NCR' }
    if (-not $cityIdx.ContainsKey($city)) { continue }

    $bq=0.0; [void][double]::TryParse($f[$iB],[ref]$bq)
    $lr=0.0; [void][double]::TryParse($f[$iL],[ref]$lr)
    if (-not $subs.ContainsKey($sub)) { $subs[$sub]=$subList.Count; $subList.Add($sub) }
    $ciX=$cityIdx[$city]; $chX=$chIdx[$ch]; $sbX=$subs[$sub]

    $dk=$d.ToString('yyyy-MM-dd'); $k="$dk|$ciX|$chX|$sbX"
    if (-not $daily.ContainsKey($k)) { $daily[$k]=@(0,0.0,0.0) }
    $a=$daily[$k]; $a[0]++; $a[1]+=$bq; $a[2]+=$lr

    $sc=$f[$iSc].Trim(); $scv=0.0
    if ($sc -eq '') { $bk='No Score' }
    elseif (-not [double]::TryParse($sc,[ref]$scv)) { $bk='No Score'; $scv=0.0 }
    elseif ($scv -le 2) { $bk='0-2' } elseif ($scv -le 5) { $bk='2-5' }
    elseif ($scv -le 7) { $bk='5-7' } elseif ($scv -le 9) { $bk='7-9' } else { $bk='9-10' }
    $bkk="$($d.Month)|$($d.Day)|$ciX|$chX|$sbX|$($bkIdx[$bk])"
    if (-not $bucket.ContainsKey($bkk)) { $bucket[$bkk]=@(0.0,0.0) }
    $b=$bucket[$bkk]; $b[0]+=$bq; $b[1]+=($scv*$bq)
}
$sr.Close()

# ---------- 3 cut at last complete day ----------
$othIdx=$chIdx['Others']
$panByDate=@{}; $panCore=@{}
foreach ($kv in $daily.GetEnumerator()) {
    $p=$kv.Key.Split('|'); $dt=$p[0]
    if (-not $panByDate.ContainsKey($dt)) { $panByDate[$dt]=0.0; $panCore[$dt]=0.0 }
    $panByDate[$dt]+=$kv.Value[1]
    if ([int]$p[2] -ne $othIdx) { $panCore[$dt]+=$kv.Value[1] }
}
$sorted=@($panByDate.Keys | Sort-Object)
$cut=$sorted[-1]
while ($cut -and $panByDate[$cut] -eq 0) {
    $sorted=$sorted[0..($sorted.Count-2)]
    if ($sorted.Count -eq 0) { Fail "every day is zero - aborting."; exit 3 }
    $cut=$sorted[-1]
}
$ref=@($sorted | Select-Object -Last 8 | Select-Object -SkipLast 1 | ForEach-Object { $panByDate[$_] } | Sort-Object)
$med = if ($ref.Count) { $ref[[int][math]::Floor($ref.Count/2)] } else { 0 }
if ($med -gt 0 -and $panByDate[$cut] -lt ($med*0.6)) {
    Say "     last day $cut = $([int]$panByDate[$cut]) vs 7-day median $([int]$med) - partial, cutting it"
    $sorted=$sorted[0..($sorted.Count-2)]; $cut=$sorted[-1]
}
Say "3/5  series cut at $cut"
$expectCut=(Get-Date).Date.AddDays(-1)
if ([datetime]$cut -lt $expectCut) {
    $age=[int]((Get-Date).Date - [datetime]$cut).TotalDays
    Say "     NOTE source has nothing newer than $cut ($age days back) - publishing it anyway"
}
$keepDates=@{}; foreach ($dt in $sorted) { $keepDates[$dt]=1 }

# ---------- 4 assemble + validate ----------
$dates=@($sorted)
$dIdx=@{}; for($i=0;$i -lt $dates.Count;$i++){$dIdx[$dates[$i]]=$i}
$dailyArr=New-Object System.Collections.Generic.List[object]
foreach ($kv in $daily.GetEnumerator()) {
    $p=$kv.Key.Split('|'); if (-not $keepDates.ContainsKey($p[0])) { continue }
    $v=$kv.Value
    $dailyArr.Add(@($dIdx[$p[0]],[int]$p[1],[int]$p[2],[int]$p[3],[int]$v[0],[int]$v[1],[int]$v[2]))
}
$cutD=[datetime]$cut
$bucketArr=New-Object System.Collections.Generic.List[object]
foreach ($kv in $bucket.GetEnumerator()) {
    $p=$kv.Key.Split('|'); $v=$kv.Value
    if ($v[0] -eq 0) { continue }
    $bd=[datetime]::new(2026,[int]$p[0],[int]$p[1])
    if ($bd -gt $cutD) { continue }
    $bucketArr.Add(@([int]$p[0],[int]$p[1],[int]$p[2],[int]$p[3],[int]$p[4],[int]$p[5],[int]$v[0],[math]::Round($v[1],2)))
}

function SpanSum($src,$pfx,$maxDay) {
    $t=0.0
    foreach ($dt in $dates) { if ($dt -like "$pfx*" -and [int]$dt.Split('-')[2] -le $maxDay) { $t+=$src[$dt] } }
    return [int]$t
}
$aug=SpanSum $panCore '2026-08-' 20
$sep=SpanSum $panCore '2026-09-' 20
Say "4/5  validate:  Aug 1-20 = $aug (expect $AUG_1_20_EXPECTED)   Sep 1-20 = $sep"

$augDrift=$aug-$AUG_1_20_EXPECTED
$augPct=[math]::Round(100*$augDrift/$AUG_1_20_EXPECTED,2)
if ([math]::Abs($augPct) -gt $AUG_TOLERANCE) {
    Fail "August baseline moved $augDrift ($augPct%), beyond $AUG_TOLERANCE%. Nothing written."
    exit 4
}
if ($augDrift -ne 0) { Say "     August drifted $augDrift ($augPct%) - within $AUG_TOLERANCE%" }

$liveDay=@{}
foreach ($r in $live.daily) {
    $dt=$live.dates[$r[0]]
    if (-not $liveDay.ContainsKey($dt)) { $liveDay[$dt]=0.0 }
    $liveDay[$dt]+=$r[5]
}
$settledBefore=(Get-Date).Date.AddDays(-3).ToString('yyyy-MM-dd')
$worstDay=$null; $worstPct=0
foreach ($dt in $dates) {
    if ($dt -ge $settledBefore) { continue }
    if (-not $liveDay.ContainsKey($dt) -or $liveDay[$dt] -le 0) { continue }
    $p=100*($panByDate[$dt]-$liveDay[$dt])/$liveDay[$dt]
    if ([math]::Abs($p) -gt [math]::Abs($worstPct)) { $worstPct=$p; $worstDay=$dt }
}
if ($worstDay -and [math]::Abs($worstPct) -gt $DAY_TOLERANCE) {
    Fail "$worstDay moved $([math]::Round($worstPct,1))% against the published file, beyond $DAY_TOLERANCE%. Nothing written."
    exit 4
}
if ($worstDay) { Say "     largest settled-day move: $worstDay $([math]::Round($worstPct,1))% - within $DAY_TOLERANCE%" }
Say "     baseline holds"

# ---------- 5 write ----------
$out=[ordered]@{
    lastDate=$dates[-1]; dates=$dates; cities=$cities; channels=$channels
    subs=$subList.ToArray(); buckets=$buckets
    daily=$dailyArr.ToArray(); bucket=$bucketArr.ToArray()
    aop=$live.aop; mopRate=$live.mopRate
}
if ($DryRun) { Say "5/5  dry run - dash.json not written"; exit 0 }
[IO.File]::WriteAllText($dashPath, ($out|ConvertTo-Json -Depth 6 -Compress), (New-Object Text.UTF8Encoding $false))
Say "5/5  wrote dash.json - data through $($dates[-1])"

$last=$dates[-1]; $prev=$dates[-2]
Say ""
Say "PAN BQL $last = $([int]$panByDate[$last])   prior day $prev = $([int]$panByDate[$prev])"
$sepMtd=0.0; foreach ($dt in $dates) { if ($dt -like '2026-09-*') { $sepMtd+=$panByDate[$dt] } }
Say "Sep MTD through ${last}: $([int]$sepMtd)"
