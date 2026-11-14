# Performance sample: "decrypts" a large blob chunk by chunk.
param([int]$SizeMB = 4, [int]$Chunks = 2000)

$rng = New-Object System.Random(42)
$blob = New-Object byte[] ($SizeMB * 1MB)
$rng.NextBytes($blob)
$key = [byte[]](1..32 | ForEach-Object { $rng.Next(256) })
$b64 = [Convert]::ToBase64String($blob)
$chunkSize = [int]($blob.Length / $Chunks)
$sw = [Diagnostics.Stopwatch]::StartNew()
$done = 0
$checksum = 0

for ($c = 0; $c -lt $Chunks; $c++) {
    $start = $c * $chunkSize
    for ($i = 0; $i -lt $chunkSize; $i++) {
        $blob[$start + $i] = $blob[$start + $i] -bxor $key[$i % 32]
    }
    $checksum = ($checksum + $blob[$start]) % 65521
    $done++
}

$elapsed = $sw.ElapsedMilliseconds
Write-Output "decrypted $done chunks of $chunkSize bytes in $elapsed ms, checksum $checksum"
