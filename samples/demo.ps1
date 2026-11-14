# Demo PowerShell script for Script Debugger
param([string]$Greeting = "Hello")

$name = "World"
$count = 0

function Add-Numbers($a, $b) {
    $sum = $a + $b
    return $sum
}

<# Block comment:
   nothing here runs #>
Write-Output "$Greeting, $name"

foreach ($i in 1..3) {
    $count += $i
    Write-Output "count is now $count"
}

$total = Add-Numbers 10 32
$items = @{ apple = 1; pear = 2 }
$items["plum"] = 3
$global:Shared = "visible everywhere"
Write-Output "total = $total, items = $($items.Count)"
exit 5
