# 2 decimal digits to 2 hexadecit conversion
# XY -> 0xXY
function hexInt {
    param (
        $val
    )
    [Int32]("0x" + $val)
}

# looking for the first available port
$ports = [System.IO.Ports.SerialPort]::getportnames()
if ($ports.Count -eq 0) {
    Write-Output "COM port is not found"
    break Script
}
$portName = $ports[0]

# port may be specified implicitly with
# $portName = "COM7"

$port = new-Object System.IO.Ports.SerialPort $portName,9600,None,8,one
$port.open()

# wait for the beginning of the next minnute
do {
    Start-Sleep -Milliseconds 100
    $now = Get-Date
    $seconds = $now.Second
} while ($seconds -ne 0)

# set the date
$y1 = hexInt([math]::Floor($now.Year / 100))
$y2 = hexInt($now.Year % 100)
$m = hexInt($now.Month)
$d = hexInt($now.Day)
[byte[]] $b = 0xFE, 0xFE, 0x94, 0xE0, 0x1A, 0x05, 0x00, 0x94, $y1, $y2, $m, $d, 0xFD
$port.Write($b, 0, $b.Count)

# set the time
$h = hexInt($now.Hour)
$m = hexInt($now.Minute)
[byte[]] $b = 0xFE, 0xFE, 0x94, 0xE0, 0x1A, 0x05, 0x00, 0x95, $h, $m, 0xFD
$port.Write($b, 0, $b.Count)

# wait until the com port buffer gets flushed before closing the port
while ($port.BytesToWrite -gt 0) {
    Start-Sleep -Milliseconds 50
}

$port.Close()
