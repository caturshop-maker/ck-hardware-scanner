[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$ResultPath)

$ErrorActionPreference = 'Stop'
$started = Get-Date
$result = [ordered]@{ schemaVersion='1.0'; provider='LibreHardwareMonitor'; providerVersion='0.9.6'; runtime='PowerShell7-net10'; status='FAIL'; startedAt=$started; completedAt=$null; hardware=@(); sensors=@(); errors=@() }
$computer = $null
function Add-Hardware { param($Items,[string]$Parent='')
    foreach($h in @($Items)) {
        $result.hardware += [pscustomobject][ordered]@{name=[string]$h.Name; hardwareName=[string]$h.Name; hardwareType=[string]$h.HardwareType; identifier=if($null -ne $h.Identifier){[string]$h.Identifier}else{$null}; parent=$Parent}
        try { $h.Update() } catch { $result.errors += ('Update '+$h.Name+': '+$_.Exception.Message) }
        foreach($s in @($h.Sensors)) {
            $unit=$null; switch([string]$s.SensorType) { 'Load'{$unit='%'} 'Temperature'{$unit='C'} 'Clock'{$unit='MHz'} 'Power'{$unit='W'} 'Fan'{$unit='RPM'} 'Voltage'{$unit='V'} 'Current'{$unit='A'} default {$unit=$null} }
            $result.sensors += [pscustomobject][ordered]@{name=[string]$s.Name; hardwareName=[string]$h.Name; hardwareType=[string]$h.HardwareType; sensorType=[string]$s.SensorType; value=$s.Value; unit=$unit; identifier=if($null -ne $s.Identifier){[string]$s.Identifier}else{$null}}
        }
        Add-Hardware $h.SubHardware ([string]$h.Name)
    }
}
try {
    $dll=Join-Path $PSScriptRoot 'LibreHardwareMonitorLib.dll'; if(!(Test-Path -LiteralPath $dll)){throw 'LibreHardwareMonitorLib.dll not found'}
    [Reflection.Assembly]::LoadFrom($dll) | Out-Null
    $computer=New-Object LibreHardwareMonitor.Hardware.Computer
    foreach($p in 'IsCpuEnabled','IsGpuEnabled','IsMemoryEnabled','IsMotherboardEnabled','IsStorageEnabled','IsControllerEnabled'){try{$computer.$p=$true}catch{}}
    $computer.Open(); Add-Hardware $computer.Hardware
    $valued=@($result.sensors | Where-Object {$null -ne $_.value}).Count
    $result.status=if($valued -gt 0){'PASS'}elseif(@($result.hardware).Count -gt 0){'PASS_NO_READABLE_SENSOR'}else{'FAIL'}
} catch { $result.errors += $_.Exception.ToString(); $result.status='FAIL' }
finally { if($null -ne $computer){try{$computer.Close()}catch{$result.errors += ('Close: '+$_.Exception.Message)}}; $result.completedAt=Get-Date; $result | ConvertTo-Json -Depth 10 -Compress | Set-Content -LiteralPath $ResultPath -Encoding UTF8; $result | ConvertTo-Json -Depth 10 -Compress }
