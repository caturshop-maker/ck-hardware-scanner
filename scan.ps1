[CmdletBinding()]
param(
    [switch]$SaveReport,
    [string]$OutputDirectory = (Join-Path (Get-Location) 'reports'),
    [switch]$NoColor,
    [switch]$RedactSensitive,
    [ValidateSet('TXT','HTML','JSON')][string]$OutputFormat = 'TXT'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function New-ScanResult {
    param([string]$Status = 'N/A', [string]$Value = '', [string]$Note = '')
    [pscustomobject]@{ Status = $Status; Value = $Value; Note = $Note }
}

function Get-PropertyValue { param($Object, [string]$Name, $Default = $null)
    if ($null -ne $Object) { $property=$Object.PSObject.Properties[$Name]; if ($null -ne $property -and $null -ne $property.Value) { return $property.Value } }
    return $Default
}
function Get-StructuredValue { param($Object, [string]$Name, $Default = 'N/A')
    $field=Get-PropertyValue $Object $Name $null; if ($null -ne $field) { $value=Get-PropertyValue $field Value $null; if ($null -ne $value) { return $value }; return $field }; return $Default
}
function New-SensorRecord { param([string]$Component,[string]$SensorName,$Value,[string]$Unit,[string]$Source,[datetime]$Timestamp,[string]$Status='Valid',[string]$Quality='Unvalidated',[string]$Error='')
    [pscustomobject][ordered]@{ Component=$Component; SensorName=$SensorName; Value=$Value; Unit=if($Unit){$Unit}else{'N/A'}; Source=if($Source){$Source}else{'N/A'}; Timestamp=if($Timestamp){$Timestamp}else{Get-Date}; Status=$Status; Quality=$Quality; Error=$Error }
}
function Test-SensorRecord { param($Record)
    if (!$Record -or [string]::IsNullOrWhiteSpace([string](Get-PropertyValue $Record SensorName ''))) { return [pscustomobject]@{Valid=$false;Reason='Sensor name is empty'} }
    $value=Get-PropertyValue $Record Value $null; if ($null -eq $value) { return [pscustomobject]@{Valid=$false;Reason='Sensor value is null'} }
    if ($value -is [double] -and ([double]::IsNaN($value) -or [double]::IsInfinity($value))) { return [pscustomobject]@{Valid=$false;Reason='Sensor value is NaN or Infinity'} }
    $unit=[string](Get-PropertyValue $Record Unit ''); if ([string]::IsNullOrWhiteSpace($unit) -or $unit -eq 'N/A') { return [pscustomobject]@{Valid=$false;Reason='Sensor unit is empty or invalid'} }
    $timestamp=Get-PropertyValue $Record Timestamp $null; if ($timestamp -isnot [datetime]) { return [pscustomobject]@{Valid=$false;Reason='Sensor timestamp is invalid'} }
    return [pscustomobject]@{Valid=$true;Reason='Valid'}
}
function Invoke-LhmSensorProvider {
    $runtime = if ($PSEdition -eq 'Desktop') { 'powershell51' } else { 'powershell7' }
    $exe = if ($PSEdition -eq 'Desktop') { 'powershell.exe' } else { 'pwsh.exe' }
    $helperRoot = Join-Path $PSScriptRoot (Join-Path 'runtimes' (Join-Path 'windows-x64' $runtime))
    $helper = Join-Path $helperRoot 'LibreHardwareMonitor.Helper.ps1'
    $resultPath = Join-Path ([IO.Path]::GetTempPath()) ('ck-lhm-result-' + [guid]::NewGuid().ToString('N') + '.json')
    $stdout = Join-Path ([IO.Path]::GetTempPath()) ('ck-lhm-out-' + [guid]::NewGuid().ToString('N') + '.txt')
    $stderr = Join-Path ([IO.Path]::GetTempPath()) ('ck-lhm-err-' + [guid]::NewGuid().ToString('N') + '.txt')
    if (!(Test-Path -LiteralPath $helper)) { return [pscustomobject]@{Status='Unavailable';Error='LHM helper not found';Result=$null} }
    try {
        $p = Start-Process -FilePath $exe -ArgumentList @('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $helper),'-ResultPath',('"{0}"' -f $resultPath)) -RedirectStandardOutput $stdout -RedirectStandardError $stderr -PassThru
        if (!$p.WaitForExit(30000)) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue; return [pscustomobject]@{Status='Timeout';Error='LHM helper exceeded 30 seconds';Result=$null} }
        if (!(Test-Path -LiteralPath $resultPath)) { return [pscustomobject]@{Status='Error';Error=('LHM helper returned no JSON (exit '+$p.ExitCode+')');Result=$null} }
        try { $json=Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json } catch { return [pscustomobject]@{Status='Error';Error='LHM helper JSON was invalid';Result=$null} }
        return [pscustomobject]@{Status=[string]$json.status;Error=if(@($json.errors).Count){[string](@($json.errors) -join '; ')}else{''};Result=$json}
    } catch { return [pscustomobject]@{Status='Error';Error=$_.Exception.Message;Result=$null} }
    finally { Remove-Item $resultPath,$stdout,$stderr -Force -ErrorAction SilentlyContinue }
}
function Get-SensorSnapshot { param([scriptblock]$Provider,[string]$ProviderName='Fixture/Optional')
    $now=Get-Date; if ($null -eq $Provider) { return [pscustomobject][ordered]@{Available=$false; Provider='None'; ProviderStatus='Unavailable'; Sensors=@(); Timestamp=$now; Note='No sensor provider was requested' } }
    try { $raw=@(& $Provider); $valid=@(); $errors=@(); $seen=@{}; foreach($item in $raw){ try { $check=Test-SensorRecord $item; if(!$check.Valid){$errors += $check.Reason; continue}; $key="$(Get-PropertyValue $item Component 'N/A')|$(Get-PropertyValue $item SensorName 'N/A')|$(Get-PropertyValue $item Source 'N/A')"; if($seen.ContainsKey($key)){continue}; $seen[$key]=$true; $valid += $item } catch {$errors += $_.Exception.Message} }; return [pscustomobject][ordered]@{Available=($valid.Count -gt 0); Provider=$ProviderName; ProviderStatus=if($errors.Count){'Partial'}else{'Available'}; Sensors=$valid; Timestamp=$now; Errors=$errors; Note='Sensor values require an optional validated provider' } } catch { return [pscustomobject][ordered]@{Available=$false; Provider=$ProviderName; ProviderStatus='Error'; Sensors=@(); Timestamp=$now; Errors=@($_.Exception.Message); Note='Sensor provider failed; scan continued' } }
}
function Convert-WindowsDate { param($Value)
    if ($Value -is [datetime]) { return $Value }
    if ($Value) { try { return [Management.ManagementDateTimeConverter]::ToDateTime([string]$Value) } catch { return 'N/A' } }
    return 'N/A'
}

function Invoke-SafeModule { param([string]$Name, [scriptblock]$ScriptBlock)
    try { & $ScriptBlock } catch { [pscustomobject]@{ Module = $Name; Error = $_.Exception.Message } }
}
function Get-ModuleErrorResults {
    param($Result)
    foreach ($item in @($Result)) {
        if ($null -eq $item) { continue }
        $propertyNames = @($item.PSObject.Properties | ForEach-Object { $_.Name })
        if ($propertyNames -contains 'Error') { ,$item }
    }
}

function Get-CimSafe { param([string]$ClassName, [string]$Filter = '')
    try {
        if ($Filter) { return @(Get-CimInstance -ClassName $ClassName -Filter $Filter -ErrorAction Stop) }
        return @(Get-CimInstance -ClassName $ClassName -ErrorAction Stop)
    } catch {
        try {
            if ($Filter) { return @(Get-WmiObject -Class $ClassName -Filter $Filter -ErrorAction Stop) }
            return @(Get-WmiObject -Class $ClassName -ErrorAction Stop)
        } catch { return @() }
    }
}

function Get-ComputerSection {
    $cs = @(Get-CimSafe 'Win32_ComputerSystem')[0]; $bios = @(Get-CimSafe 'Win32_BIOS')[0]
    $os = @(Get-CimSafe 'Win32_OperatingSystem')[0]; $board=@(Get-CimSafe 'Win32_BaseBoard')[0]
    $bootMode='N/A'; try { if (Test-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State') { $bootMode='UEFI' } else {$bootMode='Legacy/Unknown'} } catch {}
    $secureBoot='N/A'; try { if (Get-Command Confirm-SecureBootUEFI -ErrorAction SilentlyContinue) { $secureBoot=if (Confirm-SecureBootUEFI -ErrorAction Stop) {'Enabled'} else {'Disabled'} } } catch { $secureBoot='Not available' }
    $tpm=@(Get-CimSafe 'Win32_Tpm')[0]
    $lastBoot=Get-PropertyValue $os LastBootUpTime $null; $install=Get-PropertyValue $os InstallDate $null
    $lastBootDate=Convert-WindowsDate $lastBoot; $installDate=Convert-WindowsDate $install
    $uptime=if($lastBootDate -is [datetime]){[math]::Round(((Get-Date)-$lastBootDate).TotalHours,1)}else{'N/A'}
    $displayVersion='N/A'; $ubr='N/A'; try { $cv=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop; $displayVersion=Get-PropertyValue $cv DisplayVersion (Get-PropertyValue $cv ReleaseId 'N/A'); $ubr=Get-PropertyValue $cv UBR 'N/A' } catch {}
    [pscustomobject]@{
        Hostname = $env:COMPUTERNAME; Manufacturer = Get-PropertyValue $cs Manufacturer 'N/A'
        Model = Get-PropertyValue $cs Model 'N/A'; SerialNumber = Get-PropertyValue $bios SerialNumber 'N/A'
        BIOS = "$(Get-PropertyValue $bios Manufacturer 'N/A') $(Get-PropertyValue $bios SMBIOSBIOSVersion 'N/A')"
        Windows = "$(Get-PropertyValue $os Caption 'N/A') $(Get-PropertyValue $os Version 'N/A')"
        Motherboard=[pscustomobject]@{ Manufacturer=Get-PropertyValue $board Manufacturer 'N/A'; Product=Get-PropertyValue $board Product 'N/A'; Version=Get-PropertyValue $board Version 'N/A'; Chipset='N/A' }
        BIOSDetail=[pscustomobject]@{ Vendor=Get-PropertyValue $bios Manufacturer 'N/A'; Version=Get-PropertyValue $bios SMBIOSBIOSVersion 'N/A'; ReleaseDate=Get-PropertyValue $bios ReleaseDate 'N/A'; BootMode=$bootMode; SecureBoot=$secureBoot }
        TPM=[pscustomobject]@{ Present=if($tpm){$true}else{$false}; Version=if($tpm){Get-PropertyValue $tpm SpecVersion 'N/A'}else{'N/A'}; Status=if($tpm){Get-PropertyValue $tpm IsEnabled_InitialValue 'N/A'}else{'Not available'} }
        WindowsDetail=[pscustomobject]@{ ProductName=Get-PropertyValue $os Caption 'N/A'; Edition=Get-PropertyValue $os Caption 'N/A'; DisplayVersion=$displayVersion; Version=Get-PropertyValue $os Version 'N/A'; Build=Get-PropertyValue $os BuildNumber 'N/A'; UBR=$ubr; Architecture=Get-PropertyValue $os OSArchitecture 'N/A'; InstallDate=$installDate; LastBootTime=$lastBootDate; UptimeHours=$uptime }
    }
}

function Get-CpuArchitectureName { param($Architecture)
    $map=@{0='x86';1='MIPS';2='Alpha';3='PowerPC';5='ARM';6='Itanium';9='x64'}; if($null -eq $Architecture){return 'N/A'}; $key=[int]$Architecture; if($map.ContainsKey($key)){return $map[$key]}; return "Unknown ($Architecture)"
}
function Get-CpuSection {
    $cpu = @(Get-CimSafe 'Win32_Processor')[0]
    $archMap=@{0='x86';1='MIPS';2='Alpha';3='PowerPC';5='ARM';6='Itanium';9='x64'}; $arch=Get-PropertyValue $cpu Architecture $null; $archKey=if($null -ne $arch){[int]$arch}else{$null}
    $system= @(Get-CimSafe 'Win32_ComputerSystem')[0]; $hypervisor=Get-PropertyValue $system HypervisorPresent $null; $fw=Get-PropertyValue $cpu VirtualizationFirmwareEnabled $null; $vm=Get-PropertyValue $cpu VMMonitorModeExtensions $null
    $firmwareStatus=if($hypervisor -eq $true -and $fw -eq $false){'Unknown (hypervisor present; processor flag is not authoritative)'}elseif($null -eq $fw){'Unknown'}else{[string]$fw}; $hypervisorStatus=if($hypervisor -eq $true){'Detected'}elseif($hypervisor -eq $false){'Not detected'}else{'Unknown'}
    [pscustomobject]@{ Manufacturer=Get-PropertyValue $cpu Manufacturer 'N/A'; Name = Get-PropertyValue $cpu Name 'N/A'; ProcessorId=Get-PropertyValue $cpu ProcessorId 'N/A'; Architecture=Get-CpuArchitectureName $arch; ArchitectureRaw=if($null -ne $arch){$arch}else{'N/A'}; Family=Get-PropertyValue $cpu Family 'N/A'; Model=Get-PropertyValue $cpu Model 'N/A'; Stepping=Get-PropertyValue $cpu Stepping 'N/A'; Cores = Get-PropertyValue $cpu NumberOfCores 'N/A'; Threads = Get-PropertyValue $cpu NumberOfLogicalProcessors 'N/A'; LogicalProcessors=Get-PropertyValue $cpu NumberOfLogicalProcessors 'N/A'; BaseClockMHz = 'N/A'; CurrentClockMHz=Get-PropertyValue $cpu CurrentClockSpeed 'N/A'; MaxClockMHz = Get-PropertyValue $cpu MaxClockSpeed 'N/A'; L1CacheKB=Get-PropertyValue $cpu L1CacheSize 'N/A'; L2CacheKB=Get-PropertyValue $cpu L2CacheSize 'N/A'; L3CacheKB=Get-PropertyValue $cpu L3CacheSize 'N/A'; Socket=Get-PropertyValue $cpu SocketDesignation 'N/A'; VirtualizationCapability=if($vm -eq $true){'Supported'}elseif($hypervisor -eq $true){'Indeterminate (hypervisor detected)'}elseif($null -eq $vm){'Unknown'}else{'Not reported'}; VirtualizationFirmwareStatus=$firmwareStatus; HypervisorStatus=$hypervisorStatus; VirtualizationSupport=if($null -eq $fw){'Unknown'}else{$fw}; VirtualizationStatus=if($null -eq $vm){'Unknown'}else{$vm}; InstructionSets='N/A'; TDP='N/A' }
}

function Get-MemorySection {
    $cs = @(Get-CimSafe 'Win32_ComputerSystem')[0]; $os = @(Get-CimSafe 'Win32_OperatingSystem')[0]
    $total = Get-PropertyValue $cs TotalPhysicalMemory $null; $free = Get-PropertyValue $os FreePhysicalMemory $null
    if ($total) { $totalGB = [math]::Round([double]$total / 1GB, 2) } else { $totalGB = 'N/A' }
    if ($total -and $free) { $used = [math]::Round((1 - (([double]$free * 1KB) / [double]$total)) * 100, 1) } else { $used = 'N/A' }
    $modules = foreach ($m in @(Get-CimSafe 'Win32_PhysicalMemory')) { Convert-MemoryModule $m }
    [pscustomobject]@{ TotalGB = $totalGB; UsedPercent = $used; Modules = @($modules) }
}

function Get-MemoryTypeName { param($SmbiosType, $LegacyType)
    $raw = if ($null -ne $SmbiosType) { [int]$SmbiosType } elseif ($null -ne $LegacyType) { [int]$LegacyType } else { $null }
    $name = switch ($raw) { 20 {'DDR'} 21 {'DDR2'} 22 {'DDR2 FB-DIMM'} 24 {'DDR3'} 26 {'DDR4'} 27 {'LPDDR'} 28 {'LPDDR2'} 29 {'LPDDR3'} 30 {'LPDDR4'} 34 {'DDR5'} default {$null} }
    [pscustomobject]@{ Raw=if ($null -ne $raw) {$raw} else {'N/A'}; Name=if ($name) {$name} else {'Unknown'}; Status=if ($name) {'Reported'} else {'Unknown raw value'} }
}

function Get-MemoryFieldStatus { param($Object, [string]$Property)
    $value = Get-PropertyValue $Object $Property $null
    [pscustomobject]@{ Value=if ($null -ne $value -and "$value" -ne '') {$value} else {'N/A'}; Status=if ($null -ne $value -and "$value" -ne '') {'Reported'} else {'Not reported'} }
}

function Convert-MemoryModule { param($m)
    $type = Get-MemoryTypeName (Get-PropertyValue $m SMBIOSMemoryType $null) (Get-PropertyValue $m MemoryType $null)
    $manufacturer=Get-MemoryFieldStatus $m Manufacturer; $part=Get-MemoryFieldStatus $m PartNumber; $serial=Get-MemoryFieldStatus $m SerialNumber
    $speed=Get-MemoryFieldStatus $m Speed; $configured=Get-MemoryFieldStatus $m ConfiguredClockSpeed
    $voltage=Get-MemoryFieldStatus $m ConfiguredVoltage; if ($voltage.Value -ne 'N/A') { $voltage=[pscustomobject]@{Value="{0} mV" -f $voltage.Value; Status=$voltage.Status} }
    $capacity=Get-PropertyValue $m Capacity $null
    [pscustomobject][ordered]@{
        Slot=Get-MemoryFieldStatus $m DeviceLocator; BankLabel=Get-MemoryFieldStatus $m BankLabel
        CapacityGB=if ($capacity) {[math]::Round([double]$capacity / 1GB,2)} else {'N/A'}
        Manufacturer=$manufacturer; PartNumber=$part; SerialNumber=$serial
        MemoryType=$type.Name; SMBIOSMemoryType=$type.Raw; MemoryTypeStatus=$type.Status
        SpeedMHz=$speed; ConfiguredClockSpeedMHz=$configured; RatedSpeedMHz=[pscustomobject]@{Value='N/A';Status='Not reported by Win32_PhysicalMemory'}
        Voltage=$voltage; FormFactor=Get-MemoryFieldStatus $m FormFactor; Rank=Get-MemoryFieldStatus $m Rank
        ECC=Get-MemoryFieldStatus $m TotalWidth; DataWidth=Get-MemoryFieldStatus $m DataWidth
        Channel=Get-MemoryFieldStatus $m InterleaveDataDepth; FieldStatus=[pscustomobject]@{ RatedSpeed='Not reported'; DataRate='Not calculated because MHz/MT/s source is ambiguous'; Voltage='Only firmware-reported value used' }
    }
}

function Get-DiskSection {
    $disks = @(Get-CimSafe 'Win32_DiskDrive')
    if (!$disks) { return @([pscustomobject]@{ DiskNumber='N/A'; Model='N/A'; CapacityGB='N/A'; MediaType='N/A'; Health='N/A'; Partitions=@(); Note='Storage data unavailable' }) }
    foreach ($d in $disks) {
        $health = Get-PropertyValue $d Status 'N/A'
        $parts = foreach ($p in @(Get-CimSafe 'Win32_DiskPartition' ("DiskIndex = {0}" -f (Get-PropertyValue $d Index -1)))) { $logical = @(Get-CimSafe 'Win32_LogicalDisk' ("DeviceID = '{0}'" -f (Get-PropertyValue $p Name '')))[0]; $freeSpace=Get-PropertyValue $logical FreeSpace $null; $size=Get-PropertyValue $p Size $null; $freeGB=if($freeSpace){[math]::Round([double]$freeSpace/1GB,2)}else{'N/A'}; $sizeGB=if($size){[math]::Round([double]$size/1GB,2)}else{'N/A'}; $usedGB=if($size -and $freeSpace){[math]::Round(([double]$size-[double]$freeSpace)/1GB,2)}else{'N/A'}; [pscustomobject]@{ PartitionNumber=Get-PropertyValue $p Index 'N/A'; Name=Get-PropertyValue $p Name 'N/A'; DriveLetter=Get-PropertyValue $logical DeviceID 'N/A'; VolumeLabel=Get-PropertyValue $logical VolumeName 'N/A'; FileSystem=Get-PropertyValue $logical FileSystem 'N/A'; SizeGB=$sizeGB; UsedGB=$usedGB; FreeGB=$freeGB; FreePercent=if($size -and $freeSpace -and [double]$size -gt 0){[math]::Round(([double]$freeSpace/[double]$size)*100,1)}else{'N/A'} } }
        [pscustomobject]@{ DiskNumber = if ($null -ne (Get-PropertyValue $d Index $null)) { "Disk $($d.Index)" } else { 'N/A' }; Model = Get-PropertyValue $d Model 'N/A'; Manufacturer=Get-PropertyValue $d Manufacturer 'N/A'; SerialNumber=Get-PropertyValue $d SerialNumber 'N/A'; FirmwareRevision=Get-PropertyValue $d FirmwareRevision 'N/A'; CapacityGB = if ($d.Size) { [math]::Round([double]$d.Size / 1GB, 2) } else { 'N/A' }; MediaType = Get-PropertyValue $d MediaType 'N/A'; InterfaceType=Get-PropertyValue $d InterfaceType 'N/A'; BytesPerSector=Get-PropertyValue $d BytesPerSector 'N/A'; Health = $health; SMART=[pscustomobject]@{Supported='Not supported by standard WMI source'; Health='Unavailable'; Temperature='Unavailable'; PowerOnHours='Unavailable'; PowerCycleCount='Unavailable'}; Partitions=@($parts); Note = 'Health is based on Windows-reported status only; SMART/physical health is not guaranteed' }
    }
}

function Get-BatteryWear { param($DesignCapacity, $FullChargeCapacity)
    if (!$DesignCapacity -or !$FullChargeCapacity -or [double]$DesignCapacity -le 0) { return $null }
    return [math]::Round((1 - ([double]$FullChargeCapacity / [double]$DesignCapacity)) * 100, 1)
}

function Get-BatterySection {
    $b = @(Get-CimSafe 'Win32_Battery')[0]
    if (!$b) { return [pscustomobject]@{ Available=$false; DesignCapacity='N/A'; FullChargeCapacity='N/A'; WearPercent='N/A'; Note='Battery data unavailable (desktop or unsupported firmware)' } }
    $design = Get-PropertyValue $b DesignCapacity $null; $full = Get-PropertyValue $b FullChargeCapacity $null
    [pscustomobject]@{ Available=$true; Manufacturer=Get-PropertyValue $b Manufacturer 'N/A'; Name=Get-PropertyValue $b Name 'N/A'; SerialNumber=Get-PropertyValue $b DeviceID 'N/A'; Chemistry=Get-PropertyValue $b Chemistry 'N/A'; DesignCapacity=if ($design) { $design } else { 'N/A' }; FullChargeCapacity=if ($full) { $full } else { 'N/A' }; CurrentCapacity=Get-PropertyValue $b EstimatedChargeRemaining 'N/A'; WearPercent=if ($null -ne (Get-BatteryWear $design $full)) { Get-BatteryWear $design $full } else { 'N/A' }; CycleCount='N/A'; ChargingStatus=Get-PropertyValue $b BatteryStatus 'N/A'; Voltage=Get-PropertyValue $b Voltage 'N/A'; EstimatedRuntime=Get-PropertyValue $b EstimatedRunTime 'N/A'; ChargeRate='N/A'; ACStatus='N/A'; Note='Battery values depend on firmware; charge percentage is not capacity in mWh' }
}

function Get-GpuSection { foreach ($g in @(Get-CimSafe 'Win32_VideoController')) { $ram=Get-PropertyValue $g AdapterRAM $null; $integrated=(Get-PropertyValue $g VideoProcessor '') -match 'Intel|UMA|Integrated'; [pscustomobject]@{ Name=Get-PropertyValue $g Name 'N/A'; Manufacturer=Get-PropertyValue $g AdapterCompatibility 'N/A'; Processor=Get-PropertyValue $g VideoProcessor 'N/A'; PnpDeviceId=Get-PropertyValue $g PNPDeviceID 'N/A'; Driver=Get-PropertyValue $g DriverVersion 'N/A'; DriverDate=Get-PropertyValue $g DriverDate 'N/A'; Status=Get-PropertyValue $g Status 'N/A'; ProblemCode=Get-PropertyValue $g ConfigManagerErrorCode 'N/A'; DedicatedVRAMGB=if($integrated){'N/A'}else{if($ram){[math]::Round([double]$ram/1GB,2)}else{'N/A'}}; SharedMemoryGB='N/A'; AdapterMemoryReportedGB=if($ram){[math]::Round([double]$ram/1GB,2)}else{'N/A'}; MemoryNote=if($integrated){'Integrated GPU; adapter memory is not treated as dedicated VRAM'}else{'Dedicated/shared split not reported by WMI'}; Resolution="$(Get-PropertyValue $g CurrentHorizontalResolution 'N/A')x$(Get-PropertyValue $g CurrentVerticalResolution 'N/A')"; RefreshRateHz=Get-PropertyValue $g CurrentRefreshRate 'N/A'; Temperature='Unavailable'; Power='Unavailable' } } }
function Get-NetworkSection { $configs=@(Get-CimSafe 'Win32_NetworkAdapterConfiguration'); foreach ($n in @(Get-CimSafe 'Win32_NetworkAdapter')) { if ((Get-PropertyValue $n NetEnabled $false) -or (Get-PropertyValue $n MACAddress $null)) { $mac=Get-PropertyValue $n MACAddress $null; $c=$configs | Where-Object { $_.MACAddress -eq $mac } | Select-Object -First 1; [pscustomobject]@{ Name=Get-PropertyValue $n Name 'N/A'; Manufacturer=Get-PropertyValue $n Manufacturer 'N/A'; InterfaceDescription=Get-PropertyValue $n Name 'N/A'; AdapterType=Get-PropertyValue $n AdapterType 'N/A'; PhysicalAdapter=Get-PropertyValue $n PhysicalAdapter 'N/A'; MAC=if($mac){$mac}else{'N/A'}; IPv4=if($c){(@($c.IPAddress | Where-Object {$_ -match '\.'}) -join ', ')}else{'N/A'}; IPv6=if($c){(@($c.IPAddress | Where-Object {$_ -match ':'}) -join ', ')}else{'N/A'}; Subnet=if($c){(@($c.IPSubnet) -join ', ')}else{'N/A'}; Gateway=if($c){(@($c.DefaultIPGateway) -join ', ')}else{'N/A'}; DNS=if($c){(@($c.DNSServerSearchOrder) -join ', ')}else{'N/A'}; DHCPEnabled=if($c){Get-PropertyValue $c DHCPEnabled 'N/A'}else{'N/A'}; DHCPServer=if($c){Get-PropertyValue $c DHCPServer 'N/A'}else{'N/A'}; Speed=Get-PropertyValue $n Speed 'N/A'; ConnectionStatus=Get-PropertyValue $n NetConnectionStatus 'N/A'; DriverVersion=Get-PropertyValue $n DriverVersion 'N/A'; DriverDate=Get-PropertyValue $n DriverDate 'N/A'; PnpDeviceId=Get-PropertyValue $n PNPDeviceID 'N/A'; Status=Get-PropertyValue $n Status 'N/A'; ProblemCode=Get-PropertyValue $n ConfigManagerErrorCode 'N/A'; Classification=if((Get-PropertyValue $n PhysicalAdapter $false) -eq $true){'Physical'}else{'Virtual/Software'} } } } }
function Get-PnpClassification { param($Code)
    if ($null -eq $Code -or "$Code" -eq '' -or "$Code" -eq 'N/A') { return [pscustomobject]@{ Status='N/A'; Interpretation='Problem status unavailable; inspect manually if needed' } }
    switch ([int]$Code) {
        0 { return [pscustomobject]@{ Status='OK'; Interpretation='No reported problem' } }
        22 { return [pscustomobject]@{ Status='WARNING'; Interpretation='Device is disabled; this does not confirm physical failure' } }
        24 { return [pscustomobject]@{ Status='WARNING'; Interpretation='Device may be absent, not functioning correctly, or missing a required driver; further inspection required' } }
        28 { return [pscustomobject]@{ Status='WARNING'; Interpretation='Required drivers are not installed' } }
        31 { return [pscustomobject]@{ Status='WARNING'; Interpretation='Device is not working properly' } }
        default { return [pscustomobject]@{ Status='WARNING'; Interpretation='Windows reported a problem; physical failure is not confirmed' } }
    }
}
function Get-PnpProblemSection { foreach ($p in @(Get-CimSafe 'Win32_PnPEntity' 'ConfigManagerErrorCode <> 0')) { $code=Get-PropertyValue $p ConfigManagerErrorCode $null; $classification=Get-PnpClassification $code; [pscustomobject]@{ Name=Get-PropertyValue $p Name 'N/A'; Code=if ($null -ne $code) {$code} else {'N/A'}; Status=$classification.Status; Interpretation=$classification.Interpretation } } }

function Get-OverallStatus { param($Report)
    if (@($Report.ModuleErrors).Count -gt 0) { return 'ERROR' }
    if (@($Report.PnpProblems | Where-Object { $_.Status -eq 'WARNING' }).Count -gt 0) { return 'WARNING' }
    if ($Report.Battery.WearPercent -is [double] -and $Report.Battery.WearPercent -ge 30) { return 'WARNING' }
    return 'OK'
}

function Format-Section { param([string]$Title, $Items)
    $lines = @("[$Title]")
    if (!$Items) { return $lines + '  N/A - data unavailable' }
    foreach ($item in @($Items)) { $pairs = foreach ($p in $item.PSObject.Properties) { if ($p.Name -notin @('Note','Available')) { "$($p.Name): $($p.Value)" } }; $lines += "  $($pairs -join ' | ')"; if (($item.PSObject.Properties.Name -contains 'Note') -and $item.Note) { $lines += "  Note: $($item.Note)" } }
    return $lines
}

function Get-StatusCounts { param($Report)
    $counts = [ordered]@{ OK=0; WARNING=0; ERROR=0; 'N/A'=0 }
    foreach ($s in @($Report.OverallStatus)) { if ($counts.Contains($s)) { $counts[$s]++ } }
    foreach ($p in @($Report.PnpProblems)) { $s = if ($p.Status) { $p.Status } else { 'N/A' }; if ($counts.Contains($s)) { $counts[$s]++ } }
    foreach ($e in @($Report.ModuleErrors)) { $counts.ERROR++ }
    return [pscustomobject]$counts
}

function Get-HardwareRows { param($Report)
    $cpuName = Get-PropertyValue $Report.CPU Name 'N/A'; $cpuDetail = if ($cpuName -ne 'N/A') { $cpuName } else { 'Not available' }
    $ramTotal=Get-PropertyValue $Report.Memory TotalGB 'N/A'; $ramUsed=Get-PropertyValue $Report.Memory UsedPercent 'N/A'; $ramDetail = if ($ramTotal -ne 'N/A') { "$ramTotal GB"; if ($ramUsed -ne 'N/A') { $ramDetail += " | Usage $ramUsed%" } } else { 'Not available' }
    $disks = @($Report.Disks); $gpu = @($Report.GPU); $battery = $Report.Battery; $batteryAvailable=Get-PropertyValue $battery Available $false; $batteryWear=Get-PropertyValue $battery WearPercent 'N/A'
    $rows=@([pscustomobject]@{ Component='CPU'; Detail="$cpuDetail | Cores $(Get-PropertyValue $Report.CPU Cores 'N/A') | Threads $(Get-PropertyValue $Report.CPU Threads (Get-PropertyValue $Report.CPU LogicalProcessors 'N/A')) | Current $(Get-PropertyValue $Report.CPU CurrentClockMHz 'N/A') MHz | Max $(Get-PropertyValue $Report.CPU MaxClockMHz 'N/A') MHz"; Status=if ($cpuName -eq 'N/A') {'N/A'} else {'OK'} }
      [pscustomobject]@{ Component='RAM'; Detail=$ramDetail; Status=if ($ramTotal -eq 'N/A') {'N/A'} else {'OK'} }
      [pscustomobject]@{ Component='Storage'; Detail=if ($disks.Count) { "$($disks.Count) drive(s) detected" } else {'Not available'}; Status=if ($disks.Count) {'OK'} else {'N/A'} }
      [pscustomobject]@{ Component='GPU'; Detail=if ($gpu.Count) { (($gpu | ForEach-Object { Get-PropertyValue $_ Name 'N/A' }) -join ', ') } else {'Not available'}; Status=if ($gpu.Count) {'OK'} else {'N/A'} }
      [pscustomobject]@{ Component='Battery'; Detail=if (!$batteryAvailable) {'Not available'} elseif ($batteryWear -ne 'N/A') { "Wear $batteryWear%" } else {'Data incomplete'}; Status=if (!$batteryAvailable -or $batteryWear -eq 'N/A') {'N/A'} elseif ($batteryWear -ge 30) {'WARNING'} else {'OK'} })
    foreach($d in $disks){ $rows += [pscustomobject]@{ Component=$d.DiskNumber; Detail="$(Get-PropertyValue $d Model 'N/A') | $(Get-PropertyValue $d CapacityGB 'N/A') GB | $(Get-PropertyValue $d MediaType 'N/A') | Health $(Get-PropertyValue $d Health 'N/A')"; Status=if ((Get-PropertyValue $d Health 'N/A') -eq 'N/A') {'N/A'} else {'OK'} } }
    foreach($m in @(Get-PropertyValue $Report.Memory Modules @())){ $slot=Get-StructuredValue $m Slot; $capacity=Get-PropertyValue $m CapacityGB 'N/A'; $rows += [pscustomobject]@{ Component="RAM $slot"; Detail="$capacity GB | $(Get-PropertyValue $m MemoryType 'Unknown') | Speed $(Get-StructuredValue $m SpeedMHz) MHz | Config $(Get-StructuredValue $m ConfiguredClockSpeedMHz) MHz | $(Get-StructuredValue $m Manufacturer) $(Get-StructuredValue $m PartNumber)"; Status=if ($slot -eq 'N/A' -and $capacity -eq 'N/A') {'N/A'} else {'OK'} } }
    return $rows
}

function Get-Findings { param($Report)
    foreach ($p in @($Report.PnpProblems)) {
        $action = switch ([int]$p.Code) { 22 {'Periksa apakah sengaja dinonaktifkan'} 24 {'Periksa driver dan status perangkat'} 28 {'Instal atau perbarui driver'} default {'Periksa perangkat dan driver'} }
        [pscustomobject]@{ Device=$p.Name; Finding="$($p.Interpretation) - Code $($p.Code)"; Action=$action; Status=$p.Status }
    }
    foreach ($e in @($Report.ModuleErrors)) { [pscustomobject]@{ Device=$e.Module; Finding='Module scan failed'; Action='Ulangi scan dan periksa akses WMI/CIM'; Status='ERROR' } }
}

function Get-ExportReport { param($Report, [switch]$RedactSensitive)
    $computer = $Report.Computer | Select-Object *; $network = foreach ($n in @($Report.Network)) { $n | Select-Object * }; $memory = [pscustomobject][ordered]@{ TotalGB=Get-PropertyValue $Report.Memory TotalGB 'N/A'; UsedPercent=Get-PropertyValue $Report.Memory UsedPercent 'N/A'; Modules=foreach($m in @(Get-PropertyValue $Report.Memory Modules @())) { $m | Select-Object * } }
    if ($RedactSensitive) { $computer.SerialNumber='[REDACTED]'; foreach ($n in $network) { $n.MAC='[REDACTED]' } }
    $sensors=Get-PropertyValue $Report Sensors (Get-SensorSnapshot)
    [pscustomobject][ordered]@{ Title='CK Hardware Scanner | Inspection Report'; Generated=$Report.Generated; OverallStatus=$Report.OverallStatus; Computer=$computer; Summary=Get-StatusCounts $Report; Hardware=Get-HardwareRows $Report; Findings=Get-Findings $Report; Details=[pscustomobject]@{ System=$computer; CPU=$Report.CPU; Memory=$memory; Disks=$Report.Disks; Battery=$Report.Battery; GPU=$Report.GPU; Network=$network; PnpProblems=$Report.PnpProblems; ModuleErrors=$Report.ModuleErrors; Sensors=$sensors } }
}

function ConvertTo-HtmlSafe { param([object]$Value)
    if ($null -eq $Value) { return '' }; return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}
function Format-AsciiTable { param([string]$Title, $Rows)
    $items=@($Rows); $left=8; $right=5; foreach($r in $items){$left=[math]::Max($left,([string]$r.Property).Length);$right=[math]::Max($right,([string]$r.Value).Length)}; $border='+' + ('-'*($left+2)) + '+' + ('-'*($right+2)) + '+'; $out=@($Title,$border,('| {0} | {1} |' -f 'Property'.PadRight($left),'Value'.PadRight($right)),$border); foreach($r in $items){$out += '| {0} | {1} |' -f ([string]$r.Property).PadRight($left),([string]$r.Value).PadRight($right)}; $out += $border; return $out
}
function Get-ReportDetailLines { param($Export)
    $cpu=$Export.Details.CPU; $sys=$Export.Details.System; $mb=Get-PropertyValue $sys Motherboard ([pscustomobject]@{}); $bios=Get-PropertyValue $sys BIOSDetail ([pscustomobject]@{}); $tpm=Get-PropertyValue $sys TPM ([pscustomobject]@{}); $win=Get-PropertyValue $sys WindowsDetail ([pscustomobject]@{})
    $l=@('DETAILS'); $l += Format-AsciiTable 'CPU' @([pscustomobject]@{Property='Manufacturer';Value=Get-PropertyValue $cpu Manufacturer 'N/A'}, [pscustomobject]@{Property='Model';Value=Get-PropertyValue $cpu Name 'N/A'}, [pscustomobject]@{Property='Architecture';Value=Get-PropertyValue $cpu Architecture 'N/A'}, [pscustomobject]@{Property='Cores / Threads';Value="$(Get-PropertyValue $cpu Cores 'N/A') / $(Get-PropertyValue $cpu Threads 'N/A')"}, [pscustomobject]@{Property='Current / Max Clock';Value="$(Get-PropertyValue $cpu CurrentClockMHz 'N/A') / $(Get-PropertyValue $cpu MaxClockMHz 'N/A') MHz"}, [pscustomobject]@{Property='L1 / L2 / L3 Cache';Value="$(Get-PropertyValue $cpu L1CacheKB 'N/A') / $(Get-PropertyValue $cpu L2CacheKB 'N/A') / $(Get-PropertyValue $cpu L3CacheKB 'N/A') KB"}, [pscustomobject]@{Property='Socket';Value=Get-PropertyValue $cpu Socket 'N/A'}, [pscustomobject]@{Property='Virtualization capability';Value=Get-PropertyValue $cpu VirtualizationCapability 'N/A'}, [pscustomobject]@{Property='Firmware / Hypervisor';Value="$(Get-PropertyValue $cpu VirtualizationFirmwareStatus 'N/A') / $(Get-PropertyValue $cpu HypervisorStatus 'N/A')"})
    $l += Format-AsciiTable 'MOTHERBOARD' @([pscustomobject]@{Property='Manufacturer';Value=Get-PropertyValue $mb Manufacturer 'N/A'}, [pscustomobject]@{Property='Product';Value=Get-PropertyValue $mb Product 'N/A'}, [pscustomobject]@{Property='Version / Chipset';Value="$(Get-PropertyValue $mb Version 'N/A') / $(Get-PropertyValue $mb Chipset 'N/A')"})
    $l += Format-AsciiTable 'BIOS / TPM' @([pscustomobject]@{Property='BIOS Vendor / Version';Value="$(Get-PropertyValue $bios Vendor 'N/A') / $(Get-PropertyValue $bios Version 'N/A')"}, [pscustomobject]@{Property='Release Date';Value=Get-PropertyValue $bios ReleaseDate 'N/A'}, [pscustomobject]@{Property='Boot Mode / Secure Boot';Value="$(Get-PropertyValue $bios BootMode 'N/A') / $(Get-PropertyValue $bios SecureBoot 'N/A')"}, [pscustomobject]@{Property='TPM Present / Version';Value="$(Get-PropertyValue $tpm Present 'N/A') / $(Get-PropertyValue $tpm Version 'N/A')"}, [pscustomobject]@{Property='TPM Status';Value=Get-PropertyValue $tpm Status 'N/A'})
    $windowsRows=@(); foreach($p in @($win.PSObject.Properties)){$windowsRows += [pscustomobject]@{Property=$p.Name;Value=$p.Value}}; $l += Format-AsciiTable 'WINDOWS' $windowsRows
    $l += 'RAM MODULES'; foreach($m in @((Get-PropertyValue $Export.Details.Memory Modules @()))){ $l += Format-AsciiTable ("RAM Slot $(Get-StructuredValue $m Slot)") @([pscustomobject]@{Property='Capacity / Type';Value="$(Get-PropertyValue $m CapacityGB 'N/A') GB / $(Get-PropertyValue $m MemoryType 'Unknown')"}, [pscustomobject]@{Property='Speed / Configured';Value="$(Get-StructuredValue $m SpeedMHz) / $(Get-StructuredValue $m ConfiguredClockSpeedMHz) MHz"}, [pscustomobject]@{Property='Manufacturer';Value=Get-StructuredValue $m Manufacturer}, [pscustomobject]@{Property='Part Number';Value=Get-StructuredValue $m PartNumber}, [pscustomobject]@{Property='Serial';Value=Get-StructuredValue $m SerialNumber}) }
    $l += 'STORAGE'; foreach($d in @($Export.Details.Disks)){ $l += Format-AsciiTable (Get-PropertyValue $d DiskNumber 'Disk') @([pscustomobject]@{Property='Model / Manufacturer';Value="$(Get-PropertyValue $d Model 'N/A') / $(Get-PropertyValue $d Manufacturer 'N/A')"}, [pscustomobject]@{Property='Serial / Firmware';Value="$(Get-PropertyValue $d SerialNumber 'N/A') / $(Get-PropertyValue $d FirmwareRevision 'N/A')"}, [pscustomobject]@{Property='Capacity / Media / Interface';Value="$(Get-PropertyValue $d CapacityGB 'N/A') GB / $(Get-PropertyValue $d MediaType 'N/A') / $(Get-PropertyValue $d InterfaceType 'N/A')"}, [pscustomobject]@{Property='SMART';Value="$(Get-PropertyValue (Get-PropertyValue $d SMART ([pscustomobject]@{})) Health 'Unavailable')"}); foreach($p in @($d.Partitions)){ $l += "  Partition $(Get-PropertyValue $p PartitionNumber 'N/A') | $(Get-PropertyValue $p DriveLetter 'N/A') | $(Get-PropertyValue $p FileSystem 'N/A') | Size $(Get-PropertyValue $p SizeGB 'N/A') GB | Used $(Get-PropertyValue $p UsedGB 'N/A') GB | Free $(Get-PropertyValue $p FreeGB 'N/A') GB" } }
    $b=$Export.Details.Battery; $l += Format-AsciiTable 'BATTERY' @([pscustomobject]@{Property='Manufacturer / Name';Value="$(Get-PropertyValue $b Manufacturer 'N/A') / $(Get-PropertyValue $b Name 'N/A')"}, [pscustomobject]@{Property='Design / Full Charge';Value="$(Get-PropertyValue $b DesignCapacity 'N/A') / $(Get-PropertyValue $b FullChargeCapacity 'N/A') mWh"}, [pscustomobject]@{Property='Wear / Current';Value="$(Get-PropertyValue $b WearPercent 'N/A')% / $(Get-PropertyValue $b CurrentCapacity 'N/A')"}, [pscustomobject]@{Property='Status / Voltage';Value="$(Get-PropertyValue $b ChargingStatus 'N/A') / $(Get-PropertyValue $b Voltage 'N/A')"})
    $l += 'GPU'; foreach($g in @($Export.Details.GPU)){ $l += Format-AsciiTable (Get-PropertyValue $g Name 'GPU') @([pscustomobject]@{Property='Manufacturer / Processor';Value="$(Get-PropertyValue $g Manufacturer 'N/A') / $(Get-PropertyValue $g Processor 'N/A')"},[pscustomobject]@{Property='Driver / Status';Value="$(Get-PropertyValue $g Driver 'N/A') / $(Get-PropertyValue $g Status 'N/A')"},[pscustomobject]@{Property='Dedicated / Shared VRAM';Value="$(Get-PropertyValue $g DedicatedVRAMGB 'N/A') / $(Get-PropertyValue $g SharedMemoryGB 'N/A') GB"},[pscustomobject]@{Property='Resolution / Refresh';Value="$(Get-PropertyValue $g Resolution 'N/A') / $(Get-PropertyValue $g RefreshRateHz 'N/A') Hz"},[pscustomobject]@{Property='Temperature / Power';Value="$(Get-PropertyValue $g Temperature 'Unavailable') / $(Get-PropertyValue $g Power 'Unavailable')"}) }
    $l += 'NETWORK'; foreach($n in @($Export.Details.Network)){ $l += Format-AsciiTable (Get-PropertyValue $n Name 'Adapter') @([pscustomobject]@{Property='Type / Classification';Value="$(Get-PropertyValue $n AdapterType 'N/A') / $(Get-PropertyValue $n Classification 'N/A')"},[pscustomobject]@{Property='MAC / Status';Value="$(Get-PropertyValue $n MAC 'N/A') / $(Get-PropertyValue $n ConnectionStatus 'N/A')"},[pscustomobject]@{Property='IPv4 / IPv6';Value="$(Get-PropertyValue $n IPv4 'N/A') / $(Get-PropertyValue $n IPv6 'N/A')"},[pscustomobject]@{Property='Gateway / DNS';Value="$(Get-PropertyValue $n Gateway 'N/A') / $(Get-PropertyValue $n DNS 'N/A')"},[pscustomobject]@{Property='Speed / DHCP';Value="$(Get-PropertyValue $n Speed 'N/A') / $(Get-PropertyValue $n DHCPEnabled 'N/A')"}) }
    $sensor=$Export.Details.Sensors; $sensorRows=@([pscustomobject]@{Property='Provider';Value=Get-PropertyValue $sensor Provider 'None'},[pscustomobject]@{Property='Status';Value=Get-PropertyValue $sensor ProviderStatus 'Unavailable'},[pscustomobject]@{Property='Available sensors';Value=(@(Get-PropertyValue $sensor Sensors @())).Count},[pscustomobject]@{Property='Note';Value=Get-PropertyValue $sensor Note 'N/A'}); foreach($sr in @((Get-PropertyValue $sensor Sensors @()))){$sensorRows += [pscustomobject]@{Property="$(Get-PropertyValue $sr Component 'N/A') / $(Get-PropertyValue $sr SensorName 'N/A')";Value="$(Get-PropertyValue $sr Value 'N/A') $(Get-PropertyValue $sr Unit 'N/A') | $(Get-PropertyValue $sr Status 'N/A')"} }; $l += Format-AsciiTable 'SENSORS' $sensorRows
    return $l
}

function Convert-ReportToText { param($Report, [switch]$RedactSensitive)
    $export = Get-ExportReport $Report -RedactSensitive:$RedactSensitive
    $computerHost=Get-PropertyValue $export.Computer Hostname 'N/A'; $model=Get-PropertyValue $export.Computer Model 'N/A'
    $lines = @('CK Hardware Scanner | Inspection Report', "Hostname: $computerHost | Model: $model", "Scan time: $($export.Generated) | Overall: $($export.OverallStatus)", '', 'SUMMARY', "OK: $($export.Summary.OK) | WARNING: $($export.Summary.WARNING) | ERROR: $($export.Summary.ERROR) | N/A: $($export.Summary.'N/A')", '', 'HARDWARE', ('Component'.PadRight(12) + ' | ' + 'Detail'.PadRight(48) + ' | Status') )
    foreach ($r in @($export.Hardware)) { $lines += ("{0} | {1} | {2}" -f $r.Component.PadRight(12), ([string]$r.Detail).PadRight(48), $r.Status) }
    $lines += ''; $lines += Get-ReportDetailLines $export; $lines += ''; $lines += 'FINDINGS'; $findings=@($export.Findings); if (!$findings) { $lines += 'None - no actionable findings reported' } else { foreach ($f in $findings) { $lines += "$($f.Device) | $($f.Finding) | $($f.Action)" } }; $lines += ''; $lines += 'Generated by CK Hardware Scanner'; return ($lines -join [Environment]::NewLine)
}

function Convert-ReportToHtml { param($Report, [switch]$RedactSensitive)
    $e=Get-ExportReport $Report -RedactSensitive:$RedactSensitive; $rows=''
    foreach($r in @($e.Hardware)){ $rows += ('<tr><td>{0}</td><td>{1}</td><td class="status-{2}">{2}</td></tr>' -f (ConvertTo-HtmlSafe $r.Component),(ConvertTo-HtmlSafe $r.Detail),(ConvertTo-HtmlSafe $r.Status)) }
    $find=''; foreach($f in @($e.Findings)){ $find += ('<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>' -f (ConvertTo-HtmlSafe $f.Device),(ConvertTo-HtmlSafe $f.Finding),(ConvertTo-HtmlSafe $f.Action)) }; if (!$find) {$find='<tr><td colspan="3">None - no actionable findings reported</td></tr>'}
    $html='<!doctype html><html><head><meta charset="utf-8"><title>CK Hardware Scanner</title><style>body{font:14px Arial;color:#202124;max-width:980px;margin:24px auto;padding:0 18px}h1{font-size:22px;margin:0 0 6px}h2{font-size:15px;border-bottom:1px solid #ddd;padding-bottom:5px;margin-top:22px}table{border-collapse:collapse;width:100%;margin-top:8px}th,td{border-bottom:1px solid #e5e5e5;text-align:left;padding:7px 8px}th{background:#f5f6f7}.status-OK{color:#176b35;font-weight:bold}.status-WARNING{color:#9a6700;font-weight:bold}.status-ERROR{color:#b42318;font-weight:bold}.status-N/A{color:#666;font-weight:bold}.meta{color:#555}.print{float:right}@media print{.print{display:none}body{margin:0;max-width:none}}</style></head><body>'
    $html += '<button class="print" onclick="window.print()">Print</button><h1>CK Hardware Scanner | Inspection Report</h1>'
    $html += ('<div class="meta">Hostname: {0} | Model: {1} | Scan time: {2} | Overall: <span class="status-{3}">{3}</span></div>' -f (ConvertTo-HtmlSafe (Get-PropertyValue $e.Computer Hostname 'N/A')),(ConvertTo-HtmlSafe (Get-PropertyValue $e.Computer Model 'N/A')),(ConvertTo-HtmlSafe $e.Generated),(ConvertTo-HtmlSafe $e.OverallStatus))
    $detailHtml='<table class="details"><tr><th>Property</th><th>Value</th></tr>'; foreach($line in @(Get-ReportDetailLines $e)){ if($line -match '^\|\s*(.*?)\s*\|\s*(.*?)\s*\|$'){ $detailHtml += '<tr><td>'+ (ConvertTo-HtmlSafe $matches[1].Trim()) +'</td><td>'+ (ConvertTo-HtmlSafe $matches[2].Trim()) +'</td></tr>' } elseif($line -notmatch '^\+|^DETAILS$'){ $detailHtml += '<tr><th colspan="2">'+(ConvertTo-HtmlSafe $line)+'</th></tr>' } }; $detailHtml += '</table>'
    $html += ('<h2>Summary</h2><div>OK: {0} &nbsp; WARNING: {1} &nbsp; ERROR: {2} &nbsp; N/A: {3}</div><h2>Hardware</h2><table><tr><th>Component</th><th>Detail</th><th>Status</th></tr>{4}</table><h2>Details</h2><div class="details">{5}</div><h2>Findings</h2><table><tr><th>Device</th><th>Finding</th><th>Action</th></tr>{6}</table><p class="meta">Generated by CK Hardware Scanner</p></body></html>' -f $e.Summary.OK,$e.Summary.WARNING,$e.Summary.ERROR,$e.Summary.'N/A',$rows,$detailHtml,$find)
    return $html
}

function Convert-ReportToJson { param($Report, [switch]$RedactSensitive) return ((Get-ExportReport $Report -RedactSensitive:$RedactSensitive) | ConvertTo-Json -Depth 8) }

function Convert-ReportToTextLegacy { param($Report, [switch]$RedactSensitive)
    $computer = $Report.Computer; $network = $Report.Network
    if ($RedactSensitive) { $computer = $computer | Select-Object *; $computer.SerialNumber = '[REDACTED]'; $network = foreach ($n in @($network)) { $copy=$n | Select-Object *; $copy.MAC='[REDACTED]'; $copy } }
    $lines = @('CK HARDWARE SCANNER - PHASE 1', "Generated: $($Report.Generated)", "Overall status: $($Report.OverallStatus)", '')
    $lines += Format-Section 'COMPUTER' $computer; $lines += Format-Section 'CPU' $Report.CPU; $lines += Format-Section 'MEMORY' $Report.Memory; $lines += Format-Section 'DISK' $Report.Disks; $lines += Format-Section 'BATTERY' $Report.Battery; $lines += Format-Section 'GPU' $Report.GPU; $lines += Format-Section 'NETWORK' $network; $lines += Format-Section 'PLUG-AND-PLAY PROBLEMS' $Report.PnpProblems
    if ($Report.ModuleErrors.Count) { $lines += ''; $lines += Format-Section 'MODULE ERRORS' $Report.ModuleErrors }
    return ($lines -join [Environment]::NewLine)
}

function New-ScanReport {
    $errors = @(); $sensorErrors = @(); $run = { param($name, $block) $result = Invoke-SafeModule $name $block; $moduleErrors=@(Get-ModuleErrorResults $result); if ($moduleErrors.Count -gt 0) { if ($name -eq 'Sensors') { $sensorErrors += $moduleErrors; return [pscustomobject][ordered]@{Available=$false;Provider='LibreHardwareMonitor';ProviderStatus='Error';Sensors=@();Timestamp=(Get-Date);Errors=@($moduleErrors | ForEach-Object {$_.Error});Note='Optional sensor provider failed; inventory scan continued'} }; $errors += $moduleErrors; return $null }; return $result }
    $report = [ordered]@{ Generated=(Get-Date); Computer=&$run 'Computer' { Get-ComputerSection }; CPU=&$run 'CPU' { Get-CpuSection }; Memory=&$run 'Memory' { Get-MemorySection }; Disks=&$run 'Disks' { Get-DiskSection }; Battery=&$run 'Battery' { Get-BatterySection }; GPU=&$run 'GPU' { @(Get-GpuSection) }; Network=&$run 'Network' { @(Get-NetworkSection) }; PnpProblems=&$run 'PnP' { @(Get-PnpProblemSection) }; Sensors=&$run 'Sensors' { $snap=Invoke-LhmSensorProvider; if($snap.Status -in @('Error','Timeout','Unavailable')){ throw $snap.Error }; $records=@(); if($null -ne $snap.Result){ foreach($s in @($snap.Result.sensors)){ if($null -ne $s.value -and $null -ne $s.unit){ $records += New-SensorRecord $s.hardwareName $s.name $s.value $s.unit 'LibreHardwareMonitor' (Get-Date) 'Valid' 'Validated' } } }; Get-SensorSnapshot -ProviderName 'LibreHardwareMonitor' { $records } }; ModuleErrors=$errors; SensorErrors=$sensorErrors }
    $report.OverallStatus = Get-OverallStatus ([pscustomobject]$report); return [pscustomobject]$report
}

function Write-ReportFile { param($Report, [string]$Directory, [string]$Format='TXT', [switch]$RedactSensitive)
    if (!(Test-Path -LiteralPath $Directory)) { New-Item -ItemType Directory -Path $Directory -Force | Out-Null }
    $safeHost = if ($Report.Computer.Hostname -and $Report.Computer.Hostname -ne 'N/A') { $Report.Computer.Hostname } else { 'computer' }
    $extension = $Format.ToLowerInvariant(); $base = Join-Path $Directory ("CK-Hardware-Scanner_{0}_{1}" -f $safeHost, (Get-Date -Format 'yyyyMMdd-HHmmss')); $path = "$base.$extension"; $index=1
    while (Test-Path -LiteralPath $path) { $path = "{0}_{1}.{2}" -f $base,$index,$extension; $index++ }
    $content = switch ($Format) { 'HTML' { Convert-ReportToHtml $Report -RedactSensitive:$RedactSensitive } 'JSON' { Convert-ReportToJson $Report -RedactSensitive:$RedactSensitive } default { Convert-ReportToText $Report -RedactSensitive:$RedactSensitive } }
    [IO.File]::WriteAllText($path, $content, (New-Object Text.UTF8Encoding($false))); return $path
}

if ($MyInvocation.InvocationName -ne '.') {
    $report = New-ScanReport; if ($OutputFormat -eq 'TXT') { Write-Output (Convert-ReportToText $report -RedactSensitive:$RedactSensitive) }
    if ($SaveReport) { $path = Write-ReportFile $report $OutputDirectory $OutputFormat -RedactSensitive:$RedactSensitive; Write-Output "`nReport saved: $path" }
}
