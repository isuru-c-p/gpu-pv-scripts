Param (
    [Parameter(Mandatory=$true)]
    [string] $VMName,
    [string] $GPUName = "AUTO",
    [string] $Hostname = $ENV:Computername
)

Import-Module $PSSCriptRoot\Add-VMGpuPartitionAdapterFiles.psm1

function GetDriveLetter {
    $UsedDriveLetters = @(Get-WmiObject -Class Win32_LogicalDisk | %{$([char]$_.DeviceID.Trim(':'))})
    $TempDriveLetters = @((Compare-Object -DifferenceObject $UsedDriveLetters -ReferenceObject $( 67..90 | % { "$([char]$_)" } ) ) | ? { $_.SideIndicator -eq '<=' } | % { $_.InputObject })
    $AvailableDriveLetter = ($TempDriveLetters | Sort-Object)
    $TempDriveLetters[0]
}

$VM = Get-VM -VMName $VMName
$VHD = Get-VHD -VMId $VM.VMId

If ($VM.state -eq "Running") {
    [bool]$state_was_running = $true
    }

if ($VM.state -ne "Off"){
    "Attemping to shutdown VM..."
    Stop-VM -Name $VMName -Force
    } 

While ($VM.State -ne "Off") {
    Start-Sleep -s 3
    "Waiting for VM to shutdown - make sure there are no unsaved documents..."
    }

"Mounting Drive..."
$Partitions = (Mount-VHD -Path $VHD.Path -PassThru | Get-Disk | Get-Partition)

# Filter to partition with "Windows" folder
$DriveLetter = $Partitions | ForEach-Object {
    $driveLetter = $_.DriveLetter
    if (!$driveLetter) {
        $driveLetter = GetDriveLetter
        Write-Debug "Assigning $driveLetter..."
        Set-Partition -InputObject $_ -NewDriveLetter $driveLetter
    }
    if (Test-Path "${driveLetter}:\windows") {
        $driveLetter
    } else {
        $null
    }
} | Where-Object { $_ }

"Drive letter: $DriveLetter"

"Copying GPU Files - this could take a while..."
Add-VMGPUPartitionAdapterFiles -hostname $Hostname -DriveLetter $DriveLetter -GPUName $GPUName

"Dismounting Drive..."
Dismount-VHD -Path $VHD.Path

If ($state_was_running){
    "Previous State was running so starting VM..."
    Start-VM $VMName
    }

"Done..."