[CmdletBinding()] 
param(
[switch]$Eject,
[switch]$Close 
) 
try {
	$Diskmaster = New-Object -ComObject IMAPI2.MsftDiscMaster2
	$DiskRecorder = New-Object -ComObject IMAPI2.MsftDiscRecorder2
	$DiskRecorder.InitializeDiscRecorder($DiskMaster)
	if ($Eject) {
		$DiskRecorder.EjectMedia()
	} elseif($Close) {
		$DiskRecorder.CloseTray()
}

} catch {
    Write-Error "Failed to operate the disk. Details : $_"
} 