$driveEject = New-Object -comObject Shell.Application
$driveEject.Namespace(17).Items() |ForEach {If ($_.Name -Match "B:"){$_.InvokeVerb("Eject")}}