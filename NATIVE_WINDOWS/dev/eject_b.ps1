$driveEject = New-Object -comObject Shell.Application
$driveEject.Namespace(17).ParseName("b:").InvokeVerb("Eject")