$sh = New-Object -ComObject "Shell.Application"
$sh.Namespace(17).Items() |
    Where-Object { $_.Type -eq "BD" } |
        foreach { $_.InvokeVerb("Eject") }