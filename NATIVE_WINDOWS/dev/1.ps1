$sh = New-Object -ComObject "Shell.Application"
$sh.Namespace(17).Items() |
    Where-Object { $_.name -eq 'A:' } |
        foreach { $_.InvokeVerb("Eject") }