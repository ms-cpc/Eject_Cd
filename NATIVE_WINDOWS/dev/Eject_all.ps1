$cds = (New-Object -ComObject "WMPlayer.OCX").cdromCollection
for($i=0;$i -lt $cds.Count;$i++) { $cds.Item($i).Eject() }