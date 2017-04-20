# not using this as part of solution, was just some testing i was doing...

Get-Date
wmic product where "name like 'Microsoft Office 365 ProPlus%'"  call uninstall /nointeractive
Get-Date
wmic product where "name like 'Microsoft Project%'"  call uninstall /nointeractive
Get-Date
wmic product where "name like 'Microsoft Visio%'"  call uninstall /nointeractive
Get-Date
wmic product where "name like 'Microsoft OneDrive%'"  call uninstall /nointeractive
Get-Date
wmic product where "name like 'Office 16 Click-to-Run%'" call uninstall /nointeractive
Get-Date

wmic product | Out-File c:\gregtest2.txt