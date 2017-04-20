# not using this as part of solution, was just some testing i was doing...
# https://support.office.com/en-us/article/Manually-uninstall-Office-2016-or-Office-365-4e2904ea-25c8-4544-99ee-17696bb3027b

schtasks.exe /delete /tn "\Microsoft\Office\Office Automatic Updates"
schtasks.exe /delete /tn "\Microsoft\Office\Office Subscription Maintenance"
schtasks.exe /delete /tn "\Microsoft\Office\Office ClickToRun Service Monitor"
schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentLogOn2016"
schtasks.exe /delete /tn "\Microsoft\Office\OfficeTelemetryAgentFallBack2016"
Get-Process | Where {$_.Name -Like 'OfficeClickToRun'} | Stop-Process -Force
Get-Process | Where {$_.Name -Like 'OfficeC2RClient'} | Stop-Process -Force
Get-Process | Where {$_.Name -Like 'AppVShNotify'} | Stop-Process -Force
Get-Process | Where {$_.Name -Like 'Setup*.exe'} | Stop-Process -Force
sc delete ClickToRunSvc

Remove-Item 'C:\Program Files\Microsoft Office 15' -Recurse -Force
Remove-Item 'C:\Program Files\Microsoft Office 16' -Recurse -Force
Remove-Item 'C:\Program Files\Microsoft Office' -Recurse -Force
Remove-Item 'C:\Program Files (x86)\Microsoft Office 15' -Recurse -Force
Remove-Item 'C:\Program Files (x86)\Microsoft Office 16' -Recurse -Force
Remove-Item 'C:\Program Files (x86)\Microsoft Office' -Recurse -Force

Remove-Item 'C:\Program Files\Common Files\Microsoft Shared\ClickToRun' -Recurse -Force
Remove-Item 'C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun' -Recurse -Force
Remove-Item 'C:\ProgramData\Microsoft\ClickToRun' -Recurse -Force
Remove-Item 'C:\ProgramData\Microsoft\OFFICE\ClickToRUnPackageLocker' -Force


Remove-Item 'HKEY_LOCAL_MACHINE\software\Microsoft\Office\CLickToRun' -Recurse
Remove-Item 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppVISV' -Recurse
Remove-Item 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us' -Recurse


Remove-Item 'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office 2016 Tools' -Recurse -Force

MsiExec.exe /qn /X '{90160000-008F-0000-1000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-008C-0000-0000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-008C-0409-0000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-007E-0000-0000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-008C-0000-0000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-008C-0409-0000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-007E-0000-1000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-008C-0000-1000-0000000FF1CE}'

MsiExec.exe /qn /X '{90160000-008C-0409-1000-0000000FF1CE}'
