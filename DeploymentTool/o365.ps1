<#
###
The purpose of this PowerShell script is to enable the installation of Office 365 ProPlus components onto any computer 
that can support it, regardless of the existence of any other Office components.

The script attempts to uninstall existing software, then install the matching Click-to-Run components as requested.

The batch files that call this file can specify to install Office, Project, Visio or any combination of the three.

###
Copyright (c) 2017 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>
#
# Function to help gather existing Office components information...   
function Get-ExistingOfficeInfo() {
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$component
    )    
    $officeExists = $False
    $officeArchitecture = 'none'
    #$osArchitecture = [environment]::Is64BitOperatingSystem / wasn't working on some Windows 7 machines
    #If ($osArchitecture -eq $True) { $osArchitecture = '64' } Else { $osArchitecture = '32' }
    If (Test-Path "C:\Program Files (x86)") { $osArchitecture = '64'} Else { $osArchitecture = '32' }    
    $officeVersion = 'none'

    If ($component -eq 'office') {
        $exeName = 'winword.exe'
    } ElseIf ($component -eq 'project') {
        $exeName = 'winproj.exe'
    } ElseIf ($component -eq 'visio') {
        $exeName = 'visio.exe'
    }
    # detect versions and architecture for determining O365 build to use...
    If (Test-Path "C:\Program Files\Microsoft Office\Office14\$exeName") { $officeArchitecture = '64'; $officeVersion = '2010'; $officeExists = $True; }
    If (Test-Path "C:\Program Files\Microsoft Office\Office15\$exeName") { $officeArchitecture = '64'; $officeVersion = '2013'; $officeExists = $True; }
    If (Test-Path "C:\Program Files\Microsoft Office\Office16\$exeName") { $officeArchitecture = '64'; $officeVersion = '2016'; $officeExists = $True; }
    If (Test-Path "C:\Program Files\Microsoft Office\root\Office15\$exeName") { $officeArchitecture = '64'; $officeVersion = '2013C2R'; $officeExists = $True; }    
    If (Test-Path "C:\Program Files\Microsoft Office\root\Office16\$exeName") { $officeArchitecture = '64'; $officeVersion = '2016C2R'; $officeExists = $True; }    
    If (Test-Path "C:\Program Files (x86)\Microsoft Office\Office14\$exeName") { $officeArchitecture = '32'; $officeVersion = '2010'; $officeExists = $True; }
    If (Test-Path "C:\Program Files (x86)\Microsoft Office\Office15\$exeName") { $officeArchitecture = '32'; $officeVersion = '2013'; $officeExists = $True; }
    If (Test-Path "C:\Program Files (x86)\Microsoft Office\Office16\$exeName") { $officeArchitecture = '32'; $officeVersion = '2016'; $officeExists = $True; }            
    If (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office15\$exeName") { $officeArchitecture = '32'; $officeVersion = '2013C2R'; $officeExists = $True; }
    If (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\$exeName") { $officeArchitecture = '32'; $officeVersion = '2016C2R'; $officeExists = $True; }
    

    $object1 = New-Object PSObject -Property @{
        officeExists=$officeExists;
        officeArchitecture=$officeArchitecture;
        officeVersion=$officeVersion;
        osArchitecture=$osArchitecture;        
    }
    Return $object1
}
# Some variables that will be used later on...
$o365Action = $args[0]
$o365Scope = $args[1]
$o365SharedFolder = '\\server\share\Office365ProPlus'
$o365EmailFrom = 'serviceAccount@domain.com'
$o365EmailTo = 'you@domain.com'
$o365EmailServer = 'smtpServer.domain.local'


If ($o365Action -eq 'configure') {

    $startTime = Get-Date
    $computerName = Get-Content env:computername
    $thisEmailBody = "<body><div style='font-family: calibri;'>Hello there, "
    $thisEmailBody += "<ul><li>The request is to $o365Action the following components: $o365Scope</li>"
    $thisEmailBody += "<li>The process is starting at $startTime</li>"

    # Detect versions and architecture for determining O365 build to use...
    $getOfficeInfo = Get-ExistingOfficeInfo -component 'office'
    $getProjectInfo = Get-ExistingOfficeInfo -component 'project'
    $getVisioInfo = Get-ExistingOfficeInfo -component 'visio'
    $osArchitecture = $getOfficeInfo.osArchitecture
    $officeArchitecture = $getOfficeInfo.officeArchitecture
    $officeVersion = $getOfficeInfo.officeVersion
    $projectVersion = $getProjectInfo.officeVersion
    $visioVersion = $getVisioInfo.officeVersion 

    If ($officeArchitecture -eq 'none') { 
        $officeExisted = 'no'
        $officeArchitecture = '32' 
        $thisEmailBody += "<li>The computer $computerName has a $osArchitecture-bit OS architecture, but does not have Office installed currently.</li>"
    } Else {
        $officeExisted = 'yes'
        $thisEmailBody += "<li>The computer $computerName has a $osArchitecture-bit OS architecture, and an existing office architecture of Office $officeVersion, $officeArchitecture-bit</li>"
        If ($projectVersion -ne 'none') {
            $thisEmailBody += "<li>The computer $computerName has an existing Project install: $projectVersion, $officeArchitecture-bit</li>"
        }
        If ($visioVersion -ne 'none') {
            $thisEmailBody += "<li>The computer $computerName has an existing Visio install: $visioVersion, $officeArchitecture-bit</li>"
        }
    }


    # Uninstall these existing applications (if any)...

    # Had issues with C2R removal using OffScrub, so will handle natively first...
    If (( $officeVersion -Like '*C2R*') -Or ( $projectVersion -Like '*C2R*') -Or ( $visioVersion -Like '*C2R*')) {
        #C:\windows\system32\cmd /c "$o365SharedFolder\DeploymentTool\RemoveAll.bat" | Out-Null
        #$thisEmailBody += "<li>The computer $computerName has an existing Click-to-Run Office component, removing that first.</li>"
        $thisExe = "$o365SharedFolder\DeploymentTool\setup.exe"
        $thisArgument = "/$o365Action $o365SharedFolder\DeploymentTool\ConfigurationRemoveAll.xml"
        Start-Process "$thisExe" "$thisArgument" -Wait -NoNewWindow | Out-Null
        $rightNow = Get-Date
        $thisEmailBody += "<li>The removal of existing Office ProPlus components finished at $rightNow</li>"
    }

    # Import and run Microsoft's OffScrub script to remove any Office components...
    . "$o365SharedFolder\OffScrub\Remove-PreviousOfficeInstalls.ps1"
    Try {
        # First copy VBS files over to local workstation, some Windows 7 computers seem to need this...
        Get-ChildItem -Path "$o365SharedFolder\OffScrub\*.vbs" | Copy-Item -Destination 'c:\windows\system32'
        # Then run the removal commands...
        Remove-PreviousOfficeInstalls -RemoveClickToRunVersions $True -Remove2016Installs $True -NoReboot $True -Force $True
    } Catch {
        $getUninstallError = $error[0].Exception
        If ($getUninstallError -like '*You cannot call a method on a null-valued expression*') {
            $thisEmailBody += "<li>Remove-PreviousOfficeInstalls had an issue, although $computerName did not have any existing instances of Office identified.</li>"              
        } Else {
            $thisEmailBody += "<li>Remove-PreviousOfficeInstalls had an issue for the computer named $computerName. The error was: $getUninstallError</li>"              
        }
    }

    # If the removal process didn't encounter any errors, continue...
    If (!($getUninstallError)) {

        # Verify any previous builds are no longer found...
        $getOfficeInfo2 = Get-ExistingOfficeInfo -component 'office'
        If (($getOfficeInfo2.officeExists -eq $True) -And ( $getOfficeInfo2.officeVersion -NotLike '*C2R*')) {
            $thisEmailBody += "<li>The computer $computerName had an issue removing components via OffScrub.</li>"
        } Else {  
            If ($officeExisted -ne 'no') {     
                $thisEmailBody += "<li>The computer $computerName has completed the Office uninstallation via OffScrub.</li>"
            }
            # Run the new Click to Run installation file...
            $thisTime = Get-Date; $thisEmailBody += "<li>Installation will start now: $thisTime</li>"
            If (($officeArchitecture -eq '32') -Or ($osArchitecture -eq '32')) { $useThisArchitecture = '32' } Else { $useThisArchitecture = '64' }
            
            # Figure out which components to install...
            If (($o365Scope -Like '*office*') -And ($o365Scope -NotLike '*project*') -And ($o365Scope -NotLike '*visio*')) { $thisScope = 'OfficeOnly' }
            ElseIf (($o365Scope -Like '*office*') -And ($o365Scope -Like '*project*') -And ($o365Scope -Like '*visio*')) { $thisScope = 'OfficeAll' }
            ElseIf (($o365Scope -Like '*office*') -And ($o365Scope -Like '*project*') -And ($o365Scope -NotLike '*visio*')) { $thisScope = 'OfficeProject' }
            ElseIf (($o365Scope -Like '*office*') -And ($o365Scope -NotLike '*project*') -And ($o365Scope -Like '*visio*')) { $thisScope = 'OfficeVisio' }
            ElseIf (($o365Scope -NotLike '*office*') -And ($o365Scope -Like '*project*') -And ($o365Scope -Like '*visio*')) { $thisScope = 'ProjectVisio' }
            ElseIf (($o365Scope -NotLike '*office*') -And ($o365Scope -Like '*project*') -And ($o365Scope -NotLike '*visio*')) { $thisScope = 'ProjectOnly' }
            ElseIf (($o365Scope -NotLike '*office*') -And ($o365Scope -NotLike '*project*') -And ($o365Scope -Like '*visio*')) { $thisScope = 'VisioOnly' }
            Else { $thisScope = 'OfficeAll' } 

            $thisExe = "$o365SharedFolder\DeploymentTool\setup.exe"
            $thisArgument = "/$o365Action $o365SharedFolder\DeploymentTool\configuration"
            $thisArgument += "$useThisArchitecture"
            $thisArgument += "$thisScope.xml"
            Start-Process "$thisExe" "$thisArgument" -Wait -NoNewWindow | Out-Null
            $endTime= Get-Date
            $thisEmailBody += "<li>The installation attempt of Office 365 ProPlus $useThisArchitecture-bit components completed at $endTime.</li>"
            $duration = ($endTime-$startTime).Minutes
            $thisEmailBody += "<li>The entire process took this long: $duration minutes.</li>"
        
            $getOfficeInfo3 = Get-ExistingOfficeInfo -component 'office'
            If ($getOfficeInfo3.officeVersion -eq '2016C2R') {
                $thisEmailBody += "<li>The computer $computerName has successfully had Office 365 ProPlus components installed.</li>"
            } Else {
                $thisEmailBody += "<li>The computer $computerName had an issue installing Office 365 ProPlus components.</li>"
            }
        }
    }
} ElseIf ($o365Action -eq 'download') {

    $startTime = Get-Date
    $thisEmailBody += "<li>The 32-bit Office ProPlus components are starting to be downloaded at: $startTime</li>"
    $thisExe = "$o365SharedFolder\DeploymentTool\setup.exe"
    $thisArgument = "/$o365Action $o365SharedFolder\DeploymentTool\Configuration32OfficeAll.xml"
    Start-Process "$thisExe" "$thisArgument" -Wait -NoNewWindow | Out-Null
    $rightNow = Get-Date
    $thisEmailBody += "<li>The 32-bit Office ProPlus components have completed downloading at: $rightNow</li>"

    $rightNow = Get-Date
    $thisEmailBody += "<li>The 64-bit Office ProPlus components are starting to be downloaded at: $rightNow</li>"
    $thisExe = "$o365SharedFolder\DeploymentTool\setup.exe"
    $thisArgument = "/$o365Action $o365SharedFolder\DeploymentTool\Configuration64OfficeAll.xml"
    Start-Process "$thisExe" "$thisArgument" -Wait -NoNewWindow | Out-Null
    $endTime = Get-Date
    $thisEmailBody += "<li>The 64-bit Office ProPlus components have completed downloading at: $endTime</li>"
    $duration = ($endTime-$startTime).Minutes
            $thisEmailBody += "<li>The entire process took this long: $duration minutes.</li>"
}

# Send results to email...
    $thisEmailBody += "</ul><p/> Thanks!</div></body>"  
    send-mailmessage `
    -to "$o365EmailTo" `
    -from "$o365EmailFrom" `
    -subject "Office 365 ProPlus Installation Results: $computerName" `
    -body $thisEmailBody `
    -smtpserver "$o365EmailServer" `
    -BodyAsHtml
