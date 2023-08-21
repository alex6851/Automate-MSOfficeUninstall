


$email = Read-Host "Input email address to receive an email notification when this is complete:"

while (( $Email -notmatch '@.*\..*')) {
	$Email = Read-Host 'Im sorry I didnt understand....What email address do you want to send notifications to?'
}

$Continue = Read-host "This will CLOSE all of your Office applications. Make sure you have saved all your work. Are you ready to continue? Y/N"

if ($Continue -imatch "^Y$") {
    $OfficeServices = @(
        "lync",
        "winword",
        "excel",
        "msaccess",
        "mstore",
        "infopath",
        "setlang",
        "msouc",
        "ois",
        "onenote",
        "outlook",
        "powerpnt",
        "mspub",
        "groove",
        "visio",
        "winproj",
        "graph",
        "teams"
    )

    Write-Host "Checking to make sure all office applications are closed" -ForegroundColor Yellow




    [System.Collections.ArrayList]$Processes = Get-Process

    foreach ($officeService in $officeServices) {
        for ($i = 0; $i -lt $Processes.Count) {
            $Process = $Processes[$i]
            if($officeService -imatch "$($Process.Name)"){
                # Write-Host "Stopping Process $($officeService)"
                Stop-Process -Id $Process.Id -Force -Confirm:$false -ErrorAction SilentlyContinue
                $Processes.RemoveAt($i)
            }
            else {
                $i++
            }
        }
    }

    Write-Warning "After the uninstall a reboot will be suggested. You could SKIP the reboot and install Office365."
    $SkipReboot = Read-host "Do you want to to skip the reboot? Y/N"


    if (!(Test-Path C:\tools\)) {
        mkdir C:\tools\
    }

    Invoke-WebRequest -Uri "https://aka.ms/SaRA_EnterpriseVersionFiles" -OutFile C:\tools\Saracmd.zip


    set-location C:\tools\

    $items = Get-ChildItem -File

    $ZipFile = $items.Where({ $_.Name -imatch "saracmd.*\.zip" })


    if (!(Test-Path "C:\tools\$($Zipfile.BaseName)")) {
        Expand-Archive $ZipFile.Name -Force
    }
    set-Location $ZipFile.BaseName


    Write-Host "Uninstalling office you will receive an email when this is complete." -ForegroundColor Green

    .\SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All

    if($SkipReboot -imatch "^Y$"){
        \\PathToOfficeInstaller\setup.exe /configure \\PathtoOfficeXMLFile\configuration.xml
         Send-MailMessage -Subject "Office re-installation was finished on $($env:COMPUTERNAME)" -To $email -From "Office-AutomatedReinstall@mrcy.com" -SmtpServer "SMTPserverName"
    }
    $files = Get-childitem -Filter "saracmd*"

    foreach ($file in $files) {
        try {
            Remove-Item $file.FullName -Force
        }
        catch {
            if ($_.Exception.Message -match ".*The process cannot access the file.*because it is being used by another process.*") {
                stop-process -name explorer –force
				Start-Sleep -Seconds 10
				Remove-Item $file.FullName -Force
            }
        }
    }

    if($SkipReboot -imatch "^N$"){
        Send-MailMessage -Subject "Office un-installation was finished on $($env:COMPUTERNAME)" -BodyAsHtml "<h3>Waiting for Confirmation before computer will restart.</h3>" -To $email -From "Office-AutomatedUninstall@mrcy.com" -SmtpServer "mail.mrcy.com"
        $confirm = Read-Host "Are you ready to reboot?"
        if($confirm -imatch "^Y$"){
            shutdown /r
        }
    }
     
}