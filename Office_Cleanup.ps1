
# Run cleanup scripts after AD migration for clients with existing settings from the old tenant.

# Source: https://docs.microsoft.com/en-us/office/troubleshoot/activation/reset-office-365-proplus-activation-state

# Runs the following scripts: OLicenseCleanup.vbs, SignOutOfWamAccounts.ps1, and WPJCleanUp.cmd.

# Requires elevated permissions to run this script. Please restart the computer once the scripts are finished running.


function Check-IsElevated
{
   $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
   $p = New-Object System.Security.Principal.WindowsPrincipal($id)
   if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    return $true
   } else {
    return $false
   }   
}

$result = Check-IsElevated

$olicense = "https://download.microsoft.com/download/e/1/b/e1bbdc16-fad4-4aa2-a309-2ba3cae8d424/OLicenseCleanup.zip"
$wamcleanup = "https://download.microsoft.com/download/f/8/7/f8745d3b-49ad-4eac-b49a-2fa60b929e7d/signoutofwamaccounts.zip"
$wpjclean = "https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip"
$olicensePath = "$env:LOCALAPPDATA\OfficeLicenseCleanup\OLicenseCleanup.zip"
$wamcleanupPath = "$env:LOCALAPPDATA\OfficeLicenseCleanup\signoutofwamaccounts.zip"
$wpjcleanupPath = "$env:LOCALAPPDATA\OfficeLicenseCleanup\WPJCleanUp.zip"

if ($result) {

    $runScript = "$env:LOCALAPPDATA\OfficeLicenseCleanup\OLicenseCleanup.VBS"
    $wamScript = "$env:LOCALAPPDATA\OfficeLicenseCleanup\signoutofwamaccounts.ps1"
    $wpjScript = "$env:LOCALAPPDATA\OfficeLicenseCleanup\WPJCleanUp\WPJCleanUp\WPJCleanUp.cmd"
    
    if ((Test-Path $runScript) -and (Test-Path $wamScript) -and (Test-Path $wpjScript)) {
        write-host "`nRunning Office License Cleanup script..`n"
        cscript $runScript
        write-host "`nRunning Workplace Join Cleanup script..`n"
        cmd /C $wpjScript
        write-host "`nPlease restart your machine.`n" -ForegroundColor Green
        pause
    } Else {

        New-Item "$env:LOCALAPPDATA\OfficeLicenseCleanup" -ItemType "directory" -ea 0 1>$null


        Invoke-WebRequest -Uri $olicense -OutFile $olicensePath
        Invoke-WebRequest -Uri $wamcleanup -OutFile $wamcleanupPath
        Invoke-WebRequest -Uri $wpjclean -OutFile $wpjcleanupPath

        $one = Test-Path $olicensePath
        $two = Test-Path $wamcleanupPath
        $three = Test-Path $wpjcleanupPath

        if ($one -and $two -and $three) {
            Expand-Archive $olicensePath -DestinationPath "$env:LOCALAPPDATA\OfficeLicenseCleanup\"
            Expand-Archive $wamcleanupPath -DestinationPath "$env:LOCALAPPDATA\OfficeLicenseCleanup\"
            Expand-Archive $wpjcleanupPath -DestinationPath "$env:LOCALAPPDATA\OfficeLicenseCleanup\"
            write-host "`nRunning Office License Cleanup script..`n"
            cscript $runScript
            write-host "`nRunning Workplace Join Cleanup script..`n"
            cmd /C $wpjScript
            write-host "`nPlease restart your machine.`n" -ForegroundColor Green
            pause
        } Else {
            write-host "`nERROR: Cannot find downloaded files.`n" -ForegroundColor Red
        }

    }
    
} Else {
    write-Error "ERROR: Please run it as an administrator."
}
