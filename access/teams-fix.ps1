$Host.UI.RawUI.WindowTitle = "Microsoft Teams Fix"
param (
    [switch]$noPrompt = $false
)

function cleanCache {
    Write-Host "`nTerminating Teams processes..."
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Write-Host "`nRemoving cache files..." ; Start-Sleep -s 2
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\Cache\*" -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\Application Cache\Cache\*" -ErrorAction SilentlyContinue #>>> may be non-existent#>
    Write-Host "`nTeams cache files removed."  ; Start-Sleep -s 1
}

function cleanRoaming {
    Write-Host "`nOutlook needs to be terminated for this fix to be applied. If it's open, save your work and proceed." -ForegroundColor Red -BackgroundColor Black
    if (-not $noPrompt) {
        Write-Host -NoNewLine "`nPress any key to continue...`n";
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    }
    Write-Host "`nTerminating processes of Teams and Outlook..."  ; Start-Sleep -s 2
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Stop-Process -ProcessName outlook -Force -ErrorAction SilentlyContinue
    Write-Host "`nRemoving files from %appdata%"  ; Start-Sleep -s 2
    cleanCache
    Remove-Item -Recurse -Force "$ENV:Userprofile\appdata\roaming\Microsoft\Teams\*" -ErrorAction SilentlyContinue
    Write-Host "`nFiles from Teams in %appdata% removed." ; Start-Sleep -s 1   
}

function reinstall {
    Write-Host "`nOutlook needs to be terminated for this fix to be applied. If it's open, save your work and proceed." -ForegroundColor Red -BackgroundColor Black
    if (-not $noPrompt) {
        Write-Host -NoNewLine "`nPress any key to continue...`n";
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    }
    Write-Host "`nTerminating processes of Teams and Outlook..."
    Stop-Process -ProcessName teams -Force -ErrorAction SilentlyContinue
    Stop-Process -ProcessName outlook -Force -ErrorAction SilentlyContinue
    cleanCache
    cleanRoaming
    Write-Host "`nRemoving remaining files of Teams..."
    Remove-Item -Recurse -Force "$Env:userprofile\appdata\local\Microsoft\Teams\*"
    Remove-Item -Recurse -Force "$Env:userprofile\appdata\roaming\Microsoft\Teams\*"
    Remove-Item -Recurse -Force "$Env:userprofile\OneDrive - Petrobras\Desktop\Microsoft Teams.lnk"
    Remove-Item -Recurse -Force "$Env:userprofile\desktop\Microsoft Teams.lnk"
    Remove-Item -Recurse -Force "C:\ProgramData\Microsoft\Microsoft\*"
    Write-Host "`nTeams removed."; Start-Sleep -s 2
    Write-Host "`nRemoving registry entries..."
    Remove-ItemProperty -Force "HKCU:\Software\IM Providers\Teams\" -Name *
    Remove-ItemProperty -Force "HKCU:\Software\Microsoft\Office\Teams\" -Name *
    Remove-ItemProperty -Force "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Teams\" -Name *
    Remove-ItemProperty -Force "HKLM:\Software\IM Providers\Teams\" -Name *
    Remove-ItemProperty -Force "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run\" -Name "com.squirrel.Teams.Teams"
    Remove-ItemProperty -Force "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\VREGISTRY_B683C874-A67C-41B4-8750-72BE2153F84C\MACHINE\Software\Wow6432Node\IM Providers\Teams\" -Name * <#>>> may be non-existent#>
    Remove-ItemProperty -Force "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Wow6432Node\IM Providers\Teams\" -Name * <#>>> may be non-existent#>
    Write-Host "`nTeams registry entries removed."
    #>>> Download and installation of Teams
    $url = "https://go.microsoft.com/fwlink/?linkid=2187217"
    # Determine the path for downloading the Teams installer
    $downloadsPath = if (Test-Path "$Env:USERPROFILE\Downloads") {
        "$Env:USERPROFILE\Downloads\Teams_x64.exe"
    } else {
        "$Env:USERPROFILE\Teams_x64.exe"
    }

    Write-Host "`nDownloading Teams. This process may take longer than usual. Please wait..."

    # Perform the download
    Invoke-WebRequest -Uri $url -OutFile $downloadsPath
    # Continue with the rest of your script for installation
    Write-Host "`nStarting installation. Please wait..."
    Invoke-Expression $downloadsPath

    # sleep
    Start-Sleep -s 30

    # Optionally, delete the installer after installation
    Remove-Item -Path $downloadsPath -Force
    Write-Host "`nInstaller deleted after installation."

    
}

function selec{

    param (
    [string]$Title = 'Menu'
    )
    
    Write-Host "`n============================ Repair Menu ============================`n"
    
    Write-Host "	[1] to remove Teams cache files"
    Write-Host "	[2] to remove files from Roaming (%appdata%)"
    Write-Host "	[3] to remove all remnants of Microsoft Teams and reinstall it"
    Write-Host "	[q] to exit the script"
    
    Write-Host "`n============================================================================"
    
     $selection = Read-Host "`nSelect one of the options above"
     switch ($selection)
     {
       '1' {cleanCache
       return selec} 
       
       '2' {cleanRoaming
       return selec} 

       '3' {reinstall
        return selec}

       'q' {
           Write-Output "`nExiting..."
           Start-Sleep -s 1
           exit }
  
       default {
            if ($selection -ige 3 -or $selection -ne 'q'){
                 Write-Host "`n>>> Only select options that are on the menu!`n" -ForegroundColor Red -BackgroundColor Black
                 Start-Sleep -s 2
                 return selec }
                }   
       }
  }

  selec
