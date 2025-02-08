[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Naam van de shortcut
$shortcutName = "BedrijfsPortal"

# URL naar shortcut
$iconUrl = "https://raw.githubusercontent.com/DBNL88/Novoferm-Public/refs/heads/main/Scripts/CompanyPortalApp.ico"

# Pad naar locatie waar icoon wordt opgeslagen
$localIconDir = "$env:APPDATA\IntuneCustom"
$localIconPath = "$localIconDir\CompanyPortalApp.ico"

# Desktop map voor alle gebruikers
$DesktopPath = "$env:public\Desktop"

# Pad naar de shortcut
$shortcutPath = "$DesktopPath\$shortcutName.lnk"

try {
    # Controleer of de Desktop map bestaat
    if (!(Test-Path -Path $DesktopPath)) {
        Exit 1
    }

    # Controleer of de shortcut al bestaat
    if (Test-Path -Path $shortcutPath) {
        Exit 0
    }

    # Controleer of de doelmap bestaat, anders aanmaken
    if (!(Test-Path -Path $localIconDir)) {
        New-Item -ItemType Directory -Path $localIconDir -Force | Out-Null
    }

    # Download het icoon
    Invoke-WebRequest -Uri $iconUrl -OutFile $localIconPath -ErrorAction Stop

    # Maak de WScript Shell object aan
    $wshShell = New-Object -ComObject WScript.Shell

    # Maak de shortcut
    $shortcut = $wshShell.CreateShortcut($shortcutPath)

    # Stel het doelpad en icoonlocatie in
    $shortcut.TargetPath = "shell:AppsFolder\Microsoft.CompanyPortal_8wekyb3d8bbwe!App"
    $shortcut.IconLocation = "$localIconPath,0"

    # Sla de shortcut op
    $shortcut.Save()

} catch {
    Exit 1
}

Exit 0
