# Specify the display name for the shortcut
$shortcutName = "Company Portal"

# Specify the URL for the icon
$iconUrl = "https://deviceadvice.io/wp-content/uploads/2020/08/CompanyPortalApp.ico"

# Specify the local path to save the icon
$localIconPath = "$env:TEMP\CompanyPortalApp.ico"

# Download the icon
Invoke-WebRequest -Uri $iconUrl -OutFile $localIconPath

# Create a WScript Shell object
$wshShell = New-Object -ComObject WScript.Shell

# Create a shortcut object
$DesktopPath = [Environment]::GetFolderPath('Desktop')

$shortcut = $wshShell.CreateShortcut("$DesktopPath\$shortcutName.lnk")

# Set the target path for the shortcut
$shortcut.TargetPath = "shell:AppsFolder\Microsoft.CompanyPortal_8wekyb3d8bbwe!App"

# Set the icon location
$shortcut.IconLocation = "$localIconPath,0"

# Save the shortcut
$shortcut.Save()
