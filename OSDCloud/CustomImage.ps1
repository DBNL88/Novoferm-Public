
#Create a new OSDCloud Workspace
New-OSDCloudWorkspace -WorkspacePath C:\OSDCloud\CustomImage

#Cleanup OSDCloud Workspace Media
$KeepTheseDirs = @('boot','efi','en-us','nl-nl','sources','fonts','resources')
Get-ChildItem C:\OSDCloud\CustomImage\Media | Where {$_.PSIsContainer} | Where {$_.Name -notin $KeepTheseDirs} | Remove-Item -Recurse -Force
Get-ChildItem C:\OSDCloud\CustomImage\Media\Boot | Where {$_.PSIsContainer} | Where {$_.Name -notin $KeepTheseDirs} | Remove-Item -Recurse -Force
Get-ChildItem C:\OSDCloud\CustomImage\Media\EFI\Microsoft\Boot | Where {$_.PSIsContainer} | Where {$_.Name -notin $KeepTheseDirs} | Remove-Item -Recurse -Force

#Edit WinPE and rebuild ISO
Edit-OSDCloudWinPE -UseDefaultWallpaper




$WindowsImage = "C:\OSDBuilder\OSBuilds\Windows 11 Pro x64 24H2 26100.3194 nl-NL\OS\sources\install.wim"
$Destination = "$(Get-OSDCloudWorkspace)\Media\OSDCloud\OS"
New-Item -Path $Destination -ItemType Directory -Force
Copy-Item -Path $WindowsImage -Destination "$Destination\CustomImage.wim" -Force
New-OSDCloudISO
