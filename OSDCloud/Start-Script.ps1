#Set my working OSDCloud Template
Set-OSDCloudTemplate -Name ''

#Create my new OSDCloud Workspace
New-OSDCloudWorkspace -WorkspacePath D:\Odcloud

#Cleanup Languages
$KeepTheseDirs = @('boot','efi','nl-nl','sources','fonts','resources')
Get-ChildItem "$(Get-OSDCloudWorkspace)\Media" | Where {$_.PSIsContainer} | Where {$_.Name -notin $KeepTheseDirs} | Remove-Item -Recurse -Force
Get-ChildItem "$(Get-OSDCloudWorkspace)\Media\Boot" | Where {$_.PSIsContainer} | Where {$_.Name -notin $KeepTheseDirs} | Remove-Item -Recurse -Force
Get-ChildItem "$(Get-OSDCloudWorkspace)\Media\EFI\Microsoft\Boot" | Where {$_.PSIsContainer} | Where {$_.Name -notin $KeepTheseDirs} | Remove-Item -Recurse -Force

#Build WinPE to start OSDCloudGUI automatically
Edit-OSDCloudWinPE -UseDefaultWallpaper -StartOSDCloudGUI
