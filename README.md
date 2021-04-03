# Powershell Tools
A collection of Powershell snippets used in the environments I work with, but may be useful to others.
### Windows 10
- Start-Windows10Upgrade.ps1 - this uses the Windows Update Assistant to provide a silent Windows 10 Feature update, with pop-ups warning the user that the process is happening.
### MS-Office
- Get-OfficeVersion.ps1: Inherited an environment with a mish-mash of Office installs and zero documentation besides the keys? Me too! This function will output the computer name, last 5 characters of the product key, and Office version to CSV
### OneDrive
- Get-OneDriveSyncCount.ps1: We found that if the number exceeded 35 or so, performance degraded. So we decided to monitor this with our RMM, and this function counts the number of OneDrive sync locations the logged in user has