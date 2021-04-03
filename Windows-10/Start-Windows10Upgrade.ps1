# Set Variables:
$targetVersion = "10.0.19042"
$minimumSpace = 10          # Minimum drive space in GB
$downloadPath = "C:\IT"     # Folder for the download
$logPath = "C:\IT\logs"     # Folder for the Windows setup logs
$updateAssistantURL = "https://download.microsoft.com/download/2/b/b/2bba292a-21c3-42a6-8123-98265faff0b6/Windows10Upgrade9252.exe" # URL of the update assistant
$upgradeArguments = "/quietinstall /skipeula /auto upgrade /copylogs " + $logPath
$fileName = $updateAssistantURL.Substring($updateAssistantURL.LastIndexOf("/") + 1)

<#
  Prevent updates during working hours
  - if $workingHoursEnabled is set to true, if the script is run between the specified hours it will terminate
  - prevents accidental in-hours upgrades
#>
$workingHoursEnabled = $true 
$workingHoursStart = 0900
$workingHoursEnd = 1730

<#
  Show pop up to logged in user
  - if true, this will display a pop up to any logged in user
  - if timeout is set to 0, script will *need* user to click OK before script continues
  - so if you want this to be unattended, either set a time for the pop up greater than 0 seconds, or set $popupShow to $false
#>

$popupShow = $true
$popupMessage = "As agreed, a Windows Feature Update is scheduled for your device. This will take 2-4 hours depending on the speed of your machine and your internet. Please save all documents and close your work, leaving your PC turned on and connected to power"
$popupTitle = "Message from Contoso IT Department"
$popupTimeout = 20 

function Get-CurrentVersion {
  $OS = (Get-CimInstance Win32_OperatingSystem).caption
  $currentVersion = (Get-CimInstance Win32_OperatingSystem).version
  if (-Not ($OS -like "Microsoft Windows 10*"))
    {
      Write-Output "This machine does not have Windows 10 installed. Exiting."
      exit
    }
  if ($currentVersion -eq $targetVersion)
    {
      Write-Output "OS is up-to-date - no action required. Exiting"
      exit
    }
    else {
      Write-Output "OS is out of date, continuing"
    }
}

function Test-DiskSpace ($minimumSpace){
  $drive = Get-PSDrive C
  $freeSpace = $drive.free / 1GB
  if ($minimumSpace -gt $freeSpace)
      {
          Write-Output "Not enough space for the upgrade to continue. Exiting script."
          exit
      } else {
        Write-Output "There is enough free disk space. Continuing."
      }
  }


function Test-WorkingHours($enabled, $startTime, $endTime){
  if ($enabled -eq $true)
      {
        [int]$time = Get-Date -format HHMM
        if ($time -gt $startTime -And $time -lt $endTime)
        {
          Write-Output "Script has been executed within working hours. Exiting script."
          exit
        } else {
          Write-Output "Confirmed script has been run outside working hours. Continuing."
        }
      } else {
        Write-Output "Working hours flag disabled. Continuing."
      }
  }

function Get-UpdateAssistant($URL, $path, $log, $file){
# Download File
  If(!(test-path $path))
    {
      New-Item -ItemType Directory -Force -Path $path
      Write-Output "Created download path."
    }
  If(!(test-path $log))
  {
    New-Item -ItemType Directory -Force -Path $log
    Write-Output "Created log path."
  }
  Invoke-WebRequest -Uri $URL -OutFile $path\$file
  Write-Output "Downloaded Update Assistant"
}

function Show-Message($enabled, $title, $message, $time) {
  if ($enabled -eq $true) {
    $popUp = New-Object -ComObject Wscript.Shell
    $popUp.Popup($message,$time,$title,0x0)
    Write-Output "Message displayed. Continuing"
    } else {
      Write-Output "Pop up has been disabled. Continuing."
    }
  }

function Start-Upgrade($path, $file, $arguments) {
  Write-Output "Starting upgrade process..."
  Start-Process -FilePath $path\$file -ArgumentList $arguments
  Write-Output "Upgrade process has been started"
  Start-Sleep -s 120 # Pause a little, to make sure the process is started
}

Get-CurrentVersion
Test-DiskSpace $minimumSpace
Test-WorkingHours $workingHoursEnabled $workingHoursStart $workingHoursEnd
Get-UpdateAssistant $updateAssistantURL $downloadPath $logPath $fileName
Show-Message $popupShow $popupTitle $popupMessage $popupTimeout
Start-Upgrade $downloadPath $fileName $upgradeArguments