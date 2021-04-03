<#
  This function counts the number of OneDrive sync locations the current user has
  We found that if the number exceeded 35 or so, performance degraded. So we decided to monitor this with our RMM agent
#>

function Get-OneDriveSyncCount {
  $locations = Get-ItemProperty -Path HKCU:\Software\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache
  $count = 0
  $locations.PSObject.PRoperties | ForEach-Object {
    if ($_ -like '*Users*') {
      $count++
    }
  }
  return
}