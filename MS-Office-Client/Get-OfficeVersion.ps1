<#
  Inherited an environment with a mish-mash of Office installs and zero documentation besides the keys?
  Me too!
  This will output the computer name, last 5 characters of the product key, and Office version to CSV
#>

function Get-OfficeVersion {
  $OutputFile = "OfficeVersions.csv" # Path to the file on your server

  $officePaths = @(
    @{Version = "2016"; Path32 = "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"; Path64 = "C:\Program Files\Microsoft Office\Office16\OSPP.VBS"},
    @{Version = "2013"; Path32 = "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS"; Path64 = "C:\Program Files\Microsoft Office\Office15\OSPP.VBS"},
    @{Version = "2010"; Path32 = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"; Path64 = "C:\Program Files\Microsoft Office\Office14\OSPP.VBS"}
  )

  $OSPP = $null
  foreach ($office in $officePaths) {
    if (Test-Path -Path $office.Path64) {
      $OSPP = $office.Path64
      break
    }
    if (Test-Path -Path $office.Path32) {
      $OSPP = $office.Path32
      break
    }
  }

  if (-not $OSPP) {
    Write-Warning "No supported Office installation found."
    return
  }

  $productInfo = cscript $OSPP /dstatus
  $productKey = ($productInfo | Select-String 'Last 5').Line.split(":")[1]
  $productEdition = ($productInfo | Select-String 'License Name').Line.split(",")[1]
  $output = $env:computername+','+$productKey+','+$productEdition
  $output | Out-file $OutputFile -Append -Encoding ASCII
}