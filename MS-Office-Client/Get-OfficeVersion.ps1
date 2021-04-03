<#
  Inherited an environment with a mish-mash of Office installs and zero documentation besides the keys?
  Me too!
  This will output the computer name, last 5 characters of the product key, and Office version to CSV
#>

function Get-OfficeVersion {
  $OutputFile = "OfficeVersions.csv" # Path to the file on your server

  $2010Path32Bit = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"
  $2010Path64Bit = "C:\Program Files\Microsoft Office\Office14\OSPP.VBS"
  $2013Path32Bit = "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS"
  $2013Path64Bit = "C:\Program Files\Microsoft Office\Office15\OSPP.VBS"
  $2016Path32Bit = "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS"
  $2016Path64Bit = "C:\Program Files\Microsoft Office\Office16\OSPP.VBS"

  if(Test-Path -Path $2010Path32Bit) {
      $OSPP = $2010Path32Bit
      }
  if(Test-Path -Path $2010Path64Bit) {
      $OSPP = $2010Path64Bit
      }
  if(Test-Path -Path $2013Path32Bit) {
      $OSPP = $2013Path32Bit
      }
  if(Test-Path -Path $2013Path64Bit) {
      $OSPP = $2013Path64Bit
      }
  if(Test-Path -Path $2016Path32Bit) {
      $OSPP = $2016Path32Bit
      }
  if(Test-Path -Path $2016Path64Bit) {
      $OSPP = $2016Path64Bit
      }

  $productInfo = cscript $OSPP /dstatus
  $productKey = ($productInfo | Select-String 'Last 5').Line.split(":")[1]
  $productEdition = ($productInfo | Select-String 'License Name').Line.split(",")[1]
  $output = $env:computername+','+$productKey+','+$productEdition
  $output | Out-file $OutputFile -Append -Encoding ASCII
}