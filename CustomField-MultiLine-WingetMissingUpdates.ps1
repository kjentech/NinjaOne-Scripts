# this function has been yanked from Stack Overflow and converts Winget output into objects
function ConvertFrom-FixedColumnTable {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline)] [string] $InputObject
  )
    
  begin {
    Set-StrictMode -Version 1
    $lineNdx = 0
  }
    
  process {
    $lines = 
    if ($InputObject.Contains("`n")) { $InputObject.TrimEnd("`r", "`n") -split '\r?\n' }
    else { $InputObject }
    foreach ($line in $lines) {
      ++$lineNdx
      if ($lineNdx -eq 1) { 
        # header line
        $headerLine = $line 
      }
      elseif ($lineNdx -eq 2) { 
        # separator line
        # Get the indices where the fields start.
        $fieldStartIndices = [regex]::Matches($headerLine, '\b\S').Index
        # Calculate the field lengths.
        $fieldLengths = foreach ($i in 1..($fieldStartIndices.Count - 1)) { 
          $fieldStartIndices[$i] - $fieldStartIndices[$i - 1] - 1
        }
        # Get the column names
        $colNames = foreach ($i in 0..($fieldStartIndices.Count - 1)) {
          if ($i -eq $fieldStartIndices.Count - 1) {
            $headerLine.Substring($fieldStartIndices[$i]).Trim()
          }
          else {
            $headerLine.Substring($fieldStartIndices[$i], $fieldLengths[$i]).Trim()
          }
        } 
      }
      else {
        # data line
        $oht = [ordered] @{} # ordered helper hashtable for object constructions.
        $i = 0
        foreach ($colName in $colNames) {
          $oht[$colName] = 
          if ($fieldStartIndices[$i] -lt $line.Length) {
            if ($fieldLengths[$i] -and $fieldStartIndices[$i] + $fieldLengths[$i] -le $line.Length) {
              $line.Substring($fieldStartIndices[$i], $fieldLengths[$i]).Trim()
            }
            else {
              $line.Substring($fieldStartIndices[$i]).Trim()
            }
          }
          ++$i
        }
        # Convert the helper hashable to an object and output it.
        [pscustomobject] $oht
      }
    }
  }
    
}

# this function has been yanked from Stack Overflow and removes any progress bar output before the headers
function winget_outclean () {
  [CmdletBinding()]
  param (
    [Parameter(ValueFromPipeline)]
    [String[]]$lines
  )
  if ($input.Count -gt 0) { $lines = $PSBoundParameters['Value'] = $input }
  $bInPreamble = $true
  foreach ($line in $lines) {
    if ($bInPreamble) {
      if ($line -like "Name*") {
        $bInPreamble = $false
      }
    }

    if (-not $bInPreamble -and $line -notmatch "upgrades available.") {
      Write-Output $line
    }
  }
}

#########################

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new() 

if ($env:USERNAME -match "$env:COMPUTERNAME") {
  $winget = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" | select -ExpandProperty Path -Last 1
}
else {
  $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
}

$wingetUpgrade = & $winget upgrade | winget_outclean
# $wingetList = (& $winget list | winget_outclean) -replace "mssto…","msstore"
$wingetObjects = $wingetUpgrade | ConvertFrom-FixedColumnTable | sort Id

$WingetListOutput = $wingetObjects.Id

Ninja-Property-Set wingetList $WingetListOutput
