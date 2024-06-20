# Note about this: We're using winget.exe because the PowerShell module Microsoft.Winget.Client is only supported on PowerShell 7 and up. The module actually works in 5.1 as a standard user, but throws an error in 5.1 as SYSTEM.
# The winget.exe execuable has some very distinct flaws: notably that ONLY output to the console result in Winget actually displays the full rows. Redirect, pipe or catch it in a variable, and it truncates a big part of the columns.
# Winget has a COM API, but I don't know how to use it.

function ConvertTo-ObjectToHtmlTable {
# Credit for this function goes to https://github.com/freezscholte/Public-Ninja-Scripts/blob/main/WYSIWYG-Styling/ConvertTo-ObjectToHtmlTable.ps1

  param (
    [Parameter(Mandatory = $true)]
    [System.Collections.Generic.List[Object]]$Objects
  )
  
  # Start the HTML table
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<table><thead><tr>')
  
  # Add column headers based on the properties of the first object, excluding "RowColour"
  $Objects[0].PSObject.Properties.Name | Where-Object { $_ -ne 'RowColour' } | ForEach-Object { [void]$sb.Append("<th>$_</th>") }
  
  # Add rows
  [void]$sb.Append('</tr></thead><tbody>')
  foreach ($obj in $Objects) {
    # Use the RowColour property from the object to set the class for the row
    $rowClass = if ($obj.RowColour) { $obj.RowColour } else { "" }
    
    [void]$sb.Append("<tr class=`"$rowClass`">")
    # Generate table cells, excluding "RowColour"
    foreach ($propName in $obj.PSObject.Properties.Name | Where-Object { $_ -ne 'RowColour' }) {
      [void]$sb.Append("<td>$($obj.$propName)</td>")
    }
    [void]$sb.Append('</tr>')
  }
  [void]$sb.Append('</tbody></table>')
  $OutputLength = $sb.ToString() | Measure-Object -Character -IgnoreWhiteSpace | Select-Object -ExpandProperty Characters
  if ($OutputLength -gt 200000) {
    Write-Warning ('Output appears to be over the NinjaOne WYSIWYG field limit of 200,000 characters. Actual length was: {0}' -f $OutputLength)
  }
  return $sb.ToString()
}

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

# $wingetUpgrade = & $winget upgrade | winget_outclean
$wingetList = (& $winget list | winget_outclean) -replace "mssto…","msstore"
$wingetObjects = $wingetList | ConvertFrom-FixedColumnTable | sort Id

$objectWithColor = $wingetObjects | foreach {
  if ($_.Available) {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name RowColour -Value "danger"
  } else {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name RowColour -Value "other"
  }
  $_
}

$WingetListOutput = ConvertTo-ObjectToHtmlTable -Objects $objectWithColor

Ninja-Property-Set wingetList $WingetListOutput
