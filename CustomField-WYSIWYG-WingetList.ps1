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
