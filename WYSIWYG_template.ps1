# This template consists of a function that prints HTML tables with support for red/yellow/blue row coloring in NinjaOne

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

###################################

# code here
# if color is wanted, you need a property to filter on. I use $_.Defcon in this example
# $theThing = Get-Stuff

$objectWithColor = $theThing | foreach {
  if ($_.Defcon -eq 1) {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name RowColour -Value "danger"
  } else {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name RowColour -Value "warning"
    # danger, warning, other, blank
  } 
  $_
}

$theThingOutput = ConvertTo-ObjectToHtmlTable -Objects $objectWithColor

Ninja-Property-Set theThing $theThingOutput
