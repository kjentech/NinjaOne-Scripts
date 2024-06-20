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

$LocalUsers = Get-LocalUser | Where {$_.Enabled -eq $true -and $_.SID -ne "S-1-5-21-1677959656-505996033-2382117720-500"}
$PrivilegedGroups = "S-1-5-32-555","S-1-5-32-544"

$UsersAndGroups = foreach ($User in $LocalUsers) {
    $GroupMembership = @()
	$Privileged = $false

    if ($User.LastLogon) {
        $LastLogon = $User.LastLogon
    } else { $LastLogon = "Never" }

    Get-LocalGroup | foreach {
        if (Get-LocalGroupMember -Group $_ | where Name -match $User) {
            $GroupMembership += $_
            
	        if ($_.SID.Value -in $PrivilegedGroups) { $Privileged = $true }
        }
    }

    [PSCustomObject]@{
        ComputerName = $env:computername
        UserName = $User.Name
        LastLogon = $LastLogon
        Groups = $GroupMembership.Name
        Privileged = $Privileged
    }
}

$objectWithColor = $UsersAndGroups | foreach {
  if ($_.Privileged) {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name RowColour -Value "danger"
  } else {
    Add-Member -InputObject $_ -MemberType NoteProperty -Name RowColour -Value "warning"
  }
  $_
}

$LocalUsersOutput = ConvertTo-ObjectToHtmlTable -Objects $objectWithColor

Ninja-Property-Set localUsers $LocalUsersOutput
