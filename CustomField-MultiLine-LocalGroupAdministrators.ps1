$Domain = (Get-CimInstance Win32_ComputerSystem).Domain
$DomainNetBIOS = ($Domain -split "\.")[0]
$ExcludeMembers = "$Domain\Domain Admins","$DomainNetBIOS\Domain Admins","$env:computername\Administrator"
$LocalGroup = Get-LocalGroupMember -SID S-1-5-32-544 | where Name -notin $ExcludeMembers

$LocalGroupOutput = $LocalGroup.Name | Out-String
$LocalGroupOutput

Ninja-Property-Set localGroupAdministrators $LocalGroupOutput
