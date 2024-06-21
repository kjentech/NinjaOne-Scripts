$Domain = (Get-CimInstance Win32_ComputerSystem).Domain
$DomainNetBIOS = ($Domain -split "\.")[0]
$ExcludeMembers = "$Domain\Domain Admins","$DomainNetBIOS\Domain Admins"
$LocalGroup = Get-LocalGroupMember -SID S-1-5-32-555 | where Name -notin $ExcludeMembers

$LocalGroupOutput = $LocalGroup.Name | Out-String

Ninja-Property-Set localGroupRdpUsers $LocalGroupOutput
