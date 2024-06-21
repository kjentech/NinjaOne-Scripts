# This script writes to three different custom fields.
# It rewrites bitlockerEnabled with current value of C:
# It rewrites bitlockerFullDiskEncryptionEnabled with current value if Bitlocker is enabled
# It rewrites bitlockerRecoveryPassphrase with current value if Bitlocker is enabled

$bitlockerOnOff = if ((Get-BitLockerVolume -MountPoint C:).ProtectionStatus -eq "On") {$true} else {$false}
Ninja-Property-Set bitlockerEnabled $bitlockerOnOff

if ($bitlockerOnOff -eq $true) {

	# Check for full or used-only
	$bitlockerFullModeEnabled = [bool] -not (manage-bde -status $ENV:SystemDrive | Select-String "Used Space Only Encrypted").Count
	Ninja-Property-Set bitlockerFullDiskEncryptionEnabled $bitlockerFullModeEnabled

	# Set recovery passphrase
	$RecoveryKey = (Get-BitLockerVolume -MountPoint C).KeyProtector | Select -ExpandProperty RecoveryPassword  
	Ninja-Property-Set bitlockerRecoveryPassphrase $RecoveryKey
}
