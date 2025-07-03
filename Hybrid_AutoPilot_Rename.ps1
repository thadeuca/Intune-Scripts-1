##Hybrid AutoPilot renaming script
## By u/sryan2k1 @ https://www.reddit.com/r/Intune/comments/1lqt63l/hostname_in_hybrid_join_autopilot_environment/

## Write the new name to registry key to use as detection
function Detection {
  # Write registry key to indicate success for detection
  $path = 'HKLM:\SOFTWARE\COMPANYNAMEHERE\RenameToSerialNum'
  if (!(Test-Path -LiteralPath $path)) { [void] (New-Item -Path $path -Force) }
  Set-ItemProperty -Path $path -Name ConfigRev -Value 1 -Type DWord
}
 
## Function to search for computer name in AD without needing the AD powershell module
function Get-ExistingComputer ($name) {
  $Searcher = [ADSISearcher]"(&(objectClass=computer)(name=$($name)))"
  $Searcher.SearchRoot = [ADSI]''
  $Searcher.PropertiesToLoad.Add('Name')
  $SearchResult = $Searcher.FindOne()    #note: exceptions can get thrown here
  If ($SearchResult -And $SearchResult.Count -ne 0) {
    $SearchResult.Properties
  }
}
 
 
## Check for connectivity to domain
$dcInfo = [ADSI]"LDAP://your-domain-here.com"
If ($null -eq $dcInfo.Path) {
  Write-Error "No connectivity to AD"
  Exit 1
}
 
$oldName = $env:COMPUTERNAME
## Set new name to computer's serial number
$newName = Get-WmiObject Win32_bios | Select-Object -ExpandProperty SerialNumber
 
## Clean up serial, get rid of spaces, replace VMware- with VM, Max length 15
$newName = $newName.Replace(' ', '')
$newName = $newName.Replace('VMware-', 'VM')
if ($newName.Length -ge 15) {
  $newName = $newName.substring(0, 15)
}
 
## If the current hostname doesn't match the serial number, rename PC
IF ($newName -ne $oldName) {
  # See if the new name already exists in AD, if it does then exit 1. The old AD object needs to be removed. 
  IF (Get-ExistingComputer -name $newName) {
    write-Output "Computer $newName already exists!! Need to delete in AD"
    Exit 1
}
 
  Try {
    Rename-Computer -NewName $newName -erroraction stop
  } Catch {
    Write-Output "Error: $_"
    Exit 1
  }
  
 
  Write-Output "Computer renamed from $oldName to $newName"
 
  # Write detection reg key
  Detection
  # Require hard restart
  Exit 1641
} ELSE {
  ## The computer is already the correct name, exit 0 to indicate success
  # Write detection reg key
  Detection
  Exit 0
}
 
