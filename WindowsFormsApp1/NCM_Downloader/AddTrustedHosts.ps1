param([switch]$Elevated)

function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) 
    {
        # tried to elevate, did not work, aborting
    } 
    else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value 'CspmHmiC32,TCMHMIC01,TCMHMIC02,TCMHMIC03,TCMHMIC04,TCMHMIC05,TCMHMIC06,TCMHMIC07,TCMHMIC08,TCMHMIC09,TCMHMIC10,TCMHMIC11,TCMHMIC12,TCMHMIC13, TCMHMIC14,TCMHMIC15,TCMHMIC16,TCMHMIC17,TCMHMIC18,TCMHMIC19,TCMHMIC20' -Force
}

exit
}
