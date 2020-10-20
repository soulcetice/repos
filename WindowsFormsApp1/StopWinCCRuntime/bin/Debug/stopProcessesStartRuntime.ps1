
Invoke-Command -ComputerName "SpmHmiC08" -Credential $Cred -ErrorAction Stop -ScriptBlock 
{
    Get-Process -Name  'CCAgent' | Stop-Process -Force
    
}