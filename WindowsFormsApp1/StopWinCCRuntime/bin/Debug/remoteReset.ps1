$Username = 'SDI'
$Password = 'A02460'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass

try { 
    invoke-command -ComputerName "SpmHmiC08" -credential $Cred -ScriptBlock {Start-Process '"C:\Program Files (x86)\Siemens\WinCC\bin\CCCleaner.exe"' "-terminate" }  
} catch {
    Write-Host "error"
}