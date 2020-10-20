$Username = 'SDI'
$Password = 'A02460'
$pass = ConvertTo-SecureString -AsPlainText $Password -Force
$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username,$pass

try {
    Invoke-Command -ComputerName "SpmHmiC08" -credential $Cred -ErrorAction Stop -ScriptBlock {Invoke-Expression -Command:"cmd.exe /c '\\10.16.85.39\C$\Temp\StopWinCCRuntime.exe'"}
} catch {
    Write-Host "error"
}