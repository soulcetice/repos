﻿Invoke-Command -ComputerName "SpmHmiC08" -ScriptBlock {Invoke-Expression -Command:"cmd.exe /c '\\10.16.85.39\C$\Temp\AutoStart.bat'"}