﻿invoke-command -ComputerName "SpmHmiC08" -ScriptBlock {Start-Process '"C:\Program Files (x86)\Siemens\WinCC\bin\CCCleaner.exe"' "-terminate" }  