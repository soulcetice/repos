NCM_Downloader application
Date: 		22/07/2020
Author: 	MURESAN Radu-Adrian (MURA02)
Version:	1.5
Requirement:	InteroperabilityFunctions.dll

This app automates the NCM Manager to automatically download to the clients you select.
On the first download, it will generate an .ini file to keep the settings you just input.
the Remote Desktop Automation is required for allowing WinCC to also start the runtime (the user needs to be logged on)
the Remote Desktop Session to the engineering station must not be minimized in the taskbar in this version. should be visible, not necessarily maximized

************************************************************************************************************************************
REQUIRED FIELDS
	****************************************************************************************************************************
	Uppermost path field: used to check the LOAD.LOG for messages.
		\\10.16.80.31\d$\Project\SDIB_TCM\wincproj\SDIB_TCM_CLT_Ref(1)\winccom\LOAD.LOG
		
		In this case, the path to be input is: "\\10.16.80.31\d$\Project\SDIB_TCM\wincproj\SDIB_TCM_CLT_Ref"
	****************************************************************************************************************************
	LmHosts field: 		path with file name included (extension as well)
		Required for the IP and station data used for: 
					- displaying actual station names 
					- automatically opening the remote desktop sessions for the RT to be started for each client
	IMPORTANT! It needs the IPS & station names to be in the same order as they appear in the NCM Manager
	****************************************************************************************************************************
************************************************************************************************************************************
OPTIONAL FIELDS
	****************************************************************************************************************************
	IMPORTANT! IF RDP AUTOMATION IS TO BE USED, PLEASE BEFORE RUNNING CHECK TO NEVER WARN AGAIN ABOUT CERTIFICATES
AND AGAIN IMPORTANT! IF RDP AUTOMATION IS TO BE USED, PLEASE BEFORE RUNNING CHECK TO NEVER WARN AGAIN WHEN THAT YOU ARE CLOSING A REMOTE DESKTOP WINDOW
	Only fill if the start RDPs is checked:
		Username field: 	client logon username
		Password field: 	client logon password
		
**********************************************
In the end, the initialization file will contain data structured as follows:
D:\Project\SDIB_CSP\wincproj\SDIB_CSPM_CLT_Ref
6
24
C:\Windows\System32\drivers\etc\lmhosts
SDI
A02460
True