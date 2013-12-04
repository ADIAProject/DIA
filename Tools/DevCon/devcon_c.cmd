::--------------------------------------------------------------
:: Utility: Script to collect hardware IDs of specific devices
:: Author: Bashrat, Coop75 & JakeLD, & Romeo91
:: Last modification: v1.6 26/07/2010
:: Compatibility: Windows 2K-XP-2003-VISTA-7
::--------------------------------------------------------------
:: RAID HWID Tool wrapper from: OverFlow
::--------------------------------------------------------------
:: HDA Audio wrapper from: Debugger (DriverPacks Team)
::--------------------------------------------------------------

@ECHO OFF
CLS
TITLE .:: Save Harwdware IDs ::.

SET ExePath=%1
SET Out=%2
SET Mode=%3
SET HWID=%4

IF EXIST %Out% DEL /Q %Out%

IF %Mode%==1 (
	%ExePath% status * >> %Out%
)

IF %Mode%==2 (
	%ExePath% driverfiles * >> %Out%
)

IF %Mode%==3 (
	ECHO ============ >> %Out%
	ECHO ACPI Devices >> %Out%
	ECHO ============ >> %Out%
	%ExePath% find acpi* >> %Out%
	ECHO. >> %Out%
	ECHO ============ >> %Out%
	ECHO  HDA Audio >> %Out%
	ECHO ============ >> %Out%
	%ExePath% find hdaudio* >> %Out%
	ECHO. >> %Out%
	ECHO =========== >> %Out%
	ECHO PCI Devices >> %Out%
	ECHO =========== >> %Out%
	%ExePath% find pci* >> %Out%
	ECHO. >> %Out%
	ECHO =========== >> %Out%
	ECHO USB Devices >> %Out%
	ECHO =========== >> %Out%
	%ExePath% find usb* >> %Out%
	ECHO. >> %Out%
	ECHO ============= >> %Out%
	ECHO Input Devices >> %Out%
	ECHO ============= >> %Out%
	%ExePath% find hid* >> %Out%
	ECHO. >> %Out%
	ECHO ============ >> %Out%
	ECHO RAID Devices >> %Out%
	ECHO ============ >> %Out%
	%ExePath% hwids *CC_01* *Raid* >> %Out%
	ECHO. >> %Out%
	ECHO ================= >> %Out%
	ECHO BLUETOOTH Devices >> %Out%
	ECHO ================= >> %Out%
	%ExePath% find bluetooth* >> %Out%
	ECHO. >> %Out%
	ECHO =============== >> %Out%
	ECHO Monitor Devices >> %Out%
	ECHO =============== >> %Out%
	%ExePath% find monitor* >> %Out%
	ECHO. >> %Out%
	ECHO =============== >> %Out%
	ECHO Printer Devices >> %Out%
	ECHO =============== >> %Out%
	%ExePath% find Printer* >> %Out%
	%ExePath% find LPTENUM* >> %Out%
	%ExePath% find DOT4* >> %Out%
	ECHO. >> %Out%
	ECHO =============== >> %Out%
	ECHO ROOT Devices >> %Out%
	ECHO =============== >> %Out%
	%ExePath% find root* >> %Out%
	ECHO. >> %Out%
:: Show results with notepad
:: START %WinDir%\system32\Notepad.exe %Out%
)

IF %Mode%==4 (
TITLE .:: Remove device driver by Harwdware IDs ::.
ECHO ===============
ECHO 1. Scan Devices driver before delete - %HWID%
ECHO ===============
	%ExePath% status %HWID%
ECHO ===============
ECHO 2. Delete Devices driver
ECHO ===============
	%ExePath% remove %HWID% 
ECHO ===============
ECHO 3. Scan Devices driver after delete
ECHO ===============
	%ExePath% status %HWID%
	Pause
)


:: Clear variable
SET Out=

:: Delete extracted files
:: IF EXIST "%TMP%\devcon.exe" DEL \Q "%TMP%\devcon.exe"

:: Auto-Delete this file
:: DEL /F /Q %0

:: Exit
EXIT /B