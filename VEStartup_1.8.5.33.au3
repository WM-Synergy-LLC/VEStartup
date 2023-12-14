#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=SynergyNew.ico
#AutoIt3Wrapper_Outfile=VEStartup.exe
#AutoIt3Wrapper_Res_Description=Configure Infor client system settings
#AutoIt3Wrapper_Res_Fileversion=1.8.5.34
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2014 Synergy Resources, Inc.
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=asInvoker
#AutoIt3Wrapper_Res_Field=Copyright|Copyright © 2014 Synergy Resources, Inc., Central Islip, NY. All rights reserved.
#AutoIt3Wrapper_Res_Field=Company|Synergy Resources, Inc., Central Islip, NY
#AutoIt3Wrapper_Res_Field=Author|Craig D. Gunst
#AutoIt3Wrapper_Res_Field=Email|More.Info@SynergyResources.net
#AutoIt3Wrapper_Res_Field=Original filename|VEStartup.exe
#AutoIt3Wrapper_Res_Field=CompiledScript|AutoIt v3 Script
#AutoIt3Wrapper_Res_Field=CompiledDateTime|%date% - %time%
#AutoIt3Wrapper_Res_Field=AutoIt Version|%AutoItVer%
#AutoIt3Wrapper_Res_Field=Licence Agreement:|License Agreement: -------------------------------------------- This software is the intellectual property of Synergy Resources, Inc., Central Islip, NY and as such, is not to be distributed via any website domain or any other media without the prior written approval of Synergy Resources, Inc., Central Islip, NY. -------------------------------------------- In no event shall Synergy Resources, Inc., Central Islip, NY be liable to any party for direct, indirect, special, incidental, or consequential damages, including lost profits, arising out of the use of this software, even if Synergy Resources, Inc., Central Islip, NY has been advised of the possibility of such damage. Synergy Resources, Inc., Central Islip, NY specifically disclaims any warranties, including, but not limited to, the implied warranties of merchantability and fitness for a particular purpose. The software provided hereunder is on an "as is" basis, and Synergy Resources, Inc., Central Islip, NY has no obligations to provide maintenance, support, updates, enhancements, or modifications.
#AutoIt3Wrapper_Run_Tidy=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****


#cs -------------------------------------------------------------------------------------

	AutoIt Version: 3.2.8.1
	Author:         Craig D. Gunst (cdg)
	Craig.Gunst@SynergyResources.net

	Script Name:    VEStartup.au3
	Customer:       Multiple

	Script Function:
	GPO Startup script to update client files and registry


	Copyright © 2015 Synergy Resources, Inc., Central Islip, NY. All rights reserved.

	License Agreement:
	-------------------------------------------------------------------------------------
	This software is the intellectual property of Synergy Resources, Inc., Central Islip, NY
	and as such, is not to be distributed via any website domain or any other media
	without the prior written approval of Synergy Resources, Inc., Central Islip, NY.

	In no event shall Synergy Resources, Inc., Central Islip, NY be liable to any party
	for direct, indirect, special, incidental, or consequential damages, including lost
	profits, arising out of the use of this software, even if Synergy Resources, Inc.,
	Central Islip, NY has been advised of the possibility of such damage. Synergy
	Resources, Inc., Central Islip, NY specifically disclaims any warranties, including,
	but not limited to, the implied warranties of merchantability and fitness for a
	particular purpose. The software provided hereunder is on an "as is" basis, and
	Synergy Resources, Inc., Central Islip, NY has no obligations to provide
	maintenance, support, updates, enhancements, or modifictions.
	-------------------------------------------------------------------------------------

	Releases:
	2008-02-12 /cdg
	- Initial Release/cdg

	2008-02-15 /cdg
	- Added VScript sync/cdg
	- Fixed error in Update "VtaKioskConfiguration.config..."/cdg
	- Added error checking for all updates/copies/cdg

	2008-03-19	1.2.0.1 /cdg
	- Initial release for REMPRO.
	- Added standard AutoIt Wrapper with version numbering
	- Added code from VEStartupHelpFix
	- Added Array input for multiple entries for .ini lines

	2008-03-20	1.2.0.2	/cdg
	- Fix minor typos in AutoIt3Wrapper

	2008-07-18	1.2.1.0 /cdg
	- New for MEDCOR
	- Added ODBC options: AnsiNPW=No and AutoTranslate=No

	2008-07-27	1.3.1.2	/cdg
	- Update for RAMIND
	- Added recording script run info in registry
	- Removed exit from .chm help file parse
	- Syncronized version with VELogon
	- Added VQ Help file fix

	2008-09-06	1.3.2.1	/cdg
	- Added detect and update Oracle tnsnames.ora and sqlnet.ora
	- Syntax change: "If @error <> 1 Then" replaced with "If Not @error Then"

	2008-09-10	1.3.3.1	/cdg
	- Added delete of BuildInformation key in HKLM

	2008-09-29	1.3.3.2	/cdg
	- Added update of sql.ini Clientname in Win32client to computer name

	2008-10-02	1.3.3.3	/cdg
	- Truncate Clientname in sql.ini to 12 characters, 1st 7 & ... & last 2
	.	Computetr DT-PURCHASE-02 becomes Clientname DT-PURC...02
	.	Computetr DT-EXEC-01     stays   Clientname DT-EXEC-01

	2008-10-10	1.3.4.1	/cdg
	- Add immediate exit if no Infor products installed
	- Update VScript.DLL section to account for missing VScript.DLL
	- Change source for latest VScript.DLL from application folder to
	.	ClientConfigFiles folder

	2008-11-06	1.3.4.2	/cdg
	- Fixed path info in VScript.DLL update section

	2009-01-12	1.4.0.0	/cdg
	- Added GUI display

	2009-01-13	1.4.0.1	/cdg
	- Added .ini control of GUI delay

	2009-01-23	1.4.0.2	/cdg
	- Added support for unique ClientConfigFiles filders based on AD group of machines
	- "ClientConfigFiles\ADGroupConfigFiles\{ADGroupName} ConfigFiles"
	- [VEStartup] ADConfigGroups={array of Active Directory groups to test
	- Build folder heiarchy as needed
	- Included AD COM error handling
	- Fixed sqlnet.ora update test (was testing tnsnames.ora master against sqlnet.ora

	2009-01-24	1.4.0.3	/cdg
	- Added test for Lilly Software key for pre-6.5.x releases

	2009-01-25	1.4.0.4	/cdg
	- Copyright year update.

	2009-02-23	1.4.0.8	/cdg
	- Added looping through policy profiles as originally intended. Previously,
	.	only 1st hit on ADGroup would apply. Now presidence is Policy, ADGroup1,
	.	ADGroup2,... ADGroupx then base.

	2009-07-30	1.4.1.0	/cdg
	- Changed path for VScript.DLL if none exists, to %ProgramFiles%\\Infor Global Solutions\VScript

	2009-07-31	1.4.1.1	/cdg
	- Added Validate the PATH referecne to RunTime option (FixRunTimePath)
	- If VMFG or VQ installed, check if valid RunTime installed
	- Added Pre650Visual option for support of VMFG/VQ installed check

	2009-07-31	1.4.1.2	/cdg
	- Added test/create NetHASP,ini in .Net /bin folder
	- Reworked .Net local test for .config file updates.

	2009-07-31	1.4.1.3	/cdg
	- Added check for SQLBase environment variable

	2009-08-01	1.4.1.4	/cdg
	- Added HKCR and HKLM registry settings to associate .VMX files with Visual Manufacturing
	.	Note: Requires VELogon updates to setup HKCU registry settings

	2009-08-02	1.4.1.5	/cdg
	- Added updates from VELogon 1.4.0.7
	.	- Corrected logic in VMFG_Installed, VQ_Installed and CRM_Installed functions
	.	- Expanded test of registry keys in HKLM VMFG_Installed function

	2009-08-03	1.4.1.6	/cdg
	- Added HKCR and HKLM registry settings to associate .VFX files with Visual Financials
	- Added HKCR and HKLM registry settings to associate .VQX files with Visual Quality
	- Wrapped .VMX, .VFX and VQX associations in test for app.
	.	Note: File associations require VELogon updates to setup HKCU registry settings

	2009-08-04	1.4.1.7	/cdg
	- Corrected reference to $VQInstallPath and $VMInstallPath before defined.
	- Relocated Global GUI variables to Definitions section
	- Added missing ADConfigGroups definition when no parameter defined

	2009-08-06	1.4.1.8	/cdg
	- Corrected login in RunTime path test: $FirstRunTimePath

	2009-09-30	1.4.1.9	/cdg
	- Corrected issues in case and trailing backslash in PathFix compare logic
	- Removed bad logic "If NOT string = string" from PathFix

	2009-11-30	1.4.2.x	/cdg
	- Changed to auto incriment rev
	- Expanded Oracle auto discovery for all versions

	2010-02-27	1.4.3.x	/cdg
	- Changes for support of dual versions:
	-	- Added "SecondaryRunTimePath" parameter
	-	- Update Secondary RunTime sql.ini
	-	- Add "SecondaryRunTimePath" to system path

	2010-07-14	1.4.5.x	/cdg
	- Added management of ldap.ora, same as tnsnames.ora and sqlnet.ora

	2010-08-11	1.4.6.x	/cdg
	- Corrected oracle file updates where no file yet existed
	- Corrected ldap file distribution (was redistributing tnsnames.ora)
	- Updated copyright to 2010 in header

	2010-09-27	1.4.7.x	/cdg
	- Add x64 Oracle

	2010-12-10	1.4.8.x	/cdg
	- Correct .chm network file security entries for VM
	- Added .chm network file security entries for VQ in Profiles$

	2011-04-28	1.5.0.x	/cdg
	- Allow overwrite of sql.ini, tnsnames.ora, sqlnet.ora, vtakioskconfig.config,
	.	database.config, etc if the file is in read-only status.
	- Save a backup copy of the above files before updating. Keep the last 3 backups of
	.	each.

	2011-05-03	1.5.1.x	/cdg
	- Corrected Oracle x64 HOME discovery logic

	2011-05-03	1.5.1.x	/cdg
	- Disabled Oracle x32/x64 HOME discovery logic, x32 only

	2011-09-23	1.5.21.x	/cdg
	- Change Internet Zone from 1 to 2 for .CHM files
	.	1 = Intranet
	.	2 = Trusted

	2012-12-07	1.7.0.x		/cdg
	- Updated VE71x/VQ71x signatures for test of app present (i.e. "Infor 10 ERP Express...")
	- Updated CopyRight for Great River and 2012

	2012-12-13	1.7.1.x		/cdg
	- Added timeouts for all error messages

	2014-09-30	1.8.0.x		/cdg
	- Updated VScript.dll registration logic:
	-	VScript.dll now in local folder as resolved from new VMLocalPath
	-	Source for VScript.dll is now resolved from existing VMInstallPath
	-	added VMLocalPath
	- Updated CopyRight for Central Islip and 2014

	2014-10-09	1.8.1.x		/cdg
	- Update sql.ini on the fly with the "clientruntimedir=" set to the resolved Runtime path:

	2014-10-09	1.8.2.x		/cdg
	- Fix sql.ini update loop
	- Fix "clientruntimedir" logic

	2014-11-10	1.8.3.x	/cdg
	- Add optional installed clues paths
	- Move IGSKey test
	- Fixed typo in vscript.dll class ID

	2014-11-20	1.8.4.x	/cdg
	- Fixed logic when clues are defined but not matched

	2015-01-29	1.8.5.x	/cdg
	- Added correction for Spell Check by correcting HKLM keys


#ce -------------------------------------------------------------------------------------
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <StaticConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <Array.au3>



;----------------------------------------------------------------------------------------
; Definitions
;----------------------------------------------------------------------------------------
Dim $FileOverWrite = 1
Dim $ODBC_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\ODBC"
Dim $UrlAllowListKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\HTMLHelp\1.x\ItssRestrictions"
Dim $VMInstallPath = ""
Dim $VQInstallPath = ""
Dim $VMLocalPath = ""

; (+1.4.1.1)
Dim $Infor_HKCU_Base = "HKEY_CURRENT_USER\Software\Infor Global Solutions"
Dim $Infor_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\Infor Global Solutions"
Dim $Lilly_HKCU_Base = "HKEY_CURRENT_USER\Software\Lilly Software"
Dim $Lilly_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\Lilly Software"

; GUI support settings
Dim $GUI_Delay = 250 ; Slow down GUI message display (ms, 1000 = 1 sec)
Dim $ScriptName = StringLeft(@ScriptName, StringLen(@ScriptName) - 4)
Dim $VScriptDesc = "Visual Enterprise Client Startup Script"
Dim $Banner_Title = "Visual Enterprise" & @CRLF & "Client Administration Scripts"
Dim $ScriptFile_ini = @ScriptDir & "\" & $ScriptName & ".ini"
Dim $ScriptTempDir = @TempDir & "\$" & $ScriptName & "$"

; (+1.5.0.x)
Dim $ScriptRunTime = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC

; (+1.4.1.7)
Global $guiMsgBox
Global $guiLogo
Global $guiMain

Dim $oCOMError

; NOTE: DO NOT include trailing backslash ("\") character in path definitions

Dim $ScriptFile_ini = @ScriptDir & "\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 3) & "ini"

If FileExists($ScriptFile_ini) Then
	; Collect Alternate Local Installation Clues (+1.8.3.x)
	;-------------------------------------------------------
	$VEStartupVMInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "VMInstalledClue", "{key_not_defined}"))
	$VEStartupVQInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "VQInstalledClue", "{key_not_defined}"))
	$VEStartupCRMInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "CRMInstalledClue", "{key_not_defined}"))
	$VEStartupDotNetInstalledClue = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "DotNetInstalledClue", "{key_not_defined}"))

	If $VEStartupVMInstalledClue[1] = "{key_not_defined}" _
			And $VEStartupVQInstalledClue[1] = "{key_not_defined}" _
			And $VEStartupCRMInstalledClue[1] = "{key_not_defined}" _
			And $VEStartupDotNetInstalledClue[1] = "{key_not_defined}" Then

		;----------------------------------------------------------------------------------------
		; If Infor Products do not exist Exit immediately
		;----------------------------------------------------------------------------------------
		$IGSKey = RegRead("HKEY_LOCAL_MACHINE\Software\Infor Global Solutions", "")
		; Infor products do not exist
		If @error > 0 Then
			$LSKey = RegRead("HKEY_LOCAL_MACHINE\Software\Lilly Software", "")
			If @error > 0 Then
				Exit
			EndIf
		EndIf
	Else

		; If VMInstalledClue is defined look for valid path else use HKLM technique
		$VMInstalledFlag = False
		If Not ($VEStartupVMInstalledClue[0] = 1 And $VEStartupVMInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VEStartupVMInstalledClue[0]
				If FileExists($VEStartupVMInstalledClue[$i]) Then
					$VMInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		; If VQInstalledClue is defined look for valid path else use HKLM technique
		$VQInstalledFlag = False
		If Not ($VEStartupVQInstalledClue[0] = 1 And $VEStartupVQInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VEStartupVQInstalledClue[0]
				If FileExists($VEStartupVQInstalledClue[$i]) Then
					$VQInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		; If CRMInstalledClue is defined look for valid path else use HKLM technique
		$CRMInstalledFlag = False
		If Not ($VEStartupCRMInstalledClue[0] = 1 And $VEStartupCRMInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VEStartupCRMInstalledClue[0]
				If FileExists($VEStartupCRMInstalledClue[$i]) Then
					$CRMInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		; If DotNetInstalledClue is defined look for valid path else use HKLM technique
		$DotNetInstalledFlag = False
		If Not ($VEStartupDotNetInstalledClue[0] = 1 And $VEStartupDotNetInstalledClue[1] = "") Then
			; If any of the Installed Clue paths exist return true
			For $i = 1 To $VEStartupDotNetInstalledClue[0]
				If FileExists($VEStartupDotNetInstalledClue[$i]) Then
					$DotNetInstalledFlag = True
					ExitLoop
				EndIf
			Next
		EndIf

		If $VMInstalledFlag = False _
				And $VQInstalledFlag = False _
				And $CRMInstalledFlag = False _
				And $DotNetInstalledFlag = False Then
			Exit
		EndIf
	EndIf


	;----------------------------------------------------------------------------------------
	; Oracle client auto discovery (added in 1.4.2.x)
	;----------------------------------------------------------------------------------------

	; Read path for Oracle from registry
	$iRegKeyInstance = 0
	$OracleHome = ""
	$sRegKeyTemp = RegRead("HKEY_LOCAL_MACHINE\Software\wow6432node", "")
	If 1 = 1 Then ;@error > 0 Then ; 1.5.1.x - was "= 0"
		; x32
		$Oracle_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\Oracle"
	Else
		; x64
		$Oracle_HKLM_Base = "HKEY_LOCAL_MACHINE\Software\wow6432node\Oracle"
	EndIf

	While 1
		$iRegKeyInstance += 1
		$sRegKeyTemp = RegEnumKey($Oracle_HKLM_Base, $iRegKeyInstance)
		If @error <> 0 Then ExitLoop
		$OracleHome = RegRead($Oracle_HKLM_Base & "\" & $sRegKeyTemp, "ORACLE_HOME")
		If $OracleHome <> "" Then
			$OracleHomeName = RegRead($Oracle_HKLM_Base & "\" & $sRegKeyTemp, "ORACLE_HOME_NAME")
			ExitLoop
		EndIf
	WEnd


	;----------------------------------------------------------------------------------------
	; GUI Display Setup
	;----------------------------------------------------------------------------------------
	; Include GUI support files
	If FileExists($ScriptTempDir) Then FileDelete($ScriptTempDir)
	DirCreate($ScriptTempDir)
	FileInstall("SynergyPH3.JPG", $ScriptTempDir & "\")

	; Build main GUI display
	GUIBuildMain()
	GUISetState(@SW_SHOW)



	;----------------------------------------------------------------------------------------
	; Collect parameters from VEStartup.ini file
	;----------------------------------------------------------------------------------------
	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Recording script info in registry...")
	Sleep($GUI_Delay)


	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "@ScriptDir", "REG_SZ", @ScriptDir)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "@ScriptName", "REG_SZ", @ScriptName)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "RunTime", "REG_SZ", $ScriptRunTime)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "")
	$ScriptVersion = FileGetVersion(@ScriptFullPath)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ScriptVersion", "REG_SZ", $ScriptVersion)
	$IniVersion = FileGetTime($ScriptFile_ini, 0, 1)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "IniVersion", "REG_SZ", $IniVersion)



	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Reading " & $ScriptName & ".ini script configuration file...")

	;--------------------------
	; Read [VEStartup] Section
	;--------------------------




	$VEStartupClientConfigFilesPath = CleanPath(IniRead($ScriptFile_ini, "VEStartup", "ClientConfigFilesPath", ""))
	$VEStartupLocalRuntimePath = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "LocalRuntimePath", ""))
	$VEStartupLocalDotNetPath = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "LocalDotNetPath", ""))

	; {+1.8.0.x - Added VMLocalPath
	$VEStartupVMLocalPath = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "VMLocalPath", ""))
	For $p = 1 To $VEStartupVMLocalPath[0]
		If FileExists($VEStartupVMLocalPath[$p]) Then
			$VMLocalPath = $VEStartupVMLocalPath[$p]
			ExitLoop
		EndIf
	Next
	If $VMLocalPath = "" Then
		MsgBox(0, "VEStartup - Error: Cannot find VE Application local install path", _
				$ScriptFile_ini & @CRLF & _
				"No valid local application path found in any VMLocalPath locations" & @CRLF & @CRLF & _
				"Please report this error to your Administrator")
		RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot find VMLocalPath")
		Exit
	EndIf
	;}

	$VEStartupADConfigGroups = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "ADConfigGroups", ""))
	If $VEStartupADConfigGroups[1] <> "" Then

		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Checking " & @ComputerName & " account for Active Directory group membership...")

		$ADConfigGroups = _GetADMachineGroups($VEStartupADConfigGroups)
		For $ADg = 1 To $VEStartupADConfigGroups[0]
			If $VEStartupADConfigGroups[$ADg] = "" Then ContinueLoop
			If Not FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $VEStartupADConfigGroups[$ADg] & "} ConfigFiles") Then
				DirCreate($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $VEStartupADConfigGroups[$ADg] & "} ConfigFiles")
			EndIf
		Next
	Else
		$ADConfigGroups = $VEStartupADConfigGroups
	EndIf
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Reading " & $ScriptName & ".ini script configuration file...")
	Sleep($GUI_Delay)

	; GUI Delay Overide default = 250 (1/4 second) valid range = 0 to 10000
	$VEVEStartupGUIDelay = IniRead($ScriptFile_ini, "VEStartup", "VEStartupGUIDelay", "250")
	If $VEVEStartupGUIDelay >= 0 And $VEVEStartupGUIDelay <= 10000 Then
		$GUI_Delay = $VEVEStartupGUIDelay
	EndIf

	; (+1.4.1.1)
	$VEVEStartupFixRunTimePathFlag = GetIniFlag($ScriptFile_ini, "VEStartup", "FixRunTimePath", True)

	; (+1.4.1.1)
	; Setup for pre 6.5.x Visual registry
	$VEVEStartupPre650VisualFlag = GetIniFlag($ScriptFile_ini, "VEStartup", "Pre650Visual", False)
	If $VEVEStartupPre650VisualFlag Then
		$Reg_HKLM_Base = $Lilly_HKLM_Base
		$Reg_HKCU_Base = $Lilly_HKCU_Base
	Else
		$Reg_HKLM_Base = $Infor_HKLM_Base
		$Reg_HKCU_Base = $Infor_HKCU_Base
	EndIf

	; Manage dual RunTime's (+1.4.3.x)
	$VEVEStartupSecondaryRunTimePath = CleanPath(GetIniArray($ScriptFile_ini, "VEStartup", "SecondaryRunTimePath", ""))


	;---------------------------------------------------------------------------------------------
	; Process $VEStartupADConfigGroups and alter pointer to $VEStartupClientConfigFilesPath based
	; on first match.
	;---------------------------------------------------------------------------------------------

	;---------------------------
	; Read [Visual Mfg] Section
	;---------------------------
	; Added in 1.3.1.0
	;---------------------------------------------------------------------------------------------
	If VMFG_Installed() Then ;(+1.4.1.1)
		;	Search list for application location. Set first location where vm.exe is found.

		$VMInstallPathList = CleanPath(GetIniArray($ScriptFile_ini, "Visual Mfg", "VMInstallPath", ""))
		If $VMInstallPathList[1] <> "" Then
			; Update GUI message area
			GUICtrlSetFont($guiMsgBox, -1, 400)
			GUICtrlSetData($guiMsgBox, "Locating Visual Installation Path...")
			Sleep($GUI_Delay)
			$VMInstallPath = "" ; (+1.4.1.7)
			For $i = 1 To $VMInstallPathList[0]
				If FileExists($VMInstallPathList[$i] & "\vm.exe") Then
					$VMInstallPath = $VMInstallPathList[$i]
					ExitLoop
				EndIf
			Next
			If $VMInstallPath = "" Then
				MsgBox(16, "VEStartup - Error: Cannot find VM.EXE", _
						$ScriptFile_ini & @CRLF & _
						"VM.EXE not found in any VMInstallPath locations" & @CRLF & @CRLF & _
						"Please report this error to your Administrator", 60)
				RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot find VM.EXE")
				Exit
			EndIf
		EndIf
	EndIf
	;---------------------------------------------------------------------------------------------

	;-------------------------------
	; Read [Visual Quality] Section
	;-------------------------------
	; Added in 1.3.1.0
	;---------------------------------------------------------------------------------------------
	;	Search list for application location. Set first location where vm.exe is found.
	If VQ_Installed() Then ;(+1.4.1.1)
		$VQInstallPathList = CleanPath(GetIniArray($ScriptFile_ini, "Visual Quality", "VQInstallPath", ""))
		If $VQInstallPathList[1] <> "" Then
			; Update GUI message area
			GUICtrlSetFont($guiMsgBox, -1, 400)
			GUICtrlSetData($guiMsgBox, "Locating Quality Installation Path...")
			Sleep($GUI_Delay)

			$VQInstallPath = "" ; (+1.4.1.7)

			For $i = 1 To $VQInstallPathList[0]
				If FileExists($VQInstallPathList[$i] & "\vq.exe") Then
					$VQInstallPath = $VQInstallPathList[$i]
					ExitLoop
				EndIf
			Next
			If $VQInstallPath = "" Then
				MsgBox(16, "VEStartup - Error: Cannot find VQ.EXE", _
						$ScriptFile_ini & @CRLF & _
						"VQ.EXE not found in any VQInstallPath locations" & @CRLF & @CRLF & _
						"Please report this error to your Administrator", 60)
				RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Cannot find VQ.EXE")
				Exit
			EndIf
		EndIf
	EndIf
	;---------------------------------------------------------------------------------------------

	;----------------------------
	; Read [ODBC] Section
	;----------------------------
	$VEStartup_ODBC = IniReadSection($ScriptFile_ini, "ODBC")
	If Not @error Then
		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Rebuilding ODBC Connection Definitions...")
		Sleep($GUI_Delay)

		For $i = 1 To $VEStartup_ODBC[0][0]
			$ODBC_Parse = StringSplit($VEStartup_ODBC[$i][1], ",")
			$ODBC_Driver_Name = $ODBC_Parse[3]
			; Auto discover Oracle Home and driver name (added in 1.4.2.x)
			If StringLower($ODBC_Driver_Name) = "oracle" Then
				If $OracleHomeName = "" Then
					ContinueLoop
				Else
					$ODBC_Driver_Name = "Oracle in " & $OracleHomeName
					$ODBC_Driver = RegRead($ODBC_HKLM_Base & "\ODBCINST.INI\" & $ODBC_Driver_Name, "Driver")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Driver", "REG_SZ", $ODBC_Driver)
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Attributes", "REG_SZ", "W")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "FetchBufferSize", "REG_SZ", "64000")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "NumericSetting", "REG_SZ", "NLS")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "ForceWCHAR", "REG_SZ", "F")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "FailoverDelay", "REG_SZ", "10")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "FailoverRetryCount", "REG_SZ", "10")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "MetadataIdDefault", "REG_SZ", "F")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "BindAsDATE", "REG_SZ", "F")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "CloseCursor", "REG_SZ", "F")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "EXECSchemaOpt", "REG_SZ", "")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "EXECSyntax", "REG_SZ", "F")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Application Attributes", "REG_SZ", "T")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "ResultSets", "REG_SZ", "T")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "QueryTimeout", "REG_SZ", "T")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Failover", "REG_SZ", "T")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Lobs", "REG_SZ", "T")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "DisableMTS", "REG_SZ", "T")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "DisableDPM", "REG_SZ", "F")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "BatchAutocommitMode", "REG_SZ", "IfAllSuccessful")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Description", "REG_SZ", "")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "ServerName", "REG_SZ", $ODBC_Parse[2])
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Password", "REG_SZ", "")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "UserID", "REG_SZ", "SYSADM")
					RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "DSN", "REG_SZ", $VEStartup_ODBC[$i][0])


				EndIf
			Else
				$ODBC_Driver = RegRead($ODBC_HKLM_Base & "\ODBCINST.INI\" & $ODBC_Driver_Name, "Driver")

				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "AnsiNPW", "REG_SZ", "No")
				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "AutoTranslate", "REG_SZ", "No")
				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Database", "REG_SZ", $ODBC_Parse[2])
				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Description", "REG_SZ", "")
				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Driver", "REG_SZ", $ODBC_Driver)
				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "LastUser", "REG_SZ", "")
				RegWrite($ODBC_HKLM_Base & "\ODBC.INI\" & $VEStartup_ODBC[$i][0], "Server", "REG_SZ", $ODBC_Parse[1])



			EndIf
			RegWrite($ODBC_HKLM_Base & "\ODBC.INI\ODBC Data Sources", $VEStartup_ODBC[$i][0], "REG_SZ", $ODBC_Driver_Name)

			IniWrite("c:\Windows\ODBC.INI", "ODBC 32 bit Data Sources", $VEStartup_ODBC[$i][0], $ODBC_Driver_Name & " (32 bit)")
			IniWrite("c:\Windows\ODBC.INI", $VEStartup_ODBC[$i][0], "Driver32", $ODBC_Driver)

		Next
	EndIf


Else
	; Display GUI
	GUISetState(@SW_SHOW)

	MsgBox(16, "VEStartup - Error: Missing VEStartup.ini", _
			$ScriptFile_ini & @CRLF & _
			"VEStartup configuration file missing or inaccessable" & @CRLF & @CRLF & _
			"Please report this error to your Administrator", 60)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Missing VEStartup.ini")
	Exit
EndIf


;-----------------------------------
; Update Oracle client if installed
;-----------------------------------
; Added in 1.3.2.1
; Changed in 1.4.2.x to auto discover version and home
;---------------------------------------------------------------------------------------------

If $OracleHome <> "" Then
	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Updating Oracle Configuration Files...")
	Sleep($GUI_Delay)

	$OracleAdminPath = CleanPath($OracleHome) & "\network\admin"

	; Update Oracle 10g tnsnames.ora
	$TNSNamesMaster = ""
	For $ADCGPnt = 1 To $ADConfigGroups[0]
		If FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\tnsnames.ora") Then
			$TNSNamesMaster = $VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\tnsnames.ora"
			ExitLoop
		EndIf
	Next
	If $TNSNamesMaster = "" Then
		$TNSNamesMaster = $VEStartupClientConfigFilesPath & "\tnsnames.ora"
	EndIf


	$TNSNamesTime = FileGetTime($TNSNamesMaster, 0, 1)
	If Not @error Then
		$FCOK = True
		If Not FileExists($OracleAdminPath & "\tnsnames.ora") Then
			$FCOK = roFileCopy($TNSNamesMaster, $OracleAdminPath & "\", $FileOverWrite)
		Else
			If $TNSNamesTime <> FileGetTime($OracleAdminPath & "\tnsnames.ora", 0, 1) Then
				$FCOK = roFileCopy($TNSNamesMaster, $OracleAdminPath & "\", $FileOverWrite)
			EndIf
		EndIf
		If Not $FCOK Then
			MsgBox(16, @ScriptName & " - Error", "Unable to update Oracle tnsnames.ora" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update Oracle 10g tnsnames.ora")
			Exit
		EndIf
	EndIf

	; Update Oracle 10g sqlnet.ora
	$SQLNetMaster = ""
	For $ADCGPnt = 1 To $ADConfigGroups[0]
		If FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\sqlnet.ora") Then
			$SQLNetMaster = $VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\sqlnet.ora"
			ExitLoop
		EndIf
	Next
	If $SQLNetMaster = "" Then
		$SQLNetMaster = $VEStartupClientConfigFilesPath & "\sqlnet.ora"
	EndIf

	$SQLNetTime = FileGetTime($SQLNetMaster, 0, 1)
	If Not @error Then
		$FCOK = True
		If Not FileExists($OracleAdminPath & "\sqlnet.ora") Then
			$FCOK = roFileCopy($SQLNetMaster, $OracleAdminPath & "\", $FileOverWrite)
		Else
			If $SQLNetTime <> FileGetTime($OracleAdminPath & "\sqlnet.ora", 0, 1) Then
				$FCOK = roFileCopy($SQLNetMaster, $OracleAdminPath & "\", $FileOverWrite)
			EndIf
		EndIf
		If Not $FCOK Then
			MsgBox(16, @ScriptName & " - Error", "Unable to update Oracle sqlnet.ora" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update Oracle 10g sqlnet.ora")
			Exit
		EndIf
	EndIf

	; Update Oracle 10g ldap.ora (+1.4.5.x)
	$LDAPMaster = ""
	For $ADCGPnt = 1 To $ADConfigGroups[0]
		If FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\ldap.ora") Then
			$LDAPMaster = $VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\ldap.ora"
			ExitLoop
		EndIf
	Next
	If $LDAPMaster = "" Then
		$LDAPMaster = $VEStartupClientConfigFilesPath & "\ldap.ora"
	EndIf


	$LDAPTime = FileGetTime($LDAPMaster, 0, 1)
	If Not @error Then
		$FCOK = True
		If Not FileExists($OracleAdminPath & "\ldap.ora") Then
			$FCOK = roFileCopy($LDAPMaster, $OracleAdminPath & "\", $FileOverWrite)
		Else
			If $LDAPTime <> FileGetTime($OracleAdminPath & "\ldap.ora", 0, 1) Then
				$FCOK = roFileCopy($LDAPMaster, $OracleAdminPath & "\", $FileOverWrite)
			EndIf
		EndIf
		If Not $FCOK Then
			MsgBox(16, @ScriptName & " - Error", "Unable to update Oracle ldap.ora" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update Oracle 10g ldap.ora")
			Exit
		EndIf
	EndIf


EndIf

;---------------------------------------------------------------------------------------------

If VMFG_Installed() Or VQ_Installed() Then ; (+1.4.1.1)

	;----------------------------------
	; Update sql.ini file if it exists
	;----------------------------------
	$FirstRunTimePath = ""
	$SQLIniMaster = ""
	For $ADCGPnt = 1 To $ADConfigGroups[0]
		If FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\sql.ini") Then
			$SQLIniMaster = $VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\sql.ini"
			ExitLoop
		EndIf
	Next

	If $SQLIniMaster = "" Then
		$SQLIniMaster = $VEStartupClientConfigFilesPath & "\sql.ini"
	EndIf
	$SQLiniTime = FileGetTime($SQLIniMaster, 0, 1)
	If Not @error Then

		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Updating Gupta RunTime Configuration file (sql.ini)...")
		Sleep($GUI_Delay)

		; Update sql.ini file for RunTime if it exists
		For $i = 1 To $VEStartupLocalRuntimePath[0]
			If FileExists($VEStartupLocalRuntimePath[$i]) Then
				; Save first valid RunTime path to validate PATH environment variable (+1.4.1.1)
				If $FirstRunTimePath = "" Then $FirstRunTimePath = $VEStartupLocalRuntimePath[$i]

				If $SQLiniTime <> FileGetTime($VEStartupLocalRuntimePath[$i] & "\sql.ini", 0, 1) Then
					$FCOK = roFileCopy($SQLIniMaster, $VEStartupLocalRuntimePath[$i] & "\", $FileOverWrite)
					If Not $FCOK Then
						MsgBox(16, @ScriptName & " - Error", "Unable to update sql.ini" & @CRLF & _
								"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
								"Please report this error to your System Administrator", 60)
						RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update sql.ini")
						Exit
					EndIf
				EndIf

				; Update ClientName in sql.ini with local ComputerName
				$SQLiniClientName = StringUpper(@ComputerName)
				; Abreviate ClientName if greater than 12 characters
				If StringLen($SQLiniClientName) > 12 Then $SQLiniClientName = StringLeft($SQLiniClientName, 7) & "..." & StringRight($SQLiniClientName, 2)
				; Store the Client name in sql.ini then revert the file date-time back to the original
				If StringUpper(IniRead($VEStartupLocalRuntimePath[$i] & "\sql.ini", "Win32client", "Clientname", "")) <> $SQLiniClientName Then
					IniWrite($VEStartupLocalRuntimePath[$i] & "\sql.ini", "Win32client", "Clientname", $SQLiniClientName)
					FileSetTime($VEStartupLocalRuntimePath[$i] & "\sql.ini", $SQLiniTime, 0)
				EndIf
				If StringUpper(IniRead($VEStartupLocalRuntimePath[$i] & "\sql.ini", "Win32client", "ClientRuntimeDir", "")) <> $VEStartupLocalRuntimePath[$i] Then
					; {+1.8.1.x/1.8.2.x - Update clientruntimedir
					IniWrite($VEStartupLocalRuntimePath[$i] & "\sql.ini", "Win32client", "ClientRuntimeDir", $VEStartupLocalRuntimePath[$i])
					FileSetTime($VEStartupLocalRuntimePath[$i] & "\sql.ini", $SQLiniTime, 0)
					; }
				EndIf
				; {+1.8.2.x
				ExitLoop
				;}
			EndIf
		Next

		; No valid RunTime path found (+1.4.1.1)
		If $FirstRunTimePath = "" Then ; (1.4.1.8)
			MsgBox(16, @ScriptName & " - Error", "Gupta RunTime not found at any of the specified locations" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "RunTime not found")
			Exit
		EndIf

		; Update sql.ini file for Secondary RunTime if it exists (+1.4.3.x)
		$FirstSecondaryRunTimePath = ""
		For $i = 1 To $VEVEStartupSecondaryRunTimePath[0]
			If FileExists($VEVEStartupSecondaryRunTimePath[$i]) Then
				; Save first valid RunTime path to validate PATH environment variable (+1.4.1.1)
				If $FirstSecondaryRunTimePath = "" Then $FirstSecondaryRunTimePath = $VEVEStartupSecondaryRunTimePath[$i]

				If $SQLiniTime <> FileGetTime($VEVEStartupSecondaryRunTimePath[$i] & "\sql.ini", 0, 1) Then
					$FCOK = roFileCopy($SQLIniMaster, $VEVEStartupSecondaryRunTimePath[$i] & "\", $FileOverWrite)
					If Not $FCOK Then
						MsgBox(16, @ScriptName & " - Error", "Unable to update sql.ini" & @CRLF & _
								"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
								"Please report this error to your System Administrator", 60)
						RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update sql.ini")
						Exit
					EndIf
				EndIf

				; Update ClientName in sql.ini with local ComputerName
				$SQLiniClientName = StringUpper(@ComputerName)
				; Abreviate ClientName if greater than 12 characters
				If StringLen($SQLiniClientName) > 12 Then $SQLiniClientName = StringLeft($SQLiniClientName, 7) & "..." & StringRight($SQLiniClientName, 2)
				; Store the Client name in sql.ini then revert the file date-time back to the original
				If StringUpper(IniRead($VEVEStartupSecondaryRunTimePath[$i] & "\sql.ini", "Win32client", "Clientname", "")) <> $SQLiniClientName Then
					IniWrite($VEVEStartupSecondaryRunTimePath[$i] & "\sql.ini", "Win32client", "Clientname", $SQLiniClientName)
					FileSetTime($VEVEStartupSecondaryRunTimePath[$i] & "\sql.ini", $SQLiniTime, 0)
				EndIf

			EndIf
		Next

	EndIf


	;--------------------------------------------------------------
	; Validate/Fix RunTime in PATH environment variable (+1.4.1.1)
	;--------------------------------------------------------------
	; Added SecondaryRunTimePath (+1.4.3.x)
	If $VEVEStartupFixRunTimePathFlag Then
		If $FirstRunTimePath <> "" Then
			$aPathParsed = StringSplit(RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Path"), ";")
			; Check if RunTime path in PATH current environment variable
			If $aPathParsed[0] > 0 Then
				; Remove existing RunTimes in PATH
				$sPathNew = ""
				For $i = 1 To $aPathParsed[0]
					If StringLower(CleanPath($aPathParsed[$i])) <> StringLower($FirstRunTimePath) And _
							StringLower(CleanPath($aPathParsed[$i])) <> StringLower($FirstSecondaryRunTimePath) Then
						If $sPathNew <> "" Then $sPathNew &= ";"
						$sPathNew &= $aPathParsed[$i]
					EndIf
				Next

				If $FirstSecondaryRunTimePath <> "" Then
					$sPathNew = $FirstRunTimePath & ";" & $FirstSecondaryRunTimePath & ";" & $sPathNew
				Else
					$sPathNew = $FirstRunTimePath & ";" & $sPathNew
				EndIf

				RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "Path", "REG_EXPAND_SZ", $sPathNew)
			EndIf
		EndIf


		;---------------------------------------------------------
		; Added check for SQLBase environment variable (+1.4.1.3)
		;---------------------------------------------------------
		If RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", "SQLBASE") <> "" Then
			MsgBox(16, @ScriptName & " - Error", "SQLBase environment variable detected" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "SQLBase Environment variable present")
			Exit

		EndIf
	EndIf



	If VMFG_Installed() Then

		;------------------------------------------------------------------
		; Associate .VMX file type with VM.EXE in HKCR and HKLM (+1.4.1.4)
		;------------------------------------------------------------------
		RegWrite("HKEY_CLASSES_ROOT\.vmx", "", "REG_SZ", "VisualMfgExportFile")
		RegWrite("HKEY_CLASSES_ROOT\.vmx\VisualMfgExportFile")
		RegWrite("HKEY_CLASSES_ROOT\.vmx\VisualMfgExportFile\ShellNew")

		RegWrite("HKEY_CLASSES_ROOT\VisualMfgExportFile", "", "REG_SZ", "Visual Manufacturing Export File")
		RegWrite("HKEY_CLASSES_ROOT\VisualMfgExportFile", "EditFlags", "REG_DWORD", "00000000")
		RegWrite("HKEY_CLASSES_ROOT\VisualMfgExportFile\DefaultIcon", "", "REG_SZ", $VMInstallPath & "\vm.exe" & ",0")
		RegWrite("HKEY_CLASSES_ROOT\VisualMfgExportFile\shell", "", "REG_SZ", "open")
		RegWrite("HKEY_CLASSES_ROOT\VisualMfgExportFile\shell\open", "", "REG_SZ", "&Open")
		RegWrite("HKEY_CLASSES_ROOT\VisualMfgExportFile\shell\open\command", "", "REG_SZ", """" & $VMInstallPath & "\vm.exe" & """ ""%1""")

		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vmx", "", "REG_SZ", "VisualMfgExportFile")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vmx\VisualMfgExportFile")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vmx\VisualMfgExportFile\ShellNew")

		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualMfgExportFile", "", "REG_SZ", "Visual Manufacturing Export File")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualMfgExportFile", "EditFlags", "REG_DWORD", "00000000")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualMfgExportFile\DefaultIcon", "", "REG_SZ", $VMInstallPath & "\vm.exe" & ",0")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualMfgExportFile\shell", "", "REG_SZ", "open")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualMfgExportFile\shell\open", "", "REG_SZ", "&Open")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualMfgExportFile\shell\open\command", "", "REG_SZ", """" & $VMInstallPath & "\vm.exe" & """ ""%1""")

		;------------------------------------------------------------------
		; Associate .VFX file type with VF.EXE in HKCR and HKLM (+1.4.1.6)
		;------------------------------------------------------------------
		RegWrite("HKEY_CLASSES_ROOT\.vfx", "", "REG_SZ", "VisualFinancialsExportFile")
		RegWrite("HKEY_CLASSES_ROOT\.vfx\VisualFinancialsExportFile")
		RegWrite("HKEY_CLASSES_ROOT\.vfx\VisualFinancialsExportFile\ShellNew")

		RegWrite("HKEY_CLASSES_ROOT\VisualFinancialsExportFile", "", "REG_SZ", "Visual Financials Export File")
		RegWrite("HKEY_CLASSES_ROOT\VisualFinancialsExportFile", "EditFlags", "REG_DWORD", "00000000")
		RegWrite("HKEY_CLASSES_ROOT\VisualFinancialsExportFile\DefaultIcon", "", "REG_SZ", $VMInstallPath & "\vf.exe" & ",0")
		RegWrite("HKEY_CLASSES_ROOT\VisualFinancialsExportFile\shell", "", "REG_SZ", "open")
		RegWrite("HKEY_CLASSES_ROOT\VisualFinancialsExportFile\shell\open", "", "REG_SZ", "&Open")
		RegWrite("HKEY_CLASSES_ROOT\VisualFinancialsExportFile\shell\open\command", "", "REG_SZ", """" & $VMInstallPath & "\vf.exe" & """ ""%1""")

		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vfx", "", "REG_SZ", "VisualFinancialsExportFile")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vfx\VisualFinancialsExportFile")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vfx\VisualFinancialsExportFile\ShellNew")

		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualFinancialsExportFile", "", "REG_SZ", "Visual Financials Export File")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualFinancialsExportFile", "EditFlags", "REG_DWORD", "00000000")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualFinancialsExportFile\DefaultIcon", "", "REG_SZ", $VMInstallPath & "\vf.exe" & ",0")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualFinancialsExportFile\shell", "", "REG_SZ", "open")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualFinancialsExportFile\shell\open", "", "REG_SZ", "&Open")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualFinancialsExportFile\shell\open\command", "", "REG_SZ", """" & $VMInstallPath & "\vm.exe" & """ ""%1""")

	EndIf ; VMFG_Installed()


	If VQ_Installed() Then

		;------------------------------------------------------------------
		; Associate .VQX file type with VQ.EXE in HKCR and HKLM (+1.4.1.6)
		;------------------------------------------------------------------
		RegWrite("HKEY_CLASSES_ROOT\.vqx", "", "REG_SZ", "VisualQualityExportFile")
		RegWrite("HKEY_CLASSES_ROOT\.vqx\VisualQualityExportFile")
		RegWrite("HKEY_CLASSES_ROOT\.vqx\VisualQualityExportFile\ShellNew")

		RegWrite("HKEY_CLASSES_ROOT\VisualQualityExportFile", "", "REG_SZ", "Visual Financials Export File")
		RegWrite("HKEY_CLASSES_ROOT\VisualQualityExportFile", "EditFlags", "REG_DWORD", "00000000")
		RegWrite("HKEY_CLASSES_ROOT\VisualQualityExportFile\DefaultIcon", "", "REG_SZ", $VQInstallPath & "\vf.exe" & ",0")
		RegWrite("HKEY_CLASSES_ROOT\VisualQualityExportFile\shell", "", "REG_SZ", "open")
		RegWrite("HKEY_CLASSES_ROOT\VisualQualityExportFile\shell\open", "", "REG_SZ", "&Open")
		RegWrite("HKEY_CLASSES_ROOT\VisualQualityExportFile\shell\open\command", "", "REG_SZ", """" & $VQInstallPath & "\vf.exe" & """ ""%1""")

		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vqx", "", "REG_SZ", "VisualQualityExportFile")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vqx\VisualQualityExportFile")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\.vqx\VisualQualityExportFile\ShellNew")

		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualQualityExportFile", "", "REG_SZ", "Visual Financials Export File")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualQualityExportFile", "EditFlags", "REG_DWORD", "00000000")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualQualityExportFile\DefaultIcon", "", "REG_SZ", $VQInstallPath & "\vf.exe" & ",0")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualQualityExportFile\shell", "", "REG_SZ", "open")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualQualityExportFile\shell\open", "", "REG_SZ", "&Open")
		RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Classes\VisualQualityExportFile\shell\open\command", "", "REG_SZ", """" & $VQInstallPath & "\vm.exe" & """ ""%1""")

	EndIf ; VMFG_Installed()


EndIf ; VMFG_Installed() Or VQ_Installed()


;----------------------------------------------------------
; Test for the presence of VISUAL .Net apps (+1.4.1.2)
;----------------------------------------------------------
$sLocalDotNetBasePath = "" ;(+1.4.1.8)
For $i = 1 To $VEStartupLocalDotNetPath[0]
	If FileExists($VEStartupLocalDotNetPath[$i]) Then
		$sLocalDotNetBasePath = $VEStartupLocalDotNetPath[$i]
		ExitLoop
	EndIf
Next


;----------------------------------------------------------
; Update VISUAL .Net apps config (+1.4.1.2)
;----------------------------------------------------------
If $sLocalDotNetBasePath <> "" Then ;(1.4.1.8)

	;-------------------------------------------------
	; Check/Create for NetHASP.ini in bin ((+1.4.1.2)
	;-------------------------------------------------
	If Not FileExists($sLocalDotNetBasePath & "\bin\NetHASP.ini") Then
		IniWrite($sLocalDotNetBasePath & "\bin\NetHASP.ini", "NH_COMMON", "NH_TCPIP", "Disabled;")
	EndIf


	;----------------------------------------------------------
	; Update database.config file for DotNet Apps if it exists
	;----------------------------------------------------------
	$DBConfigMaster = ""
	For $ADCGPnt = 1 To $ADConfigGroups[0]
		If FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\database.config") Then
			$DBConfigMaster = $VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\database.config"
			ExitLoop
		EndIf
	Next
	If $DBConfigMaster = "" Then
		$DBConfigMaster = $VEStartupClientConfigFilesPath & "\database.config"
	EndIf

	$DBConfigTime = FileGetTime($DBConfigMaster, 0, 1)
	If Not @error Then
		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Updating Visual.Net configuration file (database.config)...")
		Sleep($GUI_Delay)

		; Removed loop in favor of $sLocalDotNetBasePath discovery (+1.4.1.2)
		$FCFlag = True
		If FileExists($sLocalDotNetBasePath & "\bin\database.config") Then
			If $DBConfigTime = FileGetTime($sLocalDotNetBasePath & "\bin\database.config", 0, 1) Then $FCFlag = False
		EndIf
		If $FCFlag Then
			$FCOK = roFileCopy($DBConfigMaster, $sLocalDotNetBasePath & "\bin\", $FileOverWrite)
			If Not $FCOK Then
				MsgBox(16, @ScriptName & " - Error", "Unable to update database.config" & @CRLF & _
						"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
						"Please report this error to your System Administrator", 60)
				RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update database.config")
				Exit
			EndIf
		EndIf
	EndIf


	;-------------------------------------------------------------------------
	; Update VtaKioskConfiguration.config file for VTA Kiosk App if it exists
	;-------------------------------------------------------------------------
	$KioskConfigMaster = ""
	For $ADCGPnt = 1 To $ADConfigGroups[0]
		If FileExists($VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\VtaKioskConfiguration.config") Then
			$KioskConfigMaster = $VEStartupClientConfigFilesPath & "\ADGroupConfigFiles\{" & $ADConfigGroups[$ADCGPnt] & "} ConfigFiles\VtaKioskConfiguration.config"
			ExitLoop
		EndIf
	Next
	If $KioskConfigMaster = "" Then
		$KioskConfigMaster = $VEStartupClientConfigFilesPath & "\VtaKioskConfiguration.config"
	EndIf

	$KioskConfigTime = FileGetTime($KioskConfigMaster, 0, 1)
	If Not @error Then
		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Updating VTA Kiosk configuration file (VtaKioskConfiguration.config)...")
		Sleep($GUI_Delay)

		; Removed loop in favor of $sLocalDotNetBasePath discovery (+1.4.1.2)
		$FCFlag = True
		If FileExists($sLocalDotNetBasePath & "\bin\VtaKioskConfiguration.config") Then
			If $DBConfigTime = FileGetTime($sLocalDotNetBasePath & "\bin\VtaKioskConfiguration.config", 0, 1) Then $FCFlag = False
		EndIf
		If $FCFlag Then
			$FCOK = roFileCopy($KioskConfigMaster, $VEStartupLocalDotNetPath[$i] & "\bin\", $FileOverWrite)
			If Not $FCOK Then
				MsgBox(16, @ScriptName & " - Error", "Unable to update VtaKioskConfiguration.config" & @CRLF & _
						"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
						"Please report this error to your System Administrator", 60)
				RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update VtaKioskConfiguration.config")
				Exit
			EndIf
		EndIf
	EndIf


EndIf ; IsDeclared($sLocalDotNetBasePath)


;-----------------------------------------
; Install / Update / Register VScript.dll
;-----------------------------------------
; { +1.8.0.x - Redefine VScript.dll install / update / register

If FileExists($VMInstallPath & "\VScript.dll") Then

	; If missing or different
	If FileGetVersion($VMInstallPath & "\VScript.dll") <> FileGetVersion($VMLocalPath & "\VScript.dll") Then

		; Update GUI message area
		GUICtrlSetFont($guiMsgBox, -1, 400)
		GUICtrlSetData($guiMsgBox, "Updating macro script library file (VScript.dll)...")
		Sleep($GUI_Delay)

		; Copy VScript.DLL to local
		$FCOK = roFileCopy($VMInstallPath & "\VScript.dll", $VMLocalPath & "\", $FileOverWrite)
		If Not $FCOK Then
			MsgBox(16, @ScriptName & " - Error", "Unable to update VScript.DLL" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to update VScript.DLL")
			Exit
		EndIf

		; Register VScript.dll
		$OSROK = RunWait("regsvr32.exe /s " & """" & $VMLocalPath & "\VScript.dll" & """")
		If $OSROK <> 0 Then
			MsgBox(16, @ScriptName & " - Error", "Unable to unregister previous VScript.DLL" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to unregister VScript.DLL")
			Exit
		EndIf
	EndIf

	; Is Vscript.dll registered?
	If $VMLocalPath & "\VScript.dll" <> RegRead("HKEY_CLASSES_ROOT\CLSID\{0A7378E2-3D6E-11D4-888D-0010A4E80B77}\InprocServer32", "") Then
		; Register VScript.dll
		$OSROK = RunWait("regsvr32.exe /s " & """" & $VMLocalPath & "\VScript.dll" & """")
		If $OSROK <> 0 Then
			MsgBox(16, @ScriptName & " - Error", "Unable to unregister previous VScript.DLL" & @CRLF & _
					"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
					"Please report this error to your System Administrator", 60)
			RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to unregister VScript.DLL")
			Exit
		EndIf
	EndIf

Else
	; Master Vscript.dll missing
	MsgBox(16, @ScriptName & " - Error", "Unable to locate master VScript.DLL in any VMInstallPath locations" & @CRLF & _
			"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
			"Please report this error to your System Administrator", 60)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to locate master VScript.DLL in any VMInstallPath locations")
	Exit

EndIf
; }


; Permit .chm files from VMInstallPath
;--------------------------------------
; Check for valid UNC beginning with 2 back-slash characters
$VMInstallPathParsed = StringSplit($VMInstallPath, "\")
If $VMInstallPathParsed[0] >= 4 And $VMInstallPathParsed[1] = "" And $VMInstallPathParsed[2] = "" Then
	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Setting browser security for Visual Mfg help files...")
	Sleep($GUI_Delay)

	; Set registry to specifically allow .CHM files from Visual share
	;-----------------------------------------------------------------
	$VMInstallPathUrl = "\\" & $VMInstallPathParsed[3] & "\" & $VMInstallPathParsed[4]
	$UrlAllowList = RegRead($UrlAllowListKey, "UrlAllowList")

	; If the list does not exist or is empty, add the paths:
	If RegRead($UrlAllowListKey, "UrlAllowList") = "" Then
		$UrlAllowList = $VMInstallPathUrl & ";file://" & $VMInstallPathUrl
		RegWrite($UrlAllowListKey, "UrlAllowList", "REG_SZ", $UrlAllowList)

		; If there is already an allow list, see if it contains the paths:
	Else
		; See if it contains part 1 of path, if not, add both:
		If Not StringInStr($UrlAllowList, $VMInstallPathUrl) Then
			$UrlAllowList = $UrlAllowList & ";" & $VMInstallPathUrl & ";file://" & $VMInstallPathUrl
		EndIf
		; See if it contains part 2 of the path, if not, add part 2:
		If Not StringInStr($UrlAllowList, "file://" & $VMInstallPathUrl) Then
			$UrlAllowList = $UrlAllowList & ";file://" & $VMInstallPathUrl
		EndIf
		; Update the allow list:
		RegWrite($UrlAllowListKey, "UrlAllowList", "REG_SZ", $UrlAllowList)
	EndIf

	; Set registry to permit .chm from Local Intranet Zone paths
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\HTMLHelp\1.x\ItssRestrictions", "MaxAllowedZone", "REG_DWORD", 2)

	; Add server share to Local Intranet Zone
	RegWrite("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\" & $VMInstallPathParsed[3], "file", "REG_DWORD", 2)
EndIf


; Permit .chm files from VQInstallPath
;--------------------------------------
; Check for valid UNC beginning with 2 back-slash characters
$VQInstallPathParsed = StringSplit($VQInstallPath, "\")
If $VQInstallPathParsed[0] >= 4 And $VQInstallPathParsed[1] = "" And $VQInstallPathParsed[2] = "" Then
	; Update GUI message area
	GUICtrlSetFont($guiMsgBox, -1, 400)
	GUICtrlSetData($guiMsgBox, "Setting browser security for Visual Quality help files...")
	Sleep($GUI_Delay)


	; Set registry to specifically allow .CHM files from Visual share
	;-----------------------------------------------------------------
	$VQInstallPathUrl = "\\" & $VQInstallPathParsed[3] & "\" & $VQInstallPathParsed[4]
	$UrlAllowList = RegRead($UrlAllowListKey, "UrlAllowList")

	; If the list does not exist or is empty, add the paths:
	If RegRead($UrlAllowListKey, "UrlAllowList") = "" Then
		$UrlAllowList = $VQInstallPathUrl & ";file://" & $VQInstallPathUrl
		RegWrite($UrlAllowListKey, "UrlAllowList", "REG_SZ", $UrlAllowList)

		; If there is already an allou list, see if it contains the paths:
	Else
		; See if it contains part 1 of path, if not, add both:
		If Not StringInStr($UrlAllowList, $VQInstallPath) Then
			$UrlAllowList = $UrlAllowList & ";" & $VQInstallPathUrl & ";file://" & $VQInstallPathUrl
		EndIf
		; See if it contains part 2 of the path, if not, add part 2:
		If Not StringInStr($UrlAllowList, "file://" & $VQInstallPathUrl) Then
			$UrlAllowList = $UrlAllowList & ";file://" & $VQInstallPathUrl
		EndIf
		; Update the allow list:
		RegWrite($UrlAllowListKey, "UrlAllowList", "REG_SZ", $UrlAllowList)
	EndIf

	; Set registry to permit .chm from Local Intranet Zone paths
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\HTMLHelp\1.x\ItssRestrictions", "MaxAllowedZone", "REG_DWORD", 2)

	; Add server share to Local Intranet Zone
	RegWrite("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\" & $VQInstallPathParsed[3], "file", "REG_DWORD", 2)
EndIf



;--------------------------------------------------------------------------------------
; Delete of BuildInformation key in HKLM
;--------------------------------------------------------------------------------------
; - Added in 1.3.3.1 /cdg
;		Pre 6.5.x releases copied the key to HKCU causing duplications
;		following upgrades. If the key is removed from earlier release
;		new, unique key will be generated.
;--------------------------------------------------------------------------------------
; Update GUI message area
GUICtrlSetFont($guiMsgBox, -1, 400)
GUICtrlSetData($guiMsgBox, "Removing duplicate BuildInformation registry key...")
Sleep($GUI_Delay)

RegDelete("HKEY_LOCAL_MACHINE\SOFTWARE\Identification\Other\BuildInformation")



;--------------------------------------------------------------------------------------
; Correct lexicon keys in HKLM to six spell check issue
;--------------------------------------------------------------------------------------
; - Added in 1.8.5.x /cdg
;		Per Infor KB 1408377, path info in HKLM is incorrect for spell check
;--------------------------------------------------------------------------------------
; Update GUI message area
GUICtrlSetFont($guiMsgBox, -1, 400)
GUICtrlSetData($guiMsgBox, "Updating Spell Check Configuration Registry Keys...")
Sleep($GUI_Delay)

RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Wintertree\SSCE", "HelpFile", "REG_SZ", $VMLocalPath & '\Ssce.hlp')
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Wintertree\SSCE", "MainLexPath", "REG_SZ", $VMLocalPath)
RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\Wintertree\SSCE", "UserLexPath", "REG_SZ", $VMLocalPath)



;--------------------------------------------------------------------------------------
; Normal Exit
;--------------------------------------------------------------------------------------
RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Normal")

Exit



;======================================================================================
; Function Definitions
;======================================================================================

;--------------------------------------------------------------------------------------
; GetIniArray - Get multi-value key from .ini file
;--------------------------------------------------------------------------------------
; Function to read multiple entries from .ini key to array
; array[0] contains the number of elements, array[1-array[0]]
; contains the values stripped of leading and trailing space

Func GetIniArray($iniFilename, $iniSection, $iniKey, $iniDefault)

	Local $IniReturn[32], $IniTemp, $i

	$IniTemp = IniRead($iniFilename, $iniSection, $iniKey, "~{KeyNotFound}")
	If $IniTemp = "~{KeyNotFound}" Then
		$IniTemp = $iniDefault
	EndIf

	$IniReturn = StringSplit($IniTemp, ",", 0)
	For $i = 1 To $IniReturn[0]
		$IniReturn[$i] = StringStripWS($IniReturn[$i], 3)
	Next

	Return $IniReturn

EndFunc   ;==>GetIniArray


;--------------------------------------------------------------------------------------
; CleanPath - Remove trailing backslash ("\") from path
;--------------------------------------------------------------------------------------
; Also strips leading and trailing space
; If variable is an array, array[0]= number of elements

Func CleanPath($InputPath)

	If IsArray($InputPath) Then
		For $i = 1 To $InputPath[0]
			$InputPath[$i] = StringStripWS($InputPath[$i], 3)
			If StringRight($InputPath[$i], 1) = "\" Then
				$InputPath[$i] = StringLeft($InputPath[$i], StringLen($InputPath[$i]) - 1)
			EndIf
		Next
	Else
		$InputPath = StringStripWS($InputPath, 3)
		If StringRight($InputPath, 1) = "\" Then
			$InputPath = StringLeft($InputPath, StringLen($InputPath) - 1)
		EndIf
	EndIf

	Return $InputPath

EndFunc   ;==>CleanPath

;============================================================================================
; Display GUI
;============================================================================================

Func GUIBuildMain()

	$iMainWidth = 500
	$iCenterHight = 0

	$iMainHight = $iCenterHight + 123 ; Minimum = 123 for no center window

	$iTitleHight = 100
	$iLogoWidth = 80

	;--------------------------------------------------------------------------------------------
	; Build Main Window
	;--------------------------------------------------------------------------------------------
	;	$GUID_Main = GUICreate($ScriptName & " - " & $VScriptDesc , $iMainWidth, $iMainHight, -1, -1, BitOR($WS_CAPTION, $WS_POPUP, $WS_SYSMENU))
	;	$GUID_Main = GUICreate($ScriptName & " - " & $VScriptDesc , $iMainWidth, $iMainHight, -1, -1, BitOR($WS_POPUPWINDOW,$WS_CAPTION),BitOR($WS_EX_TOPMOST  ,0) )
	$guiMain = GUICreate($ScriptName & " - " & $VScriptDesc, $iMainWidth, $iMainHight + 27, -1, -1, $WS_DLGFRAME)


	;--------------------------------------------------------------------------------------------
	; Build top banner
	;--------------------------------------------------------------------------------------------
	; Install Synergy Banner
	$guiLogo = GUICtrlCreatePic($ScriptTempDir & "\SynergyPH3.JPG", 0, 0, $iMainWidth, $iTitleHight)
	GUICtrlSetState(-1, $GUI_DISABLE)


	$V_Offset = 2
	$H_Offset = 5
	$Banner_Font = "Arial Black"
	GUICtrlCreateLabel($Banner_Title, $iLogoWidth + $H_Offset, 15 + $V_Offset, $iMainWidth - $iLogoWidth - $H_Offset, 70 - $V_Offset, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 400, 0, $Banner_Font)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, 0x000000)
	GUICtrlSetState(-1, $GUI_ENABLE)

	; Title Text
	GUICtrlCreateLabel($Banner_Title, $iLogoWidth, 15, $iMainWidth - $iLogoWidth, 70, $SS_CENTER)
	GUICtrlSetFont(-1, 14, 400, 0, $Banner_Font)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, 0xFFFFFF)

	; Display copyright in lower-right of banner
	GUICtrlCreateLabel("©Synergy Resources, Inc.", $iMainWidth - 150, 84, 140, 12, $SS_RIGHT)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, 0xB0B0B0)

	; Display center window
	If $iCenterHight > 0 Then
		GUICtrlCreateGraphic(2, $iTitleHight + 2, $iMainWidth - 4, ($iMainHight - $iTitleHight) - 24, $SS_SUNKEN)
		GUICtrlSetState(-1, $GUI_DISABLE)
	EndIf


	; Display program version in lower-left of window
	$ScriptRev = FileGetVersion(@ScriptFullPath)
	$iRevWidth = (StringLen($ScriptRev) * 5.5) + 24
	GUICtrlCreateGraphic(($iMainWidth - $iRevWidth) - 6, $iMainHight - 21, $iRevWidth + 5, 20, $SS_SUNKEN)
	GUICtrlCreateLabel("rev: " & $ScriptRev, ($iMainWidth - $iRevWidth) - 4, $iMainHight - 18, $iRevWidth, 14, $SS_CENTER)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetColor(-1, 0x505050)

	; Display Message Area in lower-left of window
	GUICtrlCreateGraphic(2, $iMainHight - 21, ($iMainWidth - $iRevWidth) - 9, 20, $SS_SUNKEN)
	$guiMsgBox = GUICtrlCreateLabel("", 6, $iMainHight - 18, ($iMainWidth - $iRevWidth) - 14, 14, $SS_LEFT)
	GUICtrlSetFont(-1, 8, 400, 0, "Arial")
	GUICtrlSetColor(-1, 0x505050)

EndFunc   ;==>GUIBuildMain


;--------------------------------------------------------------------------------------
; _GetADMachineGroups - Read Active Directory profile security group membership
;--------------------------------------------------------------------------------------

Func _GetADMachineGroups($ADGroups)

	Local $objComputer, $CurrentComputer
	Local $vMembershipList = ""

	$objComputer = ObjCreate("ADSystemInfo")
	$oCOMError = ObjEvent("AutoIt.Error", "_ADError")
	$CurrentComputer = ObjGet("LDAP://" & $objComputer.ComputerName)
	$avMembersList = $CurrentComputer.MemberOf
	If IsArray($avMembersList) Then
		For $vMember In $avMembersList
			$avMemberParsed = StringSplit($vMember, ",")
			$avMemberParsed = StringSplit($avMemberParsed[1], "=")
			If $vMembershipList <> "" Then $vMembershipList &= ","
			$vMembershipList &= $avMemberParsed[2]
		Next
	Else
		$avMemberParsed = StringSplit($avMembersList, ",")
		$avMemberParsed = StringSplit($avMemberParsed[1], "=")
		If $avMemberParsed[0] > 1 Then
			$vMembershipList &= $avMemberParsed[2]
		Else
			$vMembershipList = ""
		EndIf
	EndIf

	$avMembershipList = _StringSplitTrim($vMembershipList, ",")

	Local $vADGList = ""

	If Not IsArray($ADGroups) Then
		If $ADGroups = "*" Then
			Return $avMembershipList
		Else
			$ADGroups = _StringSplitTrim($ADGroups, ",")
		EndIf
	EndIf

	For $iADG = 1 To $ADGroups[0]
		For $iML = 1 To $avMembershipList[0]
			If StringLower($ADGroups[$iADG]) = StringLower($avMembershipList[$iML]) Then
				If $vADGList <> "" Then $vADGList &= ","
				$vADGList &= $ADGroups[$iADG]
			EndIf
		Next
	Next
	$avADGList = _StringSplitTrim($vADGList, ",")
	Return $avADGList

EndFunc   ;==>_GetADMachineGroups

Func _ADError()
	$HexNumber = Hex($oCOMError.number, 8)
	MsgBox(16, @ScriptName & " - Error", "Unable to read Active Directory groups" & @CRLF & _
			"Visual Enterprise applactions may not function as expected" & @CRLF & @CRLF & _
			"Please report this error to your System Administrator", 60)
	RegWrite("HKEY_LOCAL_MACHINE\Software\Synergy Resources\" & StringLeft(@ScriptName, StringLen(@ScriptName) - 4), "ExitCondition", "REG_SZ", "Unable to read Active Directory groups")
	Exit
EndFunc   ;==>_ADError


;--------------------------------------------------------------------------------------
; _StringSplitTrim - Split string and trim leading and trailing whitespace
;--------------------------------------------------------------------------------------

Func _StringSplitTrim($SST_String, $SST_Delimiter)

	Local $iSST
	$sSST_Tmp = StringSplit($SST_String, $SST_Delimiter)
	For $iSST = 1 To $sSST_Tmp[0]
		$sSST_Tmp[$iSST] = StringStripWS($sSST_Tmp[$iSST], 3)
	Next

	Return $sSST_Tmp
EndFunc   ;==>_StringSplitTrim


;--------------------------------------------------------------------------------------
; GetIniFlag - Read ini flag parameter Yes/No and return True/False (+1.4.1.1)
;--------------------------------------------------------------------------------------

Func GetIniFlag($GIFIniFile, $GIFIniSection, $GIFIniKey, $GIFDefault)

	$GIFTemp = IniRead($GIFIniFile, $GIFIniSection, $GIFIniKey, "{key_not_found}")
	If $GIFTemp = "{key_not_found}" Then
		Return $GIFDefault
	EndIf

	$GIFTemp = StringStripWS(StringLower($GIFTemp), 3)
	If StringInStr("|1|yes|y|true|t|", $GIFTemp) Then Return True
	If StringInStr("|0|no|n|false|f|", $GIFTemp) Then Return False
	SetError(1)
	Return $GIFDefault

EndFunc   ;==>GetIniFlag


;--------------------------------------------------------------------------------------
; Test for Visual Manufacturing install (+1.4.1.1, 1.4.1.5)
;--------------------------------------------------------------------------------------

Func VMFG_Installed()
	; {+1.8.3.x
	; If VMInstalledClue is defined look for valid path else use HKLM technique
	If $VEStartupVMInstalledClue[1] <> "" And $VEStartupVMInstalledClue[1] <> "{key_not_defined}" Then
		; If any of the Installed Clue paths exist return true
		For $i = 1 To $VEStartupVMInstalledClue[0]
			If FileExists($VEStartupVMInstalledClue[$i]) Then
				Return True
			EndIf
		Next
		Return False
	EndIf
	; }

	Local $i
	$i = 0
	While 1
		$i += 1
		$RegKey = RegEnumKey($Infor_HKLM_Base, $i)
		If @error <> 0 Then
			ExitLoop
		ElseIf StringInStr($RegKey, "Visual") > 0 And (StringInStr($RegKey, "Manufacturing") > 0 Or StringInStr($RegKey, "Mfg") > 0 Or StringInStr($RegKey, "Enterprise") > 0 Or (StringInStr($RegKey, "Infor") > 0 And StringInStr($RegKey, "10") > 0 And StringInStr($RegKey, "ERP") > 0 And StringInStr($RegKey, "Express") > 0)) Then
			Return True
		EndIf
	WEnd
	If $VEVEStartupPre650VisualFlag Then
		$i = 0
		While 1
			$i += 1
			$RegKey = RegEnumKey($Lilly_HKLM_Base, $i)
			If @error <> 0 Then
				ExitLoop
			ElseIf StringInStr($RegKey, "Visual") > 0 And (StringInStr($RegKey, "Manufacturing") > 0 Or StringInStr($RegKey, "Mfg") > 0 Or StringInStr($RegKey, "Enterprise") > 0) Then
				Return True
			EndIf
		WEnd
	EndIf
	Return False

EndFunc   ;==>VMFG_Installed


;--------------------------------------------------------------------------------------
; Test for Visual CRM install (+1.4.1.1, 1.4.1.5)
;--------------------------------------------------------------------------------------

Func CRM_Installed()
	; {+1.8.3.x
	; If CRMInstalledClue is defined look for valid path else use HKLM technique
	If $VEStartupCRMInstalledClue[1] <> "" And $VEStartupCRMInstalledClue[1] <> "{key_not_defined}" Then
		; If any of the Installed Clue paths exist return true
		For $i = 1 To $VEStartupCRMInstalledClue[0]
			If FileExists($VEStartupCRMInstalledClue[$i]) Then
				Return True
			EndIf
		Next
		Return False
	EndIf
	; }

	Local $i
	$i = 0
	While 1
		$i += 1
		$RegKey = RegEnumKey($Infor_HKLM_Base, $i)
		If @error <> 0 Then
			ExitLoop
		ElseIf StringInStr($RegKey, "Visual CRM") Then
			Return True
		EndIf
	WEnd
	If $VEVEStartupPre650VisualFlag Then
		$i = 0
		While 1
			$i += 1
			$RegKey = RegEnumKey($Lilly_HKLM_Base, $i)
			If @error <> 0 Then
				ExitLoop
			ElseIf StringInStr($RegKey, "Visual CRM") Then
				Return True
			EndIf
		WEnd
	EndIf
	Return False

EndFunc   ;==>CRM_Installed


;--------------------------------------------------------------------------------------
; Test for Visual Quality install (+1.4.1.1, 1.4.1.5)
;--------------------------------------------------------------------------------------

Func VQ_Installed()
	; {+1.8.3.x
	; If VQInstalledClue is defined look for valid path else use HKLM technique
	If $VEStartupVQInstalledClue[1] <> "" And $VEStartupVQInstalledClue[1] <> "{key_not_defined}" Then
		; If any of the Installed Clue paths exist return true
		For $i = 1 To $VEStartupVQInstalledClue[0]
			If FileExists($VEStartupVQInstalledClue[$i]) Then
				Return True
			EndIf
		Next
		Return False
	EndIf
	; }

	Local $i
	$i = 0
	While 1
		$i += 1
		$RegKey = RegEnumKey($Infor_HKLM_Base, $i)
		If @error <> 0 Then
			ExitLoop
		ElseIf StringInStr($RegKey, "Visual Quality") Then
			Return True
		EndIf
	WEnd
	If $VEVEStartupPre650VisualFlag Then
		$i = 0
		While 1
			$i += 1
			$RegKey = RegEnumKey($Lilly_HKLM_Base, $i)
			If @error <> 0 Then
				ExitLoop
			ElseIf StringInStr($RegKey, "Visual Quality") Then
				Return True
			EndIf
		WEnd
	EndIf
	Return False

EndFunc   ;==>VQ_Installed


;--------------------------------------------------------------------------------------
; roFileCopy - Read-Only FileCopy with backup
;--------------------------------------------------------------------------------------
; Force FileCopy to work on files with read-only flag set
; Test the read-only attribute, remove it if it is set, do the copy, then set the
; attribute back to read-only
; NOTE: Not intended for use with wildcards
;--------------------------------------------------------------------------------------
Func roFileCopy($roFCSource, $roFCdest, $roFCflag)
	Local $roFC = 0
	Local $roFCbu[1] = [0]

	; Parse Source into path, file, file name and file ext
	$a_roFCSource = StringSplit($roFCSource, "\")
	$roFCSourceFile = $a_roFCSource[$a_roFCSource[0]]
	$a_roFCFile = StringSplit($roFCSourceFile, ".")
	$roFCFileName = $a_roFCFile[1]
	If $a_roFCFile[0] = 2 Then
		$roFCFileExt = "." & $a_roFCFile[2]
	EndIf
	$roFCSourcePath = StringLeft($roFCSource, StringLen($roFCSource) - StringLen($roFCSourceFile))

	; Strip trailing "\" from dest
	$roFCdest = CleanPath($roFCdest)

	If FileExists($roFCSource) Then
		If FileExists($roFCdest & "\" & $roFCSourceFile) Then

			; Backup the old destination file
			If BitAND($roFCflag, 1) = 1 Then ; Overwrite flag set?
				If StringInStr(FileGetAttrib($roFCdest & "\" & $roFCSourceFile), "R") <> 0 Then ; Read-only flag set on destination?
					$roFC = 1
					FileSetAttrib($roFCdest & "\" & $roFCSourceFile, "-R")
				EndIf

				FileMove($roFCdest & "\" & $roFCSourceFile, $roFCdest & "\" & $roFCFileName & "_bu" & $ScriptRunTime & $roFCFileExt)
				If $roFC = 1 Then
					FileSetAttrib($roFCdest & "\" & $roFCFileName & "_bu" & $ScriptRunTime & $roFCFileExt, "+R")
				EndIf

				; Purge all but the last 3 backups
				$roFCh = FileFindFirstFile($roFCdest & "\" & $roFCFileName & "_bu*" & $roFCFileExt)
				If @error <> 1 Then
					While 1
						ReDim $roFCbu[$roFCbu[0] + 2]
						$roFCbu[0] += 1
						$roFCbu[$roFCbu[0]] = FileFindNextFile($roFCh)
						If @error = 1 Then
							$roFCbu[0] -= 1
							ExitLoop
						EndIf
					WEnd
					If $roFCbu[0] > 3 Then
						_ArraySort($roFCbu, 1, 1)
						For $roFCbuNext = 4 To $roFCbu[0]
							FileSetAttrib($roFCdest & "\" & $roFCbu[$roFCbuNext], "-R")
							FileDelete($roFCdest & "\" & $roFCbu[$roFCbuNext])
						Next
					EndIf
				EndIf
				FileClose($roFCh)

			EndIf

		EndIf
	EndIf

	; Copy the file
	$roFCResult = FileCopy($roFCSource, $roFCdest, $roFCflag)

	; Return the results
	Return $roFCResult

EndFunc   ;==>roFileCopy

