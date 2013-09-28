'AppInstall.vbs - installs application with cmdline switches. Intelligently runs updated installers on network-share.

'Copyright (C) 2013  Brenton Keegan

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

'REQUIRES: TruncatebyCharacter (available in this repository)
'REQUIRES: WriteEventLogEntry (available in this repository)
'REQUIRES: CheckforRunningProcess (available in this repository)


Option Explicit

'This script will launch any exe installer that can be run with command line switches. The goal of this script is to perform automated application installs. This is to be used in conjunction with group policy as a startup script to install applications under the SYSTEM account.
'This script assumes the installer will automatically remove previous versions of the application. If previous versions need to be removed manually it is recommend that a custom script be made.
'When the installer finishes it will set the modified date of the main exe to match that of the installer on the network share.
'To determine whether or not an application needs to be updated it will compare the modified date on the local app exe and the install exe. If the install exe modified date is different than it means the installer has not run successfuly and will attempt to update.


'strInstallerPath: UNC path of where the installer exe is located. End path with backslash (\)
'strAppName: name of application - will appear in event logs
'strAppExePath: Path of executable once installed. Note: this is just used to check to see if it's already installed. This will not affect where the installer installs the application to. DO NOT USE "Program Files (x86)" paths - path will be modified in the script if it's a 32 bit application being installed on a 64 bit machine - just make sure to set the Isx64 bit boolean coorectly.
'bolisx64: boolean - whether or not application is native 64-bit. 
'bolReinstall: Forces a reinstall of the application even if it's found that it's already installed.
'strSwitches: Install switches to perform automated install
'strOrg: name or initials of organization - will appear in logs


'stores commandline args
dim args 
Set args = WScript.Arguments

'cbool is needed because passing in "true" and "false" in at the commandline will go in as a string
If args.count = 7 then
	InstallApp args.Item(0), args.Item(1), args.Item(2), cbool(args.Item(3)), cbool(args.Item(4)), args.Item(5), args.Item(6)
else
	WriteEventLogEntry "APPLICATION", "ERROR", "Custom Script: " & Wscript.ScriptName, 1000, "Incorrect number of parameters"
end if
'Note: when entering arguements thru commandline (or gpo) do not use commas, only a single space	
'ex: RunInstaller "\\burlington.vtoxford.org\corp\Active Directory\Software Installs\WinSCP\installers\", "WinSCP", "C:\Program Files\WinSCP\WinSCP.exe", false, true, "/verysilent"
function InstallApp(strInstallerPath,strAppName,strAppExePath,bolisx64,bolReinstall,strSwitches,strOrg)

	on error resume next
	WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, strAppName & "  Installation/Update Process Begun"

	dim objFSO, objFolder, colFiles, objFile
	dim strComputer
	dim objshell, objAppShell
	dim strInstallerEXE
	dim LastDateModified
	dim bolInstallRun
	dim installreturncode
	dim objLocalExe
	dim objInstallerExe
	dim strLocalPathFolder
	dim strMainExeName
	dim bolInstallerRunning
	
	bolInstallRun = false
	LastDateModified = 0 
	strComputer = "."
	set objShell = CreateObject("WScript.Shell")
	set objShell = CreateObject("WScript.Shell")
	set objAppShell = CreateObject("Shell.Application")
	set objFSO = CreateObject("Scripting.FileSystemObject")
	set objFolder = objFSO.GetFolder(strInstallerPath)
	set colFiles = objFolder.Files
	
	'this loops through each file it finds in the installer directory and will ulatimately select the one that was most recently modified. This should in theory select the most current installer. If you wish to avoid errors you can include only one installer in the target directory
	for each objFile In colFiles
		if objfile.DateLastModified  > LastDateModified then
			LastDateModified = objfile.DateLastModified
			strInstallerEXE = objFile.Name
		end if
	next

	set objInstallerExe = objFSO.GetFile(strInstallerPath & strInstallerEXE)

	'Checks if OS is x64 by checking the existence of the program files (x86) directory (cheapest way to check). If OS is 64 but and the application is 32 bit it will modify the local install path accordingly (used to check if app is already installed)
	if objFSO.FolderExists("c:\Program Files (x86)") = true then
		WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, "Installing on a 64-bit system"
		if bolisx64 = false then
			WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " "& strAppName & " Installation", 1000, "Application is 32 bit - changing path to Program Files (x86)"
			strAppExePath = Replace(strAppExePath,"Program Files", "Program Files (x86)")				
		end if
		
		set objLocalExe = objFSO.GetFile(strAppExePath)
		strMainExeName = truncatebycharacter(strAppExePath, "\", "right", true ,1)
		strLocalPathFolder = truncatebycharacter(strAppExePath, "\", "right", false ,0)
	
	end if
	
	if err.number <> 0 then
		WriteEventLogEntry "APPLICATION", "ERROR", strOrg & " " & strAppName & " Installation", 1000, "Error:" & err.number & " " & err.description
	end if
	
	if objFSO.FileExists(strAppExePath) then
		if bolReinstall = true then
			WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, "Force Reinstall option set - Installing " & strAppName
			installreturncode = RunInstaller(strInstallerPath,strInstallerEXE,strSwitches,strOrg & " " & strAppName)
			bolInstallRun = true
			bolInstallerRunning = true
		else
			if objLocalExe.DateLastModified <> objInstallerExe.DateLastModified then
				WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, "Newer version available on server - Installing " & strAppName
				installreturncode = RunInstaller(strInstallerPath,strInstallerEXE,strSwitches,strOrg & " " & strAppName)
				bolInstallRun = true
				bolInstallerRunning = true
			else
				WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000,  strAppName & " Already Installed. Force reinstall not set and no new version"
			end if
		end if
	else
		WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, "No previous installation found - installing " & strAppName
		installreturncode = RunInstaller(strInstallerPath,strInstallerEXE,strSwitches,strOrg & " " & strAppName)
		bolInstallRun = true
		bolInstallerRunning = true
	end if
	
	if bolInstallRun = true then
		'check if original installer is still running
		Do while bolInstallerRunning = True
			bolInstallerRunning = CheckForRunningProcess(strMainExeName,strComputer)
		Loop
	
		If bolInstallerRunning = False then
			WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, "Setting modified date on file: " & strLocalPathFolder & strMainExeName
			objAppShell.NameSpace(strLocalPathFolder).ParseName(strMainExeName).ModifyDate = objInstallerExe.DateLastModified
			if err.number <> 0 then
				WriteEventLogEntry "APPLICATION", "ERROR", strOrg & " " & strAppName & " Installation", 1000, "Unable to set modified date - error:" & err.number & " " & err.description 
			end if
		end if
	
		on error goto 0
		WriteEventLogEntry "APPLICATION", "INFORMATION", strOrg & " " & strAppName & " Installation", 1000, "Installer return code:"  & installreturncode
	end if
end function
'====================================================================================================================
function RunInstaller(Installs,InstallerEXE,Switches,strLogName)
	
	dim objshell
	dim installcmd
	set objShell = CreateObject("WScript.Shell")

	installcmd = chr(34) & Installs & InstallerEXE & chr(34) & " " & Switches
	
	on error resume next
	WriteEventLogEntry "APPLICATION", "INFORMATION", strLogName & " Installation", 1000, "Started Installation"
	RunInstaller = objShell.Run(installcmd,1, true)
	WriteEventLogEntry "APPLICATION", "INFORMATION", strLogName & " Installation", 1000, "Ended Installation"
	if err.number <> 0 then
		WriteEventLogEntry "APPLICATION", "ERROR", strLogName & " Installation", 1000, "Error running installer. Error # " & Err.number
	end if
	on error goto 0
	
end function
