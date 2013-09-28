'CheckforRunningProcess.vbs - checks if a process is running on a specified machine.

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

function CheckForRunningProcess(strProc,strCompname)

	'Returns true if the specified process is running - othewise returns false
	'strProc - name of process to check
	'strCompname - name of computer to check for running process on.
	
	dim objWMIService
	dim objProcess
	CheckForRunningProcess = False
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strCompname & "\root\cimv2")

	for each objProcess in objWMIService.InstancesOf ("Win32_Process")
		
		If objProcess.Name = strProc then
			CheckForRunningProcess = True
		End If
	next
	
end function
