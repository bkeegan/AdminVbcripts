'WriteEventLogEntry.vbs - wrapper function around the windows createvent cmd. Requires elevated privilages. 

'Copyright (C) 2008  Brenton Keegan

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

function WriteEventLogEntry (ByVal strLog, ByVal strType, ByVal strSource,ByVal intID,ByVal strLogmessage)
'writes an event log entry using the 'eventcreate' command. Use of this command favored over shell.logmessage method because it allows for more control
'strLog: Log name to write to (APPLICATION, SYSTEM)
'strType: Type of log entry (SUCCESS, ERROR, WARNING, INFORMATION)
'strSource: "Source" item in event log info (E.g "VON Startup Script")
'intID: Numeric identifier (1-1000)
'strLogmessage: Actual contents of event log

	dim objShell
	dim strCommand
	
	set objShell = WScript.CreateObject( "WScript.Shell" )
	strCommand = "eventcreate /L " & strLog & " /T " & strType & " /SO " & chr(34) & strSource & chr(34) & " /ID " & intID & " /D " & chr(34) & strLogmessage & chr(34)
	
	objShell.Run strCommand

end function
