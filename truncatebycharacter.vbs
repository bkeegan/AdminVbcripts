'truncatebycharacter.vbs - chops off a text string by a specified character and returns either side of the string including (or not including) the cut-off character. 
'Designed to extract filenames from full paths or exclude filename in full path

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


function truncatebycharacter(ByVal strString,byVal strCharacter,byVal strSide,byVal bolResult,byVal intIncludeChr)
'this function returns a truncated string based on strString. It can truncate from either side, based on strCharacter
'The return value can either be characters before or after strCharacter. It can configured to include or disinclude strCharacter in the return value
'Parameters:
'	-strString: Path of a file, can be \\computer\share\file.ext or c:\directory\file.ext
'	-strCharacter: which character to truncate by
'	-strSide: can be set either "left" or "right". Determines which side of the string to start from
'	-bolResult: True = all characters before strCharacter from strSide, False = all characters after strCharacter from strSide
'	-intIncludeChr: Include the character defined in strCharacter in the return value. 0 = Include, 1 = do not include

	dim intTotalPathLength
	dim intChrLocation
	
	intTotalPathLength = len(strString)
	intChrLocation = 0 
	
	if strSide = "left" then
		intChrLocation = inStr(strString,strCharacter)
		'takes all characters before the first truncate character found from the left side.
		if bolResult = True then
			truncatebycharacter = left(strString,intChrLocation-intIncludeChr) 
		'takes all characters after the first truncate character found from the left side.
		elseif bolresult = False then
			truncatebycharacter = mid(strString,intChrLocation+intIncludechr)
		end if
		exit function
	elseif strSide = "right" then
		intChrLocation = intTotalPathLength - inStrRev(strString,strCharacter)
		'takes all characters after the first truncate character found from the right side.
		if bolResult = True then
			truncatebycharacter = right(strString,intChrLocation-intIncludeChr+1) 
		'takes all characters before the first truncate character found from the right side.
		elseif bolresult = False then
			truncatebycharacter = left(strString,intTotalPathLength-intChrLocation-intIncludeChr) 
		end if
		exit function
	end if
	
end function
