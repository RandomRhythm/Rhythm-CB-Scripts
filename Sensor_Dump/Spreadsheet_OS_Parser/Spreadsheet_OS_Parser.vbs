'Spreadsheet OS Parser for CB_Sensor_Dump csv output
'requires Microsoft Excel
'v2.5 - Identify CentOS as Linux even if it does not mention Linux.

'Copyright (c) 2022 Ryan Boyle randomrhythm@rhythmengineering.com.

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'Spreadsheet OS parser for CB feeds csv output
'requires Microsoft Excel

Const forwriting = 241
Const ForAppending = 8
Const ForReading = 1
Dim intTabCounter
Dim boolJustMajorVersion : boolJustMajorVersion = False
Dim boolUniqueOnly: boolUniqueOnly = True 'only report on a unique computer name once.
'set inital values

intTabCounter = 1
intWriteRowCounter = 1



CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strCachePath = CurrentDirectory & "\cache"

strDebugPath = CurrentDirectory & "\Debug"
wscript.echo "Please open the sensor CSV report"
OpenFilePath1 = SelectFile( )


Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")


'create sub directories 
if objFSO.folderexists(strDebugPath) = False then _
objFSO.createfolder(strDebugPath)
if objFSO.folderexists(strCachePath) = False then _
objFSO.createfolder(strCachePath)

Dim DictCompName: Set DictCompName = CreateObject("Scripting.Dictionary")'
Dim DictOSWorkversion: Set DictOSWorkversion = CreateObject("Scripting.Dictionary")'
Dim DictOSServversion: Set DictOSServversion = CreateObject("Scripting.Dictionary")
Dim DictOSWorkversionMac: Set DictOSWorkversionMac = CreateObject("Scripting.Dictionary")'
Dim DictOSWorkversionWindows: Set DictOSWorkversionWindows = CreateObject("Scripting.Dictionary")'
Dim DictOSServversionWindows: Set DictOSServversionWindows = CreateObject("Scripting.Dictionary")
Dim DictOSServversionLinux: Set DictOSServversionLinux = CreateObject("Scripting.Dictionary")
Dim DictOSconsolidated: Set DictOSconsolidated = CreateObject("Scripting.Dictionary")
Set objExcel = CreateObject("Excel.Application")
OpenFilePath1 = OpenFilePath1
Set objWorkbook = objExcel.Workbooks.Open _
    (OpenFilePath1)
    objExcel.Visible = True
mycolumncounter = 1
Do Until objExcel.Cells(1,mycolumncounter).Value = ""
    
  if objExcel.Cells(1,mycolumncounter).Value = "Computer" or objExcel.Cells(1,mycolumncounter).Value = "computer_dns_name" or objExcel.Cells(1,mycolumncounter).Value = "name" or _
  objExcel.Cells(1,mycolumncounter).Value = "Device Name" then int_hostname_location = mycolumncounter
  if objExcel.Cells(1,mycolumncounter).Value = "Hostname" or objExcel.Cells(1,mycolumncounter).Value = "FQDN" then int_hostname_location = mycolumncounter
  if objExcel.Cells(1,mycolumncounter).Value = "Operating System" or objExcel.Cells(1,mycolumncounter).Value = "OS" or _
  objExcel.Cells(1,mycolumncounter).Value = "OS Platform" then int_vuln_location = mycolumncounter
  if objExcel.Cells(1,mycolumncounter).Value = "OS Version" or objExcel.Cells(1,mycolumncounter).Value = "OS version"  or objExcel.Cells(1,mycolumncounter).Value = "osVersion" then int_vuln_location = mycolumncounter
  if objExcel.Cells(1,mycolumncounter).Value = "os_environment_display_string" then int_vuln_location = mycolumncounter
  if objExcel.Cells(1,mycolumncounter).Value = "OS Distribution" then 
    int_vuln_location = mycolumncounter 'defender overwrite. OS Version can give only a build number which need further code to support
    exit do
  end if
  mycolumncounter = mycolumncounter +1
loop

if int_vuln_location = "" then
	msgbox "Error! Unable to identify the Operating System column"
	objExcel.quit
	wscript.quit(2)
end if

intRowCounter = 2
intNonEmpty = int_hostname_location 'need to point at a column that is always populated. 

Do Until objExcel.Cells(intRowCounter,intNonEmpty).Value = "" 'loop till you hit null value (end of rows)
  strCompName = objExcel.Cells(intRowCounter,int_hostname_location).Value
  if boolUniqueOnly = False or (boolUniqueOnly = True and DictCompName.exists(strCompName) = False) then 
	  strTmpVulnInfo = objExcel.Cells(intRowCounter,int_vuln_location).Value
	  if instr(strTmpVulnInfo, "Server") > 0 or instr(strTmpVulnInfo, "CentOS") > 0 or instr(strTmpVulnInfo, "Ubuntu") > 0 then
      if DictOSServversion.exists(strTmpVulnInfo) = False then 
        DictOSServversion.add strTmpVulnInfo, 1
        if Instr(strTmpVulnInfo, "Windows") > 0 and DictOSServversionWindows.exists(strTmpVulnInfo) = False then _
          DictOSServversionWindows.add strTmpVulnInfo, 1
        if (instr(strTmpVulnInfo, "Linux") > 0 or instr(strTmpVulnInfo, "Ubuntu") > 0 or instr(strTmpVulnInfo, "CentOS") > 0) and DictOSServversionLinux.exists(ShortenOSname(strTmpVulnInfo)) = False then _
          DictOSServversionLinux.add ShortenOSname(strTmpVulnInfo), 1
        
      else
        DictOSServversion.item(strTmpVulnInfo) = DictOSServversion.item(strTmpVulnInfo) + 1
        if Instr(strTmpVulnInfo, "Windows") > 0  then _
        DictOSServversionWindows.item(strTmpVulnInfo) = DictOSServversionWindows.item(strTmpVulnInfo) + 1
        if Instr(strTmpVulnInfo, "Linux") > 0  then _
        DictOSServversionLinux.item(ShortenOSname(strTmpVulnInfo)) = DictOSServversionLinux.item(ShortenOSname(strTmpVulnInfo)) + 1	 	  
      end if  
	  else 
		if DictOSWorkversion.exists(strTmpVulnInfo) = False then
			DictOSWorkversion.add strTmpVulnInfo, 1
			if (Instr(strTmpVulnInfo, "Mac") > 0 or Instr(strTmpVulnInfo, "mac") > 0) and DictOSWorkversionMac.exists(strTmpVulnInfo) = False then _
			  DictOSWorkversionMac.add strTmpVulnInfo, 1
			if instr(strTmpVulnInfo, "Windows") > 0 and DictOSWorkversionWindows.exists(strTmpVulnInfo) = False then _
			  DictOSWorkversionWindows.add strTmpVulnInfo, 1
		else
		  DictOSWorkversion.item(strTmpVulnInfo) = DictOSWorkversion.item(strTmpVulnInfo) + 1
		  if Instr(strTmpVulnInfo, "Mac") > 0  or  Instr(strTmpVulnInfo, "mac") > 0 then _
		  DictOSWorkversionMac.item(strTmpVulnInfo) = DictOSWorkversionMac.item(strTmpVulnInfo) + 1
		  if Instr(strTmpVulnInfo, "Windows") > 0  then _
		  DictOSWorkversionWindows.item(strTmpVulnInfo) = DictOSWorkversionWindows.item(strTmpVulnInfo) + 1	  
		end if
	  end if
	  if instr(strTmpVulnInfo, "OSX") > 0 or instr(strTmpVulnInfo, "macOS") > 0 then
		strConsolidated = "Mac OS X"
	  elseif instr(strTmpVulnInfo, "Linux") > 0 and (instr(strTmpVulnInfo, "release") > 0 Or instr(strTmpVulnInfo, "SUSE") > 0) and instr(strTmpVulnInfo, ".") > 0 then
		strConsolidated = ShortenOSname(strTmpVulnInfo)
	  elseif instr(strTmpVulnInfo, "2003") then
		strConsolidated = "Windows 2003"
	  elseif instr(strTmpVulnInfo, "2008") then
		strConsolidated = "Windows 2008"
	  elseif instr(strTmpVulnInfo, "2012") then
		strConsolidated = "Windows 2012"
	  elseif instr(strTmpVulnInfo, "2016") then
		strConsolidated = "Windows 2016"
	  elseif instr(strTmpVulnInfo, "2019") then
		strConsolidated = "Windows 2019"
	  elseif instr(strTmpVulnInfo, "Windows XP") then
		strConsolidated = "Windows XP"
	  elseif instr(strTmpVulnInfo, "Vista") then
		strConsolidated = "Windows Vista"
	  elseif instr(strTmpVulnInfo, "Windows 7") > 0 or instr(strTmpVulnInfo, "Windows7") > 0 or instr(strTmpVulnInfo, "6.1.7601") > 0 then
		strConsolidated = "Windows 7"
	  elseif instr(strTmpVulnInfo, "Windows 8.1") then
		strConsolidated = "Windows 8.1"
	  elseif instr(strTmpVulnInfo, "Windows 8") then
		strConsolidated = "Windows 8"
	  elseif instr(strTmpVulnInfo, "Windows 10") >0  or instr(strTmpVulnInfo, "Windows10") >0  or instr(strTmpVulnInfo, "10.0.18363") >0  then
		if instr(strTmpVulnInfo, "Server") then
			strConsolidated = "Windows 2016"
		else
			strConsolidated = "Windows 10"
		end if
	  end if
	  if DictOSconsolidated.exists(strConsolidated) = False then
		DictOSconsolidated.add strConsolidated, 1
	  else
		DictOSconsolidated.item(strConsolidated) = DictOSconsolidated.item(strConsolidated) + 1
	  end if

	DictCompName.add strCompName, ""
  end if
  
  intRowCounter = intRowCounter +1
loop

FixUpHeader
intRowCounter = 1
Move_next_Workbook_Worksheet( "Operating Systems")
Write_Spreadsheet_line "Operating Systems|Count"
FixUpHeader
if DictOSconsolidated.count > 0 then

  for each strOSname in DictOSconsolidated
    Write_Spreadsheet_line strOSname & "|" & DictOSconsolidated.item(strOSname)
  next
end if

intRowCounter = 1
  Move_next_Workbook_Worksheet( "OS Version")
  Write_Spreadsheet_line "OS Versions|Count"
  FixUpHeader
if DictOSWorkversion.count > 0 then

  for each strOSname in DictOSWorkversion
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSWorkversion.item(strOSname)
  next
end if
if DictOSServversion.count > 0 then
  for each strOSname in DictOSServversion
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSServversion.item(strOSname)
  next
end if
intRowCounter = 1
  Move_next_Workbook_Worksheet( "Workstation OS")
  Write_Spreadsheet_line "Workstation OS|Count"
  FixUpHeader
if DictOSWorkversion.count > 0 then
  for each strOSname in DictOSWorkversion
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSWorkversion.item(strOSname)
  next
end if

intRowCounter = 1
  Move_next_Workbook_Worksheet( "Mac OS")
  Write_Spreadsheet_line "Mac OS|Count"
  FixUpHeader
if DictOSWorkversionMac.count > 0 then
  for each strOSname in DictOSWorkversionMac
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSWorkversionMac.item(strOSname)
  next
end if

intRowCounter = 1
  Move_next_Workbook_Worksheet( "Windows Workstation OS")
  Write_Spreadsheet_line "Windows Workstation OS|Count"
  FixUpHeader
if DictOSWorkversionWindows.count > 0 then
  for each strOSname in DictOSWorkversionWindows
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSWorkversionWindows.item(strOSname)
  next
end if

intRowCounter = 1
  Move_next_Workbook_Worksheet( "Server OS")
  Write_Spreadsheet_line "Server OS|Count"
  FixUpHeader
if DictOSServversion.count > 0 then
  for each strOSname in DictOSServversion
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSServversion.item(strOSname)
  next
end if

intRowCounter = 1
  Move_next_Workbook_Worksheet( "Windows Server OS")
  Write_Spreadsheet_line "Windows Server|Count"
  FixUpHeader
if DictOSServversionWindows.count > 0 then
  for each strOSname in DictOSServversionWindows
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSServversionWindows.item(strOSname)
  next
end if


intRowCounter = 1
  Move_next_Workbook_Worksheet( "Linux OS")
  Write_Spreadsheet_line "Linux Server|Count"
  FixUpHeader
if DictOSServversionLinux.count > 0 then
  for each strOSname in DictOSServversionLinux
    Write_Spreadsheet_line ShortenOSname(strOSname) & "|" & DictOSServversionLinux.item(strOSname)
  next
end if


Function ShortenOSname(strOSname)
Dim strReturnShort
Dim boolServer
strReturnShort = strOSname
if instr(strReturnShort, "Linux") > 0 and instr(strReturnShort, "release") > 0 and instr(strReturnShort, ".") > 0 then
    strReturnShort = Left(strReturnShort, instr(strReturnShort, ".") + 1)
	strReturnShort = replace(strReturnShort, "release","")
	strReturnShort = replace(strReturnShort, "Red Hat Enterprise Linux Server","RHEL")

ElseIf instr(strReturnShort, "SUSE") > 0 Then
	strReturnShort = replace(strReturnShort, "SUSE Linux Enterprise Server","SUSE")
	if instr(strReturnShort, "SUSE") > 0 and instr(strReturnShort, "\n") > 0 then
    strReturnShort = left(strReturnShort, instr(strReturnShort, "\n") - 1)  
    End if
end if
boolServer = False
if instr(strReturnShort, "Server ") > 0 then
	boolServer = True
end if
strReturnShort = replace(strReturnShort, "Windows Server 2008 R2 Windows Storage Server 2008 R2", "Storage Server 2008 R2")
strReturnShort = replace(strReturnShort, "Windows Server ", "")
strReturnShort = replace(strReturnShort, "Windows ", "")
strReturnShort = replace(strReturnShort, "Server ", "")
strReturnShort = replace(strReturnShort, ",", "")
strReturnShort = replace(strReturnShort, "Service Pack ", "SP")
strReturnShort = replace(strReturnShort, "Standard", "STD")
strReturnShort = replace(strReturnShort, "Professional", "Pro")
strReturnShort = replace(strReturnShort, "Datacenter", "DCE")
strReturnShort = replace(strReturnShort, "Enterprise Edition", "EE")
strReturnShort = replace(strReturnShort, "Enterprise", "EE")
strReturnShort = replace(strReturnShort, "Edition", "")
strReturnShort = replace(strReturnShort, "without", "w/o")
if boolServer = True then
	strReturnShort = replace(strReturnShort, "10 STD", "2016 STD")
	strReturnShort = replace(strReturnShort, "10 EE", "2016 EE")
	strReturnShort = replace(strReturnShort, "10 DCE", "2016 DCE")
end if
strReturnShort = replace(strReturnShort, "Microsoft ", "")
strReturnShort = replace(strReturnShort, "(Evaluation)", "(Eval)")
if instr(strReturnShort, "Linux") > 0 and (instr(strReturnShort, "CentOS") > 0 or instr(strReturnShort, "RHEL") > 0 Or instr(strReturnShort, "SUSE") > 0) then
    strReturnShort = replace(strReturnShort, "Linux ","")
end if
ShortenOSname = strReturnShort
end function

Sub UpdateVersionDict(strVersionNumber)
if instr(strVersionNumber, " ") then 
  arrayVN = split(strVersionNumber, " ")
  strVN = arrayVN(0)
else
  strVN = strVersionNumber
end if
if boolJustMajorVersion = True then
  if instr(strVN, ".") then
    strVN = left(strVN, instr(strVN, ".")-1) 
  elseif instr(strVN, ",") then
    strVN = left(strVN, instr(strVN, ",")-1) 
  end if
end if
if DictVersion.exists(strVN) then 
  DictVersion.item(strVN) = DictVersion.item(strVN) + 1
else
  DictVersion.add strVN, 1
end if  
end sub
Function RemoveCharsForFname(TextFileName)
'Remove unsupported characters from file name
strTmpFilName1 = right(TextFileName, len(TextFileName) - instrrev(TextFileName,"\"))
strTmpFilName2 = replace(strTmpFilName1,"/",".")
'TextFileName = replace(TextFileName,"\",".")
strTmpFilName2 = replace(strTmpFilName2,":",".")
strTmpFilName2 = replace(strTmpFilName2,"*",".")
strTmpFilName2 = replace(strTmpFilName2,"?",".")
strTmpFilName2 = replace(strTmpFilName2,chr(34),".")
strTmpFilName2 = replace(strTmpFilName2,"<",".")
strTmpFilName2 = replace(strTmpFilName2,">",".")
strTmpFilName2 = replace(strTmpFilName2,"|",".")
TextFileName = replace(TextFileName,strTmpFilName1,strTmpFilName2)
'will error if file name is to long
If Len(TextFileName) > 255 Then TextFileName = Left(TextFileName, 255)
RemoveCharsForFname = TextFileName
end function

function LogData(TextFileName, TextToWrite,EchoOn)
Dim strTmpFilName1
Dim strTmpFilName2
TextFileName = RemoveCharsForFname(TextFileName)

Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      fsoLogData.CreateTextFile TextFileName, True
  End If
on error resume next
Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
if err.number <> 0 then
  msgbox "Error writting to " & TextFileName & " perhaps the file is locked?"
  err.number = 0
  Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
  if err.number <> 0 then exit function
end if

on error goto 0
WriteTextFile.WriteLine TextToWrite
WriteTextFile.Close
Set fsoLogData = Nothing
End Function

Function GetFilePath (ByVal FilePathName)
found = False

Z = 1

Do While found = False and Z < Len((FilePathName))

 Z = Z + 1

         If InStr(Right((FilePathName), Z), "\") <> 0 And found = False Then
          mytempdata = Left(FilePathName, Len(FilePathName) - Z)
          
             GetFilePath = mytempdata

             found = True

        End If      

Loop

end Function

Function GetData(contents, ByVal EndOfStringChar, ByVal MatchString)
MatchStringLength = Len(MatchString)
x= 0

do while x < len(contents) - (MatchStringLength +1)

  x = x + 1
  if Mid(contents, x, MatchStringLength) = MatchString then
    'Gets server name for section
    for y = 1 to len(contents) -x
      if instr(Mid(contents, x + MatchStringLength, y),EndOfStringChar) = 0 then
          TempData = Mid(contents, x + MatchStringLength, y)
        else
          exit do  
      end if
    next
  end if
loop
GetData = TempData
end Function






Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&ælig;-4ba3-bca5-ec349df65ef6

    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec = Nothing
    Set wshShell = Nothing
End Function



Function isIPaddress(strIPaddress)
DIm arrayTmpquad
Dim boolReturn_isIP
boolReturn_isIP = True
if instr(strIPaddress,".") then
  arrayTmpquad = split(strIPaddress,".")
  for each item in arrayTmpquad
    if isnumeric(item) = false then boolReturn_isIP = false
  next
else
  boolReturn_isIP = false
end if
isIPaddress = boolReturn_isIP
END FUNCTION

Function ReturnHostFromHeader(strtmpline)
if instr(strtmpline, "Host: ") then
  strtmpline = getdata(strtmpline, ":", "Host: ")
  if right(strtmpline,6) = "Accept" then strtmpline = left(strtmpline,len(strtmpline)-6)
  if right(strtmpline,10) = "Connection" then strtmpline = left(strtmpline,len(strtmpline)-10)
  ReturnHostFromHeader = strtmpline
end if
End Function

  


Function RemoveDups(strRMdupsData, strSplitChar)
Dim ArrayRemoveDups
Dim strReturnRemoveDups
if instr(strRMdupsData, strSplitChar) then
  ArrayRemoveDups = split(strRMdupsData, strSplitChar)
  Dim dicTmpRemoveDuplicate: Set dicTmpRemoveDuplicate = CreateObject("Scripting.Dictionary")
  for xRP = 0 to ubound(ArrayRemoveDups)
    if not dicTmpRemoveDuplicate.Exists(ArrayRemoveDups(xRP)) then _
    dicTmpRemoveDuplicate.Add ArrayRemoveDups(xRP), dicTmpRemoveDuplicate.Count
  next
  For Each Item In dicTmpRemoveDuplicate
    if strReturnRemoveDups =  "" Then
      strReturnRemoveDups = Item
    else
      strReturnRemoveDups = strReturnRemoveDups & strSplitChar & Item
    end if
  next
    RemoveDups = strReturnRemoveDups 

else
  RemoveDups = strRMdupsData
end if



End Function  


Function IsPrivateIP(strIP)
Dim boolReturnIsPrivIp
Dim ArrayOctet
boolReturnIsPrivIp = False
if isIPaddress(strIP) = False then
  IsPrivateIP = False
  exit function
end if
if left(strIP,3) = "10." then
  boolReturnIsPrivIp = True
elseif left(strIP,4) = "172." then
  ArrayOctet = split(strIP,".")
  if ArrayOctet(1) >15 and ArrayOctet(1) < 32 then
    boolReturnIsPrivIp = True
  end if
elseif left(strIP,7) = "192.168" then
  boolReturnIsPrivIp = True
end if
IsPrivateIP = boolReturnIsPrivIp
End Function




Sub Write_Spreadsheet_line(strSSrow)
Dim intColumnCounter
if instr(strSSrow,"|") then
  strSSrow = split(strSSrow, "|")
  for intColumnCounter = 1 to ubound(strSSrow) + 1
    objExcel.Cells(intWriteRowCounter, intColumnCounter).Value = strSSrow(intColumnCounter -1)
  next
else
    objExcel.Cells(intWriteRowCounter, 1).Value = strSSrow
end if
intWriteRowCounter = intWriteRowCounter + 1
end sub

Sub Add_Workbook_Worksheet(strWorksheetName)
Set objWorkbook = objExcel.Worksheets(objExcel.Worksheets.count)
objWorkbook.Activate

objExcel.ActiveWorkbook.Worksheets.Add
intWriteRowCounter = 1
Set objSheet1 = objExcel.Worksheets(objExcel.Worksheets(objExcel.Worksheets.count -1).name)
    Set objSheet2 = objExcel.Worksheets(objExcel.Worksheets(objExcel.Worksheets.count).name)
    objSheet2.Move objSheet1


objExcel.Worksheets(objExcel.Worksheets.count).Name = strWorksheetName
Set objWorkbook = objExcel.Worksheets(objExcel.Worksheets.count)
objWorkbook.Activate


end sub

Sub Move_next_Workbook_Worksheet(strWorksheetName)
intTabCounter = intTabCounter + 1
if objExcel.Worksheets.count < intTabCounter then
  Add_Workbook_Worksheet(strWorksheetName)
else
  Set objWorkbook = objExcel.Worksheets(intTabCounter)
  objWorkbook.Activate
  if strWorksheetName <> "" then objExcel.Worksheets(intTabCounter).Name = strWorksheetName
  intWriteRowCounter = 1
end if
end sub


Function ReturnPairedListfromDict(tmpDictionary)
Dim strTmpDictList
For Each Item In tmpDictionary
  if strTmpDictList = "" then 
  
    strTmpDictList = tmpDictionary.Item(Item) & " - " & Item
  else
    strTmpDictList = strTmpDictList & ", " & tmpDictionary.Item(Item) & " - " & Item
  end if

next

ReturnPairedListfromDict = strTmpDictList
End Function

Function ReturnListfromDict(tmpDictionary)
Dim strTmpDictList
For Each Item In tmpDictionary
  if strTmpDictList = "" then 
  
    strTmpDictList =  Item
  else
    strTmpDictList = strTmpDictList & vbCrLf & Item
  end if

next

ReturnListfromDict = strTmpDictList
End Function



Sub FixUpHeader() 'https://www.experts-exchange.com/questions/23820327/Freeze-Panes-through-VBS-Script.html
With objExcel.ActiveSheet
	.Rows(1).Font.Bold = True '1.  Bold the headers (always in row 1)
	.AutoFilterMode = False 'turn off any existing autofilter just in case
on error resume next
	.Rows(1).AutoFilter '2. Turn on AutoFilter for all coloms
  if err.number <> 0 then exit sub 'row is already autofiltered?
on error goto 0
	.Columns.AutoFit '3. Set Column width to AutoFit Selection
	'4. Set a freeze under column 1 so that the header is always present at the top
	.Range("A2").Select
End With
objExcel.ActiveWindow.FreezePanes = True
end sub