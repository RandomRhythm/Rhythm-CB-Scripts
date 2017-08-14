'Spreadsheet OS Parser for CB_Sensor_Dump csv output
'requires Microsoft Excel
'v1.1 Support for reporting on MS17-010 KB4013389

'Copyright (c) 2017 Ryan Boyle randomrhythm@rhythmengineering.com.
'All rights reserved.

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

Const forwriting = 241
Const ForAppending = 8
Const ForReading = 1
Dim intTabCounter
Dim boolJustMajorVersion : boolJustMajorVersion = False
'set inital values

intTabCounter = 1
intWriteRowCounter = 1



CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strCachePath = CurrentDirectory & "\cache"


strDebugPath = CurrentDirectory & "\Debug"
wscript.echo "Please open the vuln CSV report"
OpenFilePath1 = SelectFile( )


Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")


'create sub directories 
if objFSO.folderexists(strDebugPath) = False then _
objFSO.createfolder(strDebugPath)
if objFSO.folderexists(strCachePath) = False then _
objFSO.createfolder(strCachePath)


Dim dictOutdated: Set dictOutdated = CreateObject("Scripting.Dictionary")'
Dim dictUnsupported: Set dictUnsupported = CreateObject("Scripting.Dictionary")'
Dim DictUpdated: Set DictUpdated = CreateObject("Scripting.Dictionary")'
Dim DictVersion: Set DictVersion = CreateObject("Scripting.Dictionary")'
Set objExcel = CreateObject("Excel.Application")
OpenFilePath1 = OpenFilePath1
Set objWorkbook = objExcel.Workbooks.Open _
    (OpenFilePath1)
    objExcel.Visible = True
mycolumncounter = 1
Do Until objExcel.Cells(1,mycolumncounter).Value = ""
  if objExcel.Cells(1,mycolumncounter).Value = "MD5" then int_MD5_Location = mycolumncounter 'Source IP
  if objExcel.Cells(1,mycolumncounter).Value = "Path" then int_path_Location = mycolumncounter 'Host Name of source
  if objExcel.Cells(1,mycolumncounter).Value = "Publisher" then intfileHashLocation = mycolumncounter 'File Hash
  if objExcel.Cells(1,mycolumncounter).Value = "Company" then intsnameLocation = mycolumncounter'Detection Name
  if objExcel.Cells(1,mycolumncounter).Value = "Product" then intalertTypeLocation = mycolumncounter'Alert Type
  if objExcel.Cells(1,mycolumncounter).Value = "CB Prevalence" then intactionLocation = mycolumncounter'Action taken (blocked, notify)
  if objExcel.Cells(1,mycolumncounter).Value = "Logical Size" then intoccurredLocation = mycolumncounter 'Time stamp
  if objExcel.Cells(1,mycolumncounter).Value = "Host Name" then int_hostname_location = mycolumncounter 'C&C IP address
  if objExcel.Cells(1,mycolumncounter).Value = "Info Link" then intcncportLocation = mycolumncounter 'C&C port number
  if objExcel.Cells(1,mycolumncounter).Value = "Alliance Score" then intchannelLocation = mycolumncounter 'communication
  if objExcel.Cells(1,mycolumncounter).Value = "Parent Name" then intheaderLocation = mycolumncounter 'header
  if objExcel.Cells(1,mycolumncounter).Value = "Command Line" then intobjurlLocation = mycolumncounter 'objurl
  if objExcel.Cells(1,mycolumncounter).Value = "ID GUID" then intSevLocation = mycolumncounter 'Severity
  if objExcel.Cells(1,mycolumncounter).Value = "Child Count" then intosinfoLocation = mycolumncounter 'osinfo
  if objExcel.Cells(1,mycolumncounter).Value = "Version" then int_version_location = mycolumncounter 'smtp-to
  if objExcel.Cells(1,mycolumncounter).Value = "64-bit" then intSMTPFromLocation = mycolumncounter'smtp-mail-from
  if objExcel.Cells(1,mycolumncounter).Value = "Vuln" then int_vuln_location = mycolumncounter'subject
  
  mycolumncounter = mycolumncounter +1
loop
If BoolSMTPAlert = True then
  int_scrIPAddressLocation = intSMTPTOLocation

elseif BoolHostFilter = True then
  int_scrIPAddressLocation = intshostLocation
end if

intRowCounter = 2
strTmpvalue = objExcel.Cells(intRowCounter,int_path_Location).Value
strTmpvalue = lcase(strTmpvalue)
if instr(strTmpvalue, "iexplore.exe") > 0  or instr(strTmpvalue, "internet explorer") then
  strProduct = "Internet Explorer"
  strVulnType = "Outdated " 
  strPatched = "Up to date "
  strVulnDetail = " version"
  strPatchDetail = " version"
  strChatText = " Version Support"
  boolJustMajorVersion = True
elseif instr(strTmpvalue, "macromed") > 0  or instr(strTmpvalue, "flash") then
  strProduct = "Flash Player"
  strVulnType = "Outdated " 
  strPatched = "Up to date "
  strVulnDetail = " version"
  strPatchDetail = " version"
  strChatText = " Version Support"
elseif instr(strTmpvalue, "mshtml.dll") > 0 then
  strProduct = "MS15-065 KB3065822"  
  strVulnType = "Patch " 
  strPatched = "Patch "
  strVulnDetail = " not applied"
  strPatchDetail = " applied"
  strChatText = " Patched"
elseif instr(strTmpvalue, "netapi32.dll") > 0 then
  strProduct = "MS08-067"  
  strVulnType = "Patch " 
  strPatched = "Patch "
  strVulnDetail = " not applied"
  strPatchDetail = " applied"
  strChatText = " Patched"
elseif instr(strTmpvalue, "vbscript.dll") > 0 then
  strProduct = "MS16-051 KB3155533"  
  strVulnType = "Patch " 
  strPatched = "Patch "
  strVulnDetail = " not applied"
  strPatchDetail = " applied"
  strChatText = " Patched"
elseif instr(strTmpvalue, "silverlight") > 0 then
  strProduct = "Silverlight"  
  strVulnType = "Vulnerable " 
  strPatched = ""
  strVulnDetail = ""
  strPatchDetail = " Update Applied"
  strChatText = " Patched"
elseif instr(strTmpvalue, "srv.sys") > 0 then
  strProduct = "Windows SMB Server"  
  strVulnType = "Vulnerable " 
  strPatched = ""
  strVulnDetail = ""
  strPatchDetail = " update applied"
  strChatText = " Patched"
  else
  strProduct = inputbox("enter product name")
end if  
Do Until objExcel.Cells(intRowCounter,1).Value = "" 'loop till you hit null value (end of rows)
  strTmpVulnInfo = objExcel.Cells(intRowCounter,int_vuln_location).Value
  strTmpCompNames = objExcel.Cells(intRowCounter,int_hostname_location).Value
  if instr(strTmpCompNames, "/") = 0 then strTmpCompNames = strTmpCompNames & "/"
  arrayComNames = split(strTmpCompNames, "/")
  strTmpVersionNumber = objExcel.Cells(intRowCounter,int_version_location).Value
  for each strCompName in arrayComNames
    if strCompName <> "" then
      

        if strTmpVulnInfo = "unsupported " & strProduct & " major version detected" or _ 
        instr(strTmpVulnInfo," not receive publicly released security updates") > 0 then
          if dictUnsupported.exists(strCompName) = false then 
            dictUnsupported.add strCompName, strTmpVersionNumber
            UpdateVersionDict strTmpVersionNumber
          end if 
        end if
        if strTmpVulnInfo = "outdated " & strProduct & " version detected" or instr(strTmpVulnInfo, "not applied") > 0 or _
        instr(strTmpVulnInfo, "missing patch") or instr(strTmpVulnInfo, "Silverlight flaw") then
          if dictOutdated.exists(strCompName) = false then 
            dictOutdated.add strCompName, strTmpVersionNumber
            UpdateVersionDict strTmpVersionNumber
          end if
        elseif strTmpVulnInfo = "up to date " & strProduct & " detected" or strTmpVulnInfo = "IE on a supported version" or _
        instr(strTmpVulnInfo, " applied") > 0 or instr(strTmpVulnInfo, " patched with") > 0 or _
		instr(strTmpVulnInfo, " patched for ")		then
          if DictUpdated.exists(strCompName) = false then 
            DictUpdated.add strCompName, strTmpVersionNumber         
            UpdateVersionDict strTmpVersionNumber
          end if
         end if 
    end if
  next

  intRowCounter = intRowCounter +1
loop
intRowCounter = 1
if dictUnsupported.count > 0 then
  Move_next_Workbook_Worksheet( "Unsupported")
  Write_Spreadsheet_line "Unsupported " & strProduct & " major version|Version Number"
  for each strCompName in dictUnsupported
    Write_Spreadsheet_line strCompName & "|" & dictUnsupported.item(strCompName)
  next
end if
intRowCounter = 1
if dictOutdated.count > 0 then
  Move_next_Workbook_Worksheet("Outdated")
  Write_Spreadsheet_line strVulnType & strProduct & strVulnDetail & "|Version Number"
  for each strCompName in dictOutdated 
    Write_Spreadsheet_line strCompName & "|" & dictOutdated.item(strCompName)
  next
end if
intRowCounter = 1
Move_next_Workbook_Worksheet("Up to Date")
Write_Spreadsheet_line strPatched & strProduct & strPatchDetail & "|Version Number"
for each strCompName in DictUpdated
  Write_Spreadsheet_line strCompName & "|" & DictUpdated.item(strCompName)
next

Move_next_Workbook_Worksheet("Support Chart")
Write_Spreadsheet_line strProduct & strChatText & "|" & "Count"
if dictUnsupported.count > 0 then Write_Spreadsheet_line "Unsupported|" &  dictUnsupported.count
if dictOutdated.count > 0 then Write_Spreadsheet_line "Outdated|" &  dictOutdated.count
Write_Spreadsheet_line "Updated|" &  DictUpdated.count

Move_next_Workbook_Worksheet("Version Chart")
Write_Spreadsheet_line  strProduct & "Versions" & "|" & "Count"
for each strVersionNumber in DictVersion
  Write_Spreadsheet_line strVersionNumber & "|" &  DictVersion.item(strVersionNumber)
next


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



