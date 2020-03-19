'Spreadsheet Vuln Parser for CB_feeds_dump csv output
'requires Microsoft Excel
'v1.5

'Copyright (c) 2019 Ryan Boyle randomrhythm@rhythmengineering.com.

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
Dim unsupportedTotal: unsupportedTotal = 0
Dim outdatedTotal: outdatedTotal = 0
Dim objExcel
Dim objWorkbook
Dim dictHostExclusion: Set dictHostInclusion = CreateObject("Scripting.Dictionary")'used to exclude hosts from reporting such as ones that have been decommissioned.
Dim dictOutdated: Set dictOutdated = CreateObject("Scripting.Dictionary")'
Dim dictUnsupported: Set dictUnsupported = CreateObject("Scripting.Dictionary")'
Dim DictUpdated: Set DictUpdated = CreateObject("Scripting.Dictionary")'
Dim DictVersion: Set DictVersion = CreateObject("Scripting.Dictionary")'

CurrentDirectory = GetFilePath(wscript.ScriptFullName)

'config
strPathToIncludedHosts = CurrentDirectory & "\includedHosts.txt"
'end config


'set inital values
intTabCounter = 1
intWriteRowCounter = 1


strDebugPath = CurrentDirectory & "\Debug"
wscript.echo "Please open the vuln CSV report"
OpenFilePath1 = SelectFile( )


Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")


'create sub directories 
if objFSO.folderexists(strDebugPath) = False then _
objFSO.createfolder(strDebugPath)




LoadList strPathToIncludedHosts, dictHostInclusion
if dictHostInclusion.count > 0 then
	msgbox "Hosts listed in " & strPathToIncludedHosts & " will be reported on. All other hosts will be excluded from reporting"
end if

Set objExcel = CreateObject("Excel.Application")
OpenFilePath1 = OpenFilePath1
Set objWorkbook = objExcel.Workbooks.Open _
    (OpenFilePath1)
    objExcel.Visible = True
mycolumncounter = 1
Do Until objExcel.Cells(1,mycolumncounter).Value = ""
  if objExcel.Cells(1,mycolumncounter).Value = "MD5" then int_MD5_Location = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Path" then int_path_Location = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Publisher" then intfileHashLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Company" then intsnameLocation = mycolumncounter'
  if objExcel.Cells(1,mycolumncounter).Value = "Product" then intalertTypeLocation = mycolumncounter'
  if objExcel.Cells(1,mycolumncounter).Value = "CB Prevalence" then intactionLocation = mycolumncounter'
  if objExcel.Cells(1,mycolumncounter).Value = "Logical Size" then intoccurredLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Host Name" then int_hostname_location = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Info Link" then intcncportLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Alliance Score" then intchannelLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Parent Name" then intheaderLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Command Line" then intobjurlLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "ID GUID" then intSevLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Child Count" then intosinfoLocation = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "Version" then int_version_location = mycolumncounter '
  if objExcel.Cells(1,mycolumncounter).Value = "64-bit" then intSMTPFromLocation = mycolumncounter'
  if objExcel.Cells(1,mycolumncounter).Value = "Vuln" then int_vuln_location = mycolumncounter'
  
  mycolumncounter = mycolumncounter +1
loop

if int_vuln_location = "" then
	msgbox "Error! Unable to identify the Vuln column"
	objExcel.quit
	wscript.quit(2)
end if


intRowCounter = 2
strTmpvalue = objExcel.Cells(intRowCounter,int_path_Location).Value
strTmpvalue = lcase(strTmpvalue)
if instr(strTmpvalue, "iexplore.exe") > 0  or instr(strTmpvalue, "internet explorer") then
  strProduct = "Internet Explorer"
  strVulnType = "Outdated " 
  strPatched = "Up to Date "
  strVulnDetail = " Version"
  strPatchDetail = " Version"
  strChatText = " Version Support"
  boolJustMajorVersion = True
elseif instr(strTmpvalue, "macromed") > 0  or instr(strTmpvalue, "flash") then
  strProduct = "Flash Player"
  strVulnType = "Outdated " 
  strPatched = "Up to Date "
  strVulnDetail = " Version"
  strPatchDetail = " Version"
  strChatText = " Version Support"
elseif instr(strTmpvalue, "mshtml.dll") > 0 then
  strProduct = "MS15-065 KB3065822"  
  strVulnType = "Patch " 
  strPatched = "Patch "
  strVulnDetail = " not Applied"
  strPatchDetail = " Applied"
  strChatText = " Patched"
elseif instr(strTmpvalue, "netapi32.dll") > 0 then
  strProduct = "MS08-067"  
  strVulnType = "Patch " 
  strPatched = "Patch "
  strVulnDetail = " not Applied"
  strPatchDetail = " Applied"
  strChatText = " Patched"
elseif instr(strTmpvalue, "vbscript.dll") > 0 then
  strProduct = "MS16-051 KB3155533"  
  strVulnType = "Patch " 
  strPatched = "Patch "
  strVulnDetail = " not Applied"
  strPatchDetail = " Applied"
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
  strPatchDetail = " Update Applied"
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
    if strCompName <> "" and (dictHostInclusion.count = 0 or dictHostInclusion.exists(lcase(strCompName)) = True) then 'only report on included hosts
		strTmpCompName = ucase(strCompName)
      

        if strTmpVulnInfo = "unsupported " & strProduct & " major version detected" or _ 
        instr(strTmpVulnInfo," not receive publicly released security updates") > 0 then
          if dictUnsupported.exists(strTmpCompName) = false then 
            dictUnsupported.add strTmpCompName, strTmpVersionNumber
            UpdateVersionDict strTmpVersionNumber
          end if 
        end if
        if strTmpVulnInfo = "outdated " & strProduct & " version detected" or instr(strTmpVulnInfo, "not applied") > 0 or _
        instr(strTmpVulnInfo, "missing patch") or instr(strTmpVulnInfo, "Silverlight flaw") then
          if dictOutdated.exists(strTmpCompName) = false then 
            dictOutdated.add strTmpCompName, strTmpVersionNumber
            UpdateVersionDict strTmpVersionNumber
          end if
        elseif strTmpVulnInfo = "up to date " & strProduct & " detected" or strTmpVulnInfo = "IE on a supported version" or _
        instr(strTmpVulnInfo, " applied") > 0 or instr(strTmpVulnInfo, " patched with") > 0 or _
		instr(strTmpVulnInfo, " patched for ")		then
          if DictUpdated.exists(strTmpCompName) = false then 
            DictUpdated.add strTmpCompName, strTmpVersionNumber         
            UpdateVersionDict strTmpVersionNumber
          end if
         end if 
    end if
  next

  intRowCounter = intRowCounter +1
loop
FixUpHeader
intRowCounter = 1
if dictUnsupported.count > 0 then
  Move_next_Workbook_Worksheet( "Unsupported")
  Write_Spreadsheet_line "Unsupported " & strProduct & " Computer|Version Number"
  FixUpHeader
  for each strCompName in dictUnsupported
	if DictUpdated.exists(strCompName) = False then
		Write_Spreadsheet_line strCompName & "|" & dictUnsupported.item(strCompName)
		unsupportedTotal = unsupportedTotal + 1
	end if
  next
end if

intRowCounter = 1
if dictOutdated.count > 0 then
  Move_next_Workbook_Worksheet("Outdated")
  Write_Spreadsheet_line strVulnType & strProduct & strVulnDetail & " Computer|Version Number"
  FixUpHeader
  for each strCompName in dictOutdated 
	if DictUpdated.exists(strCompName) = False then
		Write_Spreadsheet_line strCompName & "|" & dictOutdated.item(strCompName)
		outdatedTotal = outdatedTotal + 1
	end if
  next
end if
intRowCounter = 1
Move_next_Workbook_Worksheet("Up to Date")
Write_Spreadsheet_line strPatched & strProduct & strPatchDetail & " Computer|Version Number"
FixUpHeader
for each strCompName in DictUpdated
  Write_Spreadsheet_line strCompName & "|" & DictUpdated.item(strCompName)
next

Move_next_Workbook_Worksheet("Support Chart")
Write_Spreadsheet_line strProduct & strChatText & "|" & "Count"
FixUpHeader
if dictUnsupported.count > 0 then Write_Spreadsheet_line "Unsupported|" &  unsupportedTotal
if dictOutdated.count > 0 then Write_Spreadsheet_line "Outdated|" &  outdatedTotal
Write_Spreadsheet_line "Updated|" &  DictUpdated.count

Move_next_Workbook_Worksheet("Version Chart")
Write_Spreadsheet_line  strProduct & " Versions" & "|" & "Count"
FixUpHeader
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

Sub FixUpHeader() 'https://www.experts-exchange.com/questions/23820327/Freeze-Panes-through-VBS-Script.html
With objExcel.ActiveSheet
	.Rows(1).Font.Bold = True '1.  Bold the headers (always in row 1)
	.AutoFilterMode = False 'turn off any existing autofilter just in case
	.Rows(1).AutoFilter '2. Turn on AutoFilter for all coloms
	.Columns.AutoFit '3. Set Column width to AutoFit Selection
	'4. Set a freeze under column 1 so that the header is always present at the top
	.Range("A2").Select
End With
objExcel.ActiveWindow.FreezePanes = True
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



Sub LoadList(strListPath, dictToLoad)
if objFSO.fileexists(strListPath) then
  Set objFile = objFSO.OpenTextFile(strListPath)
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        strData = objFile.ReadLine
          if dictToLoad.exists(lcase(strData)) = False then 
			dictToLoad.add lcase(strData), ""
		end if
        on error goto 0
    end if
  loop
end if
end sub
