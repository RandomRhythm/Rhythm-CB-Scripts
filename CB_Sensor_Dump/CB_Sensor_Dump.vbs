'CB Sensor Dump
'This script will dump sensor information via the CB Response (Carbon Black) API

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


dim strCarBlackAPIKey
Dim intTotalQueries
Dim IntDaysQuery
Dim strStartDateQuery
Dim strEndDateQuery
Dim strIPquery
Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1
Dim DictIPAddresses: set DictIPAddresses = CreateObject("Scripting.Dictionary")'

'---Config Section
BoolDebugTrace = False 
IntDayStartQuery = "*" 'days to go back for start date of query. Example "-8". Set to "*" to query all binaries
IntDayEndQuery = "*" 'days to go back for end date of query. Example "-1". Set to "*" for no end date
strIPquery = "" 'Only dump information for sensors that held a particual IP adress. example: "10.10.10.80"
'---End Config section

if isnumeric(IntDayStartQuery) then
  strStartDateQuery = DateAdd("d",IntDayStartQuery,date)
  ' AND server_added_timestamp:[" & strStartDateQuery & "T00:00:00 TO "
  strStartDateQuery = " AND server_added_timestamp:[" & FormatDate (strStartDateQuery) & "T00:00:00 TO "
  if isnumeric(IntDayEndQuery) then
    strEndDateQuery = DateAdd("d",IntDayEndQuery,date)
    strEndDateQuery = FormatDate (strEndDateQuery) & "T00:00:00]"
  elseif IntDayEndQuery = "*" then
    IntDayEndQuery = "*]"
  end if
end if

if strIPquery <> "" then
  if isIPaddress(strIPquery) then
    strIPquery = "?ip=" & strIPquery
  else
    msgbox "Invalid IP address: " & strIPquery
    wscript.quit 998
  end if
end if

CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug\VT\"
if objFSO.folderexists(CurrentDirectory & "\debug") = false then objFSO.createfolder CurrentDirectory & "\debug"
if objFSO.folderexists(strDebugPath) = false then objFSO.createfolder strDebugPath
strSSfilePath = CurrentDirectory & "\CBSensor_" & udate(now) & ".csv"

strRandom = "4bv3nT9vrkJpj3QyueTvYFBMIvMOllyuKy3d401Fxaho6DQTbPafyVmfk8wj1bXF" 'encryption key. Change if you want but can only decrypt with same key
Set objFSO = CreateObject("Scripting.FileSystemObject")


strFile= CurrentDirectory & "\cb.dat"
strAPIproduct = "Carbon Black" 


strData = ""
StrBaseCBURL = ""
if objFSO.fileexists(strFile) then
  Set objFile = objFSO.OpenTextFile(strFile)
  if not objFile.AtEndOfStream then 'read file
      'On Error Resume Next
      strData = objFile.ReadLine 
      if not objFile.AtEndOfStream then StrBaseCBURL = objFile.ReadLine
      'on error goto 0
  end if
  if strData <> "" then
    strData = Decrypt(strData,strRandom)
      strTempAPIKey = strData
      strData = ""
  end if
end if
on error resume next
objFile.close
on error goto 0

if not objFSO.fileexists(strFile) and strData = "" then
  strTempAPIKey = inputbox("Enter your " & strAPIproduct & " api key")
  if strTempAPIKey <> "" then
    strTempEncryptedAPIKey = encrypt(strTempAPIKey,strRandom)
    logdata strFile,strTempEncryptedAPIKey,False
  end if
end if


if StrBaseCBURL = "" and strTempAPIKey <> "" then
    strTempEncryptedAPIKey = ""
    StrBaseCBURL = inputbox("Enter your " & strAPIproduct & " base URL (example: https://ryancb-example.my.carbonblack.io")
    if StrBaseCBURL <> "" then
      logdata strFile,StrBaseCBURL,False
    end if
end if  
if strTempAPIKey = "" then

    msgbox "invalid api key"
    wscript.quit(999)
end if
strCarBlackAPIKey = strTempAPIKey


if instr(lcase(StrBaseCBURL),".") <> 0 and instr(lcase(StrBaseCBURL),"http") <> 0 and instr(lcase(StrBaseCBURL),"://") <> 0 then
  if strCarBlackAPIKey <> "" and StrBaseCBURL <> "" then BoolUseCarbonBlack = True   
else
  msgbox "Invalid URL specified for Carbon Black: " & StrBaseCBURL & vbcrlf & "Delete the dat file to input new URL information: " & strFile
  StrBaseCBURL = "" 
  BoolUseCarbonBlack = False
end if

if BoolUseCarbonBlack = True then
  strTmpLogLine = chr(34) & "Computer|Operating System|Date Registered|Stored Bytes|Status|Health|Group ID|Last Checkin|Event Log Bytes|Days Reporting In|Computer Name" & Chr(34)
  strTmpLogLine = replace(strTmpLogLine, "|", chr(34) & "," & chr(34))
  LogData strSSfilePath, strTmpLogLine, false

  intTotalQueries = 10
  'loop through CB results
  intTotalQueries = DumpCarBlack()
  wscript.sleep 10
end if



Function DumpCarBlack()

Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Dim strAVEurl
Dim strReturnURL
dim strAssocWith
Dim strCBresponseText
Dim strtmpCB_Fpath

strAVEurl = StrBaseCBURL & "/api/v1/sensor" & strIPquery

objHTTP.open "GET", strAVEurl, False

objHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
  

on error resume next
  objHTTP.send 
  if err.number <> 0 then
    logdata CurrentDirectory & "\VT_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
    exit function 
  end if
on error goto 0  
'creates a lot of data. Don't enable debug logging on next line unless your going to disable it again
if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "_Sensor" & ".txt", objHTTP.responseText & vbcrlf & vbcrlf,BoolEchoLog 
strCBresponseText = objHTTP.responseText

strArrayCBresponse = split(strCBresponseText, vblf & "  {")
for each strCBResponseText in strArrayCBresponse

  if len(strCBresponseText) > 0 then
    logdata strDebugPath & "cbresponse.log", strCBresponseText, false
    if instr(strCBresponseText, "Sample not found by hash ") then
      'hash not found
    else

      if instr(strCBresponseText, "computer_dns_name")  then 
       ' msgbox strCBresponseText
       ' msgbox instr(strCBresponseText, "os_environment_display_string" & Chr(34) & ": " & chr(34))
        strTmpName = getdata(strCBresponseText, chr(34), "computer_dns_name" & Chr(34) & ": " & chr(34))
        strTmpOS = getdata(strCBresponseText, chr(34), "os_environment_display_string" & Chr(34) & ": " & chr(34) )
       ' msgbox strTmpOS
        strTmpregistered = getdata(strCBresponseText, chr(34), "registration_time" & Chr(34) & ": "& chr(34) )
        strStoredBytes = getdata(strCBresponseText, chr(34), "num_storefiles_bytes" & Chr(34) & ": "& chr(34) )
        strStatusBytes = getdata(strCBresponseText, chr(34), "status" & Chr(34) & ": "& chr(34) )
        strHealth = getdata(strCBresponseText, chr(34), "sensor_health_message" & Chr(34) & ": "& chr(34) )
        strGroup = getdata(strCBresponseText, ",", "group_id" & Chr(34) & ": " )
        strTmpLastCheckIn = getdata(strCBresponseText, chr(34), "last_checkin_time" & Chr(34) & ": "& chr(34) )
        strTmpEvtBytes = getdata(strCBresponseText, chr(34), "num_eventlog_bytes" & Chr(34) & ": "& chr(34) )
        strCompName = getdata(strCBresponseText, chr(34), "computer_name" & Chr(34) & ": "& chr(34) )
        strDaysonline = datediff("d", left(strTmpregistered, instr(strTmpregistered, ".") -1),left(strTmpLastCheckIn, instr(strTmpLastCheckIn, ".") -1))
        LogData strSSfilePath, chr(34) & strTmpName & chr(34) & "," & chr(34) & strTmpOS & chr(34) & "," & chr(34) & strTmpregistered & chr(34) & "," & chr(34) & strStoredBytes & chr(34) & "," & chr(34) & strStatusBytes & chr(34) & "," & chr(34) & strHealth & chr(34) & "," & chr(34) & strGroup & chr(34)& "," & chr(34) & strTmpLastCheckIn & chr(34) & "," & chr(34) & strTmpEvtBytes & chr(34) & "," & chr(34) & strDaysonline & chr(34) & "," & chr(34) & strCompName & chr(34), False
      end if
    end if
  end if
  
next
set objHTTP = nothing
end function

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





function LogData(TextFileName, TextToWrite,EchoOn)
Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      on error resume next
      fsoLogData.CreateTextFile TextFileName, True
      if err.number <> 0 and err.number <> 53 then msgbox err.number & " " & err.description & vbcrlf & TextFileName
      on error goto 0
  End If
if TextFileName <> "" then


  Set WriteTextFile = fsoLogData.OpenTextFile(TextFileName,ForAppending, False)
  on error resume next
  WriteTextFile.WriteLine TextToWrite
  if err.number <> 0 then 
    on error goto 0
    WriteTextFile.Close
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.CharSet = "utf-16"
    objStream.Open
    objStream.WriteText TextToWrite
    on error resume next
    objStream.SaveToFile TextFileName, 2
    if err.number <> 0 then msgbox err.number & " - " & err.message & " Problem writting to " & TextFileName
    if err.number <> 0 then msgbox "problem writting text: " & TextToWrite
    on error goto 0
    Set objStream = nothing
  end if
end if
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

function UDate(oldDate)
    UDate = DateDiff("s", "01/01/1970 00:00:00", oldDate)
end function


Function encrypt(StrText, key) 
  Dim lenKey, KeyPos, LenStr, x, Newstr 
   
  Newstr = "" 
  lenKey = Len(key) 
  KeyPos = 1 
  LenStr = Len(StrText) 
  StrText = StrReverse(StrText) 
  For x = 1 To LenStr 
       Newstr = Newstr & chr(asc(Mid(StrText,x,1)) + Asc(Mid(key,KeyPos,1))) 
       KeyPos = keypos+1 
       If KeyPos > lenKey Then KeyPos = 1 
       'if x = 4 then msgbox "error with char " & Chr(34) & asc(Mid(StrText,x,1)) - Asc(Mid(key,KeyPos,1)) & Chr(34) & " At position " & KeyPos & vbcrlf & Mid(StrText,x,1) & Mid(key,KeyPos,1) & vbcrlf & asc(Mid(StrText,x,1)) & asc(Mid(key,KeyPos,1))
  Next 
  encrypt = Newstr 
 End Function 
  
Function Decrypt(StrText,key) 
  Dim lenKey, KeyPos, LenStr, x, Newstr 
   
  Newstr = "" 
  lenKey = Len(key) 
  KeyPos = 1 
  LenStr = Len(StrText) 
   
  StrText=StrReverse(StrText) 
  For x = LenStr To 1 Step -1 
     on error resume next
     Newstr = Newstr & chr(asc(Mid(StrText,x,1)) - Asc(Mid(key,KeyPos,1))) 
     if err.number <> 0 then
        msgbox "error with char " & Chr(34) & asc(Mid(StrText,x,1)) - Asc(Mid(key,KeyPos,1)) & Chr(34) & " At position " & KeyPos & vbcrlf & Mid(StrText,x,1) & Mid(key,KeyPos,1) & vbcrlf & asc(Mid(StrText,x,1)) & asc(Mid(key,KeyPos,1))
        wscript.quit(011)
     end if
     on error goto 0
     KeyPos = KeyPos+1 
     If KeyPos > lenKey Then KeyPos = 1 
   Next 
   Newstr=StrReverse(Newstr) 
   Decrypt = Newstr 
End Function 
 
Function FormatDate(strFDate) 
Dim strTmpMonth
Dim strTmpDay
strTmpMonth = datepart("m",strFDate)
strTmpDay = datepart("d",strFDate)
if len(strTmpMonth) = 1 then strTmpMonth = "0" & strTmpMonth
if len(strTmpDay) = 1 then strTmpDay = "0" & strTmpDay

FormatDate = datepart("yyyy",strFDate) & "-" & strTmpMonth & "-" & strTmpDay


end function


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