'Cb Response Alert Dump

'Copyright (c) 2018 Ryan Boyle randomrhythm@rhythmengineering.com.
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
Dim StrCBfilePath
Dim StrCBdigSig
Dim StrCBcompanyName
Dim StrCBproductName
Dim StrCBFileSize
Dim StrCBprevalence
Dim StrCBMD5
Dim intTotalQueries
Dim IntDaysQuery
Dim strStartDateQuery
Dim strEndDateQuery
Dim strHashOutPath
Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1
Dim DictIPAddresses: set DictIPAddresses = CreateObject("Scripting.Dictionary")'
Dim DictFeedInfo: set DictFeedInfo = CreateObject("Scripting.Dictionary")'
Dim DictFeedExclude: set DictFeedExclude = CreateObject("Scripting.Dictionary")'

Dim boolHeaderWritten
Dim boolEchoInfo
Dim intSleepDelay
Dim intPagesToPull
Dim intSizeLimit
Dim intReceiveTimeout
'---Config Section
BoolDebugTrace = False
boolEchoInfo = False 
IntDayStartQuery = "-9" 'days to go back for start date of query. Set to * to query all binaries
IntDayEndQuery = "*" 'days to go back for end date of query. Set to * for no end date
strTimeMeasurement = "d" '"h" for hours "d" for days
'DictFeedExclude.add "SRSThreat", 0 'exclude feed
'DictFeedExclude.add "NVD", 0 'exclude feed
'DictFeedExclude.add "SRSTrust", 0 'exclude feed
'DictFeedExclude.add "cbemet", 0 'exclude feed
intSleepDelay = 90000 'delay between queries
intPagesToPull = 20 'Number of alerts to retrieve at a time
intSizeLimit = 20000 'don't dump more than this number of pages per feed
intReceiveTimeout = 120 'number of seconds for timeout
'---End Config section


if isnumeric(IntDayStartQuery) then
  strStartDateQuery = DateAdd(strTimeMeasurement,IntDayStartQuery,now)

  ' AND server_added_timestamp:[" & strStartDateQuery & "T00:00:00 TO "
  strStartDateQuery = " AND created_time:[" & FormatDate (strStartDateQuery) & " TO "
  if IntDayEndQuery = "*" then
    strEndDateQuery = "*]"
  elseif isnumeric(IntDayEndQuery) then
    strEndDateQuery = DateAdd(strTimeMeasurement,IntDayEndQuery,now)
    strEndDateQuery = FormatDate (strEndDateQuery) & "]"
  end if
end if


CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug\VT\"
strSSfilePath = CurrentDirectory & "\CBIP_" & udate(now) & ".csv"

strRandom = "4bv3nT9vrkJpj3QyueTvYFBMIvMOllyuKy3d401Fxaho6DQTbPafyVmfk8wj1bXF" 'encryption key. Change if you want but can only decrypt with same key
Set objFSO = CreateObject("Scripting.FileSystemObject")


if intCountMetaorVT = 0 then

  strFile= CurrentDirectory & "\cb.dat"
  strAPIproduct = "Carbon Black" 
end if

strData = ""
if objFSO.fileexists(strFile) then
  Set objFile = objFSO.OpenTextFile(strFile)
  if not objFile.AtEndOfStream then 'read file
      On Error Resume Next
      strData = objFile.ReadLine 
      if intCountMetaorVT = 0 then 
        StrBaseCBURL = objFile.ReadLine
      end if  
      on error goto 0
  end if
  if strData <> "" then
    strData = Decrypt(strData,strRandom)
      strTempAPIKey = "apikey=" & strData
      strData = ""
  end if
end if

if not objFSO.fileexists(strFile) and strData = "" then
  strTempAPIKey = inputbox("Enter your " & strAPIproduct & " api key")
  if strTempAPIKey <> "" then
  strTempEncryptedAPIKey = strTempAPIKey
    strTempEncryptedAPIKey = encrypt(strTempEncryptedAPIKey,strRandom)
    logdata strFile,strTempEncryptedAPIKey,False
    strTempEncryptedAPIKey = ""
    if intCountMetaorVT = 0 then
      StrBaseCBURL = inputbox("Enter your " & strAPIproduct & " base URL (example: https://ryancb-example.my.carbonblack.io")
      logdata strFile,StrBaseCBURL,False
    end if 
  end if
end if  
if strTempAPIKey = "" then

    msgbox "invalid api key"
    wscript.quit(999)
end if

if instr(strTempAPIKey,"apikey=") then
  strCarBlackAPIKey = replace(strTempAPIKey,"apikey=","")
else
  strCarBlackAPIKey = strTempAPIKey
end if

if strCarBlackAPIKey <> "" and StrBaseCBURL <> "" then BoolUseCarbonBlack = True   

on error resume next
objFile.close
on error goto 0
strTempAPIKey = ""




intTotalQueries = 50
'get feed info  
DumpCarBlack 0, False, intTotalQueries, "/api/v1/feed"

for each strCBFeedID in DictFeedInfo
  'msgbox "DictFeedExclude.exists(" & DictFeedInfo.item(strCBFeedID) & ")=" & DictFeedExclude.exists(strCBFeedID)
  if DictFeedExclude.exists(DictFeedInfo.item(strCBFeedID)) = False then
    strQueryFeed = "/api/v1/alert?q=feed_name:" & DictFeedInfo.item(strCBFeedID)  & strStartDateQuery & strEndDateQuery
   
    if strQueryFeed <> "" then
      wscript.sleep 10
      intCBcount = 10
      boolHeaderWritten = False
      strHashOutPath = CurrentDirectory & "\CBalert_" & DictFeedInfo.item(strCBFeedID) & "_" & udate(now) & ".csv"
      intTotalQueries = DumpCarBlack(0, True, intCBcount, strQueryFeed)
      wscript.sleep intSleepDelay
      logdata CurrentDirectory & "\CB_Alerts.log", date & " " & time & ": " & "Total number of items being retrieved for feed " & DictFeedInfo.item(strCBFeedID) & ": " & intTotalQueries ,boolEchoInfo
      
      if clng(intTotalQueries) > 0 then
        
        do while intCBcount < clng(intTotalQueries) and intCBcount < intSizeLimit
          logdata strDebugPath & "\follow_queries.log" , date & " " & time & " " & DictFeedInfo.item(strCBFeedID) & ": " & intCBcount & " < " & intTotalQueries & " and " & intCBcount & " < " & intSizeLimit, false
          DumpCarBlack intCBcount, True, intPagesToPull, strQueryFeed
          intCBcount = intCBcount + intPagesToPull
          wscript.sleep intSleepDelay
        loop
      end if
      strSSfilePath = CurrentDirectory & "\CBIP_" & DictFeedInfo.item(strCBFeedID) & "_" & udate(now) & ".csv"
      For each item in DictIPAddresses
        LogData strSSfilePath, item & "|" & DictIPAddresses.item(item), False
      next
      DictIPAddresses.RemoveAll
     
    else
      msgbox "Parser not configured for " & DictFeedInfo.item(strCBFeedID)
    end if
  end if
next


Function DumpCarBlack(intCBcount,BoolProcessData, intCBrows, strURLQuery)

Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Dim strAVEurl
Dim strReturnURL
dim strAssocWith
Dim strCBresponseText
Dim strtmpCB_Fpath
Dim StrTmpFeedIP

strAVEurl = StrBaseCBURL & strURLQuery 

if BoolProcessData = True then strAVEurl = strAVEurl & "&start=" & intCBcount & "&rows=" & intCBrows
'msgbox strAVEurl
objHTTP.SetTimeouts 600000, 600000, 600000, 900000 
objHTTP.open "GET", strAVEurl, True

objHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
  

on error resume next
  objHTTP.send
  If objHTTP.waitForResponse(intReceiveTimeout) Then 'response ready
        'success!
    Else 'wait timeout exceeded
        logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " CarBlack lookup failed due to timeout", False
        exit function  
    End If 
  if err.number <> 0 then
    logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
    logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " HTTP status code - " & objHTTP.status,False 
    exit function 
  end if
on error goto 0  
'creates a lot of data. Don't uncomment next line unless your going to disable it again
'if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "" & ".txt", objHTTP.responseText & vbcrlf & vbcrlf,BoolEchoLog 
strCBresponseText = objHTTP.responseText

if instr(strCBresponseText, "b Response Cloud is currently undergoing maintenance and will be back shortly") > 0 then
  wscript.sleep 240000 
  DumpCarBlack = DumpCarBlack(intCBcount,BoolProcessData, intCBrows, strURLQuery)
  exit function
end if
'msgbox strCBresponseText
if instr(strCBresponseText, vblf & "    {") then
  strArrayCBresponse = split(strCBresponseText, vblf & "    {")
else
  strArrayCBresponse = split(strCBresponseText, vblf & "  {")
end if
for each strCBResponseEntry in strArrayCBresponse

  if len(strCBResponseEntry) > 0 then
    'logdata strDebugPath & "cbresponse.log", strCBResponseEntry, True

      if instr(strCBResponseEntry, "provider_url" & Chr(34) & ": ") and instr(strCBresponseText, "id" & Chr(34) & ": ") then
        strTmpFeedID = getdata(strCBResponseEntry, ",", "id" & Chr(34) & ": ")
        strTmpFeedName = getdata(strCBResponseEntry, Chr(34), chr(34) & "name" & Chr(34) & ": " & Chr(34))
        if DictFeedInfo.exists(strTmpFeedID) = false then DictFeedInfo.add strTmpFeedID, strTmpFeedName
      elseif BoolProcessData = True then 
        if instr(strCBresponseText, "total_results" & Chr(34) & ": ") > 0 then
          DumpCarBlack = getdata(strCBresponseText, ",", "total_results" & Chr(34) & ": ")
        
          if instr(strCBResponseEntry, "ioc_value") then
            LogIOCdata strCBResponseEntry
          else
            logdata currentdirectory & "\ioc_value.log", "Debug - did not contain ioc_value: " & strCBResponseEntry, False
          end if
        else
             logdata currentdirectory & "\total_results.log" , "Debug - did not contain total_results: " & strCBresponseText, False
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

Sub LogIOCdata(strCBresponseText)


if instr(strCBresponseText, "ioc_value") then 

  strCBfilePath = getdata(strCBresponseText, chr(34), "process_path" & Chr(34) & ": " & chr(34))
  strioc_value = getdata(strCBresponseText, chr(34), "ioc_value" & Chr(34) & ": " & Chr(34))
  if strioc_value = "" then 
    strioc_value = getdata(strCBresponseText, "}", "ioc_value" & Chr(34) & ": " & Chr(34) & "{")
  end if
  interface_ip = getdata(strCBresponseText, chr(34), "interface_ip" & Chr(34) & ": " & Chr(34))
  sensor_id = getdata(strCBresponseText, chr(34), "sensor_id" & Chr(34) & ": " & Chr(34))
  strdescription = getdata(strCBresponseText, chr(34), "description" & Chr(34) & ": " & Chr(34))
  search_query = getdata(strCBresponseText, chr(34), "search_query" & Chr(34) & ": " & Chr(34))
  StrCBMD5 = getdata(strCBresponseText, chr(34), "md5" & Chr(34) & ": " & Chr(34))
  strCBprevalence = getdata(strCBresponseText, ",", "hostCount" & Chr(34) & ": ")
  strCBHostname = getdata(strCBresponseText, chr(34), "hostname" & Chr(34) & ": " & chr(34))
  strstatus = getdata(strCBresponseText, ",", "strstatus" & Chr(34) & ": ")
  process_name = getdata(strCBresponseText, chr(34), "process_name" & Chr(34) & ": " & chr(34))
  netconn_count = getdata(strCBresponseText, ",", "netconn_count" & Chr(34) & ": ")
  if instr(strCBresponseText,"ioc_attr") then
    iocSection = getdata(strCBresponseText, "}", "ioc_attr" & Chr(34) & ": " & chr(34) & "{")
    strDirection = getdata(iocSection, "\", "direction\" & Chr(34) & ": \" & Chr(34))
    strprotocol = getdata(iocSection, "\", "protocol\" & Chr(34) & ": \" & Chr(34))
    strlocal_port = getdata(iocSection, "\", "local_port\" & Chr(34) & ": \" & Chr(34))
    strdns_name = getdata(iocSection, "\", "dns_name\" & Chr(34) & ": \" & Chr(34))
    strlocal_ip = getdata(iocSection, "\", "local_ip\" & Chr(34) & ": \" & Chr(34))
    strport = getdata(iocSection, "\", "remote_port\" & Chr(34) & ": \" & Chr(34))
    strremote_ip = getdata(iocSection, "\", "remote_ip\" & Chr(34) & ": \" & Chr(34))
  end if  
  if strCBHostname = "" then
    strTmpCBHostname = getdata(strCBresponseText, "]", "hostnames" & Chr(34) & ": [" & vblf & "        " & chr(34))
    if instr(strTmpCBHostname, "|") then
      arrayCBHostName = split(strTmpCBHostname, "|")
      for each CBNames in arrayCBHostName
        arrayCBnames = split(CBNames, vbLf)
        for each CBhostName in arrayCBnames
          strTmpCBHostname = replace(CBhostName, chr(34), "")
          strTmpCBHostname = replace(strTmpCBHostname, " ","" )
          if isnumeric(strTmpCBHostname) = False and strTmpCBHostname <> "" then
            'msgbox strTmpCBHostname
            if strCBHostname = "" then
              strCBHostname = strTmpCBHostname
            else
              strCBHostname= strCBHostname & "/" & strTmpCBHostname
            end if
          end if
        next
      next
    end if
  end if

  alert_severity = getdata(strCBresponseText, ",", "alert_severity" & Chr(34) & ": ")

  strtmpCB_Fpath = getfilepath(strCBfilePath)
  'RecordPathVendorStat strtmpCB_Fpath 'record path vendor statistics
end if

logdata currentdirectory & "\IOCs.txt", strioc_value, false
if strioc_value = "" then msgbox "Debug - strioc_value = "": " & strCBresponseText
if strioc_value <> "" then

  strCBfilePath = AddPipe(strCBfilePath) 'CB File Path
  process_name = AddPipe(process_name) 'CB Digital Sig
  netconn_count = AddPipe(netconn_count)'CB Company Name
  strstatus = AddPipe(strstatus) 'Product Name        
  strCBFileSize = AddPipe(strCBFileSize)  
  strCBprevalence = AddPipe(strCBprevalence)
  strCBHostname = AddPipe(strCBHostname)
  interface_ip = AddPipe(interface_ip)
  strdescription = AddPipe(strdescription)
  sensor_id = AddPipe(sensor_id)
  alert_severity = AddPipe(strCBcmdline)
  StrCBMD5 = AddPipe(StrCBMD5)
  
  IOC_Entries = ""
  IOC_Head = ""

  if instr(strCBresponseText,"ioc_attr") then
    strDirection = AddPipe(strDirection)
    strprotocol = AddPipe(strprotocol)
    strlocal_port = AddPipe(strlocal_port)
    strdns_name = AddPipe(strdns_name)
    strlocal_ip = AddPipe(strlocal_ip)
    strport = AddPipe(strport)
    strremote_ip = AddPipe(strremote_ip)
    search_query = AddPipe(search_query)
    IOC_Entries = strDirection & strprotocol & strlocal_port & strdns_name & strlocal_ip & strport & strremote_ip & search_query
    IOC_Head = ",Direction, Protocol, Local Port, DNS Name, Local IP, Port, Report IP, search_query"
  end if

  if boolHeaderWritten = False then
      strSSrow = "IOC,MD5,Path," & "process_name," & "netconn_count," & "Status," & "CB Prevalence,interface_ip, sensor_id, Description, Severity" & IOC_Head & ",Host Name"
      logdata strHashOutPath, strSSrow, False
      boolHeaderWritten = True
  END IF

  strSSrow = strioc_value & StrCBMD5 & strCBfilePath & process_name & netconn_count & strstatus & strCBprevalence & interface_ip  & sensor_id & strdescription & alert_severity & IOC_Entries & strCBHostname
  strTmpSSlout = chr(34) & replace(strSSrow, "|",chr(34) & "," & Chr(34)) & chr(34)
  logdata strHashOutPath, strTmpSSlout, False
end if
strCBfilePath = ""
strCBdigSig = ""
strCBcompanyName = ""
strCBproductName = ""
strCBFileSize = ""
strCBprevalence = "" 
StrCBMD5 = "" 
strCBHostname = ""
strCBInfoLink = ""
strCBcmdline = ""
parent_name = ""
end sub




function LogData(TextFileName, TextToWrite,EchoOn)
Set fsoLogData = CreateObject("Scripting.FileSystemObject")
if EchoOn = True then wscript.echo TextToWrite
  If fsoLogData.fileexists(TextFileName) = False Then
      'Creates a replacement text file 
      on error resume next
      fsoLogData.CreateTextFile TextFileName, True
      if err.number <> 0 and err.number <> 53 then msgbox "can't create file " & Chr(34) & TextFileName & Chr(34) & ": " & err.number & " " & err.description & vbcrlf & TextFileName
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

Sub ExitExcel()
if BoolUseExcel = True then
  objExcel.DisplayAlerts = False
  objExcel.quit
end if
end sub
Function RemoveTLS(strTLS)
dim strTmpTLS
if len(strTLS) > 0 then
  for rmb = 1 to len(strTLS)
    if mid(strTLS, rmb, 1) <> " " then
      strTmpTLS = right(strTLS,len(strTLS) - RMB +1)
      exit for
    end if
  next
end if

if len(strTmpTLS) > 0 then
  for rmb = len(strTmpTLS)  to 1 step -1

    if mid(strTmpTLS, rmb, 1) <> " " then
      strTmpTLS = left(strTmpTLS,len(strTmpTLS) - (len(strTmpTLS) - RMB))
      exit for
    end if
  next
end if

RemoveTLS = strTmpTLS
end function
Function AddPipe(strpipeless)
dim strPipeAdded

if len(strpipeless) > 0 then
  if left(strpipeless, 1) <> "|" then 
    strPipeAdded = "|" & strpipeless

  else
    strPipeAdded = strpipeless
  end if  
else
  strPipeAdded = "|"
end if

AddPipe = strPipeAdded 
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


