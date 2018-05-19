'CB Hash Dump v2.8 - Dumps hashes from CB (Carbon Black) Response
'Dumps CSV "MD5|Path|Publisher|Company|Product|CB Prevalence|Logical Size|Score

'This script will write out hashes and some associated data via the CB Response (Carbon Black) API

'Copyright (c) 2018 Ryan Boyle randomrhythm@rhythmengineering.com.

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
Dim strSRSTRustQuery
Dim strHostFilter
Dim boolOutputHosts
Dim boolOutputDateAdded
Dim boolOutputCopyright
Dim boolOutputInternalName
Dim boolOutputOrigFname
Dim boolOutputFileDesc
Dim boolEnableYARA
Dim yaraFeedID
Dim tmpYaraUID
Dim ExpandYARA
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim dictYARA: Set dictYARA = CreateObject("Scripting.Dictionary")
Dim dictYARoutput: Set dictYARoutput = CreateObject("Scripting.Dictionary")
Dim boolUseSocketTools
Dim strLicenseKey
Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1


'---Config Section
BoolDebugTrace = False 'Leave this to false unless asked to collect debug logs.
IntDayStartQuery = "*" 'time to go back for start date of query. Set to "*" to query all binaries. Set to "-7" for the last week.
strTimeMeasurement = "d" '"h" for hours "d" for days
IntDayEndQuery = "*" 'days to go back for end date of query. Set to "*" for no end date. Set to "-1" to stop at yesterday.
strBoolIs_Executable = "True" 'set to "true" to query executables. Set to "false" to query resources (DLLs).
BoolExcludeSRSTRust = True 'Exclude trusted applications from the query
strSensorID = "" 'sensor_id
strHostFilter = "" 'computer name to filter to. Use uppercase, is case sensitive. If does not work try using sensor_id instead
strqueryitem = "" 'feed or other query item to filter off of example:"alliance_score_srsthreat:* "
boolOutputHosts = True ' Set to True to output hostnames for each binary
boolOutputDateAdded = True ' Set to True to output the date that the file was added to Cb Response
boolOutputDateSigned = True ' Set to True to output the date the binary was signed
boolOutputInternalName = True 'Seems uncommon for this to be populated
boolOutputCopyright = True
boolOutputComments = True
boolOutputOrigFname = True 'Seems uncommon for this to be populated
boolOutputFileDesc = True
boolEnableYARA = True'include yara
expandYARA = False 'Adds a column for each rule. Set to false to put all YARA data in one column
boolUseSocketTools = False 'Uses external library from SocketTools (needed when using old OS that does not support latest TLS standards)
strLicenseKey = "" 'Lincense key is required to use SocketTools 
'---End Config section

if strHostFilter <> "" then 
  msgbox "filtering to host " & strHostFilter
  strHostFilter = " AND hostname:" & strHostFilter
end if
if strSensorID <> "" then 
  msgbox "filtering to sensor ID " & strSensorID
  strHostFilter = " AND sensor_id:" & strSensorID
end if
if strqueryitem <> "" then
  strHostFilter = strHostFilter & " AND " & strqueryitem
end if

if isnumeric(IntDayStartQuery) then
  strStartDateQuery = DateAdd(strTimeMeasurement,IntDayStartQuery,now)

  ' AND server_added_timestamp:[" & strStartDateQuery & "T00:00:00 TO "
  strStartDateQuery = " AND server_added_timestamp:[" & FormatDate (strStartDateQuery) & " TO "
  if IntDayEndQuery = "*" then
    strEndDateQuery = "*]"
  elseif isnumeric(IntDayEndQuery) then
    strEndDateQuery = DateAdd(strTimeMeasurement,IntDayEndQuery,now)
    strEndDateQuery = FormatDate (strEndDateQuery) & "]"
  end if
elseif isnumeric(IntDayEndQuery) then
  strEndDateQuery = DateAdd(strTimeMeasurement,IntDayEndQuery,now)
  strEndDateQuery = " AND server_added_timestamp:[ * TO " & FormatDate (strEndDateQuery) & "]"
end if

msgbox "Date query: " & right(strStartDateQuery & strEndDateQuery, len(strStartDateQuery & strEndDateQuery) - instr(strStartDateQuery & strEndDateQuery,"[") +1) 

strSRSTRustQuery = ""
if BoolExcludeSRSTRust = True then
  strSRSTRustQuery = " AND -alliance_score_srstrust:*"
end if

CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug\"
strSSfilePath = CurrentDirectory & "\CB_" & udate(now) & ".csv"

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


if boolUseSocketTools = True then
on error resume next
  Set objST_Http = WScript.CreateObject("SocketTools.HttpClient.9")
  if err.number <> 0 then
on error goto 0
    msgbox "Problem loading SocketTools HTTP Client. Script will now exit. Try registering the control or disabling SocketTools"
    wscript.quit
  end if
  nError = objST_Http.Initialize(strLicenseKey) 
  If nError <> 0 Then
      WScript.Echo "Unable to initialize SocketTools component"
      WScript.Quit(1)
  End If
end if


strTempAPIKey = ""


intTotalQueries = 50
'get feed info
DumbCbFeed 0, False, intTotalQueries, "/api/v1/feed"
'process yara
if boolEnableYARA = True then 
	wscript.sleep 10 
  if dictYARA.count  = 0 then
		CbFeedQuery "feed_id:" & yaraFeedID, "YARA"
		if dictYARA.count  = 0  then 
			'wscript.echo "Nothing returned from YARA feed so disabling it."
			boolEnableYARA = False
		end if

	end if
end if	



if BoolUseCarbonBlack = True then
  ssInternalName = ""
  ssCopyright = ""
  ssComment = ""
  ssYARA = ""
  if boolOutputOrigFname = True then ssOrigFname = "|Original File Name"
  if boolOutputInternalName = True then ssInternalName = "|Internal Name"
  IF boolOutputCopyright = True then ssCopyright = "|Copyright"
  If boolOutputFileDesc = True then ssFileDesc = "|File Description"
  If boolOutputComments = True then ssComment = "|Comments"
  'build header row
  strSSrow = "MD5|Path|Publisher|Company|Product" & ssInternalName & ssOrigFname & ssCopyright & ssFileDesc & ssComment & "|CB Prevalence" & "|Logical Size|Alliance Score"
  if boolOutputDateAdded = True then strSSrow = strSSrow & "|Date Time Added"
  if boolOutputDateSigned = True then strSSrow = strSSrow & "|Date Time Signed"
  if boolOutputHosts = True then strSSrow = strSSrow & "|Computers"
  If boolEnableYARA = True then 
	if expandYARA = False then
		strSSrow = strSSrow & "|YARA"
	else
		strSSrow = strSSrow & "|" & YARAheaderrow
	end if
	end if
  strTmpSSlout = chr(34) & replace(strSSrow, "|",chr(34) & "," & Chr(34)) & chr(34)
  logdata strSSfilePath, strTmpSSlout, False
  intTotalQueries = 10
  'loop through CB results
  intTotalQueries = DumpCarBlack(0, False, intTotalQueries)
  wscript.sleep 10
  msgbox "Total number of items being retrieved : " & intTotalQueries
  'DumpCarBlack 0, True, intTotalQueries 
  intCBcount = 0
  do while intCBcount < clng(intTotalQueries)
    DumpCarBlack intCBcount, True, 10000 
    intCBcount = intCBcount +10000
  loop
end if


Function DumpCarBlack(intCBcount,BoolProcessData, intCBrows)

Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Dim strAVEurl
Dim strReturnURL
dim strAssocWith
Dim strCBresponseText
Dim strtmpCB_Fpath

'msgbox StrBaseCBURL & "/api/v1/binary?q=is_executable_image:" & strBoolIs_Executable & strSRSTRustQuery & strStartDateQuery & strEndDateQuery & "&start=" & intCBcount & "&rows=" & intCBrows
strAVEurl = StrBaseCBURL & "/api/v1/binary?q=is_executable_image:" & strBoolIs_Executable & strSRSTRustQuery & strHostFilter & strStartDateQuery & strEndDateQuery & "&start=" & intCBcount & "&rows=" & intCBrows
if boolUseSocketTools = False then
  objHTTP.open "GET", strAVEurl, False
  objHTTP.SetOption 2, 13056
  objHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
    

  on error resume next
    objHTTP.send 
    if err.number <> 0 then
      logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
      exit function 
    end if
  on error goto 0  
  'creates a lot of data. DOn't uncomment next line unless your going to disable it again
  if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "" & ".txt", objHTTP.responseText & vbcrlf & vbcrlf,BoolEchoLog 
  strCBresponseText = objHTTP.responseText
else
  strCBresponseText = SocketTools_HTTP(strAVEurl)
  
end if
if instr(strCBresponseText, "401 Unauthorized") then
  Msgbox "Carbon Black did not like the API key supplied"
  wscript.quit(997)
end if
if instr(strCBresponseText, "400 Bad Request") then
  msgbox "Server did not like the query. Try using " & chr(34) & "*" & CHr(34) & " for the start and end dates" & vbcrlf & strAVEurl
else
  strArrayCBresponse = split(strCBresponseText, vblf & "    {")
  for each strCBResponseText in strArrayCBresponse

    if len(strCBresponseText) > 0 then
      'logdata strDebugPath & "cb.log", strCBresponseText, false
      if instr(strCBresponseText, "Sample not found by hash ") then
        'hash not found
      else
        if instr(strCBresponseText, "total_results" & Chr(34) & ": ") then
          DumpCarBlack = getdata(strCBresponseText, ",", "total_results" & Chr(34) & ": ")
        elseif instr(strCBresponseText, "md5") and BoolProcessData = True then 
          'DumpCarBlack = "Carbon Black has a copy of the file for hash " & strCarBlack_ScanItem

          strCBfilePath = getdata(strCBresponseText, "]", "observed_filename" & Chr(34) & ": [")
          strCBfilePath = replace(strCBfilePath,chr(10),"")
          strCBfilePath = RemoveTLS(strCBfilePath)
          strCBfilePath = getdata(strCBfilePath, chr(34),chr(34))'just grab the fist file path listed
          if instr(strCBresponseText, "digsig_publisher") then 
            strCBdigSig = getdata(strCBresponseText, chr(34), "digsig_publisher" & Chr(34) & ": " & Chr(34))
            strCBdigSig = replace(strCBdigSig,chr(10),"")
			strCBdigIssue = getdata(strCBresponseText, chr(34), "digsig_issuer" & Chr(34) & ": " & Chr(34))
          else
            'not signed 
			strCBdigSig = ""
			strCBdigIssue = ""
          end if
          if instr(strCBresponseText, "signed" & Chr(34) & ": " & Chr(34) & "Signed") = 0 and instr(strCBresponseText, "signed" & Chr(34) & ": " & Chr(34) & "Unsigned") = 0 then
            'problem with sig
            strCBdigSig = getdata(strCBresponseText, chr(34), "signed" & Chr(34) & ": " & Chr(34)) & " - " & strCBdigSig
          end if 
          if boolOutputHosts = True then
            strCBHostname = getdata(strCBresponseText, ",", "hostname" & Chr(34) & ": ")
            if strCBHostname = "" then
              strTmpCBHostname = getdata(strCBresponseText, "]", "endpoint" & Chr(34) & ": [" & vblf & "        " & chr(34))
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
          end if
          strCBcompanyName = getdata(strCBresponseText, chr(34), "company_name" & Chr(34) & ": " & Chr(34))
          strCBcompanyName = "|" & RemoveTLS(strCBcompanyName)
          strCBproductName = getdata(strCBresponseText, chr(34), "product_name" & Chr(34) & ": " & Chr(34))
          strCBproductName = "|" &RemoveTLS(strCBproductName)
          StrCBMD5 = getdata(strCBresponseText, chr(34), "md5" & Chr(34) & ": " & Chr(34))
          strCBprevalence = getdata(strCBresponseText, ",", "host_count" & Chr(34) & ": ")
          strCBFileSize = getdata(strCBresponseText, ",", "orig_mod_len" & Chr(34) & ": ")
          strtmpCB_Fpath = getfilepath(strCBfilePath)
          strCBVTScore = getdata(strCBresponseText, ",", "alliance_score_virustotal" & Chr(34) & ": ")
		  if boolOutputDateAdded = True then
			strDateTimeAdded = getdata(strCBresponseText, chr(34), "server_added_timestamp" & Chr(34) & ": " & Chr(34))
		  end if
		  if boolOutputDateSigned = True then
			strDateTimeSigned = getdata(strCBresponseText, chr(34), "digsig_sign_time" & Chr(34) & ": " & Chr(34))
		  end if
		  if boolOutputInternalName = True then
			strInternalName = getdata(strCBresponseText, chr(34), "internal_name" & Chr(34) & ": " & Chr(34))
		  end if
		  if boolOutputCopyright = True then
			strcopyright = getdata(strCBresponseText, chr(34), "legal_copyright" & Chr(34) & ": " & Chr(34))
		  end if
		  if boolOutputComments = True then 
			strComments = getdata(strCBresponseText, chr(34), "comments" & Chr(34) & ": " & Chr(34))
		  end if
		  if boolOutputOrigFname = True then 
			strOrigFname = getdata(strCBresponseText, chr(34), "original_filename" & Chr(34) & ": " & Chr(34))
		  end if		  
		if boolOutputFileDesc = True then 
			strFileDesc = getdata(strCBresponseText, chr(34), "file_desc" & Chr(34) & ": " & Chr(34))
		  end if
          'RecordPathVendorStat strtmpCB_Fpath 'record path vendor statistics
        end if
      end if
    end if
		
	strYaraLine = ""
	if boolEnableYARA = True then
		if expandYARA = False then 
			if dictYARA.exists(StrCBMD5) then
				strYaraLine = "|" & dictYARA.item(StrCBMD5)
			else
				strYaraLine = "|" 
			end if
		else
			strYaraLine = YARAEntryrow(StrCBMD5)
		end if
	end if
	
    if StrCBMD5 <> "" then
      strCBfilePath = AddPipe(strCBfilePath) 'CB File Path
      strCBdigSig = AddPipe(strCBdigSig) 'CB Digital Sig
      strCBcompanyName = AddPipe(strCBcompanyName)'CB Company Name
      strCBproductName = AddPipe(strCBproductName) 'Product Name        
      strCBFileSize = AddPipe(strCBFileSize)  
      strCBprevalence = AddPipe(strCBprevalence)
      strCBVTScore = AddPipe(strCBVTScore)
      strCBHostname = boolAddPipe(strCBHostname, boolOutputHosts)
      strDateTimeAdded = boolAddPipe(strDateTimeAdded, boolOutputDateAdded)
      strDateTimeSigned = boolAddPipe(strDateTimeSigned, boolOutputDateSigned)
      strInternalName = boolAddPipe(strInternalName, boolOutputInternalName)
      strcopyright = boolAddPipe(strcopyright, boolOutputCopyright)
	  strComments = boolAddPipe(strComments, boolOutputComments)
	  strOrigFname  = boolAddPipe(strOrigFname, boolOutputOrigFname)
	  strFileDesc  = boolAddPipe(strFileDesc, boolOutputFileDesc)
	  
      strSSrow = StrCBMD5 & strCBfilePath & strCBdigSig & strCBcompanyName & strCBproductName & strOrigFname & strInternalName & strcopyright & strFileDesc & strComments & strCBprevalence & strCBFileSize & strCBVTScore & strDateTimeAdded & strDateTimeSigned & strCBHostname & strYaraLine
	  strTmpSSlout = chr(34) & replace(strSSrow, "|",chr(34) & "," & Chr(34)) & chr(34)
      logdata strSSfilePath, strTmpSSlout, False
    end if
    strCBfilePath = ""
    strCBdigSig = ""
    strCBcompanyName = ""
    strCBproductName = ""
    strCBFileSize = ""
    strCBprevalence = "" 
    StrCBMD5 = "" 
    strCBVTScore = ""
  next
end if
set objHTTP = nothing
end function

Function boolAddPipe(strPipeless, BooleanAddPipe)
if BooleanAddPipe = True then
	strReturnPiped = AddPipe(strPipeless)
else
	strReturnPiped = strPipeless
end if
boolAddPipe = strReturnPiped
end function

Function GetData(contents, ByVal EndOfStringChar, ByVal MatchString)
MatchStringLength = Len(MatchString)
x= instr(contents, MatchString)

  if X >0 then
    strSubContents = Mid(contents, x + MatchStringLength, len(contents) - MatchStringLength - x +1)
    if instr(strSubContents,EndOfStringChar) > 0 then
      GetData = Mid(contents, x + MatchStringLength, instr(strSubContents,EndOfStringChar) -1)
      exit function
    else
      GetData = Mid(contents, x + MatchStringLength, len(contents) -x -1)
      exit function
    end if
  end if
GetData = ""

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
    strPipeAdded = "|" & replace(strpipeless, "|", ",")

  else
    strPipeAdded = "|" & replace(right(strpipeless, len(strpipeless) -1), "|", ",")
  end if  
else
  strPipeAdded = "|"
end if

AddPipe = strPipeAdded 
end function




Function encrypt(StrText, key) 'Rafael Paran? - https://gallery.technet.microsoft.com/scriptcenter/e0d5d71c-313e-4ac1-81bf-0e016aad3cd2
  Dim lenKey, KeyPos, LenStr, x, Newstr 
   
  Newstr = "" 
  lenKey = Len(key) 
  KeyPos = 1 
  LenStr = Len(StrText) 
  StrTmpText = StrReverse(StrText) 
  For x = 1 To LenStr 
       Newstr = Newstr & chr(asc(Mid(StrTmpText,x,1)) + Asc(Mid(key,KeyPos,1))) 
       KeyPos = keypos+1 
       If KeyPos > lenKey Then KeyPos = 1 
       'if x = 4 then msgbox "error with char " & Chr(34) & asc(Mid(StrTmpText,x,1)) - Asc(Mid(key,KeyPos,1)) & Chr(34) & " At position " & KeyPos & vbcrlf & Mid(StrTmpText,x,1) & Mid(key,KeyPos,1) & vbcrlf & asc(Mid(StrTmpText,x,1)) & asc(Mid(key,KeyPos,1))
  Next 
  encrypt = Newstr 
End Function 
  
Function Decrypt(StrText,key) 'Rafael Paran? - https://gallery.technet.microsoft.com/scriptcenter/e0d5d71c-313e-4ac1-81bf-0e016aad3cd2
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
strTmpHours = datepart("h",strFDate)
strTmpMinutes = datepart("n",strFDate)
strTmpSeconds = datepart("s",strFDate)
if len(strTmpMonth) = 1 then strTmpMonth = "0" & strTmpMonth
if len(strTmpDay) = 1 then strTmpDay = "0" & strTmpDay

if len(strTmpHours) = 1 then strTmpHours = "0" & strTmpHours
if len(strTmpMinutes) = 1 then strTmpMinutes = "0" & strTmpMinutes
if len(strTmpSeconds) = 1 then strTmpSeconds = "0" & strTmpSeconds

FormatDate = datepart("yyyy",strFDate) & "-" & strTmpMonth & "-" & strTmpDay & "T" & strTmpHours & ":" & strTmpMinutes & ":" & strTmpSeconds


end function



Function CbFeedQuery(strQuery, strUniquefName)
Dim intParseCount: intParseCount = 10
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
strAppendQuery = ""
boolexit = False 
do while boolexit = False 
	strAVEurl = StrBaseCBURL & "/api/v1/threat_report?q=" & strQuery & strAppendQuery
	if boolUseSocketTools = False then
		objHTTP.open "GET", strAVEurl, False
		objHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 

		on error resume next
		  objHTTP.send 
		  if err.number <> 0 then
			logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
			exit function 
		  end if
		on error goto 0 
		
		CBresponseText = objHTTP.responseBody
	else
		strCBresponseText = SocketTools_HTTP(strAVEurl)
	end if
	if len(CBresponseText) > 0 then
	
		binTempResponse = objHTTP.responseBody
		  StrTmpResponse = RSBinaryToString(binTempResponse)
		logdata CurrentDirectory & "\Cb_TQueryResults.log", StrTmpResponse,False 

		if instr(StrTmpResponse, vblf & "    {") > 0 then
		  strArrayCBresponse = split(StrTmpResponse, vblf & "    {")
		else
		  strArrayCBresponse = split(StrTmpResponse, vblf & "  {")
		end if
		for each strCBResponseText in strArrayCBresponse
			 strTmpIOC = getdata(strCBResponseText, "]", "[")

			 strItem = getdata(strTmpIOC, chr(34) ,chr(34))
				strCBid = getdata(strCBResponseText, chr(34), chr(34) & "id" & Chr(34) & ": " & Chr(34))
        strTitle = getdata(strCBResponseText, chr(34), "title" & Chr(34) & ": " & Chr(34))

        if strTitle <> "" then
          if instr(strTitle, "Matched yara rules: ") and ishash(strItem) then
			strTitle = replace(right(strTitle,len(strTitle) -20), ",", "^")
            dictYARA.add strItem, strTitle
          end if
			if instr(strTitle, "^") = 0 then
				queryYARAdict strTitle, strItem
			else
				arrayRuleNames = split(strTitle, "^")
				for each yaraRname in arrayRuleNames
					queryYARAdict yaraRname, strItem
				next
			end if
          strRowOut = strCBid & "|" & strTitle & "|" & strItem
          strRowOut = chr(34) & replace(strRowOut,"|",chr(34) & "," & Chr(34)) & chr(34)
          if tmpYaraUID = "" then tmpYaraUID = udate(now)
          logdata CurrentDirectory & "\" & strUniquefName & "_" & tmpYaraUID & ".csv",strRowOut , false
        end if
		next
	end if
  intResultCount = getdata(StrTmpResponse, ",", "total_results" & Chr(34) & ": ")
	if isnumeric(intResultCount) then

    intAnswer = vbno 'msgbox (intParseCount & " items have been pulled down. Do you want to pull down more? There are a total of " & intResultCount & " items to retrieve",vbYesNo, "Cb Scripts")
		if intAnswer = vbno and intParseCount < clng(intResultCount) then
			
			strAppendQuery = "&start=" & intParseCount & "&rows=" & 1000
			intParseCount = intParseCount + 1000
		else
			boolexit = True
			exit function
		end if
	else
		boolexit = True
		msgbox "YARA query error"
		exit function
	end if
loop
End function


Function DumbCbFeed(intCBcount,BoolProcessData, intCBrows, strURLQuery)

Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Dim strAVEurl
Dim strReturnURL
dim strAssocWith
Dim strCBresponseText
Dim strtmpCB_Fpath
Dim StrTmpFeedIP
Dim boolProcessChildren: boolProcessChildren = False
strAVEurl = StrBaseCBURL & strURLQuery
if BoolProcessData = True and instr(strAVEurl, "?") > 0 then
  strAVEurl = strAVEurl & "&start=" & intCBcount & "&rows=" & intCBrows
end if
if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "" & ".txt", "Query URL=" & strAVEurl & vbcrlf & vbcrlf,BoolEchoLog 

if boolUseSocketTools = False then
  objHTTP.open "GET", strAVEurl, False

  objHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
    

  on error resume next
    objHTTP.send 
    if err.number <> 0 then
      logdata CurrentDirectory & "\CBF_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
      exit function 
    end if
  on error goto 0  
  'creates a lot of data. Don't uncomment next line unless your going to disable it again
  if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "" & ".txt", objHTTP.responseText & vbcrlf & vbcrlf,BoolEchoLog 
  strCBresponseText = objHTTP.responseText
else
  strCBresponseText = SocketTools_HTTP(strAVEurl)
  
end if
if instr(strCBresponseText, "401 Unauthorized") then
  Msgbox "Carbon Black did not like the API key supplied"
  wscript.quit(997)
end if
if instr(strCBresponseText, vblf & "    {") > 0 then
  strArrayCBresponse = split(strCBresponseText, vblf & "    {")
else
  strArrayCBresponse = split(strCBresponseText, vblf & "  {")
end if
for each strCBResponseText in strArrayCBresponse

  if len(strCBresponseText) > 0 then
    'logdata strDebugPath & "cbresponse.log", strCBresponseText, True
    if instr(strCBresponseText, "Sample not found by hash ") > 0 then
      'hash not found
    else
      if instr(strCBresponseText, "total_results" & Chr(34) & ": ") > 0 then
        DumbCbFeed = getdata(strCBresponseText, ",", "total_results" & Chr(34) & ": ")
      elseif instr(strCBresponseText, "provider_url" & Chr(34) & ": ") > 0 and instr(strCBresponseText, "id" & Chr(34) & ": ") > 0 then
        strTmpFeedID = getdata(strCBresponseText, ",", "id" & Chr(34) & ": ")
        strTmpFeedName = getdata(strCBresponseText, Chr(34), chr(34) & "name" & Chr(34) & ": " & Chr(34))
		if strTmpFeedName = "yara" then yaraFeedID = strTmpFeedID

      end if
    end if
  end if

next
set objHTTP = nothing
end function


Function IsHash(TestString)

    Dim sTemp
    Dim iLen
    Dim iCtr
    Dim sChar
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            if isnumeric(sChar) or "a"= lcase(sChar) or "b"= lcase(sChar) or "c"= lcase(sChar) or "d"= lcase(sChar) or "e"= lcase(sChar) or "f"= lcase(sChar)  then
              'allowed characters for hash (hex)
            else
              IsHash = False
              exit function
            end if
        Next
    
    IsHash = True
    else
      IsHash = False
    End If
    
End Function

Function RSBinaryToString(xBinary)
  'Antonin Foller, http://www.motobit.com
  'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
  'to a string (BSTR) using ADO recordset

  Dim Binary
  'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
  If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary
  
  Dim RS, LBinary
  Const adLongVarChar = 201
  Set RS = CreateObject("ADODB.Recordset")
  LBinary = LenB(Binary)
  
  If LBinary>0 Then
    RS.Fields.Append "mBinary", adLongVarChar, LBinary
    RS.Open
    RS.AddNew
      RS("mBinary").AppendChunk Binary 
    RS.Update
    RSBinaryToString = RS("mBinary")
  Else
    RSBinaryToString = ""
  End If
End Function

Function queryYARAdict(strRuleName, strCbHash)
'returns true if exists in DB otherwise returns false and adds to DB
if dictYARoutput.exists(strRuleName) = False then
  Set dictYARoutput.item(strRuleName) = CreateObject("Scripting.Dictionary")
  dictYARoutput.item(strRuleName).item(strCbHash) = ""
  queryYARAdict = False
elseif dictYARoutput.item(strRuleName).exists(strCbHash) then
  queryYARAdict = True
else
    dictYARoutput.item(strRuleName).item(strCbHash) = ""
  queryYARAdict = False
end if

end function

Function YARAheaderrow()
strTmpReturnHead = ""

for each dictName in dictYARoutput
  if strTmpReturnHead = "" then
    strTmpReturnHead = dictName
  else
    strTmpReturnHead = strTmpReturnHead & "|" & dictName
  end if
next
YARAheaderrow = strTmpReturnHead
end function

Function YARAEntryrow(strCbHash)
strTmpReturnHead = ""

for each dictName in dictYARoutput
    if dictYARoutput.item(dictName).exists(strCbHash) = True then
      strTmpReturnHead = AppendValuedList(strTmpReturnHead,dictName,"|")
    else
      strTmpReturnHead = strTmpReturnHead & "|" 
    end if

next
YARAEntryrow = strTmpReturnHead
end function

Function AppendValuedList(strAggregate,strAppend,strSeparator)
    if strAggregate = "" then
      strAggregate = strSeparator & strAppend
    else
      strAggregate = strAggregate & strSeparator & strAppend
    end if
AppendValuedList = strAggregate

end Function




Function SocketTools_HTTP(strRemoteURL)
' SocketTools 9.3 ActiveX Edition
' Copyright 2018 Catalyst Development Corporation
' All rights reserved
'
' This file is licensed to you pursuant to the terms of the
' product license agreement included with the original software,
' and is protected by copyright law and international treaties.
' Unauthorized reproduction or distribution may result in severe
' criminal penalties.
'

'
' Retrieve the specified page from a web server and write the
' contents to standard output. The parameter should specify the
' URL of the page to display


Const httpTransferDefault = 0
Const httpTransferConvert = 1

Dim objArgs
Dim objHttp
Dim strBuffer
Dim nLength
Dim nArg, nError


'
' Create an instance of the control
'
Set objHttp = WScript.CreateObject("SocketTools.HttpClient.9")

'
' Initialize the object using the specified runtime license key;
' if the key is not specified, the development license will be used
'
strLicenseKey = "" ' Should be set to the runtime license key
nError = objHttp.Initialize(strLicenseKey) 
If nError <> 0 Then
    WScript.Echo "Unable to initialize SocketTools component"
    WScript.Quit(1)
End If

objHttp.HeaderField = "X-Auth-Token"
objHttp.HeaderValue = strCarBlackAPIKey 
    
' Setup error handling since the component will throw an error
' if an invalid URL is specified

On Error Resume Next: Err.Clear
objHttp.URL = strRemoteURL

' Check the Err object to see if an error has occurred, and
' if so, let the user know that the URL is invalid

If Err.Number <> 0 Then
    WScript.echo "The specified URL is invalid"
    WScript.Quit(1)
End If

' Reset error handling and connect to the server using the
' default property values that were updated when the URL
' property was set (ie: HostName, RemotePort, UserName, etc.)
On Error GoTo 0
nError = objHttp.Connect()

If nError <> 0 Then
    WScript.echo "Error connecting to " & strRemoteURL & ". " & objHttp.LastError & ": " & objHttp.LastErrorString
    WScript.Quit(1)
End If
objHttp.timeout = 90
' Download the file to the local system
nError = objHttp.GetData(objHttp.Resource, strBuffer, nLength, httpTransferConvert)

If nError = 0 Then
    SocketTools_HTTP = strBuffer
Else
    WScript.echo "Error " & objHttp.LastError & ": " & objHttp.LastErrorString
	SocketTools_HTTP = objHttp.ResultString
End If

objHttp.Disconnect
objHttp.Uninitialize
end function
