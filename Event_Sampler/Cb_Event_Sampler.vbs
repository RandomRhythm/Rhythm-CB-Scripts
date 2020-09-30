'Cb Event Sampler v1.0.2
'Queries IOCs in Cb Response event data and provides a sampling CSV output


'Copyright (c) 2020 Ryan Boyle randomrhythm@rhythmengineering.com.

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

Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1
Const TristateTrue = -1
Const TristateFalse = 0
Dim StrBaseCBURL
Dim strUnique
Dim boolEventHeader: boolEventHeader = False
Dim boolNetworkHeader: boolNetworkHeader = False
Dim boolRegHeader: boolRegHeader = False
Dim boolModHeader: boolModHeader = False
Dim boolChildHeader: boolChildHeader = False
Dim boolFileHeader: boolFileHeader = False
Dim boolCrossHeader: boolCrossHeader = False
Dim boolNetworkEnable
Dim boolRegEnable
Dim boolModEnable
Dim boolChildEnable
Dim boolFileEnable
Dim boolCrossEnable
Dim dictRegAction: Set dictRegAction = CreateObject("Scripting.Dictionary")
Dim dictChild: Set dictChild = CreateObject("Scripting.Dictionary")
Dim dictFileAction: Set dictFileAction = CreateObject("Scripting.Dictionary")
Dim dictUID: Set dictUID = CreateObject("Scripting.Dictionary")
Dim boolDebug: boolDebug = false
Dim boolReportUserName
Dim pullAllSections
Dim intSleepDelay
Dim intPagesToPull
Dim intReceiveTimeout
Dim intAnswer: intAnswer = ""
Dim boolUseSocketTools
Dim strLicenseKey
Dim sensor_id
Dim APIVersion
Dim intClippingLevel
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim boolHeaderWritten
Dim strIOC
DIm objShellComplete
Set objShellComplete = WScript.CreateObject("WScript.Shell") 
Dim intHeaderCount
dim tmpArrayPointer() 'temporary location pointer 
Dim boolCaseSensitive 'force case sensitive matching
Dim strDelimiter 'This is the delimiter character
Dim DictHeader: Set DictHeader = CreateObject("Scripting.Dictionary") 'Maping between header text and integer column location. Header text is populated from a file. 

CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug"

'Optional config section
APIVersion = 4
boolNetworkEnable = True
boolRegEnable = False
boolModEnable = False
boolChildEnable = False
boolFileEnable = False
boolCrossEnable = False
pullAllSections = True 'set to true to grab everything
boolReportUserName = True 'Include associated user name
boolReportProcessName = True 'Include associated process name
strCbQuery = "" 'Cb Response query to run. Can be passed as an argument to the script.
intSleepDelay = 1000 'delay between queries
intPagesToPull = 1 'Number of events to retrieve at a time. Only supports a value of 1 currently.
intReceiveTimeout = 120 'number of seconds for timeout
strReportPath = "\Reports" 'directory to write report output
strInputPath = "" 'File to process for IOCs
boolCaseSensitive = True 'Setting to false forces everything to lowercase
strDelimiter =  "," 'delimiter character. Use VbTab for tab separated or "," for comma separated
strUniqueColumn = "MD5" 'exact text match of column header that contains the unique data to track. Example MD5 column
intHrowAbort = 6 'Number of rows in to abort if header has not been identified.
boolUseSocketTools = False 'Uses external library from SocketTools (needed when using old OS that does not support latest TLS standards)
strLicenseKey = "" 'License key is required to use SocketTools 
strIniPath="Cb_es.ini"
'end config section

if objFSO.FileExists(strIniPath) = false then
	If InStr(strIniPath, "\") = 0 Then 
		strIniPath = CurrentDirectory & "\" & strIniPath
	End If
End if		
if instr(strReportPath, ":") = 0 then 
	strReportPath = CurrentDirectory & "\" & strReportPath
end if


if objFSO.FileExists(strIniPath) = True then
'---Ini loading section
intSleepDelay = ValueFromINI(strIniPath, "IntegerValues", "SleepDelay", intSleepDelay)
intReceiveTimeout = ValueFromINI(strIniPath, "IntegerValues", "ReceiveTimeout", intReceiveTimeout)
APIVersion = ValueFromINI(strIniPath, "IntegerValues", "APIVersion", APIVersion)
boolUseSocketTools = ValueFromINI(strIniPath, "BooleanValues", "UseSocketTools", boolUseSocketTools)
boolNetworkEnable = ValueFromINI(strIniPath, "BooleanValues", "Network", boolNetworkEnable)
boolModEnable = ValueFromINI(strIniPath, "BooleanValues", "Modules", boolModEnable)
boolChildEnable = ValueFromINI(strIniPath, "BooleanValues", "Child", boolChildEnable)
boolFileEnable = ValueFromINI(strIniPath, "BooleanValues", "File", boolFileEnable)
boolCrossEnable = ValueFromINI(strIniPath, "BooleanValues", "Cross", boolCrossEnable)
pullAllSections = ValueFromINI(strIniPath, "BooleanValues", "AllSections", pullAllSections)
boolReportUserName = ValueFromINI(strIniPath, "BooleanValues", "ReportUserName", boolReportUserName)
boolReportProcessName = ValueFromINI(strIniPath, "BooleanValues", "ReportProcessName", boolReportProcessName)
boolDebug = ValueFromINI(strIniPath, "BooleanValues", "Debug", boolDebug)	
strDelimiter = ValueFromINI(strIniPath, "StringValues", "Delimiter", strDelimiter)
strInputPath = ValueFromINI(strIniPath, "StringValues", "InputFile", strInputPath)
strUniqueColumn = ValueFromINI(strIniPath, "StringValues", "UniqueColumn", strUniqueColumn)
'---End ini loading section
else
	if BoolRunSilent = False then WScript.Echo strIniPath & " does not exist. Using script configured/default settings instead"
end if

if len(strInputPath) > 1 and instr(strInputPath, ":") = 0 then 
	strInputPath = CurrentDirectory & "\" & strInputPath
end if

if cint(APIVersion) > 4 then
  msgbox "API version " & APIVersion & " is not supported. Changing to V4"
  APIVersion = 4
end if

strUnique = udate(now)
strRandom = "4bv3nT9vrkJpj3QyueTvYFBMIvMOllyuKy3d401Fxaho6DQTbPafyVmfk8wj1bXF" 'encryption key. Change if you want but can only decrypt with same key

'create sub directories
if objFSO.folderexists(strReportPath) = False then _
objFSO.createfolder(strReportPath)
if objFSO.folderexists(strDebugPath) = False then _
objFSO.createfolder(strDebugPath)


'RegMod field 0: operation type, an integer 1, 2, 4 or 8
'1: Created the registry key
'2: First wrote to the registry key
'4: Deleted the key
'8: Deleted the value
dictRegAction.add "1", "Created"
dictRegAction.add "2", "Written"
dictRegAction.add "4", "Deleted Key"
dictRegAction.add "8", "Deleted Value"


'field 0: operation type, an integer 1, 2, 4 or 8
'1: Created the file
'2: First wrote to the file
'4: Deleted the file
'8: Last wrote to the file
dictFileAction.add "1", "Created"
dictFileAction.add "2", "First Written"
dictFileAction.add "4", "Deleted"
dictFileAction.add "8", "Last Writen"


strFile = CurrentDirectory & "\cb.dat"
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
    wscript.echo "invalid api key"
    wscript.quit(999)
end if
strCarBlackAPIKey = strTempAPIKey

if objFSO.fileexists(CurrentDirectory & "\" & strInputPath) then
  strInputPath = CurrentDirectory & "\" & strInputPath
else

	wscript.echo "Please open the text input list or CSV file"
	strInputPath = SelectFile( )
end if

'Read list of items to query
if not objFSO.fileexists(strInputPath) then
  objFSO.CreateTextFile strInputPath, True
   objShellComplete.run "notepad.exe " & chr(34) & strInputPath & chr(34)
  msgbox "Input list (" & strInputPath & ") file was not found. The file has been created and opened in notepad. Please input the hashes or IP and domain addresses you want to scan and save the file." 
end if
Set oFile = objFSO.GetFile(strInputPath)

	If oFile.Size = 0 Then
    objFSO.CreateTextFile strInputPath, True
   objShellComplete.run "notepad.exe " & chr(34) & strInputPath & chr(34)
  msgbox "Input list (" & strInputPath & ") file was empty. The file has been opened in notepad. Please input hashes or IP addresses and domains you want to scan and save the file." 

	End If

boolHeaderWritten = False
strHeaderImport = "" 'header from CSV file we are importing
Set objRLfile = objFSO.OpenTextFile(strInputPath)
Do While Not objRLfile.AtEndOfStream
  if not objRLfile.AtEndOfStream then 'read file
	  On Error Resume Next
	  strLineIn = objRLfile.ReadLine 
	  on error goto 0
		if instr(strLineIn, strDelimiter) then
			if BoolHeaderLocSet = False then
				strHeaderImport = strLineIn
				SetHeaderLocations strHeaderImport
				BoolHeaderLocSet = True
				if UniqueColumn = "" then
					UniqueColumn = InputBox("Type the text label of the column you want to perform queries against")
				end if
				if instr(strHeaderImport, UniqueColumn) = 0 then
					msgbox "Script will now exit. The text supplied for UniqueColumn " & Chr(34) & UniqueColumn & chr(34) & " does not match a header entry: " & strHeaderImport
					wscript.quit(4)
				end if
        On Error Resume Next
        strLineIn = objRLfile.ReadLine 
        on error goto 0				
			end if
			strIOC = ReturnSpreadSheetItem(strLineIn, DictHeader.item(strUniqueColumn))
			
		else
			strIOC = strLineIn
		end if
		if isIPaddress(strIOC) = True then
			strCbQuery = "ipaddr:" & strIOC
		elseif ishash(strIOC) then
			strCbQuery = "md5:" & strIOC
		else
			strCbQuery = "domain:" & strIOC
		end if
		'wscript.echo "executing query: " & strCbQuery
		returnValues = CbQuery(strCbQuery, boolHeaderWritten, strHeaderImport)
		'msgbox "returnValues:" & returnValues
		if boolHeaderWritten = False Then
			boolHeaderWritten = True
		end if
		commaOut = ""
		
		if returnValues = "" then
			for HeadCount = 1 to intHeaderCount 'missing data is populated with empty cells
			 commaOut = AppendValues(commaOut,chr(34) & chr(34)) 
			next
			commaOut = AppendValues(chr(34) & strIOC & chr(34),commaOut) 
			returnValues = commaOut
    end if
		if strHeaderImport <> "" then
			logdata strReportPath & "\Query_" & strUnique & ".csv",strLineIn & strDelimiter & returnValues, false
		else
			logdata strReportPath & "\Query_" & strUnique & ".csv",returnValues, false
		end if


	end if
	wscript.sleep intSleepDelay
loop


Function CbQuery(strQuery, boolWriteHeader, strExistingHeader)
Dim intParseCount: intParseCount = intPagesToPull
Set objHTTP_CbQ = CreateObject("WinHttp.WinHttpRequest.5.1")
objHTTP_CbQ.SetTimeouts 600000, 600000, 600000, 900000 
strAppendQuery = ""
boolexit = False 
do while boolexit = False 
	strAVEurl = StrBaseCBURL & "/api/v1/process?q=" & strQuery & strAppendQuery
	if boolUseSocketTools = False then
		objHTTP_CbQ.open "GET", strAVEurl, False
		objHTTP_CbQ.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
		on error resume next
		  objHTTP_CbQ.send 
		  If objHTTP_CbQ.waitForResponse(intReceiveTimeout) Then 'response ready
			'success?
			if objHTTP_CbQ.Status <> 200 then
				logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Non-200 status code returned: " & objHTTP_CbQ.Status & " " & objHTTP_CbQ.StatusText, False
				If objHTTP_CbQ.Status = 504 then
					msgbox "The gateway timed out. Perhaps try using a smaller PagesToPull value in the Cb_PE.ini file. The script will now exit"
					logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Cb_Pull_Events lookup failed due to gateway timeout: " & strAppendQuery, False 
					wscript.quit (504)
				ElseIf objHTTP_CbQ.Status = 502 then
					msgbox "Bad Gateway. Perhaps try using a smaller PagesToPull value in the Cb_PE.ini file. The script will now exit"
					logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Cb_Pull_Events lookup failed due to bad gateway: " & strAppendQuery, False 
					wscript.quit (502)
				end if
			end if
		  Else 'wait timeout exceeded
			logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Cb_Pull_Events lookup failed due to timeout: " & strAppendQuery, False
			exit function  
		  End If 
		  if err.number <> 0 then
			logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Cb_Pull_Events lookup failed with HTTP error. - " & err.description,False 
			exit function 
		  end if
		on error goto 0 
		CBresponseText = objHTTP_CbQ.responseBody
	else
	  StrTmpResponse = SocketTools_HTTP(strAVEurl)
	  CBresponseText = StrTmpResponse
	end if

	if len(CBresponseText) > 0 then
	
		binTempResponse = CBresponseText

		if boolUseSocketTools = False then
		  StrTmpResponse = RSBinaryToString(binTempResponse)

		end if
		if boolDebug = true then logdata CurrentDirectory & "\Cb_QueryResults.log", StrTmpResponse,False 
		
		if instr(StrTmpResponse, "title>Maintenance - Carbon Black Response Cloud") > 0 then
		  logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Server under maintenance", False
		  Msgbox "The Cb Response server reports it is under maintenance and is not providing API results. This can occur when running large queries. Try limiting the query and see if the problem persists."
		elseif instr(StrTmpResponse, vblf & "    {") > 0 then
		  strArrayCBresponse = split(StrTmpResponse, vblf & "    {")
		else
		  strArrayCBresponse = split(StrTmpResponse, vblf & "  {")
		end if
		for each strCBResponseText in strArrayCBresponse
			strCBSegID = getdata(strCBresponseText, ",", "segment_id" & Chr(34) & ": ")
			strCBID = getdata(strCBresponseText, chr(34), chr(34) & "id" & Chr(34) & ": " & CHr(34))
			if strCBID = "" then
				strCBID = getdata(strCBresponseText, chr(34), chr(34) & "unique_id" & Chr(34) & ": " & CHr(34))
				if instr(strCBID, "-") > 0 then strCBID = left(strCBID, len(strCBID) -9)
			end if
			if strCBID <> "" then
				logdata CurrentDirectory & "\CB_UID.log", strCBID & "-" & strCBSegID ,False 
				if pullAllSections = True then
					segments = SegCheck(strCBID)
					if instr(segments, "|") > 0 then
					 arraySegment = split(segments, "|")
					 for each strSeg in arraySegment
					  'if dictUID.exists(strCBID & "-" & strCBSegID) = false then
						if len(strSeg) > 12 then
						  strTmpReturn = CBEventData (strCBID & "/" & HexToDec(right(strSeg, 12)), boolWriteHeader, strExistingHeader)
						  if strTmpReturn <> "" then 
							CbQuery = strTmpReturn
							exit function
						  end if
						end if
					  
						'dictUID.add strCBID & "-" & strCBSegID, ""
					  'end if
					 next
					end if
				end if
				if strCBSegID <> "" then
				 'segment_id: REQUIRED the process segment id; this is the segment_id field in search results. If this is set to 0
				  strTmpReturn = CBEventData(strCBID & "/" & strCBSegID, boolWriteHeader, strExistingHeader)
				  if strTmpReturn <> "" then 
					CbQuery = strTmpReturn
					exit function
				  end if
				end if
			end if
		next
	end if
	intResultCount = getdata(StrTmpResponse, ",", "total_results" & Chr(34) & ": ")

	if isnumeric(intResultCount) then
		if intResultCount = 0 then
		  'wscript.echo "Zero items were retrieved. Please double check your query and try again: " & chr(34) & strCbQuery & chr(34)
		  exit function
		end if
		if intParseCount >= clng(intResultCount) then
		  'wscript.echo intResultCount & " items retrieved for query " & chr(34) & strCbQuery & chr(34)
		  exit function
		end if
		strMessageText = ". Do you want to pull the rest down?"
		if intClippingLevel < clng(intResultCount) then strMessageText = ". Do you want to pull the rest down (up to clipping level " & intClippingLevel & ")?"
		if intAnswer = "" then intAnswer = msgbox (intParseCount & " items have been pulled down for query " & chr(34) & strCbQuery & Chr(34) & strMessageText & " There are a total of " & intResultCount & " items to retrieve. Selecting no will pull down " & intPagesToPull & " more",vbYesNoCancel, "Cb Scripts")
		if intAnswer <> vbCancel and intParseCount < clng(intResultCount) and intClippingLevel > clng(intParseCount) then
			if intAnswer = vbNo then intAnswer = ""
			strAppendQuery = "&start=" & intParseCount & "&rows=" & intPagesToPull
			intParseCount = intParseCount + intPagesToPull
		else
			boolexit = True
			exit function
		end if
	else
		boolexit = True
		msgbox "total_results is missing from HTTP Response - " & StrTmpResponse
		msgbox "The script will now exit. Try running the query with a time limitation by adding something like " & chr(34) & "AND last_update:-10080m" & chr(34) & ". This example addition will restrict the query to the last week of activity" 
		exit function
	end if
	wscript.sleep intSleepDelay
loop
set objHTTP_CbQ = nothing
End function

Function SegCheck(strIDPath)
Set objHTTP_SC = CreateObject("WinHttp.WinHttpRequest.5.1")
strAVEurl = StrBaseCBURL & "/api/v" & APIVersion & "/process/" & strIDPath & "/segment"
if boolUseSocketTools = False then
	objHTTP_SC.open "GET", strAVEurl, False
	objHTTP_SC.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 

	on error resume next
	  objHTTP_SC.send 
	  if err.number <> 0 then
		logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
		exit function 
	  end if
	on error goto 0 
	CBresponseText = objHTTP_SC.responseBody
else
	  StrTmpResponse = SocketTools_HTTP(strAVEurl)
	  CBresponseText = StrTmpResponse
end if

if len(CBresponseText) > 0 then
	if boolUseSocketTools = False then
		binTempResponse = objHTTP_SC.responseBody
		StrTmpResponse = RSBinaryToString(binTempResponse)
	end if
  if instr(StrTmpResponse, "Unhandled exception.") > 0 then exit function 
  'debug line
  if boolDebug = true then logdata CurrentDirectory & "\CBs_Download.txt", StrTmpResponse,False 
  'msgbox StrTmpResponse
  if instr(StrTmpResponse, ">The requested URL was not found on the server.<") = 0 then
  
    
  
  end if
  if instr(StrTmpResponse, "last_server_update") > 0 then
    arrayUID = split(StrTmpResponse, "last_server_update")
    for each strUID in arrayUID
      strUIDs = strUIDs & "|" & getdata(strUID, chr(34), "unique_id" & chr(34) & ": " & Chr(34))
  
    next
  end if
end if

SegCheck = strUIDs
set objHTTP_SC = nothing
end function


Function CBEventData(strIDPath, boolWriteHeader, strCsvHeader)
Set objHTTP_ED = CreateObject("WinHttp.WinHttpRequest.5.1")
Dim strAVEurl
Dim CBresponseText
Dim binTempResponse
Dim StrTmpResponse
strAVEurl = StrBaseCBURL & "/api/v" & APIVersion & "/process/" & strIDPath & "/event" 
if boolUseSocketTools = False then
	objHTTP_ED.SetTimeouts 600000, 600000, 600000, 900000 
	objHTTP_ED.open "GET", strAVEurl, False
	objHTTP_ED.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
	logdata CurrentDirectory & "\CB_Download.log", strAVEurl,False 
	on error resume next
	  objHTTP_ED.send 
	  If objHTTP_ED.waitForResponse(intReceiveTimeout) Then 'response ready
		'success!
	  Else 'wait timeout exceeded
		logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Cb_Pull_Events lookup failed due to timeout", False
		wscript.sleep intSleepDelay
	  End If 
	  if err.number <> 0 then
		  logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " Cb_Pull_Events lookup failed with HTTP error. - " & err.description,False 
		  if err.message = "The operation timed out" then
			wscript.sleep intSleepDelay
		  end if
	  end if
	err.clear
	CBresponseText = objHTTP_ED.responseBody
else
	  StrTmpResponse = SocketTools_HTTP(strAVEurl)
	  CBresponseText = StrTmpResponse
end if


if err.number <> 0 then 
	if err.message = "The data necessary to complete this operation is not yet available." then
		logdata CurrentDirectory & "\CB_Error.log", Date & " " & Time & " " & err.number & " " & err.message, False
		wscript.sleep intSleepDelay
		CBresponseText = objHTTP_ED.responseBody
	end if
End If
on error goto 0 
if len(CBresponseText) = 0 then
  logdata CurrentDirectory & "\CB_Download.log", Date & " " & Time & " Event can't be retrieved - " & strIDPath,False 
  wscript.sleep 5
  exit function
end if
if boolUseSocketTools = False then
	binTempResponse = objHTTP_ED.responseBody
	StrTmpResponse = RSBinaryToString(binTempResponse)
end if
if boolDebug = true then logdata CurrentDirectory & "\CB_EDownload.txt", StrTmpResponse,False 
if instr(StrTmpResponse, "Unhandled exception.") > 0 then exit function 

'msgbox StrTmpResponse
if instr(StrTmpResponse, ">The requested URL was not found on the server.<") = 0 then
'strTmpText = getdata(,"]", "childproc_complete" & CHr(34) & ": [")
strTmpCmd = getdata(StrTmpResponse,chr(34), "cmdline" & CHr(34) & ": " & chr(34))
strTmpemet_count = getdata(StrTmpResponse,",", "emet_count" & CHr(34) & ": " )
strTmpexec_events_count = getdata(StrTmpResponse,",", "exec_events_count" & CHr(34) & ": " )
strTmp_netconn_count = getdata(StrTmpResponse,",", "netconn_count" & CHr(34) & ": " )
strTmp_alliance_score_bit9suspiciousindicators = getdata(StrTmpResponse,",", "alliance_score_bit9suspiciousindicators" & CHr(34) & ": " )
strTmp_id = getdata(StrTmpResponse,Chr(34), chr(34) & "id" & CHr(34) & ": " )
strTmp_ExePath = getdata(StrTmpResponse,",", chr(34) & "path" & CHr(34) & ": " )
if instr(strTmp_ExePath, chr(34)) > 0 then strTmp_ExePath = replace(strTmp_ExePath, chr(34), "") 'remove quotes
 
strTmp_segment_id = getdata(StrTmpResponse,",", "segment_id" & CHr(34) & ": " )
sensor_id = getdata(StrTmpResponse,",", "sensor_id" & CHr(34) & ": " )
if boolReportUserName = True then 
  username = getdata(StrTmpResponse,Chr(34), "username" & CHr(34) & ": " & Chr(34))
  username = "," & Chr(34) & username & Chr(34) 
  userNheader = "|User Name"
end if
if boolReportProcessName = True then
  processname = getdata(StrTmpResponse,Chr(34), "process_name" & CHr(34) & ": " & Chr(34))
  processname = "," & Chr(34) & processname & Chr(34) 
  processNheader = "|Process Name"
  
end if
strTmp_host_type = getdata(StrTmpResponse, Chr(34), "host_type" & CHr(34) & ": " & Chr(34))
strTmp_group = getdata(StrTmpResponse, Chr(34), "group" & CHr(34) & ": " & Chr(34))
strTmp_fork_children_count = getdata(StrTmpResponse,",", "fork_children_count" & CHr(34) & ": " )
If strTmp_fork_children_count = "" Then strTmp_fork_children_count = getdata(StrTmpResponse,",", "childproc_count" & CHr(34) & ": " )
strTmp_fork_children_count = getdata(StrTmpResponse,",", "childproc_count" & CHr(34) & ": " )
regmod_count = getdata(StrTmpResponse,",", "regmod_count" & CHr(34) & ": " )
filemod_count  = getdata(StrTmpResponse,",", "filemod_count" & CHr(34) & ": " )
modload_count  = getdata(StrTmpResponse,",", "modload_count" & CHr(34) & ": " )
crossproc_count = getdata(StrTmpResponse,",", "crossproc_count" & CHr(34) & ": " )
if boolEventHeader = False then
  outHeadrow = "Item|Registry Modification|File Modification|Module Load|Network|Children|Cross Process|Suspicious Indicators|Host Type|sensorID|Group|Blocked Process|CMD|User Name|Process Name"
  boolEventHeader = True
end if
process_pid = getdata(StrTmpResponse,",", "process_pid" & CHr(34) & ": " )
tmpCountValues = chr(34) & regmod_count & Chr(34) & "," & chr(34) & filemod_count & Chr(34) & "," & chr(34) & modload_count & Chr(34) & "," & chr(34) & strTmp_netconn_count & Chr(34) & "," & chr(34) & strTmp_fork_children_count & Chr(34) & "," & chr(34) & crossproc_count & Chr(34) 
strOutLine = Chr(34) & strIOC & Chr(34) & "," & tmpCountValues & "," & chr(34) & _
strTmp_alliance_score_bit9suspiciousindicators & Chr(34) & "," & chr(34) & strTmp_host_type & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & "," & chr(34) & strTmp_group & Chr(34) & "," & chr(34) & _
 strTmp_processblock_count & Chr(34) & ","& chr(34) & strTmpCmd & Chr(34) & username & processname 





if boolNetworkEnable = True and APIVersion  > 1 and APIVersion < 5 then
  if boolNetworkHeader = False then
	outHeadrow = outHeadrow & "|IP Address|Local Port|Remote Port|Protocol|Domain|Outbound" 
	boolNetworkHeader = True
  end if
  PassiveDNS = ""
  networkCSV = ""
  strTmpText = getdata(StrTmpResponse,"]", "netconn_complete" & CHr(34) & ": [") 
  if instr(strTmpText, "},") = 0 then
    strTmpText = strTmpText & ","
  end if
  arrayIPinfo = split(strTmpText, "},")
  for each IPinfo in arrayIPinfo
    strDomain = getdata (IPinfo, chr(34), "domain" & Chr(34) & ": " & chr(34))
    strProtocol = getdata (IPinfo, chr(34), "proto" & Chr(34) & ": " & chr(34))
  strLport = getdata (IPinfo, ",", "local_port" & Chr(34) & ": " )
  strDirection = getdata (IPinfo, chr(34), "direction" & Chr(34) & ": " & chr(34))
  strRport = getdata (IPinfo, ",", "remote_port" & Chr(34) & ": ")
  strIP = getdata (IPinfo, chr(34), "remote_ip" & Chr(34) & ": " & chr(34))
  strDtime = getdata (IPinfo, chr(34), "timestamp" & Chr(34) & ": " & chr(34))
  if strDtime <> "" Then
    if APIVersion = 2 or APIVersion = 3 then
      strIP = IPDecToDotQuad(strIP)
    end if

	if strIP = strIOC Then 'got a match for the IP or domain we are trying to get association for so return
		PassiveDNS = strDomain
	elseif strDomain = strIOC then 'got a match for the IP or domain we are trying to get association for so return
		PassiveDNS = strIP
	end if

	'if we aren't filtering to a domain/IP or 
	if isIPaddress(strIOC) = False and instr(strIOC, ".") = 0 or ishash(strIOC) = True or PassiveDNS <> "" then
		networkCSV = Chr(34) & strIP & chr(34) & "," & _
		Chr(34) & strLport & chr(34) & "," & Chr(34) & strRport & chr(34) & "," & Chr(34) & strProtocol & chr(34) & "," & _
		Chr(34) & strDomain & chr(34) & "," & Chr(34) & strDirection & chr(34) 
		exit for
    end if
  else
    networkCSV = Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & chr(34)
  End if
  next
	If networkCSV = "" Then 'if a match wasn't made just report on the last one
		networkCSV = Chr(34) & strIP & chr(34) & "," & _
		Chr(34) & strLport & chr(34) & "," & Chr(34) & strRport & chr(34) & "," & Chr(34) & strProtocol & chr(34) & "," & _
		Chr(34) & strDomain & chr(34) & "," & Chr(34) & strDirection & chr(34) 
	End if	
strOutLine = strOutLine & "," & networkCSV	
end if

if boolNetworkEnable = True and APIVersion  = 1 then
  if boolNetworkHeader = False then
	outHeadrow =  outHeadrow & "|IP Address|Remote Port|Protocol|Domain|Outbound"
	boolNetworkHeader = True
  end if
  strTmpText = getdata(StrTmpResponse,"]", "netconn_complete" & CHr(34) & ": [") 
  If strTmpText = "" Then strTmpText = ", "
  NetConnarrayEvents = split(strTmpText, ", ")
  for each EventEntry in NetConnarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  ArrayEE = split(replace(EventEntry,chr(34), ""), "|")
	  if ubound(arrayEE) > 4 then
	   if isnumeric(arrayEE(1)) then
			dotQuadIP = IPDecToDotQuad(arrayEE(1))
		else
			dotQuadIP = arrayEE(1)
		end if
	   strOutLine = strOutLine & "," & Chr(34) & dotQuadIP & Chr(34) & "," & Chr(34) & arrayEE(2) & Chr(34) & "," & Chr(34) & arrayEE(3) & Chr(34) & "," & Chr(34) & arrayEE(4) & Chr(34) & "," & Chr(34) & arrayEE(5) & Chr(34) 
		exit for
	  end if
	else
	    strOutLine = strOutLine & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & chr(34)
	    Exit For
  end if
  next
end if

if boolRegEnable = True then 
  if boolRegHeader = False then

	outHeadrow = outHeadrow & "|Action|Date Time|Registry Key"
	boolRegHeader = True
  end if
   strTmpText = getdata(StrTmpResponse,"]", "regmod_complete" & CHr(34) & ": [")
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 3 then
		strAction = ""
		if dictRegAction.exists(arrayEE(0)) then strAction =  dictRegAction.item(arrayEE(0))
	   strOutLine = strOutLine & "," & Chr(34) & strAction & Chr(34) & "," & Chr(34) & arrayEE(1) & Chr(34) & "," & Chr(34) & arrayEE(2) & Chr(34)
	   exit for
	  end if
	else
    	strOutLine = strOutLine & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34)
    	Exit for
	end if
  next
end if



if boolModEnable = True then
  if boolModHeader = False then

	outHeadrow = outHeadrow & "|Module Date Time|Module MD5|Module File Path|Module SHA256"
	boolModHeader = True
  end if
   strTmpText = getdata(StrTmpResponse,"]", "modload_complete" & CHr(34) & ": [")
  If strTmpText = "" then strTmpText = ", "
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  if right(tmpEvent,1) = "|" then tmpEvent = left(tmpEvent, len(tmpEvent) -1) 'remove end pipe as have not seen any values after it.
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) = 2 Then tmpEvent = tmpEvent & "|"
	  If ubound(arrayEE) > 1 then
      if arrayEE(2) <> strTmp_ExePath then
        strOutLine = strOutLine & "," & chr(34) & replace(tmpEvent, "|", chr(34) & "," & Chr(34)) & Chr(34) 
        exit for
      end if

	  end if
	else
	    strOutLine = strOutLine & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34)
	    Exit for
	end if
  Next
end if

if boolChildEnable = True and APIVersion  >= 3 then 
  dictChild.RemoveAll
  if boolChildHeader = False then
	outHeadrow = outHeadrow & "|Child Start Time|Child End Time|Child Unique ID|Child MD5|Child File Path|Child PID|Suppressed|Parent PID|Parent Unique ID|Sensor ID|Child Command Line" 
	boolChildHeader = True
  end if    
  strTmpText = getdata(StrTmpResponse,"]", "childproc_complete" & CHr(34) & ": [")
  if instr(strTmpText,  "},") = 0 then strTmpText = strTmpText & "},"
  CbarrayEvents = split(strTmpText, "},")
  for each EventEntry in CbarrayEvents
	childMD5 = getdata(EventEntry,chr(34), "md5" & CHr(34) & ": " & chr(34) )
	childCommandLine = getdata(EventEntry,chr(34) & ",", "commandline" & CHr(34) & ": " & chr(34) )
	childSha256 = getdata(EventEntry,chr(34), "sha256" & CHr(34) & ": " & chr(34) )
	childProcessId = getdata(EventEntry,chr(34), "processId" & CHr(34) & ": " & chr(34) )
	childIs_suppressed = getdata(EventEntry,",", "is_suppressed" & CHr(34) & ": " )
	childDateStartTime = getdata(EventEntry,chr(34), "start" & CHr(34) & ": " & chr(34) )
	childDateEndTime = getdata(EventEntry,chr(34), "end" & CHr(34) & ": " & chr(34) )
	childIs_tampered = getdata(EventEntry,",", "is_tampered" & CHr(34) & ": " )
	childPID = getdata(EventEntry,",", "pid" & CHr(34) & ": " )
	childIDPath = getdata(EventEntry,chr(34), "path" & CHr(34) & ": " & chr(34) )

		
   strWriteLine = Chr(34) & childProcessId & Chr(34) & _
   "," & Chr(34) & childMD5 & Chr(34) & "," & Chr(34) & childIDPath & Chr(34) & _
   "," & Chr(34) & childPID & Chr(34) & "," & Chr(34) & childIs_suppressed & Chr(34) & _
   "," & Chr(34) & process_pid & Chr(34) & "," & Chr(34) & strIDPath & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & "," & Chr(34) & childCommandLine & Chr(34)
   if dictChild.exists(strWriteLine) = False then 
	dictChild.add strWriteLine, childDateEndTime
   else		
	strOutLine = strOutLine & "," & Chr(34) & childDateStartTime & Chr(34) & "," & Chr(34) & dictChild.item(strWriteLine) & Chr(34) & "," & strWriteLine
	exit for
   end if	

  next
end if

if boolFileEnable = True then 
  if boolFileHeader = False then

	outHeadrow = outHeadrow & "|File Action|File Date Time|File Path|Last Write MD5|File Type|Tamper Attempt"
	boolFileHeader = True
  end if       
  strTmpText = getdata(StrTmpResponse,chr(34) & "], ", "filemod_complete" & CHr(34) & ": [")
  If strTmpText = "" then strTmpText = ", "
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 4 then
		strAction = ""
		if dictRegAction.exists(arrayEE(0)) then strAction =  dictFileAction.item(arrayEE(0))
		strOutLine = strOutLine & "," & Chr(34) & strAction & Chr(34) & "," & Chr(34) & arrayEE(1) & Chr(34) & "," & Chr(34) & arrayEE(2) & Chr(34)  & "," & Chr(34) & arrayEE(3) & Chr(34)  & "," & Chr(34) & arrayEE(4) & Chr(34)  & "," & Chr(34) & arrayEE(5) & Chr(34) 
	   exit for
	  else
		logdata CurrentDirectory & "\CB_Pull_Error.log", Date & " " & Time & " FileMod error splitting the value into an array size greater than four: " & tmpEvent,False 
	  end if
	else
	    strOutLine = strOutLine & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34)
	    Exit for
	end if
  next
end if

if boolCrossEnable = True then 
  if boolCrossHeader = False then

	outHeadrow = outHeadrow & "|Cross Process Action|Date Time|Target Unique ID|Target MD5|Target Path|Open Type|Access Requested|Tamper|Inbound Open|PID|Process Path|CrossProc Unique ID"
	boolCrossHeader = True
  end if    
  strTmpText = getdata(StrTmpResponse,"]", "crossproc_complete" & CHr(34) & ": [")
  If strTmpText = "" then strTmpText = ", "
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 1 then
	   strOutLine = strOutLine & "," & chr(34) & replace(tmpEvent, "|", chr(34) & "," & Chr(34)) & process_pid & Chr(34) & "," & Chr(34) & strTmp_ExePath & Chr(34) & "," & Chr(34) & strIDPath & Chr(34)
	   exit for
	  end if
	else
	    strOutLine = strOutLine & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & chr(34)
	    Exit for
	end if
  next
end if



else
logdata CurrentDirectory & "\CB_Download.log", Date & " " & Time & " Event can't be retrieved - " & strIDPath,False 
wscript.sleep 5
end if
'Date Time|IP Address|Remote Port|Protocol|Domain|Outbound|Sensor ID" & userNheader & processNheader|EMET|Execute|Network|Suspicious Indicators|SegmentID|Host Type|Group|Children|Blocked Process|CMD"
if boolWriteHeader = False then
	if strCsvHeader <> "" then strCsvHeader = strCsvHeader & ","
	logdata strReportPath & "\Query_" & strUnique & ".csv", strCsvHeader & chr(34) & replace(outHeadrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	intHeaderCount = ubound(split(outHeadrow, "|"))
end if
CBEventData =  strOutLine

set objHTTP_ED = nothing
end function



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
  Set RS = nothing
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
end function




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
      if err.number <> 0 and err.number <> 53 then msgbox "Logging error: " & err.number & " " & err.description & vbcrlf & TextFileName
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
 Set WriteTextFile = nothing
Set fsoLogData = Nothing
End Function

Function IPDecToDotQuad(intDecIP)
if IsIPv6(intDecIP) = True then 
	IPDecToDotQuad = intDecIP
	exit function
end if
tmpOct = ""
y = 0
for x = 1 to 32 
y=y+1
 tmpBit = GetBit(intDecIP, x) 
 if tmpBit = True then 
  tmpOct =  "1" & tmpOct
 else
  tmpOct =  "0" & tmpOct
 end if 
  if y = 8 then 
    'msgbox tmpOct
    'msgbox Dec2Bin(tmpOct)
    strIP = Dec2Bin(tmpOct) & "." & strIP
    y=0
    tmpOct = ""
  end if
next
strIP = left(strIP,len(strIP)-1)
IPDecToDotQuad = strIP
end function

Function GetBit(lngValue, BitNum)
     Dim BitMask
     If BitNum < 32 Then BitMask = 2 ^ (BitNum - 1) Else BitMask = "&H80000000"
     GetBit =Cbool(lngValue AND BitMask)
End Function


Function Dec2Bin(binary)

For s = 1 To Len(binary)
    n = n + (Mid(binary, Len(binary) - s + 1, 1) * (2 ^ (s - 1)))
Next
Dec2Bin = n
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



Function HexToDec(strHex)'http://blog.benfinnigan.com/2008/10/handling-large-hex-in-vbscript.html
    Dim i
    Dim size
    Dim ret
    size = Len(strHex) - 1
    ret = CDbl(0)
    For i = 0 To size
        ret = ret + CDbl("&H" & Mid(strHex, size - i + 1, 1)) * (CDbl(16) ^ CDbl(i))
    Next
    HexToDec = ret
End Function

function UDate(oldDate)
    UDate = DateDiff("s", "01/01/1970 00:00:00", oldDate)
end function


Function IsIPv6(TestString)

    Dim sTemp
    Dim iLen
    Dim iCtr
    Dim sChar
    
    if instr(TestString, ":") = 0 then 
		IsIPv6 = false
		exit function
	end if
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            if isnumeric(sChar) or "a"= lcase(sChar) or "b"= lcase(sChar) or "c"= lcase(sChar) or "d"= lcase(sChar) or "e"= lcase(sChar) or "f"= lcase(sChar) or ":" = sChar then
              'allowed characters for hash (hex)
            else
              IsIPv6 = False
              exit function
            end if
        Next
    
    IsIPv6 = True
    else
      IsIPv6 = False
    End If
    
End Function


Function ValueFromIni(strFpath, iniSection, iniKey, currentValue)
returniniVal = ReadIni( strFpath, iniSection, iniKey)
if returniniVal = " " or  returniniVal = "" then 
	returniniVal = currentValue
end if 
if TypeName(returniniVal) = "String" then
	returniniVal = stringToBool(returniniVal)'convert type to boolean if needed
elseif TypeName(returniniVal) = "Integer" then
	returniniVal = int(returniniVal)'convert type to int if needed
end if
ValueFromIni = returniniVal
end function

Function stringToBool(strBoolean)
if lcase(strBoolean) = "true" then 
	returnBoolean = True
elseif lcase(strBoolean) = "false" then 
	returnBoolean = False
else
	returnBoolean = strBoolean
end if
stringToBool = returnBoolean
end function

Function ReadIni( myFilePath, mySection, myKey ) 'http://www.robvanderwoude.com/vbstech_files_ini.php
    ' This function returns a value read from an INI file
    '
    ' Arguments:
    ' myFilePath  [string]  the (path and) file name of the INI file
    ' mySection   [string]  the section in the INI file to be searched
    ' myKey       [string]  the key whose value is to be returned
    '
    ' Returns:
    ' the [string] value for the specified key in the specified section
    '
    ' CAVEAT:     Will return a space if key exists but value is blank
    '
    ' Written by Keith Lacelle
    ' Modified by Denis St-Pierre and Rob van der Woude

    Dim intEqualPos
    Dim objFSO_ini, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO_ini = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO_ini.FileExists( strFilePath ) Then
        Set objIniFile = objFSO_ini.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
                strLine = Trim( objIniFile.ReadLine )

                ' Parse lines until the next section is reached
                Do While Left( strLine, 1 ) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
                        ' Check if item is found in the current line
                        If LCase( strLeftString ) = LCase( strKey ) Then
                            ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    Else
        if BoolRunSilent = False then WScript.Echo strFilePath & " does not exist. Using script configured/default settings instead"
    End If
	Set objIniFile = nothing
	Set objFSO_ini = nothing
End Function



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
    'WScript.echo "Error connecting to " & strRemoteURL & ". " & objHttp.LastError & ": " & objHttp.LastErrorString
    logdata CurrentDirectory & "\CB_Pull_Error.log", Date & " " & Time & " Error connecting to " & strRemoteURL & ". " & objHttp.LastError & ": " & objHttp.LastErrorString, false
End If
objHttp.timeout = 90
' Download the file to the local system
nError = objHttp.GetData(objHttp.Resource, strBuffer, nLength, httpTransferConvert)

If nError = 0 Then
    SocketTools_HTTP = strBuffer
Else
    'WScript.echo "Error " & objHttp.LastError & ": " & objHttp.LastErrorString
	SocketTools_HTTP = "Error " & objHttp.ResultString
End If

objHttp.Disconnect
objHttp.Uninitialize
Set objHttp = nothing
end function


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
if boolReturn_isIP = false then
	boolReturn_isIP = isIpv6(strIPaddress)
end if
isIPaddress = boolReturn_isIP
END FUNCTION



Function IsIPv6(TestString)

    Dim sTemp
    Dim iLen
    Dim iCtr
    Dim sChar
    
    if instr(TestString, ":") = 0 then 
		IsIPv6 = false
		exit function
	end if
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            if isnumeric(sChar) or "a"= lcase(sChar) or "b"= lcase(sChar) or "c"= lcase(sChar) or "d"= lcase(sChar) or "e"= lcase(sChar) or "f"= lcase(sChar) or ":" = sChar then
              'allowed characters for hash (hex)
            else
              IsIPv6 = False
              exit function
            end if
        Next
    
    IsIPv6 = True
    else
      IsIPv6 = False
    End If
    
End Function

Function IsHash(TestString)
Dim sTemp
Dim iLen
Dim iCtr
Dim sChar

sTemp = TestString
iLen = Len(sTemp)
If iLen > 31 Then 'md5 length is 32
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


Function AppendValues(strAggregate,strAppend)
    if strAggregate = "" then
      strAggregate = strAppend
    else
      strAggregate = strAggregate & "," & strAppend
    end if
AppendValues = strAggregate
end Function


Function ReturnSpreadSheetItem(strCSVrow, intColumnLocation) 'pass this function the csv row and which column you want to get the value
Dim strSpreadSheetItem

intArrayPointer = returnCellLocation(strCSVrow, intColumnLocation)
if instr(strCSVrow, strDelimiter) > 0 Then
	strTmpHArray = split(strCSVrow, strDelimiter)
	if ubound(tmpArrayPointer) >= intColumnLocation and cint(intColumnLocation) > -1 then
		if ubound(tmpArrayPointer) = intArrayPointer then
			strSpreadSheetItem = replace(strTmpHArray(intArrayPointer), Chr(34), "")
		elseif (tmpArrayPointer(intColumnLocation) +1 <> tmpArrayPointer(intColumnLocation +1)) then
			strSpreadSheetItem = ""
			for itemCount = 0 to tmpArrayPointer(intColumnLocation +1) - (tmpArrayPointer(intColumnLocation) +1)
				strSpreadSheetItem = AppendValues(strSpreadSheetItem, replace(strTmpHArray(intArrayPointer + itemCount), Chr(34), ""), strDelimiter)
			next
		else
			strSpreadSheetItem = replace(strTmpHArray(intArrayPointer), Chr(34), "")
		end if
	
	else
		msgbox "SpreadSheet array mismatch:strCSVrow=" & strCSVrow & "&intArrayPointer=" & intArrayPointer  & "&ubound(tmpArrayPointer)=" & ubound(tmpArrayPointer)
		if cint(intArrayPointer) > -1 AND cint(intArrayPointer) <= ubound(strTmpHArray) then
			strSpreadSheetItem = replace(strTmpHArray(tmpArrayPointer(intArrayPointer)), Chr(34), "")
		end if
	end if

end if
ReturnSpreadSheetItem = strSpreadSheetItem
End Function

Function returnCellLocation(strQuotedLine, cellNumber) 'needed to support mixed quoted non-quoted csv
dim StrReturnCellL
  strTmpHArray = split(strQuotedLine, strDelimiter)
  redim tmpArrayPointer(ubound(strTmpHArray))
  boolQuoted = False
  intArrayCount = 0
  for cellCount = 0 to ubound(strTmpHArray)
	if boolQuoted = False then 
		tmpArrayPointer(intArrayCount) = cellCount
		if cellNumber = intArrayCount then StrReturnCellL = cellCount
		intArrayCount = intArrayCount + 1 
	end if

	if instr(strTmpHArray(cellCount),chr(34)) > 0 then 
		if boolQuoted = False and left(strTmpHArray(cellCount), 1) = chr(34) and right(strTmpHArray(cellCount),1) = chr(34) then
			boolQuoted = False
		elseif boolQuoted = True and right(strTmpHArray(cellCount), 1) = chr(34) then 
			boolQuoted = False
		elseif boolQuoted = False and left(strTmpHArray(cellCount), 1) = chr(34) then
			boolQuoted = True
		else
			'ignore quotes that aren't at the begening or end 
		end if
	end if
  next
returnCellLocation = StrReturnCellL  
end Function

Sub SetHeaderLocations(StrHeaderText) 'sets the integer location for the header text
if instr(StrHeaderText, strDelimiter) then
  if instr(StrHeaderText, strDelimiter) then 
    strTmpHArray = split(StrHeaderText, strDelimiter)
  else
    MsgBox "missing delimiter. Script will now exit"
    WScript.Quit (4)
  end if
  for inthArrayLoc = 0 to ubound(strTmpHArray)
	strCellData = ReturnSpreadSheetItem(StrHeaderText, inthArrayLoc)
	If boolCaseSensitive = False Then
	strCellData = LCase(strCellData)
	End If

	DictHeader.item(strCellData) = inthArrayLoc

  next
else
  Msgbox "error parsing header: " & StrHeaderText
end if
end sub


Function SelectFile( )
    ' File Browser via HTA
    ' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
    ' Features: Works in Windows Vista and up (Should also work in XP).
    '           Fairly fast.
    '           All native code/controls (No 3rd party DLL/ XP DLL).
    ' Caveats:  Cannot define default starting folder.
    '           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
    '           Dialog title says "Choose file to upload".
    ' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&?lig;-4ba3-bca5-ec349df65ef6

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
