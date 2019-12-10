'Cb Pull Events v1.4.7 - Add registry watchlist column output. Fix column allignment for cross proceess
'Pulls event data from the Cb Response API and dumps to CSV. 
'Pass the query as a parameter to the script.
'Enclose entire query in double quotes (")
'/a argument to auto accept pulling down all results.
'/b to baseline. Add letters after the "b" to tell it what to baseline. 
'/bmnc "m" - modules. "n" - network. "c" - cross process


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

Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1
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
Dim dictFileAction: Set dictFileAction = CreateObject("Scripting.Dictionary")
Dim dictChild: Set dictChild = CreateObject("Scripting.Dictionary")
Dim dictUID: Set dictUID = CreateObject("Scripting.Dictionary")
Dim dictBaselineMod: Set dictBaselineMod = CreateObject("Scripting.Dictionary") 
Dim dictBaselineCross: Set dictBaselineCross = CreateObject("Scripting.Dictionary")
Dim dictBaselineChild: Set dictBaselineChild = CreateObject("Scripting.Dictionary")
Dim dictBaselineNetwork: Set dictBaselineNetwork = CreateObject("Scripting.Dictionary")
Dim dictBaselineFile: Set dictBaselineFile = CreateObject("Scripting.Dictionary")
Dim boolBaselineMod
Dim boolBaselineCross
Dim boolBaselineChild
Dim boolBaselineNetwork
Dim boolBaselineFile
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
Dim boolUseBaseline
Dim BoolWatchLlistRegex
Dim boolRegWatchlist 'use watchlist
Dim DictRegWatchlist: Set DictRegWatchlist = CreateObject("Scripting.Dictionary")
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")

CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug"

'Optional config section
APIVersion = 4
boolNetworkEnable = True
boolRegEnable = True
boolModEnable = True
boolChildEnable = True
boolFileEnable = True
boolCrossEnable = True
pullAllSections = True 'set to true to grab everything
boolReportUserName = True 'Include associated user name
boolReportProcessName = True 'Include associated process name
boolReportParentName = True 'Include parent name
boolUseBaseline = True 'Exclude items in baseline from reporting
BoolWatchLlistRegex = False 'use regex for matching
boolRegWatchlist = True 'Use registry watchlist
strCbQuery = "" 'Cb Response query to run. Can be passed as an argument to the script.
intSleepDelay = 1000 'delay between queries
intPagesToPull = 1000 'Number of alerts to retrieve at a time
intReceiveTimeout = 120 'number of seconds for timeout
intClippingLevel = 40000 'Stop pulling results for query after hitting this amount.
strReportPath = "\Reports" 'directory to write report output
strBaselinePath = "\baselines" 'baseline save directory
strWatchlistFolder = "\data" 'location to where watchlist are stored
strRegWatchlist = "RegWatch.txt" 'Registry watchlist name
boolUseSocketTools = False 'Uses external library from SocketTools (needed when using old OS that does not support latest TLS standards)
strLicenseKey = "" 'License key is required to use SocketTools 
strIniPath="Cb_pe.ini"
'end config section

if objFSO.FileExists(strIniPath) = false then
	If InStr(strIniPath, "\") = 0 Then 
		strIniPath = CurrentDirectory & "\" & strIniPath
	End If
End if		

strReportPath = UpdatePath(strReportPath) 'add currentdirectory to path
strBaselinePath = UpdatePath(strBaselinePath) 'add currentdirectory to path
strRegWatchlist = UpdatePath(strWatchlistFolder & "\" & strRegWatchlist) 'add currentdirectory to path


if objFSO.FileExists(strIniPath) = True then
'---Ini loading section
intSleepDelay = ValueFromINI(strIniPath, "IntegerValues", "SleepDelay", intSleepDelay)
intPagesToPull = ValueFromINI(strIniPath, "IntegerValues", "PagesToPull", intPagesToPull)
intReceiveTimeout = ValueFromINI(strIniPath, "IntegerValues", "ReceiveTimeout", intReceiveTimeout)
APIVersion = ValueFromINI(strIniPath, "IntegerValues", "APIVersion", APIVersion)
boolUseSocketTools = ValueFromINI(strIniPath, "BooleanValues", "UseSocketTools", boolUseSocketTools)
boolNetworkEnable = ValueFromINI(strIniPath, "BooleanValues", "Network", boolNetworkEnable)
boolModEnable = ValueFromINI(strIniPath, "BooleanValues", "Modules", boolModEnable)
boolChildEnable = ValueFromINI(strIniPath, "BooleanValues", "Child", boolChildEnable)
boolFileEnable = ValueFromINI(strIniPath, "BooleanValues", "File", boolFileEnable)
boolCrossEnable = ValueFromINI(strIniPath, "BooleanValues", "Cross", boolCrossEnable)
boolRegEnable = ValueFromINI(strIniPath, "BooleanValues", "Registry", boolRegEnable)
boolRegWatchlist = ValueFromINI(strIniPath, "BooleanValues", "RegistryWatch", boolRegWatchlist)
pullAllSections = ValueFromINI(strIniPath, "BooleanValues", "AllSections", pullAllSections)
boolReportUserName = ValueFromINI(strIniPath, "BooleanValues", "ReportUserName", boolReportUserName)
boolReportProcessName = ValueFromINI(strIniPath, "BooleanValues", "ReportProcessName", boolReportProcessName)
boolDebug = ValueFromINI(strIniPath, "BooleanValues", "Debug", boolDebug)	
'---End ini loading section
else
	if BoolRunSilent = False then WScript.Echo strIniPath & " does not exist. Using script configured/default settings instead"
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
if objFSO.folderexists(strBaselinePath) = False then _
objFSO.createfolder(strBaselinePath)
if objFSO.folderexists(strWatchlistFolder) = False then _
objFSO.createfolder(strWatchlistFolder)


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


if WScript.Arguments.count < 1 then
  wscript.echo "No query parameter passed. Pass a CB query to the script as a argument"
  wscript.quit
end if

if WScript.Arguments(0) = "" and strCbQuery = "" then
  wscript.echo "No query parameter passed. Pass a CB query to the script as a argument"
  wscript.quit
else
	boolQuerySet = False
	if strCbQuery <> "" then
		boolQuerySet = True
	end if
	for each passedArg in WScript.Arguments
		if lcase(passedArg) = "/a" and intAnswer = "" then
			intAnswer = VbYes
		elseif lcase(left(passedArg,2)) = "/b" Then
      if instr(lcase(passedArg),"m") > 0 or boolUseBaseline = True then 
        boolBaselineMod = True
        LoadBaseline strBaselinePath & "\Modules" & ".dat",dictBaselineMod 
      end if
      if instr(lcase(passedArg),"c") > 0 or boolUseBaseline = True then 'crossprocess
        boolBaselineCross = True
        LoadBaseline strBaselinePath & "\CrossProc" & ".dat", dictBaselineCross
      end if
      if instr(lcase(passedArg),"p") > 0 or boolUseBaseline = True then 'Child Process
        boolBaselineChild = True
        LoadBaseline strBaselinePath & "\Child" & ".dat", boolBaselineChild
      end if
      if instr(lcase(passedArg),"n") > 0 or boolUseBaseline = True then 
        LoadBaseline strBaselinePath & "\Network" & ".dat",dictBaselineNetwork 
        boolBaselineNetwork = True
      end if
	  if instr(lcase(passedArg),"f") > 0 or boolUseBaseline = True then 
        LoadBaseline strBaselinePath & "\File" & ".dat",dictBaselineFile 
        boolBaselineFile = True
      end if
		elseif boolQuerySet = True then
			msgbox "A query was provided inside the script. If you want to use an external query then clear boolQuerySet in the Optional config section. Script will now exit. strCbQuery=" & strCbQuery
			wscript.quit(22)	  
		elseif strCbQuery <> "" then
			if instrRev(strCbQuery,":") > 0 then

				if instr(right(strCbQuery,  len(strCbQuery) - instrRev(strCbQuery,":")), " ") = 0 then
				  if instrRev(strCbQuery," ") > 0 then
					if instr(passedArg," ") = 0 then
					  if passedArg = WScript.Arguments(WScript.Arguments.count -1) then
						strCbQuery = left(strCbQuery, instrRev(strCbQuery,":")) & chr(34) & right(strCbQuery, len(strCbQuery) - instrRev(strCbQuery,":"))
						strCbQuery = strCbQuery & " " & strSecondWord & passedArg & chr(34)
					  elseif strSecondWord = "" then
						strSecondWord = passedArg
					  else
						strSecondWord = strSecondWord & " " & passedArg
					  end if
					else				
					  strSecondWord = strSecondWord & " " & left(passedArg, instr(passedArg," ")-1)
					  if len(passedArg) > 0 and instr(strSecondWord, ":") = 0 then
						strCbQuery = left(strCbQuery, instrRev(strCbQuery,":")) & chr(34) & right(strCbQuery, len(strCbQuery) - instrRev(strCbQuery,":"))
						strCbQuery = strCbQuery & " " & strSecondWord & chr(34) & right(passedArg, len(passedArg) - instr(passedArg," ") + 1)
					  end if
					end if
				  end if
				end if
		    end if
		else
		  strCbQuery = passedArg
		end if
		
	next
end if

if boolRegWatchlist = True then LoadCustomDict strRegWatchlist, DictRegWatchlist

msgbox "executing query: " & strCbQuery
logdata strReportPath & "\Query_" & strUnique & ".txt",strCbQuery, false
CbQuery strCbQuery



Function CbQuery(strQuery)
Dim intParseCount: intParseCount = 10
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
						  CBEventData strCBID & "/" & HexToDec(right(strSeg, 12)), strquery
						end if
					  
						'dictUID.add strCBID & "-" & strCBSegID, ""
					  'end if
					 next
					end if
				end if
				if strCBSegID <> "" then
				 'segment_id: REQUIRED the process segment id; this is the segment_id field in search results. If this is set to 0
				  CBEventData strCBID & "/" & strCBSegID, strquery
				end if
			end if
		next
	end if
	intResultCount = getdata(StrTmpResponse, ",", "total_results" & Chr(34) & ": ")

	if isnumeric(intResultCount) then
		if intResultCount = 0 then
		  wscript.echo "Zero items were retrieved. Please double check your query and try again: " & chr(34) & strCbQuery & chr(34)
		  wscript.quit (997)
		end if
		if intParseCount >= clng(intResultCount) then
		  wscript.echo intResultCount & " items retrieved for query " & chr(34) & strCbQuery & chr(34)
		  wscript.quit
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


Function CBEventData(strIDPath,strQuery)
Set objHTTP_ED = CreateObject("WinHttp.WinHttpRequest.5.1")
Dim strAVEurl
Dim strReturnURL
dim strAssocWith
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
if boolReportParentName = True then
  parentprocessname = "," & Chr(34) & getdata(StrTmpResponse,Chr(34), "parent_name" & CHr(34) & ": " & Chr(34)) & chr(34)
  parentNheader = "|Parent Name"
end if
strTmp_host_type = getdata(StrTmpResponse, Chr(34), "host_type" & CHr(34) & ": " & Chr(34))
strTmp_group = getdata(StrTmpResponse, Chr(34), "group" & CHr(34) & ": " & Chr(34))
strTmp_fork_children_count = getdata(StrTmpResponse,",", "fork_children_count" & CHr(34) & ": " )
strTmp_processblock_count = getdata(StrTmpResponse,",", "processblock_count" & CHr(34) & ": " )
if boolEventHeader = False then
  outrow = "EMET|Execute|Network|Suspicious Indicators|SegmentID|Host Type|Group|Children|Blocked Process|CMD"
  logdata strReportPath & "\Event_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
  boolEventHeader = True
end if
logdata strReportPath & "\Event_out_" & strUnique & ".csv", chr(34) & strTmpemet_count & Chr(34) & "," & chr(34) & strTmpexec_events_count & Chr(34) & "," & chr(34) & strTmp_netconn_count & Chr(34) & "," & chr(34) & _
strTmp_alliance_score_bit9suspiciousindicators & Chr(34) & ","& chr(34) & strTmp_segment_id & Chr(34) & "," & chr(34) & strTmp_host_type & Chr(34) & "," & chr(34) & strTmp_group & Chr(34) & "," & chr(34) & _
strTmp_fork_children_count & Chr(34) & ","& chr(34) & strTmp_processblock_count & Chr(34) & ","& chr(34) & strTmpCmd & Chr(34), false
process_pid = getdata(StrTmpResponse,",", "process_pid" & CHr(34) & ": " )




if boolNetworkEnable = True and APIVersion  > 1 and APIVersion < 5 then
  if boolNetworkHeader = False then
	outrow = "Date Time|IP Address|Local Port|Remote Port|Protocol|Domain|Outbound|Sensor ID" & userNheader & processNheader
	logdata strReportPath & "\IP_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolNetworkHeader = True
  end if
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
    strWriteLine = Chr(34) & strDtime & chr(34) & "," & Chr(34) & strIP & chr(34) & "," & _
    Chr(34) & strLport & chr(34) & "," & Chr(34) & strRport & chr(34) & "," & Chr(34) & strProtocol & chr(34) & "," & _
    Chr(34) & strDomain & chr(34) & "," & Chr(34) & strDirection & chr(34) & "," & Chr(34) & sensor_id & Chr(34) & username & processname
    if boolBaselineNetwork = False then
	  if boolUseBaseline = False or dictBaselineNetwork.exists(strquery & "|" & strIP & "|" & strDomain & "|" & strDirection) = False then
		logdata strReportPath & "\IP_out_" & strUnique & ".csv",strWriteLine, false
	  end if
    elseif boolBaselineNetwork = True then
      if dictBaselineNetwork.exists(strquery & "|" & strIP & "|" & strDomain & "|" & strDirection) = False then
          dictBaselineNetwork.add strquery & "|" & strIP & "|" & strDomain & "|" & strDirection, ""
          logdata strReportPath & "\IP_out_" & strUnique & ".csv",strWriteLine, false
          logdata strBaselinePath & "\Network" & ".dat",strquery & "|" & strIP & "|" & strDomain & "|" & strDirection, false
      end if
    end if
  end if
  next

end if

if boolNetworkEnable = True and APIVersion  = 1 then
  if boolNetworkHeader = False then
	outrow = "Date Time|IP Address|Remote Port|Protocol|Domain|Outbound|Sensor ID" & userNheader & processNheader
	logdata strReportPath & "\IP_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolNetworkHeader = True
  end if
  strTmpText = getdata(StrTmpResponse,"]", "netconn_complete" & CHr(34) & ": [") 
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
	   strWriteLine = Chr(34) & arrayEE(0) & Chr(34) & "," & Chr(34) & dotQuadIP & Chr(34) & "," & Chr(34) & arrayEE(2) & Chr(34) & "," & Chr(34) & arrayEE(3) & Chr(34) & "," & Chr(34) & arrayEE(4) & Chr(34) & "," & Chr(34) & arrayEE(5) & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & username & processname 
      if dictBaselineNetwork = False then
	   if boolUseBaseline = False or dictBaselineNetwork.exists(strquery & "|" & dotQuadIP & "|" & arrayEE(4) & "|" & arrayEE(5)) = False then
        logdata strReportPath & "\IP_out_" & strUnique & ".csv",strWriteLine, false
	   end if
      elseif dictBaselineNetwork = True then
        if dictBaselineNetwork.exists(strquery & "|" & dotQuadIP & "|" & arrayEE(4) & "|" & arrayEE(5)) = False then
          dictBaselineNetwork.add strquery & "|" & dotQuadIP & "|" & arrayEE(4) & "|" & arrayEE(5), ""
          logdata strReportPath & "\IP_out_" & strUnique & ".csv",strWriteLine, false
          logdata strBaselinePath & "\Network" & ".dat",strquery & "|" & dotQuadIP & "|" & arrayEE(4) & "|" & arrayEE(5), false
        end if
      end if
	  end if
	end if
  next
end if

if boolRegEnable = True then 
  if boolRegHeader = False then

	outrow = "Action|Date Time|Registry Key|Sensor ID" & userNheader & processNheader
	if boolRegWatchlist = True then outrow = outrow & "|Watchlist"
	logdata strReportPath & "\Reg_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolRegHeader = True
  end if
   strTmpText = getdata(StrTmpResponse,"]", "regmod_complete" & CHr(34) & ": [")
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 2 then
		strAction = ""
		if dictRegAction.exists(arrayEE(0)) then strAction =  dictRegAction.item(arrayEE(0))
		if boolRegWatchlist = True then 'replace the "\\" with \ in the registry event
      strWatchResult = "," & Chr(34) & MatchWatchList(replace(arrayEE(2), "\\","\"), DictRegWatchlist) & Chr(34) 
      
    else
      strWatchResult = ""
    end if
	   strWriteLine = Chr(34) & strAction & Chr(34) & "," & Chr(34) & arrayEE(1) & Chr(34) & "," & Chr(34) & arrayEE(2) & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & username & processname & strWatchResult
	   
	  logdata strReportPath & "\Reg_out_" & strUnique & ".csv",strWriteLine, false
	  end if
	end if
  next
end if



if boolModEnable = True then
  if boolModHeader = False then

	outrow = "Date Time|MD5|File Path|Sensor ID" & userNheader & processNheader
	logdata strReportPath & "\Mod_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolModHeader = True
  end if
   strTmpText = getdata(StrTmpResponse,"]", "modload_complete" & CHr(34) & ": [")
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  if right(tmpEvent,1) = "|" then tmpEvent = left(tmpEvent, len(tmpEvent) -1) 'remove end pipe as have not seen any values after it.
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 1 then
	   strWriteLine = chr(34) & replace(tmpEvent, "|", chr(34) & "," & Chr(34)) & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & username & processname
	   if boolBaselineMod = False then
		 if boolUseBaseline = False or dictBaselineMod.exists(strquery & "|" & arrayEE(1)) = False then
			logdata strReportPath & "\Mod_out_" & strUnique & ".csv",strWriteLine, false
		end if
      elseif boolBaselineMod = True then
        if dictBaselineMod.exists(strquery & "|" & arrayEE(1)) = False then
          dictBaselineMod.add strquery & "|" & arrayEE(1), arrayEE(2)
          logdata strReportPath & "\Mod_out_" & strUnique & ".csv",strWriteLine, false
          logdata strBaselinePath & "\Modules" & ".dat",strquery & "|" & arrayEE(1), false
        end if
      end if
	  end if
	end if
  next
end if

if boolChildEnable = True and APIVersion  >= 3 then 
  dictChild.RemoveAll
  if boolChildHeader = False then
	outrow = "Start Time|End Time|Unique ID|MD5|File Path|PID|Suppressed|Parent PID|Parent Unique ID|Sensor ID|Child Command Line" & userNheader & processNheader & parentNheader
	logdata strReportPath & "\Child_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolChildHeader = True
  end if    
  strTmpText = getdata(StrTmpResponse,"]", "childproc_complete" & CHr(34) & ": [")
  if instr(strTmpText,  "},") = 0 then strTmpText = strTmpText & "},"
  CbarrayEvents = split(strTmpText, "},")
  for each EventEntry in CbarrayEvents
	childMD5 = getdata(EventEntry,chr(34), "md5" & CHr(34) & ": " & chr(34) )
	childCommandLine = getdata(EventEntry,chr(34) & ",", "commandLine" & CHr(34) & ": " & chr(34) )
	childSha256 = getdata(EventEntry,chr(34), "sha256" & CHr(34) & ": " & chr(34) )
	childProcessId = getdata(EventEntry,chr(34), "processId" & CHr(34) & ": " & chr(34) )
	childIs_suppressed = getdata(EventEntry,",", "is_suppressed" & CHr(34) & ": " )
	childDateStartTime = getdata(EventEntry,chr(34), "start" & CHr(34) & ": " & chr(34) )
	childDateEndTime = getdata(EventEntry,chr(34), "end" & CHr(34) & ": " & chr(34) )
	childIs_tampered = getdata(EventEntry,",", "is_tampered" & CHr(34) & ": " )
	childPID = getdata(EventEntry,",", "pid" & CHr(34) & ": " )
	childIDPath = getdata(EventEntry,chr(34), "path" & CHr(34) & ": " & chr(34) )
	if childMD5 <> "" then
		
	   strWriteLine = Chr(34) & childProcessId & Chr(34) & _
	   "," & Chr(34) & childMD5 & Chr(34) & "," & Chr(34) & childIDPath & Chr(34) & _
	   "," & Chr(34) & childPID & Chr(34) & "," & Chr(34) & childIs_suppressed & Chr(34) & _
	   "," & Chr(34) & process_pid & Chr(34) & "," & Chr(34) & strIDPath & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & "," & Chr(34) & childCommandLine & Chr(34) & username & processname & parentprocessname
	   if dictChild.exists(strWriteLine) = False then 
		dictChild.add strWriteLine, childDateEndTime
	   else
	    if boolBaselineChild = False then
			if boolUseBaseline = False or dictBaselineChild.exists(strquery & "|" & childSha256 & "|" & childCommandLine) = false then
				logdata strReportPath & "\Child_out_" & strUnique & ".csv",Chr(34) & childDateStartTime & Chr(34) & "," & Chr(34) & dictChild.item(strWriteLine) & Chr(34) & "," & strWriteLine, false
			end if
		elseif boolBaselineChild = True then
			if dictBaselineChild.exists(strquery & "|" & childSha256 & "|" & childCommandLine) = false then
			  dictBaselineChild.add strquery & "|" & childSha256 & "|" & childCommandLine, ""
			  logdata strReportPath & "\Child_out_" & strUnique & ".csv",Chr(34) & childDateStartTime & Chr(34) & "," & Chr(34) & dictChild.item(strWriteLine) & Chr(34) & "," & strWriteLine, false
			  logdata strBaselinePath & "\Child" & ".dat",strquery & "|" & strquery & "|" & childSha256 & "|" & childCommandLine, false		
			end if
		
		end if
	   end if	
	end if
  next
end if

if boolChildEnable = True and APIVersion  < 3 then 
  if boolChildHeader = False then

	outrow = "Date Time|Unique ID|MD5|File Path|PID|Not Suppressed|Parent PID|Unique ID|Sensor ID" & userNheader & processNheader & parentNheader
	logdata strReportPath & "\Child_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolChildHeader = True
  end if    
  strTmpText = getdata(StrTmpResponse,"]", "childproc_complete" & CHr(34) & ": [")
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 4 then
	   strWriteLine = replace(tmpEvent, "|", chr(34) & "," & Chr(34)) & Chr(34) & "," & Chr(34) & process_pid & Chr(34) & "," & Chr(34) & strIDPath & Chr(34) & "," & Chr(34) & sensor_id & Chr(34) & username & processname & parentprocessname
	   
	  logdata strReportPath & "\Child_out_" & strUnique & ".csv",strWriteLine, false
	  end if
	end if
  next
end if



if boolFileEnable = True then 
  if boolFileHeader = False then

	outrow = "Action|Date Time|File Path|Last Write MD5|File Type|Tamper Attempt|Sensor ID" & userNheader & processNheader
	logdata strReportPath & "\File_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolFileHeader = True
  end if       
  strTmpText = getdata(StrTmpResponse,chr(34) & "], ", "filemod_complete" & CHr(34) & ": [")
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 4 then
		strAction = ""
		if dictRegAction.exists(arrayEE(0)) then strAction =  dictFileAction.item(arrayEE(0))
		strWriteLine = Chr(34) & strAction & Chr(34) & "," & Chr(34) & arrayEE(1) & Chr(34) & "," & Chr(34) & arrayEE(2) & Chr(34)  & "," & Chr(34) & arrayEE(3) & Chr(34)  & "," & Chr(34) & arrayEE(4) & Chr(34)  & "," & Chr(34) & arrayEE(5) & Chr(34)  & "," & Chr(34) & sensor_id & Chr(34) & username & processname
	    if boolBaselineChild = False then
			if boolUseBaseline = False or dictBaselineFile.exists(strquery & "|" & arrayEE(2) & "|" & arrayEE(3)) = false then
				logdata strReportPath & "\File_out_" & strUnique & ".csv",strWriteLine, false
			end if
		elseif dictBaselineFile = True then
			if dictBaselineFile.exists(strquery & "|" & arrayEE(2) & "|" & arrayEE(3)) = false then
			  dictBaselineFile.add strquery & "|" & arrayEE(2) & "|" & arrayEE(3), ""
			  logdata strReportPath & "\File_out_" & strUnique & ".csv",strWriteLine, false
			  logdata strBaselinePath & "\File" & ".dat",strquery & "|" & arrayEE(3), false		
			end if
		end if
	  else
		logdata CurrentDirectory & "\CB_Pull_Error.log", Date & " " & Time & " FileMod error splitting the value into an array size greater than four: " & tmpEvent,False 
	  end if
	end if
  next
end if

if boolCrossEnable = True then 
  if boolCrossHeader = False then

	outrow = "Action|Date Time|Target Unique ID|Target MD5|Target Path|Open Type|Access Requested|Tamper|Inbound Open|PID|Process Path|Unique ID|Sensor ID" & userNheader & processNheader
	logdata strReportPath & "\Cross_out_" & strUnique & ".csv", chr(34) & replace(outrow, "|", chr(34) & "," & Chr(34)) & Chr(34), false
	boolCrossHeader = True
  end if    
  strTmpText = getdata(StrTmpResponse,"]", "crossproc_complete" & CHr(34) & ": [")
  CbarrayEvents = split(strTmpText, ", ")
  for each EventEntry in CbarrayEvents
	if instr(EventEntry, "|") > 0 then 
	  tmpEvent = replace(EventEntry,chr(34), "")
	  ArrayEE = split(tmpEvent, "|")
	  if ubound(arrayEE) > 1 then
      if right(tmpEvent, 1) = "|" and ubound(arrayEE) >7 then 'not sure what this extra column is for 
        tmpEvent = left(tmpEvent, len(tmpEvent) -1) 'remove extra column
      end if
      accessRequested = ArrayEE(6) '!need to add interpretation for the numbers in this variable
	   strWriteLine = chr(34) & replace(tmpEvent, "|", chr(34) & "," & Chr(34)) & Chr(34) & "," & Chr(34) & process_pid & Chr(34) & "," & Chr(34) & strTmp_ExePath & Chr(34) & "," & Chr(34) & strIDPath & Chr(34)  & "," & Chr(34) & sensor_id & Chr(34) & username & processname
      if boolBaselineCross = False then
		if boolUseBaseline = False or dictBaselineCross.exists(strquery & "|" & arrayEE(3)) = false then
			logdata strReportPath & "\Cross_out_" & strUnique & ".csv",strWriteLine, false
		end if
      elseif boolBaselineCross = True then
        if dictBaselineCross.exists(strquery & "|" & arrayEE(3)) = false then
          dictBaselineCross.add strquery & "|" & arrayEE(3), ""
          logdata strReportPath & "\Cross_out_" & strUnique & ".csv",strWriteLine, false
          logdata strBaselinePath & "\CrossProc" & ".dat",strquery & "|" & arrayEE(3), false
        end if
      end if
    end if
   end if 
  next
end if

strTmpText = getdata(StrTmpResponse,"]", "exec_events" & CHr(34) & ": [")
arrayEvents = split(strTmpText, ", ")
for each EventEntry in arrayEvents
  if instr(EventEntry, "|") > 0 then 
	ArrayEE = split(replace(EventEntry,chr(34), ""), "|")
	if ubound(arrayEE) > 4 then
	  logdata strReportPath & "\proc_guid_" & strUnique & ".txt",EventEntry, false

	  CBEventData arrayEE(1) & "/1"
	
	end if
  end if
next

else
logdata CurrentDirectory & "\CB_Download.log", Date & " " & Time & " Event can't be retrieved - " & strIDPath,False 
wscript.sleep 5
end if

set objHTTP_ED = nothing
end function

Function SaveBinaryDataTextStream(FileName, responseBody)
set oStream = createobject("Adodb.Stream")
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const adSaveCreateNotExist = 1 

oStream.type = adTypeBinary
oStream.open
oStream.write responseBody

' Do not overwrite an existing file
oStream.savetofile FileName, adSaveCreateNotExist

' Use this form to overwrite a file if it already exists
' oStream.savetofile DestFolder & ImageFile, adSaveCreateOverWrite

oStream.close

set oStream = nothing
Set xml = Nothing
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

Sub LoadBaseline(strListPath, dictToLoad)
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


Sub LoadCustomDict(strListPath, dictToLoad)
if objFSO.fileexists(strListPath) then
  Set objFile = objFSO.OpenTextFile(strListPath)
  Do While Not objFile.AtEndOfStream
    if not objFile.AtEndOfStream then 'read file
        On Error Resume Next
        strData = objFile.ReadLine
        if instr(strData, "|") then
          strTmpArrayDDNS = split(strData, "|")
          if dictToLoad.exists(lcase(strTmpArrayDDNS(0))) = False then _
          dictToLoad.add lcase(strTmpArrayDDNS(0)), strTmpArrayDDNS(1)
        else
          if dictToLoad.exists(lcase(strData)) = False then _
          dictToLoad.add lcase(strData), ""
        end if
        on error goto 0
    end if
  loop
end if
end sub


Function MatchWatchList (strWLcheck, dictMatchWatchList) 
Dim WLreturnValue
WLreturnValue = ""
'msgbox "strWLstoredResults=" & strWLstoredResults
if dictMatchWatchList.count = 0 then exit function
for each WatchItem in dictMatchWatchList
	'msgbox "WatchItem=" & WatchItem
  if BoolWatchLlistRegex = True then
    Set re = new regexp  'Create the RegExp object 'more info at https://msdn.microsoft.com/en-us/library/ms974570.aspx

    re.Pattern = WatchItem
    re.IgnoreCase = true
    WLRegXresult = re.Test(strWLcheck)
	'msgbox "regex match=" & WLRegXresult & " for " & WatchItem
    if WLRegXresult = True then
      WLreturnValue = dictMatchWatchList.item(WatchItem)
      exit for
    end if
  else
    'msgbox strWLcheck & " | " & WatchItem
    if instr(strWLcheck, WatchItem) > 0 then
      'msgbox dictMatchWatchList.item(WatchItem)
      WLreturnValue = dictMatchWatchList.item(WatchItem)
      exit for
    end if
  end if
next
MatchWatchList = WLreturnValue
end function

Function UpdatePath(strPath)
if instr(strPath, ":") = 0 then 
	strPath = CurrentDirectory & "\" & strPath
end if
UpdatePath = strPath
end function