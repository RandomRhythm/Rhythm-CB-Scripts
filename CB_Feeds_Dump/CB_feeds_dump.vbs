'CB Feed Dump v4.3 'Added check for CVE-2017-11826
'Pulls data from the CB Response feeds and dumps to CSV. Will pull parent and child data for the process alerts in the feeds.

'additional queries can be run via aq.txt in the current directory.
'name|query
'Example:
'knowndll|/api/v1/binary?q=observed_filename:known.dll&digsig_result:Unsigned

'More information on querying the CB Response API https://github.com/carbonblack/cbapi/tree/master/client_apis

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
Dim Dicthash: set Dicthash  = CreateObject("Scripting.Dictionary")'
Dim DictAdhocQuery: set DictAdhocQuery  = CreateObject("Scripting.Dictionary")'
Dim DictChildQuery: set DictChildQuery  = CreateObject("Scripting.Dictionary")'
Dim DictLimitedOut: set DictLimitedOut  = CreateObject("Scripting.Dictionary")'
Dim DictAdditionalQueries: set DictAdditionalQueries  = CreateObject("Scripting.Dictionary")'
Dim boolHeaderWritten
Dim boolEchoInfo
dim boolEnableabusech
dim boolEnablealienvault
dim boolEnableBit9AdvancedThreats
dim boolEnableBit9EndpointVisibility
dim boolEnableBit9SuspiciousIndicators
dim boolEnablecbbanning
dim boolEnablecbemet
dim boolEnablecbtamper
dim boolEnablefbthreatexchange
dim boolEnableiconmatching
dim boolEnablemdl
dim boolEnableNVD
dim boolEnablesans
dim boolEnableSRSThreat
dim boolEnableSRSTrust
dim boolEnableThreatConnect
dim boolEnabletor
dim boolEnableVirusTotal
Dim strFlashVersion
Dim strCachePath
Dim boolEnableNetAPI32Check
Dim boolEnableFlashCheck
Dim boolEnableMshtmlCheck
Dim boolEnableSilverlightCheck
Dim boolEnableIexploreCheck
Dim strStaticFPversion
Dim boolEnableOptivCheck
Dim boolEnableCbKnownIOCsCheck
Dim boolEnableCbFileAnalysisCheck
Dim BoolEnableCbCommunityCheck
Dim BoolEnableBit9EarlyAccessCheck
Dim boolDebugVersionCompare
Dim boolDebugFlash
Dim boolEnableYARA
Dim boolEnableCbInspection
Dim boolMS17010Check
Dim yaraFeedID
Dim tmpYaraUID
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim dictYARA: Set dictYARA = CreateObject("Scripting.Dictionary")
Dim intParseCount: intParseCount = 10
Dim BoolDebugTrace
Dim boolCVE_2017_11826

'debug
BoolDebugTrace = False
boolDebugFlash = False
boolDebugVersionCompare = False
'end debug

'---Query Config Section
boolEchoInfo = False 
IntDayStartQuery = "*" 'days to go back for start date of query. Set to "*" to query all binaries or set to -24 to query last 24 time measurement
IntDayEndQuery = "*" 'days to go back for end date of query. Set to * for no end date
strTimeMeasurement = "d" '"h" for hours "d" for days
strHostFilter = "" 'computer name to filter to. Typically uppercase and is case sensitive.
'---End Query Config Section


'---Script Settings
boolAddYARAtoReports = True 'Combines binary reports to include the YARA rules column
boolEnableabusech = True
boolEnablealienvault = True
boolEnableBit9AdvancedThreats = True
boolEnableBit9EndpointVisibility = True
boolEnableBit9SuspiciousIndicators = True
boolEnablecbbanning = True
boolEnablecbemet = True
boolEnablecbtamper = True
boolEnablefbthreatexchange = True
boolEnableiconmatching = True
boolEnablemdl = True
boolEnableNVD = True
boolEnablesans = True
boolEnableSRSThreat = True
boolEnableSRSTrust = True
boolEnableThreatConnect = True
boolEnabletor = True
boolEnableVirusTotal = True
boolEnableNetAPI32Check = True
boolEnableFlashCheck = True
boolEnableMshtmlCheck = True
boolEnableSilverlightCheck = True
boolEnableIexploreCheck = True
boolEnableOptivCheck = True
boolEnableCbKnownIOCsCheck = True
boolEnableCbFileAnalysisCheck = True
BoolEnableCbCommunityCheck = True
BoolEnableBit9EarlyAccessCheck = True
bool3155533Check = True
boolAdditionalQueries = True
boolEnableYARA = True
boolEnableCbInspection = True
boolMS17010Check = True
boolCVE_2017_11826 = True
strStaticFPversion = "28.0.0.137"
'strLTSFlashVersion = "18.0.0.383" 'support ended October 11, 2016 with version 18.0.0.382 
'---End script settings section

if strHostFilter <> "" then 
  msgbox "filtering to host " & strHostFilter
  strHostFilter = " AND hostname:" & strHostFilter
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
end if

CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug"
strSSfilePath = CurrentDirectory & "\CBIP_" & udate(now) & ".csv"
strCachePath = CurrentDirectory & "\cache"

strRandom = "4bv3nT9vrkJpj3QyueTvYFBMIvMOllyuKy3d401Fxaho6DQTbPafyVmfk8wj1bXF" 'encryption key. Change if you want but can only decrypt with same key
Set objFSO = CreateObject("Scripting.FileSystemObject")
'create sub directories
if objFSO.folderexists(CurrentDirectory & "\Debug") = False then _
objFSO.createfolder(CurrentDirectory & "\Debug")
if objFSO.folderexists(strDebugPath) = False then _
objFSO.createfolder(strDebugPath)
if objFSO.folderexists(strCachePath) = False then _
objFSO.createfolder(strCachePath)

' Store the arguments in a variable:
 Set objArgs = Wscript.Arguments
 For Each strArg in objArgs
     
    if strAdditionalQueryPath = "" then
      if objFSO.fileexists(strArg) then
        strAdditionalQueryPath = lcase(strArg)
      else
        msgbox "invalid argument: " & strArg
        
      end if
    else
      msgbox "invalid argument: " & strArg
    end if
Next

if strAdditionalQueryPath = "" and objFSO.fileexists(CurrentDirectory &"\aq.txt") then
  strAdditionalQueryPath = CurrentDirectory &"\aq.txt"
else
  boolAdditionalQueries = False
end if

if boolAdditionalQueries = True then
  'load additional queries
  if objFSO.fileexists(strAdditionalQueryPath) then
    Set objFile = objFSO.OpenTextFile(strAdditionalQueryPath)
    Do While Not objFile.AtEndOfStream
      if not objFile.AtEndOfStream then 'read file
          On Error Resume Next
          strData = objFile.ReadLine 
          if instr(strData, "|") then
            strTmpArrayAQ = split(strData, "|")
            if DictAdditionalQueries.exists(lcase(strTmpArrayAQ(0))) = False then 
				if instr(strTmpArrayAQ(1), "?") > 0 and instr(strTmpArrayAQ(1), "/") then
					DictAdditionalQueries.add lcase(strTmpArrayAQ(0)), strTmpArrayAQ(1)
				else
					msgbox "invalid additional query: " &  strData
				end if
			end if
          end if
          on error goto 0
      end if
    loop
  end if
end if

strFlashVersion = ReturnLatestFlashVer
if boolDebugFlash = true then msgbox "flash version:" & strFlashVersion


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

    msgbox "invalid api key"
    wscript.quit(999)
end if
strCarBlackAPIKey = strTempAPIKey



intTotalQueries = 50
'get feed info
DumpCarBlack 0, False, intTotalQueries, "/api/v1/feed"

if boolEnableNetAPI32Check = True then DictFeedInfo.Add "netapi32.dll", "netapi32.dll"
if boolEnableFlashCheck = True then DictFeedInfo.Add "Flash Player", "Flash Player"
if boolEnableMshtmlCheck = True then DictFeedInfo.Add "mshtml.dll", "mshtml.dll"
if boolEnableSilverlightCheck = True then DictFeedInfo.Add "silverlight", "silverlight"
if boolEnableIexploreCheck = True then DictFeedInfo.Add "iexplore.exe", "iexplore.exe"
if bool3155533Check = True then DictFeedInfo.Add "MS16-051", "vbscript.dll"
if boolMS17010Check = true then DictFeedInfo.Add "MS17-070", "srv.sys"
if boolCVE_2017_11826 = True then DictFeedInfo.Add "Microsoft Word", "winword.exe"
if boolAdditionalQueries = True then 
  for each strAquery in DictAdditionalQueries
    if DictFeedInfo.exists(DictAdditionalQueries.item(strAquery)) = False then DictFeedInfo.Add DictAdditionalQueries.item(strAquery), strAquery
  next
end if  
  
for each strCBFeedID in DictFeedInfo
  'msgbox strCBFeedID & "|" & DictFeedInfo.item(strCBFeedID)
  strQueryFeed = ""
  strCBFeedName = DictFeedInfo.item(strCBFeedID)
  select case strCBFeedName
    case "VirusTotal"
      if boolEnableVirusTotal = True then strQueryFeed = "/api/v1/binary?q=alliance_score_virustotal:*"
    case "SRSTrust"
      if boolEnableSRSTrust = True then strQueryFeed = "/api/v1/binary?q=alliance_score_srstrust:*"
    case "SRSThreat"
     if boolEnableSRSThreat = True then strQueryFeed = "/api/v1/binary?q=alliance_score_srsthreat:*"
    case "abusech"
      if boolEnableabusech = True then strQueryFeed = "/api/v1/process?q=alliance_score_abusech:*"
    case "cbbanning"
      if boolEnablecbbanning = True then strQueryFeed = "/api/v1/binary?q=alliance_score_cbbanning:*"      
    case "Bit9EndpointVisibility"
      if boolEnableBit9EndpointVisibility = True then strQueryFeed = "/api/v1/binary?q=alliance_score_bit9endpointvisibility:*"
    case "alienvault"
      if boolEnablealienvault = True then strQueryFeed = "/api/v1/process?q=alliance_score_alienvault:*"
    case "fbthreatexchange"
      if boolEnablefbthreatexchange = True then strQueryFeed = "/api/v1/process?q=alliance_score_fbthreatexchange:*"
    case "iconmatching"
      if boolEnableiconmatching = True then strQueryFeed = "/api/v1/binary?q=alliance_score_iconmatching:*"
    case "sans"
      if boolEnablesans = True then strQueryFeed = "/api/v1/process?q=alliance_score_sans:*"            
    case "NVD"
      if boolEnableNVD = True then strQueryFeed = "/api/v1/binary?q=alliance_score_nvd:*"
    case "cbemet"
      if boolEnablecbemet = True then strQueryFeed = "/api/v1/process?q=alliance_score_cbemet:*"  
    case "cbtamper"
      if boolEnablecbtamper = True then strQueryFeed = "/api/v1/process?q=alliance_score_cbtamper:*"
    case "mdl"
      if boolEnablemdl = True then strQueryFeed = "/api/v1/process?q=alliance_score_mdl:*"
    case "ThreatConnect"
      if boolEnableThreatConnect = True then strQueryFeed = "/api/v1/process?q=alliance_score_threatconnect:*"
    case "tor"
      if boolEnabletor = True then strQueryFeed = "/api/v1/process?q=alliance_score_tor:*"
    case "Bit9AdvancedThreats"
      if boolEnableBit9AdvancedThreats = True then strQueryFeed = "/api/v1/process?q=alliance_score_bit9advancedthreats:*"
    case "Bit9SuspiciousIndicators"
      if boolEnableBit9SuspiciousIndicators = True then strQueryFeed = "/api/v1/process?q=alliance_score_bit9suspiciousindicators:*"
    Case "OptivizedIntelFeedDomain"
      if boolEnableOptivCheck = True then strQueryFeed = "/api/v1/process?q=alliance_score_optivizedintelfeeddomain:*"
    Case "OptivizedIntelFeedIP"
      if boolEnableOptivCheck = True then strQueryFeed = "/api/v1/process?q=alliance_score_optivizedintelfeedip:*"
    Case "CbKnownIOCs"
      if boolEnableCbKnownIOCsCheck = True then strQueryFeed = "/api/v1/process?q=alliance_score_cbknowniocs:*"
    Case "CbFileAnalysis"
      if boolEnableCbFileAnalysisCheck = True then strQueryFeed = "/api/v1/binary?q=alliance_score_cbfileanalysis:*"
    Case "CbCommunity"
      if BoolEnableCbCommunityCheck = True then strQueryFeed = "/api/v1/process?q=alliance_score_cbcommunity:*"
    Case "Bit9EarlyAccess"
      if BoolEnableBit9EarlyAccessCheck = True then strQueryFeed = "/api/v1/binary?q=alliance_score_bit9earlyaccess:*"
	Case "yara"
      if boolEnableYARA = True then strQueryFeed = "/api/v1/binary?q=alliance_score_yara:*"	  
	Case "CbInspection"
      if boolEnableCbInspection = True then strQueryFeed = "/api/v1/binary?q=alliance_score_cbinspection:*"	
	  case "Flash Player"
      strQueryFeed = "/api/v1/binary?q=flash&digsig_publisher:Adobe  Systems  Incorporated"
    case "mshtml.dll"
      strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "mshtml.dll" & chr(34) & "&digsig_publisher:Microsoft Corporation"
    case "netapi32.dll"
      strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "netapi32.dll" & chr(34) & "&digsig_publisher:Microsoft Corporation"
    case "silverlight"
      strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "silverlight.configuration.exe" & chr(34) & "& digsig_publisher:Microsoft Corporation"
    Case "iexplore.exe"
      strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "iexplore.exe" & chr(34) & "& digsig_publisher:Microsoft Corporation"
    Case "MS16-051"
      strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "vbscript.dll" & chr(34) & "& digsig_publisher:Microsoft Corporation"
	Case "MS17-070"
      strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "srv.sys" & chr(34) & "& digsig_publisher:Microsoft Corporation"
	Case "winword.exe"
	  strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "winword.exe" & chr(34) & "& digsig_publisher:Microsoft Corporation"
    Case else
      if DictAdditionalQueries.exists(strCBFeedName) then 
        strQueryFeed = strCBFeedID
      end if
  end select
  if strQueryFeed <> "" then
	wscript.sleep 10 
    if instr(strQueryFeed, "/api/v1/binary?q=") > 0 and (boolEnableYARA = True or boolAddYARAtoReports = True) and dictYARA.count  = 0 then
		CbFeedQuery "feed_id:" & yaraFeedID, "YARA"
		if dictYARA.count  = 0  then 
			'wscript.echo "Nothing returned from YARA feed so disabling it."
			boolAddYARAtoReports = False
			boolEnableYARA = False
		end if
	end if
	wscript.sleep 10
    intTotalQueries = 10
    intTotalQueries = DumpCarBlack(0, False, intTotalQueries, strQueryFeed)
    logdata CurrentDirectory & "\CB_Feeds.log", date & " " & time & ": " & "Total number of items being retrieved for feed " & DictFeedInfo.item(strCBFeedID) & ": " & intTotalQueries ,boolEchoInfo
    'DumpCarBlack 0, True, intTotalQueries
    'DumpCarBlack 0, True, 10
    boolHeaderWritten = False
    if clng(intTotalQueries) > 0 then
      intCBcount = 0
      if BoolDebugTrace = True then logdata strDebugPath & "\CarBlacktext" & "" & ".txt", strCBFeedID & vbcrlf & "-------" & vbcrlf,BoolEchoLog 
      strUniquefName = DictFeedInfo.item(strCBFeedID) & "_" & udate(now) & ".csv"
      strHashOutPath = CurrentDirectory & "\CBmd5_" & strUniquefName
      do while intCBcount < clng(intTotalQueries)
        DumpCarBlack intCBcount, True, 10000, strQueryFeed & strStartDateQuery & strEndDateQuery & strHostFilter 
        intCBcount = intCBcount +10000
      loop
      if DictAdhocQuery.count > 0 then
        if BoolDebugTrace = True then logdata strDebugPath & "\CarBlacktext" & "" & ".txt", "Child processes " & DictAdhocQuery.count & vbcrlf & "-------" & vbcrlf,BoolEchoLog 
        for each strChildQuery in DictAdhocQuery
          strQueryFeed = "/api/v1/process/" & strChildQuery & strStartDateQuery & strEndDateQuery 
          if BoolDebugTrace = True then logdata strDebugPath & "\CarBlacktext" & "" & ".txt", "Parent Query=" & strQueryFeed & vbcrlf & "-------" & vbcrlf,BoolEchoLog 
          DumpCarBlack 0, False, 10000, strQueryFeed
        next
        
        for each strChildQuery in DictChildQuery
          strQueryFeed = "/api/v1/process/" & strChildQuery & strStartDateQuery & strEndDateQuery
          if BoolDebugTrace = True then logdata strDebugPath & "\CarBlacktext" & "" & ".txt", "Child Query=" & strQueryFeed & vbcrlf & "-------" & vbcrlf,BoolEchoLog 
          DumpCarBlack 0, True, 10000, strQueryFeed
        next        
        DictAdhocQuery.RemoveAll
        DictChildQuery.RemoveAll
        if BoolDebugTrace = True then logdata strDebugPath & "\CarBlacktext" & "" & ".txt", "End child processes" & vbcrlf & "-------" & vbcrlf,BoolEchoLog 
      end if
      
      'limited CSV output
      if DictLimitedOut.count > 0 then
        strHashOutPath = CurrentDirectory & "\Limited_CBmd5_" & strUniquefName      
         
        if left(lcase(strQueryFeed), 15) = "/api/v1/binary?" then
          'not using Parent Name,Command Line,TOR IP,ID GUID,Child Count
          strSSrow = "MD5,Path," & "Publisher," & "Company," & "Product," & "CB Prevalence," & "Logical Size,Info Link,Alliance Score,Dup Count"
        ELSE
          'not using Publisher	Company	Product	CB Prevalence	Logical Size Version,64-bit,Vuln
          strSSrow = "MD5,Path,Info Link,Alliance Score,Parent Name,Command Line,TOR IP,Dup Count"          
        END IF
        logdata strHashOutPath, strSSrow, False
        for each strRowCSV in DictLimitedOut
          logdata strHashOutPath, strRowCSV & "," & Chr(34) & DictLimitedOut.item(strRowCSV) & Chr(34), False
        
        next
        DictLimitedOut.RemoveAll
      end if
    end if
    'strSSfilePath = CurrentDirectory & "\CBIP_" & DictFeedInfo.item(strCBFeedID) & "_" & udate(now) & ".csv"
    'For each item in DictIPAddresses
    '  LogData strSSfilePath, item & "|" & DictIPAddresses.item(item), False
    'next
    'DictIPAddresses.RemoveAll
   
  else
    logdata CurrentDirectory & "\CB_Feeds.log", date & " " & time & ": " & "Parser not configured for " & DictFeedInfo.item(strCBFeedID) ,boolEchoInfo
  end if
next

'msgbox DumpCarBlack("EDD800F2A7F82E43392CEF00391109BE")
Function DumpCarBlack(intCBcount,BoolProcessData, intCBrows, strURLQuery)

Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Dim strAVEurl
Dim strReturnURL
dim strAssocWith
Dim strCBresponseText
Dim strtmpCB_Fpath
Dim StrTmpFeedIP
Dim boolProcessChildren: boolProcessChildren = False
'msgbox StrBaseCBURL & "/api/v1/binary?q=is_executable_image:true AND server_added_timestamp:[" & strStartDateQuery & "T00:00:00 TO " & strEndDateQuery & "T00:00:00]&start=" & intCBcount & "&rows=" & intCBrows
'msgbox StrBaseCBURL & "/api/v1/binary?q=is_executable_image:true" & strStartDateQuery & strEndDateQuery & "&start=" & intCBcount & "&rows=" & intCBrows
strAVEurl = StrBaseCBURL & strURLQuery
if BoolProcessData = True and instr(strAVEurl, "?") > 0 then
  strAVEurl = strAVEurl & "&start=" & intCBcount & "&rows=" & intCBrows
end if
if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "" & ".txt", "Query URL=" & strAVEurl & vbcrlf & vbcrlf,BoolEchoLog 
objHTTP.open "GET", strAVEurl, False

objHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
  

on error resume next
  objHTTP.send 
  if err.number <> 0 then
    logdata CurrentDirectory & "\CBF_Error.log", Date & " " & Time & " CarBlack lookup failed with HTTP error. - " & err.description,False 
    exit function 
  end if
on error goto 0  
'creates a lot of data. DOn't uncomment next line unless your going to disable it again
if BoolDebugTrace = True then logdata strDebugPath & "\CarBlack" & "" & ".txt", objHTTP.responseText & vbcrlf & vbcrlf,BoolEchoLog 
strCBresponseText = objHTTP.responseText
if instr(objHTTP.responseText, "401 Unauthorized") then
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
        DumpCarBlack = getdata(strCBresponseText, ",", "total_results" & Chr(34) & ": ")
      elseif instr(strCBresponseText, "provider_url" & Chr(34) & ": ") > 0 and instr(strCBresponseText, "id" & Chr(34) & ": ") > 0 then
        strTmpFeedID = getdata(strCBresponseText, ",", "id" & Chr(34) & ": ")
        strTmpFeedName = getdata(strCBresponseText, Chr(34), chr(34) & "name" & Chr(34) & ": " & Chr(34))
		if strTmpFeedName = "yara" then yaraFeedID = strTmpFeedID
        if DictFeedInfo.exists(strTmpFeedID) = false then DictFeedInfo.add strTmpFeedID, strTmpFeedName
      elseif instr(strAVEurl, "?") = 0 then 'Specific process query for children and parent
        
        if boolProcessChildren = True and BoolProcessData = False then
          strCBSegID = getdata(strCBresponseText, ",", "segment_id" & Chr(34) & ": ")
          strCBID = getdata(strCBresponseText, chr(34), chr(34) & "id" & Chr(34) & ": " & CHr(34))
          if strCBID = "" then
            strCBID = getdata(strCBresponseText, chr(34), chr(34) & "unique_id" & Chr(34) & ": " & CHr(34))
            if instr(strCBID, "-") > 0 then strCBID = left(strCBID, len(strCBID) -9)
          end if
          
          if BoolDebugTrace = True then logdata strDebugPath & "\CarBlackchild" & "" & ".txt", "strCBSegID=" & strCBSegID,BoolEchoLog 
          if BoolDebugTrace = True then logdata strDebugPath & "\CarBlackchild" & "" & ".txt", "strCBID=" & strCBID,BoolEchoLog 

          if strCBSegID <> "" and strCBID <> "" then
            if DictChildQuery.exists(strCBID & "/" & strCBSegID) = False then
              DictChildQuery.add strCBID & "/" & strCBSegID, ""
            end if
          end if
        elseif BoolProcessData = True then
          'msgbox strCBresponseText
          LogMD5Data strCBresponseText
        end if
        if instr(strCBResponseText, "children") > 0 then boolProcessChildren = True
      elseif BoolProcessData = True then 
        if instr(strCBresponseText, "md5") > 0 then
          LogMD5Data strCBresponseText
        end if
      end if
    end if
  end if
	
next
set objHTTP = nothing
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

Sub LogMD5Data(strCBresponseText)

if BoolDebugTrace = True then logdata strDebugPath & "\CarBlacktext" & "" & ".txt", strCBresponseText & vbcrlf & "-------" & vbcrlf,BoolEchoLog 

if instr(strCBresponseText, "md5") > 0 then 
  'DumpCarBlack = "Carbon Black has a copy of the file for hash " & strCarBlack_ScanItem

  strCBfilePath = getdata(strCBresponseText, "]", "observed_filename" & Chr(34) & ": [")
  strCBfilePath = replace(strCBfilePath,chr(10),"")
  strCBfilePath = RemoveTLS(strCBfilePath)
  strCBfilePath = getdata(strCBfilePath, chr(34),chr(34))'just grab the fist file path listed
  if strCBfilePath = "" then
    strCBfilePath = getdata(strCBresponseText, Chr(34), "path" & Chr(34) & ": " & Chr(34))
  end if
  if instr(strCBresponseText, "digsig_publisher") > 0 then 
    strCBdigSig = getdata(strCBresponseText, chr(34), "digsig_publisher" & Chr(34) & ": " & Chr(34))
    strCBdigSig = replace(strCBdigSig,chr(10),"")
  else
    'not signed 
  end if
  if instr(strCBresponseText, "signed" & Chr(34) & ": " & Chr(34) & "Signed") = 0 and instr(strCBresponseText, "signed" & Chr(34) & ": " & Chr(34) & "Unsigned") = 0 then
    'problem with sig
    strCBdigSig = getdata(strCBresponseText, chr(34), "signed" & Chr(34) & ": " & Chr(34)) & " - " & strCBdigSig
  end if 
  strCBcompanyName = getdata(strCBresponseText, chr(34), "company_name" & Chr(34) & ": " & Chr(34))
  strCBcompanyName = "|" & RemoveTLS(strCBcompanyName)
  strCBproductName = getdata(strCBresponseText, chr(34), "product_name" & Chr(34) & ": " & Chr(34))
  strCBproductName = RemoveTLS(strCBproductName)
  strCBproductName = "|" & replace(strCBproductName, "|", " ")
  StrCBMD5 = getdata(strCBresponseText, chr(34), "md5" & Chr(34) & ": " & Chr(34))
  strCBprevalence = getdata(strCBresponseText, ",", "host_count" & Chr(34) & ": ")
  strCBcmdline = getdata(strCBresponseText, Chr(34), "cmdline" & Chr(34) & ": " & Chr(34))
  strCBis64 = getdata(strCBresponseText, ",", "is_64bit" & Chr(34) & ": " )
  strCBVersion = getdata(strCBresponseText, Chr(34), "file_version" & Chr(34) & ": " & Chr(34))
  if strCBVersion = "" then
    strCBVersion = getdata(strCBresponseText, Chr(34), "product_version" & Chr(34) & ": " & Chr(34))

  end if
  strCBparent_name = getdata(strCBresponseText, Chr(34), "parent_name" & Chr(34) & ": " & Chr(34))
  strCBStartTime = getdata(strCBresponseText, Chr(34), "start" & Chr(34) & ": " & Chr(34))
  strCBUserName = getdata(strCBresponseText, Chr(34), "username" & Chr(34) & ": " & Chr(34))
  strCbEndTime = getdata(strCBresponseText, Chr(34), "last_server_update" & Chr(34) & ": " & Chr(34))
  strCbDuration = ""
  if len(strCBStartTime) > 7 then
    strtmpStart = replace(strCBStartTime, "T", " ")
    if instrrev(strtmpStart, ".") > 0 then
        strtmpStart = left(strtmpStart, instrrev(strtmpStart, ".") - 1)
    else
      strtmpStart = left(strtmpStart, len(strtmpStart) - 1)
    end if
  end if
  if len(strCbEndTime) > 7 then
    strtmpEnd = replace(strCbEndTime, "T", " ")
    if instrrev(strtmpEnd, ".") > 0 then
        strtmpEnd = left(strtmpEnd, instrrev(strtmpEnd, ".") - 1)
    else
      strtmpEnd = left(strtmpEnd, len(strtmpEnd) - 1)
    end if
    if (isdate(strtmpStart) = false and strtmpStart <> "") or isdate(strtmpEnd) = false then
      msgbox "invalid date:" & strCBStartTime &"|" & strtmpStart & "|" & strCbEndTime & "|" & strtmpEnd
    end if
    'msgbox isdate(strtmpEnd)
    strCbDuration = datediff("n",strtmpStart,strtmpEnd)
    if strCbDuration = 0 then
      strCbDuration = datediff("n",strtmpStart,strtmpEnd) & " sec"
    else
      strCbDuration = strCbDuration & " min"
    end if
  end if
  
  strCBHostname = getdata(strCBresponseText, Chr(34), "hostname" & Chr(34) & ": " & Chr(34))
  if strCBHostname = "" then
    strTmpCBHostname = getdata(strCBresponseText, "]", "endpoint" & Chr(34) & ": [" & vblf & "        " & chr(34))
    if instr(strTmpCBHostname, "|") > 0 then
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
  CBhostName = replace(CBhostName, chr(34), "")
  strCBAllianceScore = getdata(strCBresponseText, ",", Chr(34) & "alliance_score_")
  'set alliance score to integer only
	for intLen = 1 to len(strCBAllianceScore)
		if isnumeric(mid(strCBAllianceScore,intLen, 1)) = false and mid(strCBAllianceScore,intLen, 1) <> "-" then
			strCBAllianceScore = left(strCBAllianceScore, intLen -1)
		end if
	next
  strCBInfoLink = getdata(strCBresponseText, ",", "alliance_link_nvd" & Chr(34) & ": ")
  if strCBInfoLink = "" then
    strCBInfoLink = getdata(strCBresponseText, ",", "alliance_link_srstrust" & Chr(34) & ": ")
    if strCBInfoLink = "" then
      strCBInfoLink = getdata(strCBresponseText, ",", "alliance_link_srsthreat" & Chr(34) & ": ")
    end if
    if strCBInfoLink = "" then
      strCBInfoLink = getdata(strCBresponseText, ",", "alliance_link_virustotal" & Chr(34) & ": ")
    end if
    if strCBInfoLink = "" then
      strCBInfoLink = getdata(strCBresponseText, ",", "alliance_link_bit9endpointvisibility" & Chr(34) & ": ")
    end if
  end if
  strCBFileSize = getdata(strCBresponseText, ",", "orig_mod_len" & Chr(34) & ": ")
  strCBSegID = getdata(strCBresponseText, ",", "segment_id" & Chr(34) & ": ")
  
  strCBChildCount = getdata(strCBresponseText, ",", "childproc_count" & Chr(34) & ": ")
  strCBID = getdata(strCBresponseText, chr(34), chr(34) & "id" & Chr(34) & ": " & CHr(34))
  strtmpCB_Fpath = getfilepath(strCBfilePath)

  strTmpCBHostname = getdata(strCBresponseText, "]", "alliance_data_tor" & Chr(34) & ": [" & vblf & "        " & chr(34))
  if instr(strTmpCBHostname, vblf) > 0 then
    arrayTorIPaddresses = split(strTmpCBHostname, vblf)
    for each strTORip in arrayTorIPaddresses
      strTmpTorIP = getdata(strTORip, chr(34), "TOR-Node-")
      if strTORIPaddresses = "" then
        strTORIPaddresses = strTmpTorIP
      else
        strTORIPaddresses= strTORIPaddresses & "/" & strTmpTorIP
      end if        

    next
  end if
if BoolDebugTrace = True then logdata strDebugPath & "\CarBlackchild" & "" & ".txt", "strCBChildCount=" & strCBChildCount,BoolEchoLog 
if isnumeric(strCBChildCount) then
  if clng(strCBChildCount) > 0 then
    if strCBID <> "" then 
      if BoolDebugTrace = True then logdata strDebugPath & "\CarBlackchild" & "" & ".txt", "parent_id:" & strCBID,BoolEchoLog 
      if DictAdhocQuery.exists(strCBID & "/" & strCBSegID) = false then
        DictAdhocQuery.add strCBID & "/" & strCBSegID, strCBfilePath
      end if
    end if
  end if
 
end if
  'RecordPathVendorStat strtmpCB_Fpath 'record path vendor statistics
end if

if StrCBMD5 <> "" then
  if strQueryFeed = "/api/v1/binary?q=flash&digsig_publisher:Adobe  Systems  Incorporated" then
    if instr(lcase(strCBfilePath), ".ocx") = 0 and instr(lcase(strCBfilePath), "flashplayerplugin") = 0 then 
      exit sub
    else
      strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
    end if
  end if
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "mshtml.dll" & chr(34) & "&digsig_publisher:Microsoft Corporation" then
    if instr(lcase(strCBfilePath), "\system32\") = 0 and instr(lcase(strCBfilePath), "\syswow64\") = 0 then 
      exit sub
    else
      strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
    end if
  end if
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "netapi32.dll" & chr(34) & "&digsig_publisher:Microsoft Corporation" then
    if instr(lcase(strCBfilePath), "\system32\") = 0 and instr(lcase(strCBfilePath), "\syswow64\") = 0 then 
      exit sub
    else
      strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
    end if
  end if  
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "silverlight.configuration.exe" & chr(34) & "& digsig_publisher:Microsoft Corporation" then
    if instr(lcase(strCBfilePath), "silverlight.configuration.exe") = 0 and instr(lcase(strCBfilePath), "microsoft silverlight") = 0 then 
      exit sub
    else
      strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
    end if
  end if  
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "iexplore.exe" & chr(34) & "& digsig_publisher:Microsoft Corporation" then
    if instr(lcase(strCBfilePath), "\program files") = 0 and instr(lcase(strCBfilePath), "internet explorer") = 0 then 
      exit sub
    else
      strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
    end if
  end if    
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "uxtheme.dll" & chr(34)  & "& digsig_result:Unsigned"then
      strCBVuln = "Suspicious uxtheme.dll"
  end if  
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "vbscript.dll" & chr(34)  & "& digsig_publisher:Microsoft Corporation" then
       strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
  end if  
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "srv.sys" & chr(34) & "& digsig_publisher:Microsoft Corporation" then
	strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
  end if
  if strQueryFeed = "/api/v1/binary?q=observed_filename:" & chr(34) & "winword.exe" & chr(34) & "& digsig_publisher:Microsoft Corporation" then
	strCBVuln = ParseVulns(replace(strCBfilePath,"\\","\"), strCBVersion)
  end if
  'monitor for IP addresses in command lines
  if len(strCBcmdline) > 5 then
   Set re = new regexp  'Create the RegExp object 'more info at https://msdn.microsoft.com/en-us/library/ms974570.aspx
	boolLogIP = False
    re.Pattern = "\b(?:(?:25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\.){3}(?:25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])\b"'http://www.regular-expressions.info/ip.html
    re.IgnoreCase = true
    on error resume next
	WLRegXresult = re.Test(strCBcmdline)
	if err.number <> 0 then msgbox "problem with regex: " & WatchItem
	on error goto 0
	'msgbox "regex match=" & WLRegXresult & " for " & WatchItem
    if WLRegXresult = True then
		 boolLogIP = True
	end if
  end if
  strCBfilePath = AddPipe(strCBfilePath) 'CB File Path
  strCBdigSig = AddPipe(strCBdigSig) 'CB Digital Sig
  strCBcompanyName = AddPipe(strCBcompanyName)'CB Company Name
  strCBproductName = AddPipe(strCBproductName) 'Product Name        
  strCBFileSize = AddPipe(strCBFileSize)  
  strCBprevalence = AddPipe(strCBprevalence)
  strCBHostname = AddPipe(strCBHostname)
  strCBInfoLink = AddPipe(strCBInfoLink)
  strCBAllianceScore = AddPipe(strCBAllianceScore)
  strCBparent_name = AddPipe(strCBparent_name)
  strCBStartTime = AddPipe(strCBStartTime)
  strCBcmdline = AddPipe(strCBcmdline)
  strTORIPaddresses = AddPipe(strTORIPaddresses)
  strCBID = AddPipe(strCBID)
  strCBChildCount = AddPipe(strCBChildCount)
  strCBVersion = AddPipe(strCBVersion)
  strCBis64 = AddPipe(strCBis64)
  strCBVuln = AddPipe(strCBVuln)
  strCbUserName = AddPipe(strCbUserName)
  strCbDuration = AddPipe(strCbDuration)
  if boolHeaderWritten = False then
      'strSSrow = "MD5,Path," & "Publisher," & "Company," & "Product," & "CB Prevalence," & "Logical Size,Host Name,Info Link,Alliance Score,Parent Name,Command Line,TOR IP,ID GUID,Child Count,Version,64-bit,Vuln"
    if left(lcase(strQueryFeed), 15) = "/api/v1/binary?" then
      strYaraLine = ""
      if (boolEnableYARA = True or boolAddYARAtoReports = True) then strYaraLine = ",YARA"
       'not using Parent Name,Command Line,TOR IP,ID GUID,Child Count
       strSSrow = "MD5,Path," & "Publisher," & "Company," & "Product," & "CB Prevalence," & "Logical Size,Host Name,Info Link,Alliance Score,Version,64-bit,Vuln" & strYaraLine
    else 'process
      
       'not using Publisher	Company	Product	CB Prevalence	Logical Size Version,64-bit,Vuln
      strSSrow = "MD5,Path," & "Host Name,Info Link,Alliance Score,Parent Name,Command Line,TOR IP,ID GUID,Child Count,Start Time,User Name,Duration"
    end if
    logdata strHashOutPath, strSSrow, False
	  if boolLogIP = True then logdata left(strHashOutPath, len(strHashOutPath) -4) & "_IP.txt", strSSrow, False
	  
      boolHeaderWritten = True
  END IF
  'limited output
  'strSSrow = StrCBMD5 & strCBfilePath & strCBdigSig & strCBcompanyName & strCBproductName & strCBprevalence & strCBFileSize & strCBInfoLink & strCBAllianceScore & strCBparent_name & strCBcmdline & strTORIPaddresses
  if left(lcase(strQueryFeed), 15) = "/api/v1/binary?" then
    strSSrow = StrCBMD5 & strCBfilePath & strCBdigSig & strCBcompanyName & strCBproductName & strCBprevalence & strCBFileSize & strCBInfoLink & strCBAllianceScore
  else
    strSSrow = StrCBMD5 & strCBfilePath & strCBInfoLink & strCBAllianceScore & strCBparent_name & strCBcmdline & strTORIPaddresses
  end if
  strSSrow = chr(34) & replace(strSSrow, "|",chr(34) & "," & Chr(34)) & chr(34)
  if DictLimitedOut.exists(strSSrow) = False then 
    DictLimitedOut.add strSSrow, 1
  else
    DictLimitedOut.item(strSSrow) = DictLimitedOut.item(strSSrow) + 1
  end if
  
  'strSSrow = StrCBMD5 & strCBfilePath & strCBdigSig & strCBcompanyName & strCBproductName & strCBprevalence & strCBFileSize & strCBHostname & strCBInfoLink & strCBAllianceScore & strCBparent_name & strCBcmdline & strTORIPaddresses & strCBID & strCBChildCount & strCBVersion & strCBis64 & strCBVuln
  if left(lcase(strQueryFeed), 15) = "/api/v1/binary?" then
	strYaraLine = ""
	if boolAddYARAtoReports = True then
		if dictYARA.exists(StrCBMD5) then
			strYaraLine = "|" & dictYARA.item(StrCBMD5)
		else
			strYaraLine = "|" 
		end if
	end if
    'not using Parent Name,Command Line,TOR IP,ID GUID,Child Count
    strSSrow = StrCBMD5 & strCBfilePath & strCBdigSig & strCBcompanyName & strCBproductName & strCBprevalence & strCBFileSize & strCBHostname & strCBInfoLink & strCBAllianceScore & strCBVersion & strCBis64 & strCBVuln & strYaraLine
  else
    'not using Publisher	Company	Product	CB Prevalence	Logical Size Version,64-bit,Vuln
    strSSrow = StrCBMD5 & strCBfilePath & strCBHostname & strCBInfoLink & strCBAllianceScore & strCBparent_name & strCBcmdline & strTORIPaddresses & strCBID & strCBChildCount & strCBStartTime & strCBUserName & strCbDuration
  end if
  strTmpSSlout = chr(34) & replace(strSSrow, "|",chr(34) & "," & Chr(34)) & chr(34)
  logdata strHashOutPath, strTmpSSlout, False
  if boolLogIP = True then logdata left(strHashOutPath, len(strHashOutPath) -4) & "_IP.txt", strTmpSSlout, False
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
strCBis64 = ""
strCBVersion = ""
strCBVuln = ""
end sub




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
  
Function Decrypt(StrText,key) 'Rafael Paraná - https://gallery.technet.microsoft.com/scriptcenter/e0d5d71c-313e-4ac1-81bf-0e016aad3cd2
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

Function ParseVulns(strTmpVulnPath, StrTmpVulnVersion)
StrVulnVersion = removeInvalidVersion(StrTmpVulnVersion)
strVulnPath = lcase(strTmpVulnPath)
if instr(StrVulnVersion, ".") then
	intWinMajor = left(StrVulnVersion, instr(StrVulnVersion, ".") -1)
	if instr(right(StrVulnVersion, len(StrVulnVersion) - instr(StrVulnVersion, ".")), ".") then
		intWinMinor = left(right(StrVulnVersion, len(StrVulnVersion) - instr(StrVulnVersion, ".")), instr(StrVulnVersion, ".") -1)
	end if
end if
'msgbox "StrVulnVersion=" & StrVulnVersion & "|intWinMajor=" & intWinMajor & "|intWinMinor=" & intWinMinor
'msgbox "strVulnPath=" & strVulnPath
Dim StrVersionCompare
Dim ArrayVulnVer
if instr(lcase(strVulnPath), "c:\windows\syswow64\macromed\flash\") > 0 or instr(lcase(strVulnPath), "c:\windows\system32\macromed\flash\") > 0 then
  if instr(lcase(strVulnPath), ".ocx") > 0 or instr(lcase(strVulnPath), ".dll") > 0  or instr(lcase(strVulnPath), ".exe") > 0 then
    'check version number
    if boolDebugFlash = true then msgbox "Flash version assess: " & StrVulnVersion & vbcrlf & _
    "patched version is " & strFlashVersion & vbcrlf & "version patched = " & FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, strFlashVersion)
    if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, strFlashVersion) = True then
      ParseVulns = "up to date Flash Player detected"
    else 'out of date
      if isnumeric(left(StrVulnVersion, 2)) then
        if left(StrVulnVersion,2) <>  left(strStaticFPversion,2) then
          ParseVulns = "unsupported Flash Player major version detected"
        else
          ParseVulns = "outdated Flash Player version detected"
        end if
      else
        ParseVulns = "outdated Flash Player version detected"
      end if
    end if
  end if
elseif instr(lcase(strVulnPath), "c:\windows\syswow64\mshtml.dll") > 0 or instr(lcase(strVulnPath), "c:\windows\system32\mshtml.dll") > 0 then
if instr(strVulnVersion, ".") > 0 then
  ArrayVulnVer = split(strVulnVersion, ".")
  if ubound(ArrayVulnVer) > 2 then
    select case ArrayVulnVer(0)
      Case "6"
      StrVersionCompare = "6.0.3790.5662"
      Case "7"
         if ArrayVulnVer(2) = "6000" then
            StrVersionCompare = "7.0.6000.21481"
        elseif instr(strVulnVersion, "7.0.6002.1") > 0 then
          StrVersionCompare = "7.0.6002.19421"
        else
          StrVersionCompare = "7.0.6002.23728"
        end if
      Case "8"
        if ArrayVulnVer(2) = "6001" then
          if instr(strVulnVersion, "8.0.6001.2") > 0 then
            StrVersionCompare = "8.0.6001.23707"
          else
            StrVersionCompare = "8.0.6001.19652"
          end if
        else
          if instr(strVulnVersion, "8.0.7601.1") > 0 then
            StrVersionCompare = "8.0.7601.18896"
          else
            StrVersionCompare = "8.0.7601.23099"
          end if
        end if
      Case "9"
        if instr(strVulnVersion, "9.0.8112.1") > 0 then
          StrVersionCompare = "9.0.8112.16669"
        else
          StrVersionCompare = "9.0.8112.20784"
        end if
      Case "10"
        if instr(strVulnVersion, "10.0.9200.1") > 0 then
          StrVersionCompare = "10.0.9200.17412"
        else
          StrVersionCompare = "10.0.9200.21523"
        end if
      Case "11"
        if Bool64bit = False then '32-bit version
          StrVersionCompare = "11.0.9600.17905" 'x86
        else
          StrVersionCompare = "11.0.9600.17915" 'x64
        end if
    end select

    if intWinMajor = 5 then
      if intWinMinor = 2 or intWinMinor = 1 then 'windows XP/2003
        ParseVulns = "Unsupported OS Windows XP/2003"
      elseif intWinMinor = 0 then
        ParseVulns = "Unsupported OS Windows 2000"
      end if
    elseif StrVersionCompare <> "" then
      if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
        ParseVulns = "MS15-065 KB3065822 applied"
      else
        ParseVulns = "MS15-065 KB3065822 not applied"
      end if
    end if
  end if
end if
elseif instr(lcase(strVulnPath), "c:\windows\syswow64\lpk.dll") > 0 or instr(lcase(strVulnPath), "c:\windows\system32\lpk.dll") > 0 then
  'atm*.dll does not show in all results 
  'so suplimented with lpk.dll which isn't a good indication of being patched for MS15-078 
  'but can indicate a vulnerable system if really outdated
  if intWinMajor = 6 then 
    if intWinMinor = 0 then 
    '6.0.6002.23749 Windows Vista and Windows Server 2008
      if instr(StrVulnVersion, "6.0.6002.1") > 0 then
        if Bool64bit = False then '32-bit version
          StrVersionCompare = "6.0.6002.18051"
        else'64bit version
          StrVersionCompare = "6.0.6002.18005"
        end if
      elseif  instr(StrVulnVersion, "6.0.6001.1") > 0 then
        StrVersionCompare = "6.0.6001.18000"
      else
        StrVersionCompare = "6.0.6002.23749"
      end if
    
    elseif intWinMinor = 1 then 
      '6.1.7601.23126 Windows 7 and Windows Server 2008 R2
      if instr(StrVulnVersion, "6.1.7601.2") > 0 then
        StrVersionCompare = "6.1.7601.23126"
      else
        StrVersionCompare = "6.1.7601.18923"
      end if
    elseif intWinMinor = 2 then 
      '6.2.9200.16384 Windows 8 and Windows Server 2012
      StrVersionCompare = "6.2.9200.16384"
    elseif intWinMinor = 3 then 
      '6.3.9600.17415 Windows 8.1 and Windows Server 2012 R2
      StrVersionCompare = "6.3.9600.17415"
    end if
    
    
    if instr(strVulnVersion, "6.1.7600.") > 0 then
      ParseVulns = "Unsupported OS. Missing Windows 7 SP1"
    elseif StrVersionCompare <> "" then
      if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
            'System may still be vulnerable so don't return anything
            'ParseVulns = "MS15-078 KB3079904 applied"
      else
        ParseVulns = "MS15-078 KB3079904 not applied"
      end if
    end if
  end if
elseif instr(lcase(strVulnPath), "c:\windows\syswow64\netapi32.dll") > 0 or instr(lcase(strVulnPath), "c:\windows\system32\netapi32.dll") > 0 then

  if intWinMajor = 5 then
    if intWinMinor = 0 then 'windows 2000
      StrVersionCompare = "5.0.2195.7203"

    elseif intWinMinor = 1 Then
      if instr(StrVulnVersion, "5.1.2600.3") > 0 then
        StrVersionCompare = "5.1.2600.3462"
      else
        StrVersionCompare = "5.1.2600.5694"
      end if
    elseif intWinBuild = 2 then 'windows XP/2003
       if instr(StrVulnVersion, "5.2.3790.3") > 0 then
          StrVersionCompare = "5.2.3790.3229"
       else
          StrVersionCompare = "5.2.3790.4392"
       end if
    end if
  elseif  intWinMajor = 6 then 
    if intWinMinor = 0 then 'windows vista/2008
      if intWinBuild = 6000 then 'sp0
       if instr(StrVulnVersion, "6.0.6000.16") > 0 then
          StrVersionCompare = "6.0.6000.16764"
       else
          StrVersionCompare = "6.0.6000.20937"
       end if      
      elseif intWinBuild = 6001 then 'sp0
       if instr(StrVulnVersion, "6.0.6000.18") > 0 then
          StrVersionCompare = "6.0.6001.18157"
       else
          StrVersionCompare = "6.0.6001.18157"
       end if      
      end if
    end if
  end if
  if StrVersionCompare <> "" then
    if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
      ParseVulns = "MS08-067 applied"
    else
      ParseVulns = "MS08-067 not installed"
    end if
  end if
elseif instr(lcase(strVulnPath), "\microsoft silverlight\") > 0 and _
instr(lcase(strVulnPath), "\silverlight.configuration.exe") > 0 and instr(lcase(strVulnPath), "\program files") > 0 then
  StrVersionCompare = "5.1.41212.0"
    if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
      ParseVulns = "Silverlight patched with MS16-006 critical bulletin"
    else
      ParseVulns = "Silverlight flaw, identified as CVE-2016-0034, patched under MS16-006 critical bulletin is missing"
    end if
elseif instr(lcase(strVulnPath), "\internet explorer\iexplore.exe") > 0 and instr(lcase(strVulnPath), "\program files") > 0 then
StrVersionCompare = "11"
    if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
      ParseVulns = "IE on a supported version"
    else
      ParseVulns = "Internet Explorer (IE) is at a version that may not receive publicly released security updates. IE version 11 is the only version still receiving updates for Windows 7/Windows Server 2008 R2 and most newer operating systems."
    end if
elseif instr(lcase(strVulnPath), "\vbscript.dll") > 0 and instr(lcase(strVulnPath), "\windows") > 0 and instr(lcase(strVulnPath), "\winsxs\") = 0 then
    'Internet Explorer 9 on all supported x86-based versions of Windows Vista and Windows Server 2008
    if instr(StrVulnVersion, "5.8.7601.1") > 0 then
      StrVersionCompare = "5.8.7601.17295"

    elseif instr(StrVulnVersion, "5.8.7601.2") > 0 then
      ''nternet Explorer 9 on all supported x64-based versions of Windows Vista and Windows Server 2008

        StrVersionCompare = "5.8.7601.20906"
    'Internet Explorer 10 on all supported x64-based versions of Windows Server 201
    elseif instr(StrVulnVersion, "5.8.9200.2") > 0 then
      StrVersionCompare = "5.8.9200.21841"
    'Internet Explorer 11 on all supported Windows RT 8.1 & Internet Explorer 11 on all supported x86-based versions of Windows 8.1 & Internet Explorer 11 on all supported x64-based versions of Windows 8.1 and Windows Server 2012 R2
    elseif instr(StrVulnVersion, "5.8.9600.1") > 0 then
      StrVersionCompare = "5.8.9600.18321"      
    'disabling the following to prevent false-reporting on vulnerable versions (have to go with the higher version number above)
    'Windows 7 and Windows Server 2008 R2 & Internet Explorer 11 on all supported x64-based versions of Windows 7 and Windows Server 2008 R2
    'elseif instr(StrVulnVersion, "5.8.9600.1") then
    '  StrVersionCompare = "5.8.9600.18315" 
    end if
    if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
      ParseVulns = "Internet Explorer patched with MS16-051 KB3155533"
    else
      ParseVulns = "Internet Explorer missing patch released under MS16-051 KB3155533"
    end if
elseif lcase(strVulnPath) = "c:\windows\system32\drivers\srv.sys" then

	if instr(StrVulnVersion, "6.1.7601.") > 0 then
		  StrVersionCompare = "6.1.7601.23689" '6.1.7601.23689 Win7/Server2008R2 x64/ia-64/x86
    elseif instr(StrVulnVersion, "6.1.7600.") > 0 then
		ParseVulns = "Windows missing patch released under MS17-010 KB4013389" 'no SP1 for Windows 7
		exit function
	elseif instr(StrVulnVersion, "6.0.6002.19") > 0 then
		StrVersionCompare = "6.0.6002.19743"  '6.0.6002.19743 vista/2008 x64
    elseif instr(StrVulnVersion, "6.0.6000.") > 0 then
		ParseVulns = "Windows missing patch released under MS17-010 KB4013389"
		exit function
	elseif instr(StrVulnVersion, "6.0.6002.2") > 0 then
		StrVersionCompare = "6.0.6002.24067"  '6.0.6002.24067 vista/2008 x86
    elseif instr(StrVulnVersion, "6.2.9200.") > 0 then
		StrVersionCompare = "6.2.9200.22099"  'Server 2012		
	elseif instr(StrVulnVersion, "6.3.9600.") > 0 then
		StrVersionCompare = "6.3.9600.18604"  '6.3.9600.18604 Win8.1/rt/Server2012r2 x64/x86		
    elseif instr(StrVulnVersion, "10.0.14393.") > 0 then
		StrVersionCompare = "10.0.14393.953"  '10.0.14393.953 win10
	end if
    if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
      ParseVulns = "Windows has been patched for MS17-010 KB4013389"
    else
      ParseVulns = "Windows missing patch released under MS17-010 KB4013389"
    end if
elseif  ((instr(lcase(strVulnPath),":\program files (x86)\microsoft office") > 0 and instr(lcase(strVulnPath), "\office") > 0) or _
(instr(lcase(strVulnPath),":\program files\microsoft office") > 0 and instr(lcase(strVulnPath), "\office") > 0) or _
(instr(lcase(strVulnPath),":\program files\windowsapps\microsoft.office.desktop.word_") > 0 and instr(lcase(strVulnPath), "\office") > 0)) and _
instr(lcase(strVulnPath), "\winword.exe") > 0 and instr(lcase(strVulnPath), "\microsoft office\\updates\\download\") = 0 then
	if instr(StrVulnVersion, "12.0.") > 0 then
		StrVersionCompare = "12.0.6779.5000" 
	elseif instr(StrVulnVersion, "14.0.") > 0 then
		StrVersionCompare = "14.0.7189.5001" 
	elseif instr(StrVulnVersion, "15.0.") > 0 then
		StrVersionCompare = "15.0.4971.1002" 
	elseif instr(StrVulnVersion, "16.0.") > 0 then
		StrVersionCompare = "16.0.4600.1002" 
	end if
	if FirstVersionSupOrEqualToSecondVersion(StrVulnVersion, StrVersionCompare) then
      ParseVulns = "Windows has been patched for CVE-2017-11826"
    else
      ParseVulns = "Windows missing patch released for CVE-2017-11826"
    end if		




end if
end function

Function removeInvalidVersion(strVersionNumber)
Dim StrReturnValidVersion

if instr(strVersionNumber, " ") > 0 then
    StrReturnValidVersion = left(strVersionNumber, instr(strVersionNumber, " "))
else
  StrReturnValidVersion = strVersionNumber
end if
if instr(StrReturnValidVersion, ",") > 0 then
  StrReturnValidVersion = replace(StrReturnValidVersion, ",", ".")
end if
removeInvalidVersion = StrReturnValidVersion
end function

Function FirstVersionSupOrEqualToSecondVersion(strTmpFirstVersion, strTmpSecondVersion)
StrTmpVersionNumber = removeInvalidVersion(strTmpFirstVersion)	
strFirstVersion = StrTmpVersionNumber
StrTmpVersionNumber = removeInvalidVersion(strTmpSecondVersion)	
strSecondVersion = StrTmpVersionNumber
if boolDebugVersionCompare = True then msgbox "version compare " & strFirstVersion & vbcrlf & strSecondVersion
Dim arrFirstVersion,  arrSecondVersion, i, iStop, iMax
Dim iFirstArraySize, iSecondArraySize
Dim blnArraySameSize : blnArraySameSize = False

If strFirstVersion = strSecondVersion Then
  FirstVersionSupOrEqualToSecondVersion = True
  Exit Function
End If

If strFirstVersion = "" Then
  FirstVersionSupOrEqualToSecondVersion = False
  Exit Function
End If
If strSecondVersion = "" Then
  FirstVersionSupOrEqualToSecondVersion = True
  Exit Function
End If
if isnumeric(replace(strFirstVersion, ".", "")) = false then
  msgbox "Error converting version number due to non numeric value in the fist listed version: " & strFirstVersion
  exit function
end if
if isnumeric(replace(strSecondVersion, ".", "")) = false then
  msgbox "Error converting version number due to non numeric value in the second listed version: " & strSecondVersion
  exit function
end if
arrFirstVersion = Split(strFirstVersion, "." )
arrSecondVersion = Split(strSecondVersion, "." )
iFirstArraySize = UBound(arrFirstVersion)
iSecondArraySize = UBound(arrSecondVersion)

If iFirstArraySize = iSecondArraySize Then
  blnArraySameSize = True
  iStop = iFirstArraySize
  For i=0 To iStop
    'msgbox "arrFirstVersion=" & arrFirstVersion(i) & vbcrlf & "arrSecondVersion=" & arrSecondVersion(i)
    If clng(arrFirstVersion(i)) < clng(arrSecondVersion(i)) Then
      FirstVersionSupOrEqualToSecondVersion = False
      Exit Function
    elseif clng(arrFirstVersion(i)) > clng(arrSecondVersion(i)) then
      FirstVersionSupOrEqualToSecondVersion = True
      Exit Function			
    End If
  Next
  FirstVersionSupOrEqualToSecondVersion = True
Else
  If iFirstArraySize > iSecondArraySize Then
    iStop = iSecondArraySize
  Else
    iStop = iFirstArraySize
  End If
  For i=0 To iStop
    If clng(arrFirstVersion(i)) < clng(arrSecondVersion(i)) Then
      FirstVersionSupOrEqualToSecondVersion = False
      Exit Function
    End If
  Next
  If iFirstArraySize > iSecondArraySize Then
    FirstVersionSupOrEqualToSecondVersion = True
    Exit Function
  Else
    For i=iStop+1 To iSecondArraySize
      If clng(arrSecondVersion(i)) > 0 Then
        FirstVersionSupOrEqualToSecondVersion = False
        Exit Function
      End If
    Next
    FirstVersionSupOrEqualToSecondVersion = True
  End If
End If
End Function



Function ReturnLatestFlashVer

  ReturnLatestFlashVer = strStaticFPversion

end function






Function CbFeedQuery(strQuery, strUniquefName)
Dim intParseCount: intParseCount = 10
Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
strAppendQuery = ""
boolexit = False 
do while boolexit = False 
	strAVEurl = StrBaseCBURL & "/api/v1/threat_report?q=" & strQuery & strAppendQuery
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
            dictYARA.add strItem, replace(right(strTitle,len(strTitle) -20), ",", "^")
            strTitle = right(strTitle, len(strTitle)-20)
          end if
          strRowOut = strCBid & "|" & strTitle & "|" & strItem
          strRowOut = chr(34) & replace(strRowOut,"|",chr(34) & "," & Chr(34)) & chr(34)
          if tmpYaraUID = "" then 
            tmpYaraUID = udate(now)
            logdata CurrentDirectory & "\" & strUniquefName & "_" & tmpYaraUID & ".csv","CB ID, YARA Rules, MD5" , false
          end if
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
		msgbox "Error running query: " & strQuery
		exit function
	end if
loop
End function


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
