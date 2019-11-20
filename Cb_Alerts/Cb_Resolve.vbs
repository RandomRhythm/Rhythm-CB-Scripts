'Cb Response Resolve Alerts

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
dim strCarBlackAPIKey
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strIgnore
Dim strQuery
Dim strIDs 'quoted comma separated: & Chr(34) & "id1" & Chr(34) & ", " & Chr(34) & "id2" & Chr(34)

'---Config Section
strQuery = "" 'set to "" to disable
BoolDebugTrace = False
IgnoreReport = False 'not implemented
intReceiveTimeout = 120 'number of seconds for timeout (not implemented)
boolUseSocketTools = False 'Uses external library from SocketTools (needed when using old OS that does not support latest TLS standards) (not currently implemented)
strLicenseKey = "" 'Lincense key is required to use SocketTools 
strIniPath="Cb_Alerts.ini"
'---End Config section

if objFSO.FileExists(strIniPath) = True then
'---Ini loading section
		boolUseSocketTools = ValueFromINI(strIniPath, "BooleanValues", "UseSocketTools", boolUseSocketTools)
		BoolDebugTrace = ValueFromINI(strIniPath, "BooleanValues", "Debug", BoolDebugTrace)
'---End ini loading section
else
	if BoolRunSilent = False then WScript.Echo strFilePath & " does not exist. Using script configured/default settings instead"
end if

if IgnoreReport = True then
  strIgnore = "true"
else
  strIgnore = "false"
end if

CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strDebugPath = CurrentDirectory & "\Debug\"
if objFSO.folderexists(strDebugPath) = False then _
objFSO.createfolder(strDebugPath)


strFile= CurrentDirectory & "\cb.dat"
strAPIproduct = "Carbon Black" 
strRandom = "4bv3nT9vrkJpj3QyueTvYFBMIvMOllyuKy3d401Fxaho6DQTbPafyVmfk8wj1bXF" 'encryption key. Change if you want but can only decrypt with same key

strData = ""
if objFSO.fileexists(strFile) then
  Set objFile = objFSO.OpenTextFile(strFile)
  if not objFile.AtEndOfStream then 'read file
      On Error Resume Next
      strData = objFile.ReadLine 
        StrBaseCBURL = objFile.ReadLine
      on error goto 0
  end if
  objFile.close

  if strData <> "" then
    strData = Decrypt(strData,strRandom)
      strCarBlackAPIKey = strData
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
    StrBaseCBURL = inputbox("Enter your " & strAPIproduct & " base URL (example: https://ryancb-example.my.carbonblack.io")
    logdata strFile,StrBaseCBURL,False
  end if
end if  


if strQuery = "" then
  strQuery = inputbox("Filter alerts in the Cb Response console to the ones you want resolved. Then copy and paste the URL here and click ok.")
end if

if instr(strQuery, "://")  = 0 then
  strQuery = StrBaseCBURL & strQuery
end if

query = "{"
if strQuery <> "" then bQuery Chr(34) & "query" & Chr(34) & ": " & Chr(34) & strQuery & Chr(34) 
if strIDs <> "" then bQuery Chr(34) & "alert_ids" & Chr(34) & ": [" & Chr(34) & "id1" & Chr(34) & ", " & Chr(34) & "id2" & Chr(34) & ", ...]"
bQuery Chr(34) & "requested_status" & Chr(34) & ": " & Chr(34) & "Resolved" & Chr(34)
'bQuery Chr(34) & "set_ignored" & Chr(34) & ": " & strIgnore
'bQuery Chr(34) & "assigned_to" & Chr(34) & ": " & Chr(34) & "ahnold" & Chr(34)
query = query & "}"

StrResponse = HTTPPost(StrBaseCBURL & "/api/v1/alerts", query)
msgbox StrResponse

sub bQuery(queryAdd)
if query <> "{" then 
  query = query & "," & queryAdd
else
  query = query & queryAdd
end if
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


Function ValueFromIni(strFpath, iniSection, iniKey, currentValue)
returniniVal = ReadIni( strFpath, iniSection, iniKey)
if returniniVal = " " then 
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
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
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








Function HTTPPost(sUrl, sRequest)
set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
oHTTP.SetTimeouts 600000, 600000, 600000, 900000 
oHTTP.open "POST", sUrl,false
oHTTP.setRequestHeader "X-Auth-Token", strCarBlackAPIKey 
oHTTP.send sRequest
HTTPPost = oHTTP.responseText
End Function