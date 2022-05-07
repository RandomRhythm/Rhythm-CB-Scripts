'Extract CB Zips v1.5 (works with CB_File_Downloader) 
'parameter is the folder path containing the zip files to extract (prompt for folder path if not provided. )

'Copyright (c) 2021 Ryan Boyle randomrhythm@rhythmengineering.com.
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

Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1

DIm objShellComplete
Set objShellComplete = WScript.CreateObject("WScript.Shell") 
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objShell
Dim BoolSilent
Dim strFDname
Dim strExtension
Const WshRunning = 0
Const WshFinished = 1


'Config section
strExtension = "zip"
BoolSilent = True
boolSubDir = False 'create sub directories for each compressed file
'end Config 

Set objShell = WScript.CreateObject( "WScript.Shell" )


strFDname = "filedata"
CurrentDirectory = GetFilePath(wscript.ScriptFullName)

on error resume next
ProcessDirectory = WScript.Arguments(0)
if err.number <> 0 then 
  wscript.echo "Error getting arguments. Must pass the path to the folder containing zip files to extract as parameter."
  ProcessDirectory = fnShellBrowseForFolderVB
end if  
on error goto 0

if objFSO.FileExists("C:\Program Files\7-Zip\7z.exe") then
  str7zPath = "C:\Program Files\7-Zip\7z.exe"
elseif objFSO.FileExists("c:\Program Files (x86)\7-Zip\7z.exe") then
  str7zPath = "c:\Program Files (x86)\7-Zip\7z.exe"
else
  msgbox "7z not installed: File does not exist - " & chr(34) &  "C:\Program Files\7-Zip\7z.exe" & chr(34) & vbcrlf & "script will now exit"
  wscript.quit(888)
end if

if objFSO.FolderExists(ProcessDirectory) = False then
  msgbox "Folder does not exist " & chr(34) & ProcessDirectory & chr(34)
  wscript.quit (882)
end if

Set f = objFSO.GetFolder(ProcessDirectory)
Set fc = f.files
For Each f1 in fc
  if lcase(right(f1.name, len(strExtension) + 1)) = "." & strExtension then
    strOutputdir = CurrentDirectory
    if objFSO.FileExists(ProcessDirectory & "\" & f1.name) then
      if instr(f1.name, ".") then
        if boolSubDir = TRue then
          if objFSO.folderexists(CurrentDirectory & "\" & f1.name)  = False then 
            objFSO.createfolder(CurrentDirectory & "\" & f1.name)
            strOutputdir = CurrentDirectory & "\" & f1.name
          end if
        end if
        set execResults = objShell.Exec(chr(34) & str7zPath & Chr(34) & " x -y -o" & Chr(34) & strOutputdir & Chr(34) & " " & Chr(34) & ProcessDirectory & "\" & f1.name & Chr(34))
        wscript.sleep 700
        intExistLoop = 0
        'wait for file to be created
        Do while exitFileExistsLoop = False and WshRunning = execResults.Status
          if objFSO.FileExists(CurrentDirectory & "\" & strFDname) = True then 
            exitFileExistsLoop = True
          else
            wscript.Sleep 2500
            if intExistLoop > 11 then exitFileExistsLoop = True
            intExistLoop = intExistLoop +1
          end if
        loop
        wscript.Sleep 800
        if objFSO.FileExists(CurrentDirectory & "\" & strFDname) = False then 
          if execResults.Status = WshRunning then
            logdata CurrentDirectory & "\extract.log", "7z is still running and extraction was not verified: " & ProcessDirectory & "\" & f1.name, False
          elseif execResults.ExitCode <> 0 then
            if BoolSilent = False then msgbox "failed extraction with exit code " & str(execResults.ExitCode) & ": " & ProcessDirectory & "\" & f1.name
            logdata CurrentDirectory & "\extract.log", "failed extraction with exit code " & cstr(execResults.ExitCode) & ": " & ProcessDirectory & "\" & f1.name, False
            if BoolSilent = False then msgbox CurrentDirectory & "\" & ReturnFnameNoExt(f1.name)
          else
            if BoolSilent = False then msgbox "7z did not error but extraction was not verified: " & ProcessDirectory & "\" & f1.name
            logdata CurrentDirectory & "\extract.log", "7z did not error but extraction was not verified: " & ProcessDirectory & "\" & f1.name, False
            if BoolSilent = False then msgbox CurrentDirectory & "\" & ReturnFnameNoExt(f1.name)

          end if
        else
            logdata CurrentDirectory & "\extract.log", "Successful extraction: " & ProcessDirectory & "\" & f1.name, False
            if objFSO.FolderExists(CurrentDirectory & "\" & ReturnFnameNoExt(f1.name)) = True then
              StrAddmodifier = "_extracted"
            else
              StrAddmodifier = ""
            end if
            if objFSO.FileExists(CurrentDirectory & "\" & ReturnFnameNoExt(f1.name) & StrAddmodifier) = False then
              
              on error resume next
              objFSO.MoveFile CurrentDirectory & "\" & strFDname, CurrentDirectory & "\" & ReturnFnameNoExt(f1.name) & StrAddmodifier
              if err.number = 0 then
                logdata CurrentDirectory & "\extract.log", "Moved  " & CurrentDirectory & "\" & ReturnFnameNoExt(f1.name) & StrAddmodifier, False
              else
                logdata CurrentDirectory & "\extract.log", "Error Moving  " & CurrentDirectory & "\" & strFDname & " to " & CurrentDirectory & "\" & ReturnFnameNoExt(f1.name) & StrAddmodifier, False
                msgbox "Error moving file - " & err.number & " " & err.description 
              end if
              on error goto 0
              
              wscript.sleep 700
            else
              logdata CurrentDirectory & "\extract.log", "Already Exists: " & CurrentDirectory & "\" & ReturnFnameNoExt(f1.name) & StrAddmodifier, False
            end if
        end if  
      else
        wscript.echo "zip file missing extension"
      end if
    end if
  end if
Next


Function ReturnFnameNoExt(strFNWE)

if instr(strFNWE,".") then
tmpArrayFName = split(strFNWE, ".")

for intFNEcount = 0 to ubound(tmpArrayFName) -1
strReturnNoExt = strReturnNoExt & tmpArrayFName(intFNEcount)

next

else
  strReturnNoExt = strFNWE
end if
ReturnFnameNoExt = strReturnNoExt
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
Dim strTmpFilName1
Dim strTmpFilName2
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


function fnShellBrowseForFolderVB()
    dim objShell
    dim ssfWINDOWS
    dim objFolder
    
    ssfWINDOWS = 36
    set objShell = CreateObject("shell.application")
        set objFolder = objShell.BrowseForFolder(0, "Example", 0, ssfDRIVES)
            if (not objFolder is nothing) then
               set oFolderItem = objFolder.items.item
               fnShellBrowseForFolderVB = oFolderItem.Path 
            end if
        set objFolder = nothing
    set objShell = nothing
end function