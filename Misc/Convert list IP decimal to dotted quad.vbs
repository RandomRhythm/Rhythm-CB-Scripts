Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")


CurrentDirectory = GetFilePath(wscript.ScriptFullName)
strinFile = CurrentDirectory & "\dotquad.txt"
Set objFile = objFSO.OpenTextFile(strinFile)
Do While Not objFile.AtEndOfStream

    strData = objFile.ReadLine
    logdata CurrentDirectory & "\decout.txt", Dotted2LongIP(strData), false
    if strCBout= "" then
      strCBout = "ipaddr:" & Dotted2LongIP(strData)
    else
      strCBout = strCBout & " OR ipaddr:" & Dotted2LongIP(strData)
    end if
loop
    
logdata CurrentDirectory & "\cbcout.txt", strCBout, false
msgbox IPDecToDotQuad(intDec)

Public Function Dotted2LongIP(DottedIP) 'http://www.freevbcode.com/ShowCode.asp?ID=938
    ' errors will result in a zero value
    On Error Resume Next

    Dim i, pos
    Dim PrevPos, num

    ' string cruncher
    For i = 1 To 4
        ' Parse the position of the dot
        pos = InStr(PrevPos + 1, DottedIP, ".", 1)

        ' If its past the 4th dot then set pos to the last
        'position + 1

        If i = 4 Then pos = Len(DottedIP) + 1

       ' Parse the number from between the dots

        num = Int(Mid(DottedIP, PrevPos + 1, pos - PrevPos - 1))

        ' Set the previous dot position
        PrevPos = pos

        ' No dot value should ever be larger than 255
        ' Technically it is allowed to be over 255 -it just
        ' rolls over e.g.
         '256 => 0 -note the (4 - i) that's the 
         'proper exponent for this calculation


      Dotted2LongIP = ((num Mod 256) * (256 ^ (4 - i))) + _
         Dotted2LongIP

    Next
    on error goto 0

End Function


Function IPDecToDotQuad(intDecIP)
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