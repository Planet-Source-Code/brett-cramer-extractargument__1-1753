<div align="center">

## ExtractArgument


</div>

### Description

I use ExtractArgument (written by my friend Mike Carper) all the time. It returns an argument or token from a string based on its position within another string and a delimiter. For example: I want the "2" in the following string: "1,2,3,4,5,6,7,8,9,10".

'Sample call

'Dim sList as string

'Dim sTown as string

'sList = "POB 145,Dexter Street,Anytown,USA"

'sTown = ExtractArgument(3, sList, ",")

'sTown will be "Anytown"

I find this very useful in working with delimited files and strings, and have implemented it in INI settings as well.
 
### More Info
 
ArgNum As Integer

srchstr As String

Delim As String

The argument desired in a string format


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brett Cramer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brett-cramer.md)
**Level**          |Unknown
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brett-cramer-extractargument__1-1753/archive/master.zip)





### Source Code

```
Function ExtractArgument (ArgNum As Integer, srchstr As String, Delim As String) As String
  'Extract an argument or token from a string based on its position
  'and a delimiter.
  On Error GoTo Err_ExtractArgument
  Dim ArgCount As Integer
  Dim LastPos As Integer
  Dim Pos As Integer
  Dim Arg As String
  Arg = ""
  LastPos = 1
  If ArgNum = 1 Then Arg = srchstr
   Do While InStr(srchstr, Delim) > 0
    Pos = InStr(LastPos, srchstr, Delim)
    If Pos = 0 Then
      'No More Args found
      If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
      Exit Do
    Else
      ArgCount = ArgCount + 1
      If ArgCount = ArgNum Then
        Arg = Mid(srchstr, LastPos, Pos - LastPos)
        Exit Do
      End If
    End If
    LastPos = Pos + 1
  Loop
  '---------
  ExtractArgument = Arg
  Exit Function
Err_ExtractArgument:
  MsgBox "Error " & Err & ": " & Error
  Resume Next
End Function
```

