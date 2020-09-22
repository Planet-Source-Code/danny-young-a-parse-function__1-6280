<div align="center">

## A 'Parse' function\.


</div>

### Description

To split a string into pieces using a certain character as a delimiter. I do not want to get messages saying, "use the Split() function" as this isn't present in VB5.

Example of this is "hello to you", with the delimiter as " ". You'll get back 3 variables, one containing "hello", one containing "to" and one containing "you"
 
### More Info
 
The string thats going to be split and the delimiter in which to split it with.

Put a button on a form, and leave it as the default name

An array of parsed words from the string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Danny Young](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/danny-young.md)
**Level**          |Intermediate
**User Rating**    |4.3 (78 globes from 18 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/danny-young-a-parse-function__1-6280/archive/master.zip)





### Source Code

```
Option Explicit
Private Sub Command1_Click()
 Dim A As Variant
 Dim i As Integer
 i = 1
 A = Parse("hello to you", " ")
 Do While A(i) <> ""
 MsgBox A(i)
 i = i + 1
 Loop
End Sub
Public Function Parse(sIn As String, sDel As String) As Variant
 Dim i As Integer, x As Integer, s As Integer, t As Integer
 i = 1: s = 1: t = 1: x = 1
 ReDim tArr(1 To x) As Variant
 If InStr(1, sIn, sDel) <> 0 Then
  Do
   ReDim Preserve tArr(1 To x) As Variant
   tArr(i) = Mid(sIn, t, InStr(s, sIn, sDel) - t)
   t = InStr(s, sIn, sDel) + Len(sDel)
   s = t
   If tArr(i) <> "" Then i = i + 1
   x = x + 1
  Loop Until InStr(s, sIn, sDel) = 0
  ReDim Preserve tArr(1 To x) As Variant
  tArr(i) = Mid(sIn, t, Len(sIn) - t + 1)
 Else
  tArr(1) = sIn
 End If
 Parse = tArr
End Function
```

