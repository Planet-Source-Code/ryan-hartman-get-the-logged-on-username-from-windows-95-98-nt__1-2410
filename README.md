<div align="center">

## Get the logged on username from Windows 95/98/NT


</div>

### Description

Get the logged on username from Windows 95/98 and NT
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ryan Hartman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ryan-hartman.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ryan-hartman-get-the-logged-on-username-from-windows-95-98-nt__1-2410/archive/master.zip)

### API Declarations

```
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal
 lpBuffer As String, nSize As Long) As Long
'inl as 1
```


### Source Code

```
gsUserId = ClipNull(GetUser())
Function GetUser() As String
 Dim lpUserID As String
 Dim nBuffer As Long
 Dim Ret As Long
 lpUserID = String(25, 0)
 nBuffer = 25
 Ret = GetUserName(lpUserID, nBuffer)
 If Ret Then
 GetUser$ = lpUserID$
 End If
End Function
Function ClipNull(InString As String) As String
 Dim intpos As Integer
 If Len(InString) Then
 intpos = InStr(InString, vbNullChar)
 If intpos > 0 Then
 ClipNull = Left(InString, intpos - 1)
 Else
 ClipNull = InString
 End If
 End If
End Function
```

