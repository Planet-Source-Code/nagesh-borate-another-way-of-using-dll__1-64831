<div align="center">

## another way of using dll


</div>

### Description

another way of using dll

generally i used adding a dll from project&gt;references and so on. but that gave a lot of errors wen i take my project to another pc, so using api, i found this new way of using my dll. hope people here like it, plz give comments !!!

base64 dll came from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64297&amp;lngWId=1

and resource function from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64634&amp;lngWId=1
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[nagesh borate](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/nagesh-borate.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/nagesh-borate-another-way-of-using-dll__1-64831/archive/master.zip)





### Source Code

```
<pre>
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Sub Form_Load()
On Error GoTo erh
FileLen App.Path & "\base64.dll"
continue:
lb = LoadLibrary(App.Path & "\base64.dll")
Dim o As Object
Set o = CreateObject("base64.b64")
MsgBox o.encode("nagesh")
Set o = Nothing
FreeLibrary lb
End
erh:
If Err.Number = 53 Then
GetFromRes 101, App.Path & "\base64.dll"
GoTo continue
End If
End
End Sub
Private Sub GetFromRes(id As Integer, fp As String)
  Dim fn As Integer
  Dim fb() As Byte
  fb = LoadResData(id, "CUSTOM")
  fn = FreeFile
  Open fp For Binary Access Write As #fn
  Put #fn, , fb
  Close #fn
End Sub
</pre>
```

