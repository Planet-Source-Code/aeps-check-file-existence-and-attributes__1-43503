<div align="center">

## Check file existence and attributes


</div>

### Description

Check the existence of a file and it's attributes (DateCreated, DateLastModified, DateLastAccessed)
 
### More Info
 
Path and file name.

You must have a textbox called txtFileName, two command buttons called cmdFileExist and cmdFileAttributes.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[AepS](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aeps.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aeps-check-file-existence-and-attributes__1-43503/archive/master.zip)





### Source Code

```
Private Sub cmdFileExist_Click()
 Dim FSO, _
  FileName As String, _
  DoExist As Boolean
 Set FSO = CreateObject("Scripting.FileSystemObject")
 FileName = txtFileName
 DoExist = FSO.FileExists(FileName)
 MsgBox DoExist, , "Check Existence"
End Sub
Private Sub cmdFileAttributes_Click()
 Dim FSO, F
 Dim FileName As String
 FileName = txtFileName
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set F = FSO.GetFile(FileName)
 MsgBox "Created : " & F.DateCreated & Chr(13) & _
   "Last Modified : " & F.DateLastModified & Chr(13) & _
   "Last Accessed : " & F.DateLastAccessed, , "Check Attributes"
End Sub
```

