<div align="center">

## Read / Load contents of a textfile into a listbox


</div>

### Description

Reads / Loads contents of a textfile into a listbox
 
### More Info
 
filepath+filename and listbox name

returns contents of a textfile into a listbox


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcel Wijnands](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcel-wijnands.md)
**Level**          |Beginner
**User Rating**    |4.0 (24 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcel-wijnands-read-load-contents-of-a-textfile-into-a-listbox__1-12429/archive/master.zip)

### API Declarations

```
usage:
File2ListBox "c:\yourfile.txt", ListBox1
```


### Source Code

```
Public Sub File2ListBox(sFile As String, oList As ListBox)
Dim fnum As Integer
Dim sTemp As String
 fnum = FreeFile()
 oList.Clear
 Open sFile For Input As fnum
  While Not EOF(fnum)
   Line Input #fnum, sTemp
   oList.AddItem sTemp
  Wend
 Close fnum
End Sub
```

