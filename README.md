<div align="center">

## Floppy Disk Format


</div>

### Description

This code is a demonstation on how to format a floppy with out using format.com and with out calling the SHFormatDrive API. This code allows you to format a floppy without a popup window. Currently it is set for 1.44mb only, but all the code is there if you want to play around with 702kb and 2.88mb. From what i understand this code only works on NT based systems.
 
### More Info
 
One side effect if you do not close the handle is your floppy will become unresponive till reboot. in other words...be sure to close the handle.


<span>             |<span>
---                |---
**Submitted On**   |2004-04-06 00:19:28
**By**             |[Benjamin Cordingley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/benjamin-cordingley.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Floppy\_Dis172955462004\.zip](https://github.com/Planet-Source-Code/benjamin-cordingley-floppy-disk-format__1-52920/archive/master.zip)

### API Declarations

```
GetVolumeInformation
CreateFile
ReadFile
CloseHandle
SetErrorMode
SetFilePointer
WriteFile
FlushFileBuffers
DeviceIoControl
LockFile
UnlockFile
```





