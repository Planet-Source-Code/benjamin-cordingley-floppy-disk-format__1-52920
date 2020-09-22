Attribute VB_Name = "basFormat"
Option Explicit
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetVolumeInformation Lib "KERNEL32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function ReadFile Lib "KERNEL32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function SetErrorMode Lib "KERNEL32" (ByVal wMode As Long) As Long
Private Declare Function SetFilePointer Lib "KERNEL32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function WriteFile Lib "KERNEL32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function FlushFileBuffers Lib "KERNEL32" (ByVal hFile As Long) As Long
Private Declare Function DeviceIoControl Lib "KERNEL32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Declare Function LockFile Lib "KERNEL32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare Function UnlockFile Lib "KERNEL32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long


Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOOPENFILEERRORBOX = &H8000&
Private Const CREATE_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_EXECUTE = &H20000000
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_ALL = &H10000000
Private Const FILE_ANY_ACCESS As Long = 0
Private Const FILE_READ_ACCESS  As Long = &H1
Private Const FILE_WRITE_ACCESS As Long = &H2
Private Const FILE_BEGIN = 0
Private Const INVALID_HANDLE_VALUE = -1
Private Const METHOD_BUFFERED   As Long = 0

Private Const MAX_PATH As Long = 260
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const IOCTL_DISK_GET_DRIVE_GEOMETRY As Long = &H70000

Private Const FILE_DEVICE_DISK  As Long = &H7
Private Const IOCTL_DISK_BASE   As Long = FILE_DEVICE_DISK
Private Const IOCTL_DISK_GET_MEDIA_TYPES As Long = ((IOCTL_DISK_BASE * (2 ^ 16)) Or (FILE_ANY_ACCESS * (2 ^ 14)) Or (&H300 * (2 ^ 2)) Or METHOD_BUFFERED)


Public Enum FDFormat
    FD144mb = 2
    FD288mb = 3
    FD720kb = 5
End Enum

Private Type DISK_GEOMETRY
   Cylinders         As Currency  'LARGE_INTEGER (8 bytes)
   MediaType         As Long
   TracksPerCylinder As Long
   SectorsPerTrack   As Long
   BytesPerSector    As Long
End Type

Private Type Geometry
    Sectors As Long
    Heads As Long
    SectorsPerFat As Long
    SectorsPerCluster As Long
    SectorsPerTrack As Long
    BytesPerSector As Long
    RootDirEntries As Long
    DirectorySectors As Long
    MediaDescriptor As Long
End Type
Public Function FormatFloppy(Device As String, ByVal FDF As FDFormat) As Boolean
    Dim bIMG() As Byte
    Dim Geo As Geometry
    If Not GetGeometry(FDF, Geo) Then Exit Function
    
    bIMG = StrConv(BootSector(Geo) & BlankFAT(Geo) & DirectorySectors(Geo) & DataSectors(Geo), vbFromUnicode)
    FormatFloppy = WriteIMG(Device, bIMG, Form1.ProgressBar1)
End Function
Public Function GetGeometry(FDF As FDFormat, ByRef Geo As Geometry) As Boolean
    With Geo
        Select Case FDF
            Case 2              '1.44mb
                .BytesPerSector = 512
                .Sectors = 2880
                .SectorsPerFat = 9
                .SectorsPerCluster = 1
                .RootDirEntries = 224   '112
                .MediaDescriptor = &HF0
                .Heads = 2
                .SectorsPerTrack = 18
                .DirectorySectors = 14
            Case 3              '2.88mb
                .BytesPerSector = 512
                .Sectors = 5760
                .SectorsPerFat = 9
                .SectorsPerCluster = 2
                .RootDirEntries = 240
                .MediaDescriptor = &HF0
                .Heads = 2
                .SectorsPerTrack = 36
                .DirectorySectors = 15
    
            Case 5              '720kb
                .BytesPerSector = 512
                .Sectors = 1440
                .SectorsPerFat = 3
                .SectorsPerCluster = 2
                .RootDirEntries = 112
                .MediaDescriptor = &HF9
                .Heads = 2
                .SectorsPerTrack = 9
                .DirectorySectors = 7
            Case Else: Exit Function
        End Select
    End With
    GetGeometry = True
End Function
Public Function WriteIMG(Device As String, ByRef bIMG() As Byte, Optional PB As ProgressBar) As Boolean
    On Error GoTo ErrHandler
    Dim ret As Long, Sector As Long, hDevice As Long
    Dim i As Integer
    Dim Geo As Geometry
    Dim FDF As FDFormat
    
    'Find the size of the device to determine the format
    Select Case UBound(bIMG) + 1
        Case 1474560: FDF = FD144mb
        Case 2949120: FDF = FD288mb
        Case 737280:  FDF = FD720kb
    End Select
    
    If Not GetGeometry(FDF, Geo) Then Exit Function
    With Geo
        hDevice = CreateFile("\\?\" & Device, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0, OPEN_ALWAYS, 0, 0)
        'hDevice = CreateFile("\\?\" & Device, GENERIC_WRITE Or GENERIC_READ, 0, ByVal 0, OPEN_ALWAYS, 0, 0)
        If hDevice = INVALID_HANDLE_VALUE Then Exit Function
        Call LockFile(hDevice, LoWord(1 * .BytesPerSector), HiWord(1 * .BytesPerSector), LoWord(.Sectors * .BytesPerSector), HiWord(.Sectors * .BytesPerSector))
        
        If Not PB Is Nothing Then PB.Max = .Sectors
        
        Call SetFilePointer(hDevice, 0, 0, FILE_BEGIN)
        i = 32 'Sectors to write per WriteFile sub
        For Sector = 0 To .Sectors - 1 Step i
            DoEvents
            WriteFile hDevice, bIMG(Sector * .BytesPerSector), .BytesPerSector * i, ret, ByVal 0&
            If Not PB Is Nothing Then PB.Value = Sector
        Next Sector
        If Not PB Is Nothing Then PB.Value = .Sectors
        WriteIMG = True

ErrHandler:
        On Error Resume Next
        Call FlushFileBuffers(hDevice)
        Call UnlockFile(hDevice, LoWord(1 * .BytesPerSector), HiWord(1 * .BytesPerSector), LoWord(.Sectors * .BytesPerSector), HiWord(.Sectors * .BytesPerSector))
        CloseHandle hDevice
        On Error GoTo 0
    End With
End Function
Function DirectorySectors(ByRef Geo As Geometry) As String
    With Geo
        DirectorySectors = String(.BytesPerSector * .DirectorySectors, 0)
    End With
End Function
Function DataSectors(ByRef Geo As Geometry) As String
    With Geo
        DataSectors = String(.BytesPerSector * (.Sectors - .DirectorySectors - (.SectorsPerFat * 2) - 1), 246)
    End With
End Function
Private Function BlankFAT(ByRef Geo As Geometry) As String
    Dim BF As String
    Dim i As Integer
    With Geo
        BF = Chr(240) & String(2, 255) & String(.BytesPerSector * .SectorsPerFat - 3, 0)
    End With
    BlankFAT = BF & BF
    BF = ""
End Function
Public Function BootSector(ByRef Geo As Geometry, Optional BSVer As String = "XP", Optional VolumeName As String = "NO NAME") As String
    Dim BS As String
    With Geo
        BS = "EB3C90"                                               'Jump Instruction
        BS = BS & StrToHex("MSDOS5.0")                              'OEM ID
        BS = BS & Invert(Right("0000" & Hex(.BytesPerSector), 4))   'Bytes Per sector
        BS = BS & "0" & .SectorsPerCluster                          'Sectors per cluster
        BS = BS & "0100"                                            'Number of reserved sectors
        BS = BS & "02"                                              'Number of FAT copies
        BS = BS & Hex(.RootDirEntries) & "00"                       'Number of Max Root Directory Entries
        BS = BS & Invert(Right("0000" & Hex(.Sectors), 4))          'Total number of sectors
        BS = BS & Hex(.MediaDescriptor)                             'Media Descriptor
        BS = BS & Invert(Right("0000" & Hex(.SectorsPerFat), 4))    'Number of sectors per FAT
        BS = BS & Invert(Right("0000" & Hex(.SectorsPerTrack), 4))  'Number of sectors per FAT
        BS = BS & Right("00" & .Heads, 2) & "00"                    'Number of heads
        BS = BS & "0000"                                            'Number of hidden sectors
        BS = BS & "0000"                                            'Number of large sectors
        BS = BS & "000000000000"                                      'Extended BIOS Parameter Block
        BS = BS & "29" & RandomSerial
        BS = BS & StrToHex(Left(VolumeName & String(11, 32), 11))
        BS = BS & StrToHex("FAT12" & String(3, 32))
        Select Case BSVer
            Case "XP"
                BS = BS & "33C98ED1BCF07B8ED9B800208EC0FCBD007C384E247D248BC199E83C01721C83EB3A66A11C7C26663B07268A57FC750680CA0288560280C31073EB33C98A461098F7661603461C13561E03460E13D18B7611608946FC8956FEB82000F7E68B5E0B03C348F7F30146FC114EFE61BF0000E8E600723926382D741760B10BBEA17DF3A66174324E740983C7203BFB72E6EBDCA0FB7DB47D8BF0AC9840740C487413B40EBB0700CD10EBEFA0FD7DEBE6A0FC7DEBE1CD16CD19268B551A52B001BB0000E83B0072E85B8A5624BE0B7C8BFCC746F03D7DC746F4297D8CD9894EF2894EF6"
                BS = BS & "C606967DCBEA030000200FB6C8668B46F86603461C668BD066C1EA10EB5E0FB6C84A4A8A460D32E4F7E20346FC1356FEEB4A525006536A016A10918B4618969233D2F7F691F7F64287CAF7761A8AF28AE8C0CC020ACCB80102807E020E7504B4428BF48A5624CD136161720B40750142035E0B497506F8C341BB000060666A00EBB04E544C44522020202020200D0A52656D6F7665206469736B73206F72206F74686572206D656469612EFF0D0A4469736B206572726F72FF0D0A507265737320616E79206B657920746F20726573746172740D0A00000000000000ACCBD855AA"
            Case "98se"
                BS = BS & "33C98ED1BCFC7B1607BD7800C576001E561655BF2205897E00894E02B10BFCF3A4061FBD007CC645FE0F384E247D208BC199E87E0183EB3A66A11C7C663B078A57FC750680CA0288560280C31073ED33C9FE06D87D8A461098F7661603461C13561E03460E13D18B7611608946FC8956FEB82000F7E68B5E0B03C348F7F30146FC114EFE61BF0007E82801723E382D741760B10BBED87DF3A661743D4E740983C7203BFB72E7EBDDFE0ED87D7BA7BE7F7DAC9803F0AC9840740C487413B40EBB0700CD10EBEFBE827DEBE6BE807DEBE1CD165E1F668F04CD19BE817D8B7D1A8D45"
                BS = BS & "FE8A4E0DF7E10346FC1356FEB104E8C20072D7EA00027000525006536A016A10918B4618A22605969233D2F7F691F7F64287CAF7761A8AF28AE8C0CC020ACCB80102807E020E7504B4428BF48A5624CD136161720A40750142035E0B497577C3031801270D0A496E76616C69642073797374656D206469736BFF0D0A4469736B20492F4F206572726F72FF0D0A5265706C61636520746865206469736B2C20616E64207468656E20707265737320616E79206B65790D0A0000494F2020202020205359534D53444F532020205359537F010041BB000760666A00E93BFF000055AA"
        End Select
    End With
    BootSector = HexToStr(BS)
    BS = ""
End Function

Private Function StrToHex(ByVal Uncooked As String) As String
    Dim l As Long
    Dim i As Integer
    Dim cooked As String
    For l = 1 To Len(Uncooked)
        i = Asc(Mid(Uncooked, l, 1))
        cooked = cooked & Right("00" & Hex(i), 2)
    Next l
    StrToHex = cooked
    cooked = ""
End Function

Private Function HexToStr(ByVal Uncooked As String) As String
    Dim l As Long
    Dim cooked As String
    For l = 1 To Len(Uncooked) Step 2
        cooked = cooked & Chr(CLng("&H" & Mid(Uncooked, l, 2)))
    Next l
    HexToStr = cooked
    cooked = ""
End Function

Private Function Invert(ByVal Uncooked As String) As String
    Dim l As Long, cooked As String
    For l = Len(Uncooked) To 1 Step -2
        cooked = cooked & Mid(Uncooked, l - 1, 2)
    Next l
    Invert = cooked
    cooked = ""
End Function

Private Function RandomSerial() As String
    Randomize
    RandomSerial = Right("00" & Hex(Int((255 * Rnd) + 1)), 2) & Right("00" & Hex(Int((255 * Rnd) + 1)), 2) & Right("00" & Hex(Int((255 * Rnd) + 1)), 2) & Right("00" & Hex(Int((255 * Rnd) + 1)), 2)
End Function
Private Function GeoSupport(Device As String, FDF As FDFormat) As Boolean
    Dim ret As Long, l As Long, hDevice As Long
    Dim Geos(0 To 20) As DISK_GEOMETRY
    Dim i As Integer
    
    hDevice = CreateFile("\\.\" & Device, 0, FILE_SHARE_READ, 0, OPEN_ALWAYS, 0, 0)
    If (hDevice = INVALID_HANDLE_VALUE) Then Exit Function

    If DeviceIoControl(hDevice, IOCTL_DISK_GET_MEDIA_TYPES, ByVal 0&, 0, Geos(0), Len(Geos(0)) * 21, ret, 0) Then
        i = ret / Len(Geos(0))
    End If
    
    Call CloseHandle(hDevice)
    
    For l = 0 To i - 1
        If Geos(l).MediaType = FDF Then GeoSupport = True
    Next l
End Function
Public Function LoWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(LoWord, LongIn, 2)
End Function
Public Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

