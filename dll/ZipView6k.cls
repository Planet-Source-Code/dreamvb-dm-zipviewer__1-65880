VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ZipView6k"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Hi this is just a small zip viewer DLL I made for using on a webserver to
' allow uses to view files inside a zip. anyway this only took about 30 min to write
' so there maybe some bugs or two anyway hope you like it
'See the example project also in the folder for an example

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private Const SIG = &H4034B50
Private Const ENDSIG = &H2014B50

Private Type ZFHeader
    Signature As Long           'PKZip Signature
    Version As Integer          'Version
    GPBFlag As Integer          'General Purpose Bit Flag
    Compress As Integer         'Compression Type
    DateTime As Long            'Last Modified Date and Time
    CRC32 As Long               'CRC32
    CSize As Long               'Compressed Size
    USize As Long               'Uncompressed Size
    FileNameLen As Long         'Filename Length
    Junk As Integer             'Not sure what this is for yet
End Type

'This is the same as the above ZFHeader but a cleaned up version and some extra info
Public Type FileInfo
    zFileCRC As String          'CRC32
    zFilename As String         'Filename
    zFileSize As Long           'FileSize Uncompressed
    zPackedSize As Long
    zFileRatio As String        'Compressed Ratio
    zFileLastMod As Date        'Last Modified Date and Time
    zCompType As String         'Compression Type
    zFileExt As String          'Filenames Extection
    zFilePath As String         'Holds File Path if any is found
End Type

Private Type ZFileContents
    fCount As Long              'Number of files in the zip
    Files() As FileInfo         'Info about the files inside the zip
End Type

Public Type ZipHeader
    SIG As Long
    Version As Long
    GPBFlag As Long
    NoOfFiles As Long
End Type

Private Zip_Contents As ZFileContents
Private zHead As ZFHeader

Private Sub zFileAddToContents(iContents As FileInfo)

    With Zip_Contents
        ReDim Preserve .Files(.fCount)
            With .Files(.fCount)
                .zFileCRC = iContents.zFileCRC
                .zFileLastMod = iContents.zFileLastMod
                .zFilename = iContents.zFilename
                .zFileRatio = iContents.zFileRatio
                .zFileSize = iContents.zFileSize
                .zPackedSize = iContents.zPackedSize
                .zCompType = iContents.zCompType
                .zFileExt = iContents.zFileExt
                .zFilePath = iContents.zFilePath
            End With
        .fCount = .fCount + 1
    End With
    
End Sub

Private Sub ClearZInfo()
    Zip_Contents.fCount = 0
    Erase Zip_Contents.Files
    
    zHead.Compress = 0
    zHead.CRC32 = 0
    zHead.CSize = 0
    zHead.CSize = 0
    zHead.DateTime = 0
    zHead.FileNameLen = 0
    zHead.GPBFlag = 0
    zHead.Junk = 0
    zHead.Signature = 0
    zHead.Version = 0

End Sub

Private Function ChrPos(lzStr As String, iChr As String, Optional iStart As Integer = 1, Optional iFindFirst As Boolean = True)
Dim x As Long, idx As Long, c As String

    idx = -1
    For x = iStart To Len(lzStr)
        c = Mid$(lzStr, x, 1)
        If (c = iChr) Then
            idx = x
            If (iFindFirst) Then Exit For
        End If
    Next
    ChrPos = idx
    
End Function

Private Function GetFileExt(lFilename As String) As String
Dim x_pos As Integer
    x_pos = ChrPos(lFilename, ".", , False)
    
    If (x_pos > -1) And (x_pos < Len(lFilename)) Then
        GetFileExt = UCase(Mid(lFilename, x_pos + 1, Len(lFilename)))
    End If
End Function

Private Function GetFileName(lzExpr As String) As String
Dim x As Integer
    x = ChrPos(lzExpr, "/", , False)
    If (x <> -1) Then
        GetFileName = Mid(lzExpr, x + 1)
    Else
        'Just return lzExpr
        GetFileName = lzExpr
    End If
End Function

Private Function GetFilePath(lzExpr As String) As String
Dim x As Integer
    x = ChrPos(lzExpr, "/", , False)
    If (x <> -1) Then
        GetFilePath = Left(lzExpr, x)
    Else
        GetFilePath = "\"
    End If
End Function

Private Function SHL(iByte As Long, iShift As Long) As Long
    SHL = (iByte * (2 ^ iShift))
End Function

Private Function SHR(iByte As Long, iShift As Long) As Long
    SHR = (iByte \ (2 ^ iShift))
End Function

Public Sub OpenZip(Filename As String)
Dim fp As Long
Dim zFileInfo As FileInfo
Dim fName As String, iRatio As String

Dim OffSet As Long
Dim iCompType As Variant
Dim iDate, iTime As Date
    
    
    If LenB(Dir(Filename)) = 0 Then
        Err.Raise 10, "OpenZip", "File Not Found."
        Exit Sub
    End If
    
    fp = FreeFile
    Call ClearZInfo
    
    'Compression Types There maybe more so you can add them if you find anymore
    iCompType = Array("Stored", "Shrunk", "Reduced1", "Reduced2", "Reduced3", _
    "Reduced4", "Imploded", "Reserved", "Deflated", "Deflated", "PKWARE")
    
    Open Filename For Binary As #fp
        'Keep the current offset
        OffSet = Seek(fp)
        Get #fp, , zHead
        If (zHead.Signature <> SIG) Then
            'Check for a vaild SIG
            Err.Raise 10, "OpenZip", "Invaild Achieve Found."
            ZeroMemory zHead, Len(zHead)
            Close #fp
            Exit Sub
        Else
            'Push back to our last position in the file
           Seek #fp, OffSet
        End If
        
        With zHead
            Do
                Get #fp, , zHead
                'Check for end Signature
                If (.Signature = ENDSIG) Then Exit Do
                
                If (.Signature = SIG) Then
                    fName = Space(.FileNameLen)
                    OffSet = Seek(fp) - 2

                    Get fp, OffSet, fName
                    
                    'Last Update Date
                    iDate = DateSerial((SHR(.DateTime, 25) And &H7F) + 1980, _
                    (SHR(.DateTime, 21) And &HF), _
                    SHR(.DateTime, 16) And &H1F)
                    '
                    'Last Update Time
                    iTime = TimeSerial((SHR(.DateTime, 11) And &H1F), _
                    (SHR(.DateTime, 5) And &H3F), _
                    SHR(.DateTime, 16) And &H1F)
                    
                    'Find out the Ratio
                    If (.USize <> 0) Then
                        iRatio = CStr(Round((zHead.USize - zHead.CSize) / zHead.USize * 100) & "%")
                    Else
                        'Go here if we get a nasty divsion by zero
                        iRatio = "0%"
                    End If
                    
                    'Fill zipFileInfo Type
                    With zFileInfo
                        .zFileCRC = Hex(zHead.CRC32)
                        .zFileLastMod = (iDate & " " & iTime)
                        .zFilename = GetFileName(fName)
                        .zFilePath = Replace(GetFilePath(fName), "/", "\")
                        .zFileRatio = iRatio
                        .zFileSize = zHead.USize
                        .zPackedSize = zHead.CSize
                        .zCompType = iCompType(zHead.Compress)
                        .zFileExt = GetFileExt(.zFilename)
                        Call zFileAddToContents(zFileInfo)
                    End With
                Else
                    MsgBox ""
                End If
            
                OffSet = (Seek(fp) - 1) + .CSize + 1
                Seek #fp, OffSet
            Loop Until (.Signature <> SIG)
            
        End With
    Close #fp
    
    'Clear up
    OffSet = 0
    iRatio = ""
    fName = ""
    
    ZeroMemory zFileInfo, Len(zFileInfo)

    Erase iCompType

End Sub

Public Function GetZipInfo(ByVal Index As Long) As FileInfo
Dim zInfo As FileInfo, x As Integer
    'This Just returns information on a selected filename form the zip
    If (Index < 0) Or (Index > Zip_Contents.fCount) Then
        Err.Raise vbObject + 9, "GetZipInfo", "Invaild Index"
        Exit Function
    Else
        GetZipInfo = Zip_Contents.Files(Index)
    End If
End Function

Public Property Get ZipHeaderInfo() As ZipHeader
Dim zHeader As ZipHeader
    'Returns information about the zip Header
    With zHeader
        .GPBFlag = zHead.GPBFlag
        .SIG = SIG
        .Version = zHead.Version
        .NoOfFiles = Zip_Contents.fCount
    End With
    
    ZipHeaderInfo = zHeader
    ZeroMemory zHeader, Len(zHeader)
End Property

Private Sub Class_Terminate()
    ClearZInfo
End Sub
