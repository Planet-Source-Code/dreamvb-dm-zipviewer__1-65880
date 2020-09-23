VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmmain 
   Caption         =   "Zip View"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   390
      Left            =   8550
      TabIndex        =   7
      Top             =   255
      Width           =   630
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   390
      Left            =   7680
      TabIndex        =   6
      Top             =   255
      Width           =   765
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   6030
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8370
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0352
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   8400
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   5940
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   4350
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7673
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name    "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Modified                          "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size  "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ratio"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Packed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CRC"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Compression"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7755
      Top             =   5370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   855
      TabIndex        =   2
      Top             =   240
      Width           =   5415
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "View"
      Height          =   390
      Left            =   6375
      TabIndex        =   0
      Top             =   255
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filename"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   315
      Width           =   630
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsZipView As New ZipView6k

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Function GetFileType(FleExt As String) As String
Dim sType As String
    'First we need to get the default file Type
    sType = RegReadString(HKEY_CLASSES_ROOT, "." & FleExt, vbNullString, REG_EXPAND_SZ)
    'Next we can now get the File Type
    sType = RegReadString(HKEY_CLASSES_ROOT, sType, vbNullString, REG_EXPAND_SZ)
    
    If (sType <> "") Then
        'Return FileType found
        GetFileType = sType
    Else
        'Just return the File's Ext
        GetFileType = FleExt & " file"
    End If
    
    sType = ""
    
End Function

Function GetFileTypeIcon(FleExt As String) As String
Dim sType As String
    'First we need to get the default file Type
    sType = RegReadString(HKEY_CLASSES_ROOT, "." & FleExt, vbNullString, REG_EXPAND_SZ)
    'Next File icon
    GetFileTypeIcon = RegReadString(HKEY_CLASSES_ROOT, "\" & sType & "\DefaultIcon\", vbNullString, REG_EXPAND_SZ)
    sType = ""
End Function

Function AddIcon(lExpr As String) As Integer
Dim epos As Integer
Dim iRet As Long, File As String
    If Len(lExpr) = 0 Then lExpr = " "
    epos = InStrRev(lExpr, ",", Len(lExpr), vbBinaryCompare)
    
    If (epos > 0) Then
        'Get Filename
        File = Left(lExpr, epos - 1)
        'Get icon index
        Idx = Val(Mid(lExpr, epos + 1))
        
        'Extract the icon
        iRet = ExtractIcon(hWnd, File, Idx)
        'Draw the icon on pIcon
        iRet = DrawIconEx(pIcon.hdc, 0, 0, iRet, 0, 0, 0, 0, &H3)
        
        'DestroyIcon Icon we have now finished
        DestroyIcon iRet
        'Add the image to the imagelistbox
        ImageList1.ListImages.Add , , pIcon.Image
        'Return the Index were it was added
        AddIcon = ImageList1.ListImages.Count
        'Clear up
        pIcon.Cls
        Exit Function
    ElseIf (lExpr = "%1") Then
        'Looks like an old exe type for this I just used a default exe file
        ImageList1.ListImages.Add , , ImageList2.ListImages(2).Picture
        AddIcon = ImageList1.ListImages.Count
        Exit Function
    ElseIf lExpr = " " Then
        'Not sure so we just use the default icon
        ImageList1.ListImages.Add , , ImageList2.ListImages(1).Picture
        AddIcon = ImageList1.ListImages.Count
    Else
        iRet = ExtractIcon(hWnd, lExpr, 0)
        iRet = DrawIconEx(pIcon.hdc, 0, 0, iRet, 0, 0, 0, 0, &H3)
        ImageList1.ListImages.Add , , pIcon.Image
        AddIcon = ImageList1.ListImages.Count
        DestroyIcon iRet
    End If

End Function

Private Sub cmdExit_Click()
    Set clsZipView = Nothing
    Unload frmmain
End Sub

Private Sub cmdOpen_Click()
Dim fType As String
Dim Counter As Long
Dim zFileInfo As FileInfo
Dim Idx As Long
Dim isCreated As Boolean

    'All this code does is use MyZip viewer and fill up a Listview with all the items
    isCreated = False
    Set LstV.Icons = Nothing
    Set LstV.SmallIcons = Nothing
    
    'Setup ImageList
    ImageList1.ListImages.Clear
    ImageList1.ImageHeight = 16
    ImageList1.ImageWidth = 16
    
    With LstV
        .ListItems.Clear
        'Set up ColumnHeaders
        .ColumnHeaders(1).Width = 2500
        .ColumnHeaders(2).Width = 1500
        .ColumnHeaders(3).Width = 2000
        .ColumnHeaders(4).Width = 1000
        .ColumnHeaders(5).Width = 800
        .ColumnHeaders(6).Width = 1000
        .ColumnHeaders(7).Width = 1300
        .ColumnHeaders(8).Width = 1000
        .ColumnHeaders(9).Width = 2000
        'Load ZipFile
        clsZipView.OpenZip Text1.Text
        
        'While we have files loop
        For Counter = 0 To clsZipView.ZipHeaderInfo.NoOfFiles - 1
            zFileInfo = clsZipView.GetZipInfo(Counter)
            'Above line gets the file info for the selected index Counter
            'Check that we have a file
            If Len(zFileInfo.zFilename) > 0 Then
                 'Get the Filenames Icon
                 Idx = AddIcon(GetFileTypeIcon(zFileInfo.zFileExt))
                 
                 If (Not isCreated) Then
                    'This only needs to be done once
                    Set LstV.SmallIcons = ImageList1
                    isCreated = True
                 End If
                 
                 'Add the filename and icon found
                .ListItems.Add 1, "a" & Counter, zFileInfo.zFilename, , Idx
                
                'Add the rest of the sub-headers
                .ListItems(1).SubItems(1) = GetFileType(zFileInfo.zFileExt)
                .ListItems(1).SubItems(2) = zFileInfo.zFileLastMod
            
                If (zFileInfo.zFileSize <> 0) Then
                    .ListItems(1).SubItems(3) = Format(zFileInfo.zFileSize, "#,#")
                Else
                    .ListItems(1).SubItems(3) = 0
                End If
            
                .ListItems(1).SubItems(4) = zFileInfo.zFileRatio
                .ListItems(1).SubItems(5) = Format(zFileInfo.zPackedSize, "#,#")
                .ListItems(1).SubItems(6) = "&H" & zFileInfo.zFileCRC
                .ListItems(1).SubItems(7) = zFileInfo.zCompType
                .ListItems(1).SubItems(8) = Format(zFileInfo.zFilePath)
            End If
            
        Next Counter
    End With
    
    StatusBar1.Panels(1).Text = "Files [" & clsZipView.ZipHeaderInfo.NoOfFiles & "]"
    
End Sub

Private Sub Command1_Click()
    MsgBox "Basic Zip Viewer by deamvb"
    
End Sub

Private Sub Form_Load()
    pIcon.Width = (32 * Screen.TwipsPerPixelX)
    pIcon.Height = (32 * Screen.TwipsPerPixelY)
End Sub

Private Sub Form_Resize()
    LstV.Width = (frmmain.ScaleWidth - LstV.Left - 30)
    LstV.Height = (frmmain.ScaleHeight - StatusBar1.Height - LstV.Top)
End Sub
