VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MP3 Digital Audio Extractor"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Browse For Folder"
      Height          =   315
      Left            =   3600
      TabIndex        =   17
      Top             =   1980
      Width           =   1455
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   315
      Left            =   6660
      TabIndex        =   16
      Top             =   1980
      Width           =   1455
   End
   Begin VB.CommandButton cmdRip 
      Caption         =   "Extract Tracks"
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Top             =   1980
      Width           =   1395
   End
   Begin VB.ListBox TrackList 
      Height          =   1860
      Left            =   3600
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   60
      Width           =   4515
   End
   Begin VB.TextBox txtOutPath 
      Height          =   315
      Left            =   1020
      TabIndex        =   6
      Top             =   1980
      Width           =   2535
   End
   Begin VB.ComboBox DriveList 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   60
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Grab to"
      Height          =   1440
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   3510
      Begin VB.ComboBox cmbQuality 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   840
         List            =   "frmMain.frx":0464
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2595
      End
      Begin VB.OptionButton optWAV 
         Caption         =   "WAV"
         Height          =   255
         Left            =   1620
         TabIndex        =   14
         Top             =   0
         Width           =   675
      End
      Begin VB.OptionButton optMP3 
         Caption         =   "MP3"
         Height          =   255
         Left            =   780
         TabIndex        =   13
         Top             =   0
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.ComboBox cmbBitrate 
         Height          =   315
         ItemData        =   "frmMain.frx":04A8
         Left            =   2460
         List            =   "frmMain.frx":04D6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   900
      End
      Begin VB.CheckBox chkPrivate 
         Caption         =   "Private"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1020
         Width           =   855
      End
      Begin VB.CheckBox chkOriginal 
         Caption         =   "Original"
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   780
         Width           =   855
      End
      Begin VB.CheckBox chkCRC 
         Caption         =   "CRC"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CheckBox chkCopyright 
         Caption         =   "Copyrighted"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Bitrate"
         Height          =   195
         Left            =   2460
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Quality"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Output Path"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeName 
         Caption         =   "Rename"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This project and modifications to LAME_ENC.DLL and
' AKRRIP32.DLL was made by Arto Rusanen
' http://www.4dsoftware.8m.com



' Credits...

' LAME was originally developed by Mike Cheng (www.uq.net.au/~zzmcheng).
' Now maintained by Mark Taylor (www.sulaco.org/mp3).

' You can find LAME and its source from
' http://www.mp3dev.org

' AKRip was orginally made by Andy Key and you can find AKRip and its source from
' http://akrip.sourceforge.net/


Option Explicit



Private Sub Form_Load()
  ' Initialize Exception Filter
  SetUnhandledExceptionFilter AddressOf MyExceptionFilter

  cmbQuality.ListIndex = 1
  cmbBitrate.ListIndex = 9
  
  'Find CD Drive adapters
  Dim DriveCount As Long
  Dim MyInfo As CDREC
  ChDrive App.Path
  ChDir App.Path

  DriveCount = GetNumAdapters + 1
  
  Dim i As Long
  For i = 1 To DriveCount
    MyInfo = GetDriveInformation(i - 1)
    DriveList.AddItem StripNullsArray(MyInfo.id)
  Next i
  
  DriveList.ListIndex = 0
  txtOutPath.Text = App.Path
End Sub

Private Sub DriveList_Click()
  ' Init selected drive and read its TOC
  On Error Resume Next
  TrackList.Clear
  
  Call DeInitCDDrive
  If Not InitCDDrive(DriveList.ListIndex) Then Exit Sub
  
  Dim i As Long
  i = 1
  Do While MSB2LONG(DiscToc.tracks(i + 1).addr) <> 0
    TrackList.AddItem "Track " & i
    i = i + 1
  Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' Remove exception filter...
  SetUnhandledExceptionFilter 0
End Sub


Private Sub cmdRip_Click()
  If optMP3.Value Then
    ' Fill beConfig structure....
    Dim beConfig As PBE_CONFIG
    beConfig.dwConfig = BE_CONFIG_LAME
    
    With beConfig.format.LHV1
      .dwStructVersion = 1
      .dwStructSize = Len(beConfig)
      .dwSampleRate = 44100         '// INPUT FREQUENCY
      .dwReSampleRate = 0           '// DON"T RESAMPLE
      .nMode = BE_MP3_MODE_JSTEREO  '// OUTPUT IN STREO
      .dwBitrate = val(cmbBitrate.Text)
      .dwMpegVersion = MPEG1        '// MPEG VERSION (I or II)
      .dwPsyModel = 0               '// USE DEFAULT PSYCHOACOUSTIC MODEL
      .dwEmphasis = 0               '// NO EMPHASIS TURNED ON
      .bNoRes = True                '// No Bit resorvoir
      
      .bCopyright = chkCopyright.Value = 1
      .bCRC = chkCRC.Value = 1
      .bOriginal = chkOriginal.Value = 1
      .bPrivate = chkPrivate.Value = 1
      
      Select Case cmbQuality.ListIndex  '// QUALITY PRESET SETTING
        Case 0: .nPreset = LQP_LOW_QUALITY
        Case 1: .nPreset = LQP_NORMAL_QUALITY
        Case 2: .nPreset = LQP_HIGH_QUALITY
        Case 3: .nPreset = LQP_VOICE_QUALITY
        Case 4: .nPreset = LQP_PHONE
        Case 5: .nPreset = LQP_RADIO
        Case 6: .nPreset = LQP_TAPE
        Case 7: .nPreset = LQP_HIFI
        Case 8: .nPreset = LQP_CD
        Case 9: .nPreset = LQP_STUDIO
      End Select
    End With
  End If
  
  ' Rip all tracks that are selected
  Dim TrackNo As Long
  For TrackNo = 1 To TrackList.ListCount
    If Cancelled Then Exit For
    If TrackList.Selected(TrackNo - 1) Then
      If optMP3.Value Then
        Call RipMP3(AddSlash(txtOutPath.Text) & TrackList.List(TrackNo - 1) & ".mp3", MSB2LONG(DiscToc.tracks(TrackNo).addr), MSB2LONG(DiscToc.tracks(TrackNo + 1).addr), beConfig)
      Else
        Call RipWAV(AddSlash(txtOutPath.Text) & TrackList.List(TrackNo - 1) & ".wav", MSB2LONG(DiscToc.tracks(TrackNo).addr), MSB2LONG(DiscToc.tracks(TrackNo + 1).addr))
      End If
    End If
  Next TrackNo
End Sub

Private Sub cmdQuit_Click()
  End
End Sub

Private Sub Command1_Click()
  txtOutPath.Text = BrowseForFolder
End Sub

' Change Name
Private Sub mnuChangeName_Click()
  If TrackList.ListCount > 0 Then _
    TrackList.List(TrackList.ListIndex) = InputBox("Give new name...")
End Sub

Private Sub TrackList_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then Call mnuChangeName_Click
End Sub

Private Sub TrackList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

