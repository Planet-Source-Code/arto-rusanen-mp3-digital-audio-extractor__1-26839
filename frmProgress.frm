VERSION 5.00
Begin VB.Form frmProgress 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picProgress1 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   5115
      TabIndex        =   2
      Top             =   360
      Width           =   5175
      Begin VB.PictureBox picProgress2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   5085
         TabIndex        =   3
         Top             =   0
         Width           =   5115
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1860
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Extracting track"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   5175
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This form is only used to inform Status of Audio Extracting

Option Explicit

Private Sub Command1_Click()
  Cancelled = True
End Sub

Public Sub ChangeProgress(Status As String, CurVal As Single, TotVal As Single)
  Me.Caption = "Progress " & format(CurVal / TotVal * 100, "00.00") & " %"
  picProgress2.Width = picProgress1.ScaleWidth * (CurVal / TotVal)
  lblStatus.Caption = Status
End Sub

