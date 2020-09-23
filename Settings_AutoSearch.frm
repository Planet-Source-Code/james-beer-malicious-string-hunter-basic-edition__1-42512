VERSION 5.00
Begin VB.Form Settings_AutoSearch 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Advanced Scanning Settings:"
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      Begin VB.PictureBox AutoSearchFrame 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   120
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   337
         TabIndex        =   1
         Top             =   240
         Width           =   5055
         Begin VB.TextBox ASDirectory 
            Height          =   285
            Left            =   480
            TabIndex        =   10
            Text            =   "C:\"
            Top             =   600
            Width           =   3375
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00000080&
            Caption         =   "Run Scan Automatically Afterwards"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000080&
            Caption         =   "Look in Specified Directory"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   7
            Top             =   360
            Width           =   3735
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000080&
            Caption         =   "Scan Whole HDD (C:\)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   6
            Top             =   120
            Width           =   3975
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000080&
            Caption         =   "Look in Subdirectories"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   1320
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000080&
            Caption         =   "Let me Specify Every time"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   8
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label AutoSearchSettings_Button 
            Alignment       =   2  'Center
            Caption         =   "Browse..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   11
            Top             =   600
            Width           =   975
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   2
            Left            =   3960
            Top             =   600
            Width           =   975
         End
         Begin VB.Label AutoSearchSettings_Button 
            Alignment       =   2  'Center
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   4
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label AutoSearchSettings_Button 
            Alignment       =   2  'Center
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   3
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   0
            Left            =   2400
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   1
            Left            =   3720
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Image IconBox 
            Height          =   495
            Left            =   120
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   495
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   5520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   1
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   5520
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   4
      X1              =   0
      X2              =   5520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label AutoSearchSettingsLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Automatic Searching Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Settings_AutoSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub AutoSearchFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Form_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub AutoSearchSettings_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Index
Case Is = 0
    Options.Enabled = True
    SaveOptions 4
    Settings_AutoSearch.Hide
Case Is = 1
    Options.Enabled = True
    Settings_AutoSearch.Hide
Case Is = 2
    DirDialogFlag = "AutoSearchOptions"
    Settings_AutoSearch.Enabled = False
    DirDialog.Show
End Select
End Sub

Private Sub AutoSearchSettings_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
If AutoSearchSettings_Button(Index).BackColor <> C_Button_Select Then
    For i = 0 To 1
        AutoSearchSettings_Button(i).BackColor = C_Button_Normal
        AutoSearchSettings_Button(i).ForeColor = C_Button_NormalText
    Next i
    AutoSearchSettings_Button(Index).BackColor = C_Button_Select
    AutoSearchSettings_Button(Index).ForeColor = C_Button_SelectText
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
    For i = 0 To AutoSearchSettings_Button.UBound
        AutoSearchSettings_Button(i).BackColor = C_Button_Normal
        AutoSearchSettings_Button(i).ForeColor = C_Button_NormalText
    Next i
End Sub

Private Sub Option1_Click(Index As Integer)
    If Option1(1).Value = True Then
        ASDirectory.BackColor = vbWhite
        ASDirectory.Enabled = True
    Else
        ASDirectory.BackColor = &HC0C0C0
        ASDirectory.Enabled = False
    End If
End Sub
