VERSION 5.00
Begin VB.Form Settings_Advanced 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Advanced"
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form2"
   ScaleHeight     =   2430
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Advanced Scanning Settings:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      Begin VB.PictureBox AdvancedFrame 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   120
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   281
         TabIndex        =   2
         Top             =   240
         Width           =   4215
         Begin VB.CheckBox Check3 
            BackColor       =   &H00000080&
            Caption         =   "Erase All Destructive Class Strings"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   3855
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00000080&
            Caption         =   "Scan Comments"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000080&
            Caption         =   "Disable Syntax/Context Checking (Faster Speed)"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   3855
         End
         Begin VB.Image IconBox 
            Height          =   495
            Left            =   120
            Stretch         =   -1  'True
            Top             =   900
            Width           =   495
         End
         Begin VB.Label AdvancedSettings_Button 
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
            Left            =   2880
            TabIndex        =   6
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label AdvancedSettings_Button 
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
            Left            =   1440
            TabIndex        =   7
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   1
            Left            =   2880
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   0
            Left            =   1440
            Top             =   1080
            Width           =   1215
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   4
      X1              =   0
      X2              =   4680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   4680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   4680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   1
      X1              =   4680
      X2              =   4680
      Y1              =   0
      Y2              =   2400
   End
   Begin VB.Label AdvancedSettingsLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Advanced"
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
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Settings_Advanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AdvancedFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub AdvancedSettings_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    AdvancedSettings_Button(Index).BackColor = vbRed
Select Case Index
Case Is = 0
    Options.Enabled = True
    SaveOptions 3
    Settings_Advanced.Hide
Case Is = 1
    Options.Enabled = True
    Settings_Advanced.Hide
End Select
End Sub

Private Sub AdvancedSettings_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If AdvancedSettings_Button(Index).BackColor <> C_Button_Select Then
    For i = 0 To 1
        AdvancedSettings_Button(i).BackColor = C_Button_Normal
        AdvancedSettings_Button(i).ForeColor = C_Button_NormalText
    Next i
    AdvancedSettings_Button(Index).BackColor = C_Button_Select
    AdvancedSettings_Button(Index).ForeColor = C_Button_SelectText
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
For i = 0 To 1
AdvancedSettings_Button(i).BackColor = C_Button_Normal
AdvancedSettings_Button(i).ForeColor = C_Button_NormalText
Next i
End Sub
