VERSION 5.00
Begin VB.Form DirDialog 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Choose Search Location:"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5055
      End
      Begin VB.PictureBox DirDialogFrame 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         ScaleHeight     =   185
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   337
         TabIndex        =   1
         Top             =   240
         Width           =   5055
         Begin VB.DirListBox Dir1 
            Height          =   1890
            Left            =   0
            TabIndex        =   5
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label DirDialog_Button 
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
            TabIndex        =   2
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label DirDialog_Button 
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
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   1
            Left            =   3720
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Shape AdvS_Button_OL 
            BorderWidth     =   3
            Height          =   255
            Index           =   0
            Left            =   2400
            Top             =   2400
            Width           =   1215
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   4
      X1              =   0
      X2              =   5520
      Y1              =   3720
      Y2              =   3720
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
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   1
      X1              =   5520
      X2              =   5520
      Y1              =   0
      Y2              =   3720
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
   Begin VB.Label AutoSearchSettingsLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Browse for Directory"
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
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "DirDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub DirDialog_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Index
Case Is = 0
    Select Case DirDialogFlag
    Case Is = "AutoSearchOptions"
        Settings_AutoSearch.ASDirectory = Dir1 & "\"
        Settings_AutoSearch.Enabled = True
        DirDialog.Hide
    Case Is = "AutoSearchPromptDir"
        Form1.Enabled = True
        DirDialog.Hide
        Form1.AutoSearch Dir1 & "\"
    End Select
Case Is = 1
    Select Case DirDialogFlag
    Case Is = "AutoSearchOptions"
        Settings_AutoSearch.Enabled = True
        DirDialog.Hide
    Case Is = "AutoSearchPromptDir"
        Form1.Enabled = True
        DirDialog.Hide
    End Select
End Select
End Sub

Private Sub DirDialog_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
If DirDialog_Button(Index).BackColor <> C_Button_Select Then
    For i = 0 To 1
        DirDialog_Button(i).BackColor = C_Button_Normal
        DirDialog_Button(i).ForeColor = C_Button_NormalText
    Next i
    DirDialog_Button(Index).BackColor = C_Button_Select
    DirDialog_Button(Index).ForeColor = C_Button_SelectText
End If
End Sub

Private Sub DirDialogFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Form_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub Drive1_Change()
On Error GoTo errs
Dir1 = Drive1
Exit Sub
errs:
MsgBox "Drive is not Ready, or an Access Error Occurred", vbExclamation, "Error: " & Err.Description
Drive1 = "C:\"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
    For i = 0 To 1
        DirDialog_Button(i).BackColor = C_Button_Normal
        DirDialog_Button(i).ForeColor = C_Button_NormalText
    Next i
End Sub
