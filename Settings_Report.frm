VERSION 5.00
Begin VB.Form Settings_Report 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Frame 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      ScaleHeight     =   241
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      Begin VB.Frame AutoSaveOptions 
         BackColor       =   &H00000080&
         Caption         =   "Auto-Save Report Files:"
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3615
         Begin VB.PictureBox Frame1 
            BackColor       =   &H00000080&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   1
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   3255
            TabIndex        =   7
            Top             =   240
            Width           =   3255
            Begin VB.OptionButton ReportFileOps 
               BackColor       =   &H00000080&
               Caption         =   "Create Increments (e.g File0001.log)"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   10
               Top             =   840
               Width           =   3255
            End
            Begin VB.OptionButton ReportFileOps 
               BackColor       =   &H00000080&
               Caption         =   "Overwrite Existing File"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   9
               Top             =   600
               Width           =   3255
            End
            Begin VB.TextBox Report_Filename 
               Height          =   285
               Left            =   1680
               TabIndex        =   8
               Text            =   "ScanReportLog"
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Default Log Filename:"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   11
               Top             =   120
               Width           =   1575
            End
         End
      End
      Begin VB.CheckBox Report_AutoSave 
         BackColor       =   &H00000080&
         Caption         =   "Create Report Log Automatically"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3495
      End
      Begin VB.OptionButton ReportType 
         BackColor       =   &H00000080&
         Caption         =   "Basic - Summary of Scan Only"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton ReportType 
         BackColor       =   &H00000080&
         Caption         =   "Advanced - Files/Malicious Strings List Only"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3495
      End
      Begin VB.OptionButton ReportType 
         BackColor       =   &H00000080&
         Caption         =   "Full"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.Image IconBox 
         Height          =   495
         Left            =   120
         Top             =   3060
         Width           =   495
      End
      Begin VB.Label ReportSettings_Button 
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
         Left            =   2520
         TabIndex        =   12
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label ReportSettings_Button 
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
         Left            =   1200
         TabIndex        =   13
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Shape Options_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   0
         Left            =   1200
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Shape Options_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   1
         Left            =   2520
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Log Type:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   3375
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   520
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   1
      X1              =   280
      X2              =   280
      Y1              =   0
      Y2              =   296
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   296
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   280
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label ReportSettingsLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Report Options"
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
      TabIndex        =   14
      Top             =   0
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   4
      X1              =   0
      X2              =   280
      Y1              =   280
      Y2              =   280
   End
End
Attribute VB_Name = "Settings_Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
For i = 0 To (ReportSettings_Button.Count - 1)
If ReportSettings_Button(i).BackColor <> C_Button_Normal Then
ReportSettings_Button(i).BackColor = C_Button_Normal
ReportSettings_Button(i).ForeColor = C_Button_NormalText
End If
Next i
End Sub

Private Sub Frame_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Form_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Form_MouseMove(Button, Shift, x, Y)
End Sub

Private Sub Report_AutoSave_Click()
If Report_AutoSave.Value = 0 Then
Frame1(1).Enabled = False
Report_Filename.BackColor = &HC0C0C0
ReportFileOps(0).ForeColor = &HC0C0C0
ReportFileOps(1).ForeColor = &HC0C0C0
Else
Frame1(1).Enabled = True
Report_Filename.BackColor = vbWhite
ReportFileOps(0).ForeColor = vbBlack
ReportFileOps(1).ForeColor = vbBlack
End If
End Sub


Private Sub ReportSettings_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Select Case Index
Case Is = 0
    SaveOptions 1
    Options.Enabled = True
    Settings_Report.Hide
Case Is = 1
    Options.Enabled = True
    Settings_Report.Hide
End Select
End Sub

Private Sub ReportSettings_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
If ReportSettings_Button(Index).BackColor <> C_Button_Select Then
    For i = 0 To 1
        ReportSettings_Button(i).BackColor = C_Button_Normal
        ReportSettings_Button(i).ForeColor = C_Button_NormalText
    Next i
    ReportSettings_Button(Index).BackColor = C_Button_Select
    ReportSettings_Button(Index).ForeColor = C_Button_SelectText
End If
End Sub
