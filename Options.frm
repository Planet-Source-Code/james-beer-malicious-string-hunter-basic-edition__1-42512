VERSION 5.00
Begin VB.Form Options 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "General Settings"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7830
   LinkTopic       =   "Form2"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox OptionsFrame 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3855
      ScaleWidth      =   7575
      TabIndex        =   2
      Top             =   480
      Width           =   7575
      Begin VB.Frame Frame2 
         BackColor       =   &H00000080&
         Caption         =   "Automatic Searching:"
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   3735
         Begin VB.PictureBox Frame1 
            BackColor       =   &H00000080&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   1
            Left            =   120
            ScaleHeight     =   73
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   233
            TabIndex        =   16
            Top             =   240
            Width           =   3495
            Begin VB.Label Options_Button 
               Alignment       =   2  'Center
               Caption         =   "Auto-Search Options"
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
               Index           =   3
               Left            =   720
               TabIndex        =   17
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label FormLabel 
               BackStyle       =   0  'Transparent
               Caption         =   "Make MSH Search Files for you every time you click Automatic search, however to tweak it goto these settings"
               ForeColor       =   &H00FFFFFF&
               Height          =   615
               Index           =   1
               Left            =   600
               TabIndex        =   18
               Top             =   0
               Width           =   2895
            End
            Begin VB.Image Image3 
               Appearance      =   0  'Flat
               Height          =   495
               Left            =   60
               Top             =   60
               Width           =   495
            End
            Begin VB.Shape Shape4 
               BorderColor     =   &H00000000&
               Height          =   615
               Left            =   0
               Top             =   0
               Width           =   3495
            End
            Begin VB.Shape Options_Button_OL 
               BorderWidth     =   3
               Height          =   255
               Index           =   3
               Left            =   720
               Top             =   720
               Width           =   2055
            End
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00000080&
         Caption         =   "Scan Options:"
         ForeColor       =   &H00FFFFFF&
         Height          =   3015
         Index           =   2
         Left            =   3960
         TabIndex        =   7
         Top             =   120
         Width           =   3495
         Begin VB.PictureBox Frame1 
            BackColor       =   &H00000080&
            BorderStyle     =   0  'None
            Height          =   2655
            Index           =   2
            Left            =   120
            ScaleHeight     =   177
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   217
            TabIndex        =   8
            Top             =   240
            Width           =   3255
            Begin VB.CheckBox Option_ExcludeClass 
               BackColor       =   &H00000080&
               Caption         =   "Exclude Potential Class Strings"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   2655
            End
            Begin VB.CheckBox Option_ExcludeClass 
               BackColor       =   &H00000080&
               Caption         =   "Exclude Suspicious Class Strings"
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   2655
            End
            Begin VB.CheckBox Option_CommentOut 
               BackColor       =   &H00000080&
               Caption         =   "Comment Out and Tag Strings"
               Enabled         =   0   'False
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   1200
               Width           =   2655
            End
            Begin VB.Label Options_Button 
               Alignment       =   2  'Center
               Caption         =   "Advanced"
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
               Index           =   4
               Left            =   600
               TabIndex        =   12
               Top             =   2280
               Width           =   2055
            End
            Begin VB.Image Image4 
               Appearance      =   0  'Flat
               Height          =   480
               Left            =   60
               Top             =   1632
               Width           =   480
            End
            Begin VB.Label FormLabel 
               BackStyle       =   0  'Transparent
               Caption         =   "These settings allow you to control the scan function itself."
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   3
               Left            =   720
               TabIndex        =   14
               Top             =   1680
               Width           =   2295
            End
            Begin VB.Shape Shape5 
               BorderColor     =   &H00000000&
               Height          =   615
               Left            =   0
               Top             =   1560
               Width           =   3135
            End
            Begin VB.Shape Options_Button_OL 
               BorderWidth     =   3
               Height          =   255
               Index           =   4
               Left            =   600
               Top             =   2280
               Width           =   2055
            End
            Begin VB.Shape Shape3 
               BorderColor     =   &H00000000&
               Height          =   615
               Left            =   0
               Top             =   0
               Width           =   3255
            End
            Begin VB.Label FormLabel 
               BackStyle       =   0  'Transparent
               Caption         =   "The Scan can be customized to exclude classes and perform actions."
               ForeColor       =   &H00FFFFFF&
               Height          =   615
               Index           =   2
               Left            =   720
               TabIndex        =   13
               Top             =   0
               Width           =   2535
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   495
               Left            =   60
               Top             =   60
               Width           =   495
            End
         End
      End
      Begin VB.Frame Frame 
         BackColor       =   &H00000080&
         Caption         =   "Report Settings:"
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3735
         Begin VB.PictureBox Frame1 
            BackColor       =   &H00000080&
            BorderStyle     =   0  'None
            Height          =   1095
            Index           =   0
            Left            =   120
            ScaleHeight     =   73
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   233
            TabIndex        =   4
            Top             =   240
            Width           =   3495
            Begin VB.Label Options_Button 
               Alignment       =   2  'Center
               Caption         =   "Report Options"
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
               Left            =   720
               TabIndex        =   5
               Top             =   720
               Width           =   2055
            End
            Begin VB.Shape Shape2 
               BorderColor     =   &H00000000&
               Height          =   615
               Left            =   0
               Top             =   0
               Width           =   3495
            End
            Begin VB.Label FormLabel 
               BackStyle       =   0  'Transparent
               Caption         =   "The report options has 3 types of logs you can create and settings for the auto-save feature."
               ForeColor       =   &H00FFFFFF&
               Height          =   615
               Index           =   0
               Left            =   720
               TabIndex        =   6
               Top             =   0
               Width           =   2775
            End
            Begin VB.Image Image2 
               Appearance      =   0  'Flat
               Height          =   495
               Left            =   60
               Top             =   60
               Width           =   495
            End
            Begin VB.Shape Options_Button_OL 
               BorderWidth     =   3
               Height          =   255
               Index           =   2
               Left            =   720
               Top             =   720
               Width           =   2055
            End
         End
      End
      Begin VB.Label Options_Button 
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
         Left            =   6120
         TabIndex        =   19
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Options_Button 
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
         Left            =   4680
         TabIndex        =   20
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label FormLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Like the new layout?, This allows you to view only the information and settings you want, "
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   3360
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Shape Options_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   0
         Left            =   4680
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Shape Options_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   1
         Left            =   6120
         Top             =   3480
         Width           =   1215
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   4
      X1              =   0
      X2              =   520
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   312
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   1
      X1              =   520
      X2              =   520
      Y1              =   0
      Y2              =   312
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
      Index           =   3
      X1              =   0
      X2              =   520
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Label GeneralSettings 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "General Settings"
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
      Width           =   7815
   End
   Begin VB.Label Options_Tool_Tip 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   4410
      Width           =   7770
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
Options_Tool_Tip = Empty
For i = 0 To (Options_Button.Count - 1)
If Options_Button(i).BackColor <> C_Button_Normal Then
Options_Button(i).BackColor = C_Button_Normal
Options_Button(i).ForeColor = C_Button_NormalText
End If
Next i
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Option_CommentOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Options_Tool_Tip = "Check to Comment out and Tag Strings"
End Sub

Private Sub Option_ExcludeClass_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
    Options_Tool_Tip = "Check to exclude all POTENTIAL Class strings"
ElseIf Index = 1 Then
    Options_Tool_Tip = "Check to exclude all SUSPICIOUS Class strings"
End If
End Sub

Private Sub Options_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
'On Error Resume Next
    Options_Button(Index).BackColor = vbRed
Select Case Index
Case Is = 0
    SaveOptions 2
    Form1.Enabled = True
    Options.Hide
Case Is = 1
    Form1.Enabled = True
    Options.Hide
Case Is = 2
    Options.Enabled = False
    
        For i = 0 To Settings_Report.ReportSettings_Button.UBound
            Settings_Report.ReportSettings_Button(i).BackColor = C_Button_Normal
            Settings_Report.ReportSettings_Button(i).ForeColor = C_Button_NormalText
        Next i

    LoadOptions 1
    Settings_Report.Show
Case Is = 3
    Options.Enabled = False
    
        For i = 0 To Settings_AutoSearch.AutoSearchSettings_Button.UBound
            Settings_AutoSearch.AutoSearchSettings_Button(i).BackColor = C_Button_Normal
            Settings_AutoSearch.AutoSearchSettings_Button(i).ForeColor = C_Button_NormalText
        Next i

    LoadOptions 4
    Settings_AutoSearch.Show
Case Is = 4
    Options.Enabled = False
    
        For i = 0 To Settings_Advanced.AdvancedSettings_Button.UBound
            Settings_Advanced.AdvancedSettings_Button(i).BackColor = C_Button_Normal
            Settings_Advanced.AdvancedSettings_Button(i).ForeColor = C_Button_NormalText
        Next i
    
    LoadOptions 3
    Settings_Advanced.Show
End Select

End Sub

Private Sub Options_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
If Options_Button(Index).BackColor <> vbYellow Then
    For i = 0 To (Options_Button.Count - 1)
        Options_Button(i).BackColor = C_Button_Normal
        Options_Button(i).ForeColor = C_Button_NormalText
    Next i
    Options_Button(Index).BackColor = C_Button_Select
    Options_Button(Index).ForeColor = C_Button_SelectText

Select Case Index
Case Is = 0
    Options_Tool_Tip = "Save General settings and Return"
Case Is = 1
    Options_Tool_Tip = "Discard General settings and Return"
Case Is = 2
    Options_Tool_Tip = "Opens report options"
Case Is = 3
    Options_Tool_Tip = "Opens Auto-Search Options"
Case Is = 4
    Options_Tool_Tip = "Opens Advanced Settings"
End Select

End If
End Sub

Private Sub OptionsFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(Button, Shift, X, Y)
End Sub
