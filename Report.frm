VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form ReportWindow 
   BackColor       =   &H00000020&
   BorderStyle     =   0  'None
   Caption         =   "Scan Report"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9045
   Icon            =   "Report.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9045
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox ReportFrame 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4095
      ScaleWidth      =   8775
      TabIndex        =   2
      Top             =   480
      Width           =   8775
      Begin VB.Frame Report_Information 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   0
         TabIndex        =   4
         Top             =   960
         Width           =   8655
         Begin MSComctlLib.ListView Report_MSL 
            Height          =   1815
            Left            =   3960
            TabIndex        =   5
            Top             =   480
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "MSH String"
               Object.Width           =   3422
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Class"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Occurs"
               Object.Width           =   1835
            EndProperty
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   2880
            TabIndex        =   24
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "Total Bytes Processed:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "Malicious/Accessory Strings Found:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   21
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "POTENTIAL Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "SUSPICIOUS Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "CAUTION Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   2655
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "WARNING Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "DANGER Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "DESTRUCTIVE Class:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   2655
         End
         Begin VB.Label FormLabel 
            BackColor       =   &H00000080&
            Caption         =   "Total Lines Processed:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   2880
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   2880
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   2880
            TabIndex        =   11
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   2880
            TabIndex        =   10
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   5
            Left            =   2880
            TabIndex        =   9
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   8
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Result 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   7
            Left            =   2880
            TabIndex        =   7
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label FormLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Strings Found:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   3960
            TabIndex        =   6
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.TextBox Report_Header 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   120
         Width           =   7815
      End
      Begin MSComctlLib.ProgressBar Log_Percent 
         Height          =   255
         Left            =   1680
         TabIndex        =   25
         Top             =   3720
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Report_Button 
         Alignment       =   2  'Center
         Caption         =   "Done"
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
         Left            =   7440
         TabIndex        =   27
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Report_Button 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Save Scan Report"
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
         Left            =   5280
         TabIndex        =   26
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label FormLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Saving Report..."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape Report_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   1
         Left            =   7440
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Report_Button_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   0
         Left            =   5280
         Top             =   3720
         Width           =   1935
      End
   End
   Begin VB.Label Report_Tool_Tip 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   4650
      Width           =   8945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   9000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   1
      X1              =   9000
      X2              =   9000
      Y1              =   0
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      Index           =   4
      X1              =   0
      X2              =   9000
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   9000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   4920
      Y2              =   0
   End
   Begin VB.Label ScanDoneLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Scan Completed - Report Information"
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
      Width           =   9015
   End
End
Attribute VB_Name = "ReportWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim i As Long
Report_Tool_Tip = Empty
For i = 0 To 1
    If Report_Button(i).BackColor <> C_Button_Normal Then
    Report_Button(i).BackColor = C_Button_Normal
    Report_Button(i).ForeColor = C_Button_NormalText
    End If
Next i
End Sub

Private Sub Report_Button_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Form1.Path_Label = Empty
Form1.File_Label = Empty
Form1.FileType_Label = Empty
Form1.Size_Label = Empty

Report_Button(Index).BackColor = C_Button_Select
Report_Button(Index).ForeColor = C_Button_SelectText

If Index = 0 Then
    SaveLogAs
ElseIf Index = 1 Then
    Report_Button(Index).BackColor = &HC0C0C0
    Form1.MSL.Refresh
    Form1.ProjectFiles.Refresh

If Form1.ProjectFiles.ListItems.Count = 0 Then
    Form1.MSHStatus.Panels.Item(1) = "Not Ready"
    Form1.MSHStatus.Panels.Item(1).Picture = Form1.ImageList1.ListImages.Item(3).Picture
Else
    Form1.MSHStatus.Panels.Item(1) = "Ready for next Scan/Re-Scan"
    Form1.MSHStatus.Panels.Item(1).Picture = Form1.ImageList1.ListImages.Item(2).Picture
End If

Form1.Enabled = True
ReportWindow.Hide
Report_Button(Index).BackColor = C_Button_Normal
Report_Button(Index).ForeColor = C_Button_NormalText
End If
End Sub

Private Sub Report_Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
Dim i As Long
If Report_Button(Index).BackColor <> C_Button_Select Then
For i = 0 To 1
Report_Button(i).BackColor = C_Button_Normal
Report_Button(i).ForeColor = C_Button_NormalText
Next i
Report_Button(Index).BackColor = C_Button_Select
Report_Button(Index).ForeColor = C_Button_SelectText

Select Case Index
Case Is = 0
Report_Tool_Tip = "Saves a log of everything this scan found for future reference"
Case Is = 1
Report_Tool_Tip = "Closes this window"
End Select
End If

End Sub

Function SaveLogAs()
FormLabel(10).Visible = True
Log_Percent.Visible = True
Form1.CD1.Filename = Empty
Form1.CD1.Filter = "Malicious String Hunter - Scan Report|*.log"
Form1.CD1.DialogTitle = "Save Scan Report"
Form1.CD1.ShowSave
CreateLog
End Function

Function CreateLog()
Dim i As Long
If Form1.CD1.Filename <> Empty And Right$(Form1.CD1.Filename, 4) = ".log" Then
'create a huge log file for safe keeping

'The total byte count is how much data was process, not the buffer
'size after completion

Close #1 'make sure it's not in use
Open Form1.CD1.Filename For Output As #1

Print #1, "Made by James Beer, BeApp Programs - Malicious String Hunter v.1.1"

If MSH_Settings.Report_LogCreateType = 0 Then
Print #1, "Created Log Type: Basic Summary - Scan Performed data"
ElseIf MSH_Settings.Report_LogCreateType = 1 Then
Print #1, "Created Log Type: Advanced Summary - Includes all Strings Found"
ElseIf MSH_Settings.Report_LogCreateType = 2 Then
Print #1, "Created Log Type: Full Summary - Everything the Scan Recorded"
End If

Print #1, "Date/Time of Log: " & Now
Print #1, "Best viewed without Word Wrap enabled"
Print #1, " "

If MSH_Settings.Report_LogCreateType <> 1 Then
With ReportI
    Print #1, "Total Malicious Strings:" & Chr(9) & .TotalMaliciousCount
    Print #1, "PROTENTIAL CLASS:       " & Chr(9) & .POTENTIAL
    Print #1, "SUSPICIOUS CLASS:       " & Chr(9) & .SUSPICIOUS
    Print #1, "CAUTION CLASS:          " & Chr(9) & .CAUTION
    Print #1, "WARNING CLASS:          " & Chr(9) & .WARNING
    Print #1, "DANGER CLASS:           " & Chr(9) & .DANGER
    Print #1, "DESTRUCTIVE CLASS:      " & Chr(9) & .DESTRUCTIVE
    Print #1, " "
    Print #1, "TOTAL LINES:            " & Chr(9) & .TOTALLines
    Print #1, "TOTAL BYTES:            " & Chr(9) & .TOTALBYTES
    Print #1, " "
End With
End If

Print #1, " "

Log_Percent.Value = 20: DoEvents

Print #1, "----- Files Selected for Scan -----"

For i = 1 To UBound(FilesToScan)
    Print #1, FilesToScan(i).Pathname & FilesToScan(i).Filename
Next i
Print #1, "Total of " & UBound(FilesToScan) & " selected file(s)"
Print #1, " "

Log_Percent.Value = 40: DoEvents

Print #1, "----- Project Files -----"
Print #1, "Project" & Chr(9) & "File" & Chr(9) & "Strings Found"
    With Form1
        For i = 1 To .ProjectFiles.ListItems.Count
            Print #1, .ProjectFiles.ListItems.Item(i) & Chr(9) & _
                      .ProjectFiles.ListItems.Item(i).SubItems(1) & Chr(9) & _
                      .ProjectFiles.ListItems.Item(i).SubItems(2)
        Next i
    End With
Print #1, "Total of " & Form1.ProjectFiles.ListItems.Count & " Linked Project File(s)"
Print #1, " "

Log_Percent.Value = 60: DoEvents

If MSH_Settings.Report_LogCreateType <> 1 Then
'A Basic overview of what was found, it only shows the
'string found as how many times it occurs
    Print #1, "----- Basic: Reference Strings Found -----"
    Print #1, "String" & Chr(9) & "Class" & Chr(9) & Chr(9) & "Occurs"
        For i = 1 To Report_MSL.ListItems.Count
        Print #1, Report_MSL.ListItems.Item(i) & Chr(9) & _
                  Report_MSL.ListItems.Item(i).SubItems(1) & Chr(9) & _
                  Report_MSL.ListItems.Item(i).SubItems(2)
        Next i
    Print #1, Report_MSL.ListItems.Count & " Reference String(s) were found"
    Print #1, " "
End If

Log_Percent.Value = 80: DoEvents

If MSH_Settings.Report_LogCreateType >= 1 Then
    Print #1, "----- Advanced: Overall Strings Found -----"
    Print #1, "Source File" & Chr(9) & _
          "Malicious/Accessory String" & Chr(9) & _
          "Line" & Chr(9) & _
          "Class"
    With Form1
        For i = 1 To .MSL.ListItems.Count
        Print #1, .MSL.ListItems.Item(i) & Chr(9) & _
                  .MSL.ListItems.Item(i).SubItems(1) & Chr(9) & Chr(9) & Chr(9) & _
                  .MSL.ListItems.Item(i).SubItems(2) & Chr(9) & _
                  .MSL.ListItems.Item(i).SubItems(3)
        Next i
    End With
End If

Log_Percent.Value = 100: DoEvents

Close #1
End If
Report_Button(0).BackColor = &HC0C0C0
FormLabel(10).Visible = False
Log_Percent.Visible = False
End Function

Private Sub ReportFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Form_MouseMove(Button, Shift, x, Y)
End Sub
