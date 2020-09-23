VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BeApp Programs - Malicious String Hunter v1.2 [Project Scanner 2003 Project] by James Beer"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   11775
   Icon            =   "MSH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   785
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add_File"
            Object.ToolTipText     =   "Adds a file to the scan list"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Remove_File"
            Object.ToolTipText     =   "Remove a selected file from the scan list"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Scan_Files"
            Object.ToolTipText     =   "Scan every file listed"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Show_Report"
            Object.ToolTipText     =   "Shows the existing report"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Erase_Info"
            Object.ToolTipText     =   "Clears the lists"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Change the way the scan and report works"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Auto_Search"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar MSHStatus 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8640
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   13124
            MinWidth        =   13124
            Picture         =   "MSH.frx":0442
            Text            =   "Not Ready"
            TextSave        =   "Not Ready"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox MainScreen 
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   120
      Picture         =   "MSH.frx":0A18
      ScaleHeight     =   537
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   769
      TabIndex        =   2
      Top             =   480
      Width           =   11535
      Begin VB.TextBox Path_Label 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Timer LightEffect1 
         Interval        =   25
         Left            =   9840
         Top             =   1560
      End
      Begin VB.TextBox CodeLine_Label 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   10
         Top             =   7320
         Width           =   10455
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   6840
         Picture         =   "MSH.frx":701F
         ScaleHeight     =   107
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   310
         TabIndex        =   9
         Top             =   120
         Width           =   4650
      End
      Begin VB.Frame ShowWork 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   4200
         TabIndex        =   3
         Top             =   3000
         Visible         =   0   'False
         Width           =   3735
         Begin VB.PictureBox WorkForeGround 
            BackColor       =   &H00000080&
            BorderStyle     =   0  'None
            ForeColor       =   &H00000000&
            Height          =   1095
            Left            =   120
            ScaleHeight     =   1095
            ScaleWidth      =   3495
            TabIndex        =   4
            Top             =   360
            Width           =   3495
            Begin VB.Label Working_Button 
               Alignment       =   2  'Center
               BackColor       =   &H00000080&
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
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   2160
               TabIndex        =   6
               Top             =   720
               Width           =   1215
            End
            Begin VB.Shape Working_Button_OL 
               BorderWidth     =   3
               Height          =   255
               Left            =   2160
               Top             =   720
               Width           =   1215
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   9
               Left            =   1800
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   8
               Left            =   1680
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   7
               Left            =   1560
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   6
               Left            =   1440
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   5
               Left            =   1320
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   4
               Left            =   1200
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   3
               Left            =   1080
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   2
               Left            =   960
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   1
               Left            =   840
               Top             =   720
               Width           =   135
            End
            Begin VB.Shape Shape2 
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   0
               Left            =   720
               Top             =   720
               Width           =   135
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   120
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Scanning in Progress...  Please Wait"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Left            =   720
               TabIndex        =   5
               Top             =   120
               Width           =   2655
            End
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Malicious String Hunter is Working"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   3735
         End
      End
      Begin MSComctlLib.ListView MSL 
         Height          =   3015
         Left            =   240
         TabIndex        =   8
         Top             =   4200
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   4497
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Malicious/Accessory String"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Line"
            Object.Width           =   7232
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Class"
            Object.Width           =   2664
         EndProperty
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   9840
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9360
         Top             =   2640
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":851F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":9171
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":977E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":9D54
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":A385
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":AFD7
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":BC29
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":C87B
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":D4CD
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":E11F
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":ED71
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":F1C3
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":F615
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":FA67
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView FTSList 
         Height          =   1695
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Path"
            Object.Width           =   5010
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Filename"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9360
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   60
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":FEB9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":10B0B
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":11118
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":116EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":11D1F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":12971
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":135C3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":14215
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":14E67
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":15AB9
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":1670B
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":16B5D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":16FAF
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MSH.frx":17401
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ProjectFiles 
         Height          =   1695
         Left            =   6720
         TabIndex        =   11
         Top             =   2160
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Project"
            Object.Width           =   2752
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File"
            Object.Width           =   2778
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Strings Found"
            Object.Width           =   2099
         EndProperty
      End
      Begin VB.Label FormLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         Caption         =   "Scan Information/Progress:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   29
         Top             =   0
         Width           =   2415
      End
      Begin VB.Image ListOptionsIcon 
         Height          =   240
         Index           =   3
         Left            =   4650
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image ListOptionsIcon 
         Height          =   240
         Index           =   1
         Left            =   4650
         Stretch         =   -1  'True
         Top             =   2340
         Width           =   255
      End
      Begin VB.Image ListOptionsIcon 
         Height          =   240
         Index           =   0
         Left            =   4650
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image ListOptionsIcon 
         Height          =   240
         Index           =   2
         Left            =   4650
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   240
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Process File:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label File_Label 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   23
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Located At:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Size_Label 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0 KB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "File Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label FileType_Label 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "<No File>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Files to Scan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label FormLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Files (If Project File Loaded): "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   14
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   7320
         Width           =   495
      End
      Begin VB.Label FormLabel 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Malicious String List:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label ListOptions 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Add File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Shape ListOptions_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   0
         Left            =   4920
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label ListOptions 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Scan File(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   28
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Shape ListOptions_OL 
         BorderWidth     =   3
         FillColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   2
         Left            =   4920
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label ListOptions 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Auto-Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   27
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Shape ListOptions_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   3
         Left            =   4920
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label ListOptions 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Remove File"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   26
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Shape ListOptions_OL 
         BorderWidth     =   3
         Height          =   255
         Index           =   1
         Left            =   4920
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000C0&
         Height          =   3735
         Left            =   120
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_AddFile 
         Caption         =   "Add &File/Project"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu_File_RemoveFile 
         Caption         =   "Remove Selected File/Project"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnu_File_RemoveAll 
         Caption         =   "Remove All File(s)"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_File_SearchForFiles 
         Caption         =   "Automatic Searching"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu_File_Scan 
         Caption         =   "&Scan File(s)/Project(s)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnu_Info 
      Caption         =   "&Information"
      Begin VB.Menu mnu_Info_Clear 
         Caption         =   "Clear Information"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_Info_Report 
         Caption         =   "Display Existing Report"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnu_Options 
      Caption         =   "&Options"
      Begin VB.Menu mnu_Options_Settings 
         Caption         =   "General Settings"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_Options_AllowSorting 
         Caption         =   "Allow Column Sorting"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu_About 
      Caption         =   "&Credits"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ScanFilesListed()
Dim i As Long

Image1(0).Picture = ImageList2.ListImages.Item(5).Picture
Label7.Caption = "Scanning in progress..." & Chr(13) & "Please Wait"
Working_Button.Caption = "Cancel"
MSHStatus.Panels(1).Picture = Form1.ImageList1.ListImages.Item(5).Picture
MSHStatus.Panels(1) = "Scanning in Progress..."

If InUse = False Then
MenuAndToolBarEnabled False

    InUse = True
    EndScan = False

    ReDim GetErrors(0)

        For i = 0 To MS_Total
            ReportI.MSOccur(i) = 0
        Next i

    ReportI.TOTALLines = 0
    ReportI.TotalMaliciousCount = 0
    ReportI.POTENTIAL = 0
    ReportI.SUSPICIOUS = 0
    ReportI.CAUTION = 0
    ReportI.WARNING = 0
    ReportI.DANGER = 0
    ReportI.DESTRUCTIVE = 0
    ReportI.TOTALBYTES = 0

    DoEvents
    
    CodeLine_Label = Empty
    ClearList 'clears the lists
        
        i = 0
        Do Until i >= UBound(FilesToScan)
        i = i + 1
            Path_Label = FilesToScan(i).Pathname
            Scan FilesToScan(i).Pathname, FilesToScan(i).Filename
        Loop

    If EndScan = False Then
        CreateReport
        InUse = False
    Else
        InUse = False
        Path_Label = Empty
        File_Label = Empty
        ClearList
        MSHStatus.Panels.Item(1) = "Scan Cancelled - Awaiting new scan"
        MSHStatus.Panels.Item(1).Picture = Form1.ImageList1.ListImages.Item(3).Picture
    
    End If
    
Else

    MSHStatus.Panels.Item(1) = "Scan is in Use or Failed to Shutdown - Cannot Reinitiate"
    MSHStatus.Panels.Item(1).Picture = Form1.ImageList1.ListImages.Item(4).Picture
    
End If

MenuAndToolBarEnabled True
End Sub

Private Sub MenuAndToolBarEnabled(ByVal OnOff As Boolean)
'Another piece of code from Roger Gilchrist
Dim i As Long
    With Toolbar1.Buttons
    For i = 1 To .Count
            .Item(i).Enabled = OnOff
        Next i
    End With
        mnu_File_AddFile.Enabled = OnOff
        mnu_File_RemoveFile.Enabled = OnOff
        mnu_File_Scan.Enabled = OnOff
        mnu_File_SearchForFiles.Enabled = OnOff
        mnu_File_RemoveAll.Enabled = OnOff
        mnu_Info_Clear.Enabled = OnOff
        mnu_Info_Report.Enabled = OnOff
        mnu_Options_Settings.Enabled = OnOff
        ProjectFiles.Visible = OnOff
        MSL.Visible = OnOff
        FTSList.Enabled = OnOff
        For i = 5 To 7
            FormLabel(i).Visible = OnOff
        Next i
        CodeLine_Label.Visible = OnOff
        
        If OnOff = False Then
            ProgressLight = 0
            ShowWork.Visible = True
            LightEffect1.Enabled = True
        Else
            ProgressLight = 0
            ShowWork.Visible = False
            LightEffect1.Enabled = False
        End If
'Roger I modified this sub slightly
'and added the command boxs (Well Now they're labels to
'make it look better and have those images in them)

        For i = 0 To 3
        ListOptions(i).Enabled = OnOff
        ListOptions(i).BackColor = C_Button_Normal
        ListOptions(i).ForeColor = C_Button_NormalText
        Next i
        
If UBound(FilesToScan) = 0 Then
    ListOptions(1).Enabled = False
    mnu_File_RemoveFile.Enabled = False
End If
End Sub

Private Sub Form_Load()
Dim i As Long

Open App.Path & "\Settings.MPC" For Binary Access Read Lock Read As #3
Get #3, , MSH_Settings
Close #3
LoadOptions 1
LoadOptions 2
mnu_Options_AllowSorting.Checked = MSH_Settings.GUI_AllowSorting

ReDim FilesToScan(0)

For i = 0 To ListOptions.UBound
    ListOptions(i).BackColor = C_Button_Normal
    ListOptions(i).ForeColor = C_Button_NormalText
Next i


ListOptionsIcon(0).Picture = ImageList1.ListImages(11).Picture
ListOptionsIcon(1).Picture = ImageList1.ListImages(12).Picture
ListOptionsIcon(2).Picture = ImageList2.ListImages(5).Picture
ListOptionsIcon(3).Picture = ImageList2.ListImages(9).Picture


MSHStatus.Panels.Item(1) = "Not Ready"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
Load_MSData 'Obtains the String Reference file to this program

If MSH_Settings.ShowSplash = 0 Then
Form1.Hide
Splash.Show
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
TerminateApp
End Sub

Private Sub LightEffect1_Timer()
Dim i As Long

'This progress bar effect is for program that have undetermined
'processing times, it indicates only to the user that it's
'not frozen

ProgressLight = ProgressLight + 1

If ProgressLight = 10 Then ProgressLight = 0
For i = 0 To 9
    Form1.Shape2(i).FillColor = vbBlack
Next i

If ProgressLight > 0 Then
    Shape2(ProgressLight - 1).FillColor = &HC000&
Else
    Shape2(9).FillColor = &HC000&
End If

If ProgressLight > 1 Then
    Shape2(ProgressLight - 2).FillColor = &H8000&
Else
    Shape2(9 - (1 - ProgressLight)).FillColor = &HC000&
End If

Shape2(ProgressLight).FillColor = vbGreen
End Sub

Private Sub ListOptions_Click(Index As Integer)
Select Case Index
    Case Is = 0
        AddFile
    Case Is = 1
        RemoveFile
    Case Is = 2
        ScanFilesListed
    Case Is = 3
    
        GetAutoSearchInfo
End Select

ListOptions(Index).BackColor = C_Button_Normal
ListOptions(Index).ForeColor = C_Button_NormalText
End Sub

Private Sub ListOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
If ListOptions(Index).BackColor <> C_Button_Select Then
For i = 0 To (ListOptions.Count - 1)
ListOptions(i).BackColor = C_Button_Normal
ListOptions(i).ForeColor = C_Button_NormalText
Next i
ListOptions(Index).BackColor = C_Button_Select
ListOptions(Index).ForeColor = C_Button_SelectText
End If
End Sub

Private Sub ListOptionsIcon_Click(Index As Integer)
Call ListOptions_Click(Index)
End Sub

Private Sub ListOptionsIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If ListOptions(Index).Enabled = True Then
Call ListOptions_MouseMove(Index, Button, Shift, x, Y)
End If
End Sub


Private Sub MainScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i As Long
For i = 0 To ListOptions.Count - 1
If ListOptions(i).BackColor = C_Button_Select Then
ListOptions(i).BackColor = C_Button_Normal
ListOptions(i).ForeColor = C_Button_NormalText
End If
Next i
If Working_Button.BackColor <> C_Button_Normal Then
    Working_Button.BackColor = C_Button_Normal
    Working_Button.ForeColor = C_Button_NormalText
End If
End Sub

Private Sub mnu_About_Click()
ShowSplash
End Sub

Private Sub mnu_File_AddFile_Click()
AddFile
End Sub

Private Sub mnu_File_Exit_Click()
Dim x As Integer

If InUse = True Then
MsgBox "A Scan is still running, Are you sure?", vbExclamation + vbYesNo, "Scanning In Progress"
If x = vbYes Then TerminateApp
Else
TerminateApp
End If

End Sub


Private Sub mnu_File_RemoveAll_Click()
Dim x As Long
x = MsgBox("Are you sure you want to Remove all files from scan list?", vbYesNo + vbExclamation, "Removal Notify")
If x = vbNo Then Exit Sub
FTSList.ListItems.Clear
ReDim FilesToScan(0)
mnu_File_RemoveFile.Enabled = False
mnu_File_RemoveAll.Enabled = False
ListOptions(1).Enabled = False
End Sub

Private Sub mnu_File_RemoveFile_Click()
RemoveFile
End Sub

Private Sub mnu_File_Scan_Click()
ScanFilesListed
End Sub

Private Sub mnu_File_SearchForFiles_Click()
GetAutoSearchInfo
End Sub

Private Sub mnu_Info_Clear_Click()
If InUse = False Then ClearList
End Sub

Private Sub mnu_Info_Report_Click()
If InUse = False Then
ReDim GetErrors(0)
CreateReport
End If
End Sub

Private Sub mnu_Options_AllowSorting_Click()

If mnu_Options_AllowSorting.Checked = False Then
mnu_Options_AllowSorting.Checked = True
MSH_Settings.GUI_AllowSorting = True
Else
mnu_Options_AllowSorting.Checked = False
MSH_Settings.GUI_AllowSorting = False
End If

Kill App.Path & "\Settings.MPC"
Open App.Path & "\Settings.MPC" For Binary Access Write Lock Write As #3
Put #3, , MSH_Settings
Close #3

End Sub

Private Sub mnu_Options_Settings_Click()
If InUse = False Then
ShowOptions
End If
End Sub

Private Sub MSL_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If MSH_Settings.GUI_AllowSorting = True Then MSL.Sorted = True Else MSL.Sorted = False
MSL.SortKey = (ColumnHeader.Index - 1)
If MSL.SortOrder = lvwAscending Then
   MSL.SortOrder = lvwDescending
Else
   MSL.SortOrder = lvwAscending
End If
MSL.Refresh
End Sub

Private Sub MSL_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
CodeLine_Label = Item.SubItems(2)
CodeLine_Label.ForeColor = Item.ListSubItems(2).ForeColor
End Sub


Private Sub ProjectFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
If MSH_Settings.GUI_AllowSorting = True Then ProjectFiles.Sorted = True Else ProjectFiles.Sorted = False
ProjectFiles.SortKey = (ColumnHeader.Index - 1)
If ProjectFiles.SortOrder = lvwAscending Then
   ProjectFiles.SortOrder = lvwDescending
Else
   ProjectFiles.SortOrder = lvwAscending
End If
ProjectFiles.Refresh
End Sub

Private Sub ProjectFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim i As Long
For i = 1 To ProjectFiles.ListItems.Count
If i <> Item.Index Then
    ProjectFiles.ListItems.Item(i).ForeColor = vbBlack
    ProjectFiles.ListItems.Item(i).Bold = False
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).ForeColor = vbBlack
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).Bold = False
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbBlack
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).Bold = False
Else
    ProjectFiles.ListItems.Item(i).ForeColor = vbBlue
    ProjectFiles.ListItems.Item(i).Bold = True
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).ForeColor = vbBlue
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(1).Bold = True
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).ForeColor = vbBlue
    ProjectFiles.ListItems.Item(i).ListSubItems.Item(2).Bold = True
End If
Next i
For i = 1 To MSL.ListItems.Count
    If MSL.ListItems.Item(i) = Item.SubItems(1) Then
        MSL.ListItems.Item(i).ForeColor = vbBlue
    Else
        MSL.ListItems.Item(i).ForeColor = vbBlack
    End If
Next i
ProjectFiles.Refresh
MSL.Refresh
End Sub


Private Sub ShowWork_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Working_Button.BackColor <> C_Button_Normal Then
    Working_Button.BackColor = C_Button_Normal
    Working_Button.ForeColor = C_Button_NormalText
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "Add_File"
    AddFile
Case Is = "Remove_File"
    RemoveFile
Case Is = "Scan_Files"
    ScanFilesListed
Case Is = "Show_Report"
    ReDim GetErrors(0)
    CreateReport
Case Is = "Erase_Info"
    ClearList
Case Is = "Options"
    ShowOptions
Case Is = "Auto_Search"
    GetAutoSearchInfo
End Select
End Sub

Private Sub WorkForeGround_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Working_Button.BackColor <> C_Button_Normal Then
    Working_Button.BackColor = C_Button_Normal
    Working_Button.ForeColor = C_Button_NormalText
End If
End Sub

Private Sub Working_Button_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Working_Button.BackColor <> C_Button_Normal Then
    Working_Button.BackColor = C_Button_Normal
    Working_Button.ForeColor = C_Button_NormalText
End If
EndScan = True
End Sub

Private Sub Scan(ByVal PathN As String, ByVal FileN As String)
Dim i As Long
Dim x As Long

'I seperated the path and file data into two variables because
'I wanted the list to only include file names of where that
'string orginated from
On Error GoTo Err_Check
If UCase(Right$(FileN, 4)) = ".VBP" Then
ProjectFile = FileN
GetVBPFiles PathN & FileN
i = 0
    Do Until i >= UBound(SubFiles)
    'since the loop is based on a dynamic array which
    'expands as it's loops, DO...LOOP is the only other
    'fastest way other than FOR...NEXT
    i = i + 1
        Filename = SubFiles(i) 'The files of which this
        'project file has
        
        Size_Label = GetFileSizeProper(FileLen(PathN & Filename))
        ReportI.TOTALBYTES = ReportI.TOTALBYTES + FileLen(PathN & Filename)
        'NOTE: The buffer only handles the amount of bytes
        'for each file as they're processed, then it's
        'erased and refilled with the next file's data
        'the buffer doesn't expand to accompany all files.
        'Only the one it's scanning.
        
        LoadFile PathN, Filename 'loads and scans the file instantly
        ProjectFiles.ListItems.Add , , ProjectFile
        ProjectFiles.ListItems.Item(ProjectFiles.ListItems.Count).SubItems(1) = SubFiles(i)
        ProjectFiles.ListItems.Item(ProjectFiles.ListItems.Count).SubItems(2) = MaliciousCount
        ProjectFiles.ListItems.Item(ProjectFiles.ListItems.Count).SmallIcon = IIf(MaliciousCount > 0, 3, 1)
    If i > 500 Then Exit Do ' A failsafe if it loops
    'out of control, Doevents slows it down too much
    'so this is a alternative way to prevent lock-ups.
    'NO Visual Basic project should ever excess this amount
    'Well I've never seen one that does anyway.
    Loop
    Filename = ProjectFile 'Since the scan is primarily based
    'on the project file it should give it's name back in case
    'of a re-scan.
Else
    Filename = FileN
    Size_Label = GetFileSizeProper(FileLen(PathN & FileN))
    ReportI.TOTALBYTES = ReportI.TOTALBYTES + FileLen(PathN & FileN)
    LoadFile PathN, FileN 'loads and scans the file instantly
End If
Exit Sub

Err_Check:
AddErrorMsg Err.Description, " " & Filename
Resume Next
End Sub

Function GetFileSizeProper(ByVal Size As Single) As String
Select Case Size
Case Is < 1024
    GetFileSizeProper = Size & " Bytes"
Case 1024 To ((1024 ^ 2) - 1)
    GetFileSizeProper = CCur(Size / 1024) & " KB"
Case Is >= (1024 ^ 2)
    GetFileSizeProper = CCur(Size / (1024 ^ 2)) & " MB"
End Select
End Function

Public Sub AutoSearch(ByVal StartingPath As String)
Dim Directories() As String
Dim Projects() As String
Dim a As String
Dim i As Long
Dim ua As String
Dim AddedDirs As Long
Dim GAR As Integer

MenuAndToolBarEnabled False

ReDim Directories(0)
ReDim Projects(0)


a = Dir(StartingPath, vbDirectory)
AddedDirs = 1

Directories(0) = StartingPath

'This searchs for every (or most) directory in a main directory you specify
'it might not be 100% as this is my first attempt at a file searching engine
'but by shock it works pretty damn well!

Do Until AddedDirs = 0 Or EndScan = True
restart:
    Do While a <> Empty And EndScan = False
        'Get Every Single Sub Directory
        'NOTE: 16 = vbDirectory Flag
        ua = UCase(Right(a, 4))
        GAR = GetAttr(StartingPath & a)
        If a <> "." And a <> ".." Then
            If GAR = vbDirectory Or GAR = (vbDirectory + vbReadOnly) Or _
               GAR = (vbDirectory + vbHidden) Then
               'the Attributes I had problems with got furstrated for a
               'while but it was so simple to fix, don't you hate when that
               'happens?
               'The GAR means Get Attributes Results, and if the name is
               'directory with read-only, hidden or neither, it's listed.
               
            
           AddedDirs = AddedDirs + 1 'Adds one to this to indicate how
           'many sub directories have been found under this directory.
           'If this returns zero then the outer loop will restart the
           'inner loop with one of the currently listed dirs to see
           'if there's other sub-directories.
           
            ReDim Preserve Directories(UBound(Directories) + 1)
            Directories(UBound(Directories)) = StartingPath & a & "\"
            End If
        End If
        If a <> "." And a <> ".." And GetAttr(StartingPath & a) <> vbDirectory And _
            ua = ".VBP" Then
            'Now this gets all the project files in the directory as it
            'goes, saves time.
            ReDim Preserve Projects(UBound(Projects) + 1)
            Projects(UBound(Projects)) = StartingPath & a
            ReDim Preserve FilesToScan(UBound(FilesToScan) + 1)
            FilesToScan(UBound(FilesToScan)).Pathname = StartingPath
            FilesToScan(UBound(FilesToScan)).Filename = a
            FTSList.ListItems.Add , , FilesToScan(UBound(FilesToScan)).Pathname
            FTSList.ListItems.Item(FTSList.ListItems.Count).SubItems(1) = FilesToScan(UBound(FilesToScan)).Filename
        End If
        
        MSHStatus.Panels(1) = "Searching... " & StartingPath

        a = Dir
        DoEvents
    Loop
AddedDirs = 0

If MSH_Settings.AutoSearch_IncludeSubs = False Then Exit Do

On Error Resume Next
i = i + 1
a = Dir(Directories(i), vbDirectory)
StartingPath = Directories(i)
If a <> Empty And EndScan = False Then GoTo restart

Loop

If MSH_Settings.AutoSearch_AutoScanAfter = 1 Then ScanFilesListed
MSHStatus.Panels(1).Picture = Form1.ImageList1.ListImages.Item(2).Picture
MSHStatus.Panels(1) = "Search Complete"

MenuAndToolBarEnabled True

End Sub

Function AddFile()
Dim x As Integer
Dim i As Long
Dim Existing As Boolean
CD1.Filename = Empty 'So if you cancel it, it will have
'no filename string pre-stored therefore won't execute
'the add file process
CD1.DialogTitle = "Select File"
CD1.Filter = "Visual Basic 6.0 Projects|*.vbp|" & _
             "Visual Basic 6.0 Forms|*.frm|" & _
             "Visual Basic 6.0 Modules|*.bas|" & _
             "Visual Basic 6.0 Class Modules|*.cls|" & _
             "Visual Basic 6.0 Resource File|*.res|" & _
             "Visual Basic 6.0 Form Binary Files|*.frx|" & _
             "Text Document|*.txt|" & _
             "VB Script|*.vbs"
'File filters, in case you're a newbie interested these define
'the file types that should be allowed to be added.
CD1.ShowOpen 'Easy isn't it, one line give you about
             '30 lines of action plus even more
If CD1.FilterIndex <> 1 And CD1.FilterIndex <> 7 And _
   CD1.FilterIndex <> 8 And CD1.Filename <> Empty Then
'For if you chose a VB file which really doesn't
'need to be added as a project file could be referenced to it
x = MsgBox("You can scan it's project file instead to scan every linked file in it" & Chr(13) & _
           "this is more efficient and save adding every file, Do you still want to do this?", vbExclamation + vbYesNo, "Optimize Scanning Tip")
If x = vbNo Then Exit Function
End If

If CD1.Filename <> Empty And Existing = False Then
For i = 1 To UBound(FilesToScan)
'to find if an existing file is already added
'this prevent wasting memory on the same file
If Left$(CD1.Filename, Len(CD1.Filename) - Len(CD1.FileTitle)) = FilesToScan(i).Pathname And _
   CD1.FileTitle = FilesToScan(i).Filename Then
Existing = True
End If
If Existing = True Then Exit For
Next i

If Existing = True Then MsgBox "Duplicate file and path found, you don't need to scan it twice", vbExclamation, "File/Path conflict": Exit Function

'Add the file to the array which is processed
If UBound(FilesToScan) = 0 Then
    ReDim FilesToScan(1 To 1)
Else
    ReDim Preserve FilesToScan(1 To (UBound(FilesToScan) + 1))
End If

FilesToScan(UBound(FilesToScan)).Pathname = Left$(CD1.Filename, Len(CD1.Filename) - Len(CD1.FileTitle))
FilesToScan(UBound(FilesToScan)).Filename = CD1.FileTitle
FTSList.ListItems.Add , , FilesToScan(UBound(FilesToScan)).Pathname
FTSList.ListItems.Item(FTSList.ListItems.Count).SubItems(1) = FilesToScan(UBound(FilesToScan)).Filename

ListOptions(1).Enabled = True
ListOptions(1).BackColor = C_Button_Normal
ListOptions(1).ForeColor = C_Button_NormalText

mnu_File_RemoveFile.Enabled = True
mnu_File_RemoveAll.Enabled = True

MSHStatus.Panels.Item(1) = "Ready - " & UBound(FilesToScan) & " File(s) Selected"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(2).Picture
End If
End Function

Private Sub RemoveFile()
Dim i As Long
If FTSList.ListItems.Count > 0 Then
FTSList.ListItems.Remove (FTSList.SelectedItem.Index)

ReDim FilesToScan(FTSList.ListItems.Count)
For i = 1 To FTSList.ListItems.Count
FilesToScan(i).Pathname = FTSList.ListItems.Item(i)
FilesToScan(UBound(FilesToScan)).Filename = FTSList.ListItems.Item(i).SubItems(1)
Next i
End If

If FTSList.ListItems.Count = 0 Then
ListOptions(1).Enabled = False
mnu_File_RemoveFile.Enabled = False
mnu_File_RemoveAll.Enabled = False
ReDim FilesToScan(0)
MSHStatus.Panels.Item(1) = "Not Ready"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
End If
End Sub

Function CreateReport()
Dim i As Long
'This generate a brief report which is easier to
'understand, the option to save a log of this
'scan is optional as long as there's malicious code
'found.
'This acts like a error handler and debug too
ReportWindow.Show
Form1.Enabled = False

With ReportWindow
For i = 0 To MS_Total
If ReportI.MSOccur(i) > 0 Then
.Report_MSL.ListItems.Add , , MS(i).MStr
.Report_MSL.ListItems.Item(.Report_MSL.ListItems.Count).SubItems(1) = MS(i).Class
.Report_MSL.ListItems.Item(.Report_MSL.ListItems.Count).SubItems(2) = ReportI.MSOccur(i)
End If
Next i
.Result(0) = ReportI.TotalMaliciousCount
.Result(1) = ReportI.POTENTIAL
.Result(2) = ReportI.SUSPICIOUS
.Result(3) = ReportI.CAUTION
.Result(4) = ReportI.WARNING
.Result(5) = ReportI.DANGER
.Result(6) = ReportI.DESTRUCTIVE
.Result(7) = ReportI.TOTALLines
.Result(8) = GetFileSizeProper(ReportI.TOTALBYTES)

If MSL.ListItems.Count > 0 Then
'this show the report that the scan worked successfully but
'found some malicious strings :(
.ScanDoneLabel = "Scan Completed Successfully - Malicious/Accessory Strings Found"
.Image2.Picture = Form1.ImageList2.ListImages.Item(3).Picture
.Report_Header = "This code contains malicious/Accessory code and should not be executed without Investigating"
.Report_Information.Visible = True
.Report_Button(0).Visible = True
.Report_Button_OL(0).Visible = True
MSHStatus.Panels.Item(1) = "Scan Completed Successfully - Malicous/Accessory Strings Found"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
Else
'same thing except no malicious strings were found :)
.ScanDoneLabel = "Scan Completed Successfully - Data is Clean"
.Image2.Picture = Form1.ImageList2.ListImages.Item(1).Picture
.Report_Header = "This code is considered clean by MSH program resources"
.Report_Information.Visible = True
.Report_Button(0).Visible = False
.Report_Button_OL(0).Visible = False
MSHStatus.Panels.Item(1) = "Scan Completed Successfully - All Clean"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(2).Picture
End If

If UBound(FilesToScan) = 0 Then
'If the filestoscan array has no files listed
    .ScanDoneLabel = "Scan Cannot Initiate - No File Specified"
    .Image2.Picture = Form1.ImageList2.ListImages.Item(2).Picture
    .Report_Header = "No projects or files were selected to scan, please add files by using the Add File button"
    .Report_Information.Visible = False
    .Report_Button(0).Visible = False
    .Report_Button_OL(0).Visible = False
    Beep
    MSHStatus.Panels.Item(1) = "Not Ready - Error"
    MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(3).Picture
End If

If UBound(GetErrors) > 0 Then
    'if it completed but encountered Visual Basic Run-time errors
    .ScanDoneLabel = "Scan Completed but Errors Occurred"
    .Image2.Picture = Form1.ImageList2.ListImages.Item(4).Picture
    .Report_Header = "The Follow Errors Occurred while Scanning:"
        For i = 1 To UBound(GetErrors)
            .Report_Header = .Report_Header & Chr(13) & Chr(10) & _
            GetErrors(i).Def & ":" & GetErrors(i).File
            .Report_Information.Visible = True
            .Report_Button(0).Visible = True
            .Report_Button_OL(0).Visible = True
        Next i

MSHStatus.Panels.Item(1) = "Scan Complete but with Errors"
MSHStatus.Panels.Item(1).Picture = ImageList1.ListImages.Item(4).Picture
End If

For i = 0 To ReportWindow.Report_Button.UBound
    .Report_Button(i).BackColor = C_Button_Normal
    .Report_Button(i).ForeColor = C_Button_NormalText
Next i

On Error Resume Next
If EndScan = True Or UBound(FilesToScan) = 0 Then Exit Function
If MSH_Settings.Report_AutoCreateLog = True Then
    .Report_Button(0).Visible = False
    .Report_Button_OL(0).Visible = False
    
    CD1.Filename = App.Path & "\" & MSH_Settings.Report_DefaultLogFilename & ".log"
    FileLen App.Path & "\" & MSH_Settings.Report_DefaultLogFilename & ".log"
    
    If Err.Number <> 53 And MSH_Settings.Report_OverWriteLog = False Then
    Err.Clear 'Using this clear the last error that was
    'created so the next function can regenerate it.
        i = 0
            Do Until Err.Number = 53 Or i > 500
            'Obtains which increment number has been used
            '* Instead of open statement use filelen because
            'it generates the same error and save you time
            i = i + 1
            FileLen App.Path & "\" & MSH_Settings.Report_DefaultLogFilename & i & ".log"
                DoEvents
            Loop
        CD1.Filename = App.Path & "\" & MSH_Settings.Report_DefaultLogFilename & i & ".log"
    End If
.CreateLog
.Report_MSL.Refresh
End If
End With
End Function

Private Sub GetAutoSearchInfo()
Dim x As Long
x = MsgBox("All Previous Files selected to scan will be discarded and this process could take a while, Are you Sure you want to perform Auto-Search?", vbYesNo + vbExclamation)
If x = vbNo Then GoTo endit
EndScan = False
Image1(0).Picture = ImageList2.ListImages.Item(9).Picture
Label7.Caption = "MSH is searching for projects..."
Working_Button.Caption = "Stop"

        FTSList.ListItems.Clear
        ReDim FilesToScan(0)

        MSHStatus.Panels(1).Picture = Form1.ImageList1.ListImages.Item(9).Picture
        MSHStatus.Panels(1) = "Searching in Progress..."

        Select Case MSH_Settings.AutoSearch_TypeOfSearch
            Case Is = 0
                AutoSearch "C:\"
            Case Is = 1
                AutoSearch MSH_Settings.AutoSearch_DefaultDir
            Case Is = 2
            
            For x = 0 To DirDialog.DirDialog_Button.UBound
                DirDialog.DirDialog_Button(x).BackColor = C_Button_Normal
                DirDialog.DirDialog_Button(x).ForeColor = C_Button_NormalText
            Next x
            
                DirDialogFlag = "AutoSearchPromptDir"
                Form1.Enabled = False
                DirDialog.Show
            End Select

Exit Sub
endit:
ListOptions(3).BackColor = C_Button_Normal
ListOptions(3).ForeColor = C_Button_NormalText
End Sub


Private Sub ShowOptions()
Dim i As Long
Form1.Enabled = False

For i = 0 To Options.Options_Button.UBound
    Options.Options_Button(i).BackColor = C_Button_Normal
    Options.Options_Button(i).ForeColor = C_Button_NormalText
Next i

LoadOptions 2
Options.Show
End Sub

Private Sub ShowSplash()
Form1.Enabled = False
Splash.Show
Splash.Check1 = MSH_Settings.ShowSplash
End Sub

Private Sub Working_Button_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Working_Button.BackColor = C_Button_Select
Working_Button.ForeColor = C_Button_SelectText
End Sub
