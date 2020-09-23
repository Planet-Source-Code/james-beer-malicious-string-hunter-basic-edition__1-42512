VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00000040&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don't Show this Again"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1605
      Left            =   1080
      Picture         =   "Splash.frx":0000
      ScaleHeight     =   1605
      ScaleWidth      =   4650
      TabIndex        =   0
      Top             =   240
      Width           =   4650
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "Alias 'Just My Stuff'"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   17
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "Billy Corner"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "Mike Howell"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   15
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "James Beer"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   14
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "James Beer"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   2160
      TabIndex        =   13
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "Roger Gilchrist"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "James Beer"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H005EC6F2&
      BackStyle       =   0  'Transparent
      Caption         =   "James Beer"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ideas/Thanks:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Automatic Search Engine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GUI:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MSH Scan Engine:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Leader:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Splash.frx":1500
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "[Project Scanner 2003] Project Credits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   5535
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4815
      Left            =   120
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
On Error Resume Next
MSH_Settings.ShowSplash = Check1.Value
Kill App.Path & "\Settings.MPC" 'To make sure it doesn't just
'reallocate data into the file instead create a new one.
'Another way would be to buffer the data but that's pointless
'for this kind of file
Open App.Path & "\Settings.MPC" For Binary Access Write Lock Write As #3
Put #3, , MSH_Settings
Close #3
End Sub

Private Sub Command1_Click()
Form1.Enabled = True
Form1.Show
Splash.Hide
End Sub

