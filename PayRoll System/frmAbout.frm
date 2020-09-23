VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About ..."
   ClientHeight    =   4725
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   5475
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3261.279
   ScaleMode       =   0  'User
   ScaleWidth      =   5141.309
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblUnload 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here 4 Exit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1883
      Left            =   3770
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1704
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "About Me"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   360
      Left            =   -270
      TabIndex        =   10
      Top             =   1800
      Width           =   5835
   End
   Begin VB.Label lblvoting 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't forget about voting, im waiting your comments please leave your comments whatever . . . . . . . . . . ."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   810
      Left            =   120
      TabIndex        =   9
      Top             =   3765
      Width           =   3735
   End
   Begin VB.Label lblmobile 
      BackStyle       =   0  'Transparent
      Caption         =   "+92-300-6410758."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   2160
      TabIndex        =   8
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label lblmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "methoomirza@hotmail.com"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label lblcontactinfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Information : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label lblDesignation 
      BackStyle       =   0  'Transparent
      Caption         =   "Working in the I.T Department                   as Assistant  Programmer."
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3615
   End
   Begin VB.Label lblDescribtion 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":35D1
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   210
      TabIndex        =   4
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label lblMyName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M Shahid Aslam Mughal."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Label lblProgrammedby 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designing and Programmed by : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2850
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   4095
      TabIndex        =   1
      Top             =   480
      Width           =   1140
   End
   Begin VB.Label lblFirstLine 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Created in mostly using Windows Development kit "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5025
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "About PayRoll System"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   -360
      TabIndex        =   11
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormslblValues()
    lblFirstLine = " Created in mostly using Windows Development kit "
    lblDescribtion = "This software was design for Managing the Pay Roll System purpose only. It will Save / Manage the Record of Employees, Create limited report can also Add more on demand."
    lblProgrammedby = "Designing and Programmed by : ": lblMyName = "M Shahid Aslam Mughal."
    lblDesignation = "Working in the I.T Department                   as Data Processor Additional Programmer."
    lblcontactinfo = "Contact Information : ": lblmail = "methoomirza@hotmail.com": lblmobile = "+92-300-6410758."
    lblvoting.Caption = "Don't forget about voting, im waiting your comments please leave your comments whatever . . . . . . . . . . ."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload frmAbout
    End If
End Sub

Private Sub Form_Load()
    frmAbout.Move (FrmMain.Width / 3), (FrmMain.Height / 6):  Call FormslblValues 'For Label's Entries.
End Sub

Private Sub lblUnload_Click()
   Unload frmAbout
End Sub
