VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "flash.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   4485
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   120
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   5880
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash FlashImage 
      Height          =   1575
      Left            =   3600
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
      _cx             =   3836
      _cy             =   2778
      FlashVars       =   ""
      Movie           =   "E:\Shahid\Visual Basic\PayRoll System\Images\Payroll.swf"
      Src             =   "E:\Shahid\Visual Basic\PayRoll System\Images\Payroll.swf"
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   3240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblUppermsg 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EVERYONE SHOULD LIVE WITH PEACE AND LOVE BECAUSE IT IS THE BASIS OF THE UNIVERS FROM WE CAN GET OUR AIM."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   -4440
      TabIndex        =   7
      Top             =   0
      Width           =   9330
   End
   Begin VB.Label lblLowermsg 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "feel free! cast vote it will wait for your favour. . . "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   4240
      Width           =   4215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cel # : +923006410758."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Rights.  All Rights Reserved to the Author."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3540
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   3240
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trying to handle the payroll system maybe it is or not plz check and leve ur comments at methoomirza@hotmail.com"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "It designed on special request all favour goes to Requester who need it more."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME . . . . . . . . ."
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      X1              =   1680
      X2              =   1680
      Y1              =   1560
      Y2              =   240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      X1              =   3480
      X2              =   5880
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   2  'Dash
      X1              =   3120
      X2              =   120
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FlashImage.Movie = App.Path & "\Images\Payroll.swf"
    
    frmSplash.Move (FrmMain.Width / 3), (FrmMain.Height / 4)
    lblUppermsg.Caption = "IS THERE ANY VACANT JOB/GIRL. I HAVE VANCAT SEAT FOR BOTH SPECIALLY 4 2nd. hehehehe IF U LIKE . . . . . ."
    lblLowermsg.Caption = "feel free! cast vote it will wait for your favour. . . "
    lblLowermsg.Left = 6000: lblUppermsg.Left = 6100
    Timer2.Interval = 10000: Timer1.Interval = 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Load frmIcons: frmIcons.Show
    frmIcons.lblUserName.Caption = LogIn_UID
End Sub

Private Sub Timer1_Timer()
    If lblLowermsg.Left > -4000 Then
        lblLowermsg.Left = lblLowermsg.Left - 20
    ElseIf lblLowermsg.Left <= -4000 Then
        lblLowermsg.Left = 6000
    End If
    
    If lblUppermsg.Left > -9100 Then
        lblUppermsg.Left = lblUppermsg.Left - 20
    ElseIf lblUppermsg.Left <= -9100 Then
        lblUppermsg.Left = 6100
    End If

End Sub

Private Sub Timer2_Timer()
    Unload frmSplash
End Sub
