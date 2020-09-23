VERSION 5.00
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmUser_Create 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create New User . . . . . . ."
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmUser_Create.frx":0000
   ScaleHeight     =   5085
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   18
      Top             =   3600
      Width           =   1215
      Begin VB.OptionButton Opt_Status 
         BackColor       =   &H8000000E&
         Caption         =   "Later"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   540
         Width           =   975
      End
      Begin VB.OptionButton Opt_Status 
         BackColor       =   &H8000000E&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Previligies"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   14
      Top             =   3600
      Width           =   3015
      Begin VB.OptionButton Opt_Privilige 
         BackColor       =   &H8000000E&
         Caption         =   "Guest"
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
         Index           =   2
         Left            =   1920
         TabIndex        =   17
         Top             =   315
         Width           =   855
      End
      Begin VB.OptionButton Opt_Privilige 
         BackColor       =   &H8000000E&
         Caption         =   "Limited User"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   560
         Width           =   1575
      End
      Begin VB.OptionButton Opt_Privilige 
         BackColor       =   &H8000000E&
         Caption         =   "Administrator"
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
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   280
         Width           =   1575
      End
   End
   Begin LVbuttons.LaVolpeButton CmdExit 
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   4635
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "E&xit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   12582912
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmUser_Create.frx":519D
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   3
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdCreate 
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4635
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Create"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   12582912
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmUser_Create.frx":51B9
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   3
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.TextBox txtConfirmPwd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   9
      Text            =   "txtConfirmPwd"
      Top             =   3135
      Width           =   3015
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   8
      Text            =   "txtPwd"
      Top             =   2685
      Width           =   3015
   End
   Begin VB.TextBox txtUserID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   3840
      TabIndex        =   7
      Text            =   "txtUserID"
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   3840
      TabIndex        =   6
      Text            =   "txtUserName"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Remember ID && Password. If forgot Please Contact Administrator."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   3495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   4320
      Width           =   660
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   5
      Top             =   3135
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   4
      Top             =   2685
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter New User ID : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter User Full Name : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmUser_Create.frx":51D5
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Create New User "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   -45
      Width           =   3615
   End
End
Attribute VB_Name = "frmUser_Create"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstCrtUsr As New ADODB.Recordset
Dim Opt_End As String: Dim MsgRep As String

Private Sub CreateNewUser() 'For Saving the Record in Database.
    With RstCrtUsr
        .AddNew
            If txtUserName.Text = "" Then .Fields(2).Value = "Unknown"
            If txtUserName.Text <> "" Then .Fields(2).Value = txtUserName.Text
            
            .Fields(0).Value = txtUserID.Text: .Fields(1).Value = txtPwd.Text
            .Fields(4).Value = Opt_Stat: .Fields(3).Value = Opt_Priv
        .Update
        MsgBox "New User has been Created Now . . . .", vbInformation, "New User . . ."
    End With: Opt_Priv = ""
End Sub

Private Sub Populate_User_Existance()
    With RstCrtUsr
        .Close: .Open "SELECT * FROM tblUsers WHERE UID='" & txtUserID.Text & "'"
        If .RecordCount > 0 Then
            MsgBox "User Name : " & txtUserID.Text & " already exist." & vbCrLf & _
                   "Please Select Unique User ID.", vbCritical, "Error! User Existance"
                   txtUserName.SetFocus
        Else
            Call CreateNewUser 'Call to Create New User that not Exist.
        End If
    End With
End Sub

Private Sub CmdCreate_Click()
    'Already exist User ID will not Create in Data base Table.
    Call Populate_User_Existance 'Call Function for User Existance if Not Save Record To Create NewUser.
    
    Call Ctrl_PayRoll.Populate_Text_Clear(frmUser_Create) 'Initiate TextBoxes(Clear Textboxes).
    CmdCreate.Enabled = False ': MsgBox "Privilige is as : " & Opt_Priv
    
    For IntI = 0 To 2 'Clearing the Option Buttons.
        Opt_Privilige(IntI).Value = False
    Next
    For IntI = 0 To 1 'Clearing the Option Buttons.
        Opt_Status(IntI).Value = False
    Next
End Sub

Private Sub CmdExit_Click()
    Opt_End = MsgBox("Do you want to cancel the Action." & vbCrLf & _
                     "Please Verify. . . . . . . . . .", vbYesNo + vbCritical, "Cancel User Create")
    If Opt_End = vbYes Then
        Unload frmUser_Create 'Unload User Create Form.
    ElseIf Opt_End = vbNo Then
        Call Ctrl_PayRoll.Populate_Text_Clear(frmUser_Create) 'Call for Clearing the Text Boxes.
        CmdCreate.Enabled = False: SendKeys "{Home}+{End}": txtUserName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    frmUser_Create.Move (FrmMain.Width / 4), (FrmMain.Height / 10)
    Call Ctrl_PayRoll.Populate_Text_Clear(frmUser_Create) 'Call for Clearing the Text Boxes.
    
    RstCrtUsr.Open "SELECT * FROM tblUsers", DB_Conect, adOpenStatic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RstCrtUsr.Close
End Sub

Private Sub Opt_Privilige_DblClick(Index As Integer)
    For IntI = 0 To 2
        If Opt_Privilige(IntI).Value = True Then
            Opt_Priv = Opt_Privilige(IntI).Caption
            CmdCreate.SetFocus
        End If
    Next
End Sub

Private Sub Opt_Privilige_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        For IntI = 0 To 2
            If Opt_Privilige(IntI).Value = True Then
                Opt_Priv = Opt_Privilige(IntI).Caption
                CmdCreate.SetFocus
            End If
        Next
    End If
End Sub

Private Sub Opt_Status_DblClick(Index As Integer)
    For IntI = 0 To 1
        If Opt_Status(IntI).Value = True Then
            Opt_Stat = Opt_Status(IntI).Caption
            Opt_Privilige(0).SetFocus
        End If
    Next
End Sub

Private Sub Opt_Status_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        For IntI = 0 To 1
            If Opt_Status(IntI).Value = True Then
                Opt_Stat = Opt_Status(IntI).Caption
                Opt_Privilige(0).SetFocus
            End If
        Next
    End If
End Sub

Private Sub txtConfirmPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtConfirmPwd.Text <> "" Then
            If (txtPwd.Text = txtConfirmPwd.Text) Then
                Opt_Status(0).SetFocus 'CmdCreate.SetFocus
            ElseIf (txtPwd.Text <> txtConfirmPwd.Text) Then
                MsgBox "Please enter the valid Confirm Password." & vbCrLf & _
                       "Password and Confirm Password must be same", vbCritical, "Error! UnMatched Password. . ."
                SendKeys "{Home}+{End}": txtConfirmPwd.SetFocus
            End If
        ElseIf txtConfirmPwd = "" Then
'            MsgBox "Please enter the valid Confirm Password." & vbCrLf & _
'                   "Without it you can't able to create User", vbCritical + vbYesNo, "Error! Confirm Password. . ."
            Opt_Status(0).SetFocus: CmdCreate.Enabled = True
        End If
    End If
End Sub

Private Sub txtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPwd.Text <> "" Then
            txtConfirmPwd.SetFocus: CmdCreate.Enabled = True
        ElseIf txtPwd = "" Then
            MsgRep = MsgBox("Please enter the Password." & vbCrLf & _
                   "OR Do you want blank Password?", vbCritical + vbYesNo, "Error! User Password. . .")
            If MsgRep = vbNo Then: SendKeys "{Home}+{End}": txtPwd.SetFocus
            If MsgRep = vbYes Then: txtConfirmPwd.SetFocus
        End If
    End If
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If (txtUserID.Text <> "") Then
            If (txtUserName.Text <> "") Then txtUserID.SetFocus: CmdCreate.Enabled = True
            txtPwd.SetFocus: CmdCreate.Enabled = True
        ElseIf txtUserID = "" Then
            MsgBox "Please must be enter the User ID." & vbCrLf & _
                   "Without UserID System Can't Create ID.", vbCritical, "Error! User ID. . ."
            SendKeys "{Home}+{End}": txtUserID.SetFocus
        End If
    End If
End Sub

Private Sub txtUserID_LostFocus()
    If txtUserID.Text <> "" Then CmdCreate.Enabled = True
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_Alpha_Char(KeyAscii, frmUser_Create, txtUserName) 'Call For Alpha Character(Capital).
    If KeyAscii = 13 Then
        If txtUserName.Text <> "" Then
            CmdCreate.Enabled = True: txtUserID.SetFocus 'txtUserID.SetFocus
        ElseIf txtUserName = "" Then
            MsgBox "User Name use for User Identification . . . . . " & vbCrLf & _
                   "If you want to proceed without it . . . . . . . Countinue.", vbInformation, "User Name . . . ."
            txtUserID.SetFocus 'txtUserID.SetFocus
        End If
    End If
End Sub

Private Sub txtUserName_LostFocus()
    If txtUserName.Text <> "" Then CmdCreate.Enabled = True
End Sub
