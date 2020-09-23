VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5355
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
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
   ScaleHeight     =   4785
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   495
      Left            =   3840
      TabIndex        =   12
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      UpColorTop1     =   -2147483628
      UpColorTop2     =   -2147483633
      UpColorTop3     =   -2147483633
      UpColorTop4     =   -2147483633
      UpColorButtom1  =   -2147483627
      UpColorButtom2  =   -2147483633
      UpColorButtom3  =   -2147483633
      UpColorButtom4  =   -2147483633
      UpColorLeft1    =   -2147483628
      UpColorLeft2    =   -2147483633
      UpColorLeft3    =   -2147483633
      UpColorLeft4    =   -2147483633
      UpColorRight1   =   -2147483627
      UpColorRight2   =   -2147483633
      UpColorRight3   =   -2147483633
      UpColorRight4   =   -2147483633
      DownColorTop1   =   -2147483627
      DownColorTop2   =   -2147483633
      DownColorTop3   =   -2147483633
      DownColorTop4   =   -2147483633
      DownColorButtom1=   -2147483628
      DownColorButtom2=   -2147483633
      DownColorButtom3=   -2147483633
      DownColorButtom4=   -2147483633
      DownColorLeft1  =   -2147483627
      DownColorLeft2  =   -2147483633
      DownColorLeft3  =   -2147483633
      DownColorLeft4  =   -2147483633
      DownColorRight1 =   -2147483628
      DownColorRight2 =   -2147483633
      DownColorRight3 =   -2147483633
      DownColorRight4 =   -2147483633
      HoverColorTop1  =   -2147483628
      HoverColorTop2  =   -2147483633
      HoverColorTop3  =   -2147483633
      HoverColorTop4  =   -2147483633
      HoverColorButtom1=   -2147483627
      HoverColorButtom2=   -2147483633
      HoverColorButtom3=   -2147483633
      HoverColorButtom4=   -2147483633
      HoverColorLeft1 =   -2147483628
      HoverColorLeft2 =   -2147483633
      HoverColorLeft3 =   -2147483633
      HoverColorLeft4 =   -2147483633
      HoverColorRight1=   -2147483627
      HoverColorRight2=   -2147483633
      HoverColorRight3=   -2147483633
      HoverColorRight4=   -2147483633
      FocusColorTop1  =   -2147483628
      FocusColorTop2  =   -2147483633
      FocusColorTop3  =   -2147483633
      FocusColorTop4  =   -2147483633
      FocusColorButtom1=   -2147483627
      FocusColorButtom2=   -2147483632
      FocusColorButtom3=   -2147483633
      FocusColorButtom4=   -2147483633
      FocusColorLeft1 =   -2147483628
      FocusColorLeft2 =   -2147483633
      FocusColorLeft3 =   -2147483633
      FocusColorLeft4 =   -2147483633
      FocusColorRight1=   -2147483627
      FocusColorRight2=   -2147483632
      FocusColorRight3=   -2147483633
      FocusColorRight4=   -2147483633
      DisabledColorTop1=   -2147483628
      DisabledColorTop2=   -2147483633
      DisabledColorTop3=   -2147483633
      DisabledColorTop4=   -2147483633
      DisabledColorButtom1=   -2147483627
      DisabledColorButtom2=   -2147483633
      DisabledColorButtom3=   -2147483633
      DisabledColorButtom4=   -2147483633
      DisabledColorLeft1=   -2147483628
      DisabledColorLeft2=   -2147483633
      DisabledColorLeft3=   -2147483633
      DisabledColorLeft4=   -2147483633
      DisabledColorRight1=   -2147483627
      DisabledColorRight2=   -2147483633
      DisabledColorRight3=   -2147483633
      DisabledColorRight4=   -2147483633
      Caption         =   ""
      BackColorUp     =   -2147483634
      BackColorDown   =   -2147483634
      BackColorHover  =   -2147483634
      BackColorFocus  =   -2147483634
      BackColorDisabled=   -2147483634
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmLogin.frx":0000
      PictureDown     =   "frmLogin.frx":0749
      PictureHover    =   "frmLogin.frx":0E82
      PictureFocus    =   "frmLogin.frx":15BB
      PictureDisabled =   "frmLogin.frx":1CF4
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBorderLevel1=   0   'False
      ShowBorderLevel2=   0   'False
   End
   Begin StyleButtonX.StyleButton CmdOk 
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      UpColorTop1     =   -2147483628
      UpColorTop2     =   -2147483633
      UpColorTop3     =   -2147483633
      UpColorTop4     =   -2147483633
      UpColorButtom1  =   -2147483627
      UpColorButtom2  =   -2147483633
      UpColorButtom3  =   -2147483633
      UpColorButtom4  =   -2147483633
      UpColorLeft1    =   -2147483628
      UpColorLeft2    =   -2147483633
      UpColorLeft3    =   -2147483633
      UpColorLeft4    =   -2147483633
      UpColorRight1   =   -2147483627
      UpColorRight2   =   -2147483633
      UpColorRight3   =   -2147483633
      UpColorRight4   =   -2147483633
      DownColorTop1   =   -2147483627
      DownColorTop2   =   -2147483633
      DownColorTop3   =   -2147483633
      DownColorTop4   =   -2147483633
      DownColorButtom1=   -2147483628
      DownColorButtom2=   -2147483633
      DownColorButtom3=   -2147483633
      DownColorButtom4=   -2147483633
      DownColorLeft1  =   -2147483627
      DownColorLeft2  =   -2147483633
      DownColorLeft3  =   -2147483633
      DownColorLeft4  =   -2147483633
      DownColorRight1 =   -2147483628
      DownColorRight2 =   -2147483633
      DownColorRight3 =   -2147483633
      DownColorRight4 =   -2147483633
      HoverColorTop1  =   -2147483628
      HoverColorTop2  =   -2147483633
      HoverColorTop3  =   -2147483633
      HoverColorTop4  =   -2147483633
      HoverColorButtom1=   -2147483627
      HoverColorButtom2=   -2147483633
      HoverColorButtom3=   -2147483633
      HoverColorButtom4=   -2147483633
      HoverColorLeft1 =   -2147483628
      HoverColorLeft2 =   -2147483633
      HoverColorLeft3 =   -2147483633
      HoverColorLeft4 =   -2147483633
      HoverColorRight1=   -2147483627
      HoverColorRight2=   -2147483633
      HoverColorRight3=   -2147483633
      HoverColorRight4=   -2147483633
      FocusColorTop1  =   -2147483628
      FocusColorTop2  =   -2147483633
      FocusColorTop3  =   -2147483633
      FocusColorTop4  =   -2147483633
      FocusColorButtom1=   -2147483627
      FocusColorButtom2=   -2147483632
      FocusColorButtom3=   -2147483633
      FocusColorButtom4=   -2147483633
      FocusColorLeft1 =   -2147483628
      FocusColorLeft2 =   -2147483633
      FocusColorLeft3 =   -2147483633
      FocusColorLeft4 =   -2147483633
      FocusColorRight1=   -2147483627
      FocusColorRight2=   -2147483632
      FocusColorRight3=   -2147483633
      FocusColorRight4=   -2147483633
      DisabledColorTop1=   -2147483628
      DisabledColorTop2=   -2147483633
      DisabledColorTop3=   -2147483633
      DisabledColorTop4=   -2147483633
      DisabledColorButtom1=   -2147483627
      DisabledColorButtom2=   -2147483633
      DisabledColorButtom3=   -2147483633
      DisabledColorButtom4=   -2147483633
      DisabledColorLeft1=   -2147483628
      DisabledColorLeft2=   -2147483633
      DisabledColorLeft3=   -2147483633
      DisabledColorLeft4=   -2147483633
      DisabledColorRight1=   -2147483627
      DisabledColorRight2=   -2147483633
      DisabledColorRight3=   -2147483633
      DisabledColorRight4=   -2147483633
      Caption         =   ""
      BackColorUp     =   -2147483634
      BackColorDown   =   -2147483634
      BackColorHover  =   -2147483634
      BackColorFocus  =   -2147483634
      BackColorDisabled=   -2147483634
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmLogin.frx":242D
      PictureDown     =   "frmLogin.frx":2C6A
      PictureHover    =   "frmLogin.frx":34A7
      PictureFocus    =   "frmLogin.frx":3CE4
      PictureDisabled =   "frmLogin.frx":4521
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBorderLevel1=   0   'False
      ShowBorderLevel2=   0   'False
   End
   Begin VB.CheckBox Chk_Unmask 
      BackColor       =   &H8000000E&
      Caption         =   "Unmasked Password"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Txt_Password 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "Txt_Password"
      Top             =   2565
      Width           =   2655
   End
   Begin VB.ComboBox CmbUserID 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblUserName 
      BackStyle       =   0  'Transparent
      Caption         =   "lblUserName"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Name : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblTries 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   2040
      TabIndex        =   10
      Top             =   4200
      Width           =   120
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remaining Tries : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   270
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   1635
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Login Window"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Note : "
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
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Checking the Unmasked Password (Upper Option) will retrive the entered Password in Alpha Numeric Character."
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1680
      TabIndex        =   6
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLogin.frx":4D5E
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
      Height          =   1215
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2565
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Select User ID : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1365
      Left            =   3960
      Picture         =   "frmLogin.frx":4E07
      Top             =   4920
      Width           =   1320
   End
   Begin VB.Image Image3 
      Height          =   1005
      Left            =   120
      Picture         =   "frmLogin.frx":AC21
      Top             =   720
      Width           =   1200
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_UID As New ADODB.Recordset: Dim LogInTries As Integer

Private Sub CmbUserId_Click()
    If CmbUserID.Text <> "Select User" Then
    
        Rst_UID.Close: Rst_UID.Open "SELECT * FROM tblUsers WHERE UID='" & CmbUserID.Text & "'"
        If Rst_UID.RecordCount > 0 Then lblUserName = Rst_UID.Fields(2).Value
        If Rst_UID.RecordCount <= 0 Then lblUserName = "Unknown"
        
    ElseIf CmbUserID.Text = "Select User" Then
        lblUserName = "Unknown"
    End If
End Sub

Private Sub CmbUserId_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Txt_Password.SetFocus
    End If
End Sub


Private Sub CmdExit_Click()
    Opt_End = MsgBox("Sure you Shuting Down the System." & vbCrLf & _
                     "Please Verify. . . . . . . . . .", vbYesNo + vbCritical, "Shut Down")
    If Opt_End = vbYes Then
        Call Ctrl_PayRoll.Deplode(frmLogin): End
    ElseIf Opt_End = vbNo Then
        Call Ctrl_PayRoll.Populate_Text_Clear(frmLogin) 'Call for Clearing the Text Boxes.
        CmdOk.Enabled = True: Chk_Unmask.Value = Unchecked
        CmbUserID.SetFocus
    End If
End Sub

Private Sub CmdOk_Click()
    LogInTries = LogInTries - 1
    
    If Trim(CmbUserID.Text) <> "Select User" Then
        Rst_UID.Close: Rst_UID.Open "SELECT * FROM tblUsers WHERE UID='" & CmbUserID.Text & "'"
        If Rst_UID.EOF = False Then
            If Trim(Rst_UID.Fields(1).Value) <> Trim(Txt_Password.Text) Then
                MsgBox "Password is not correct!" & vbCrLf & _
                       "Please try again ... " & vbCrLf & "You have " & LogInTries & _
                       " Tries Left.", vbCritical, "Password! Error": lblTries = LogInTries
                Txt_Password.Text = "": SendKeys "{Home}+{End}"
                Txt_Password.SetFocus: Exit Sub
                
            ElseIf Trim(Rst_UID.Fields(1).Value) = Trim(Txt_Password.Text) Then
                Load FrmMain: FrmMain.Show: Load frmSplash: frmSplash.Show
                LogIn_UID = lblUserName: LogIn_Time = Date & " " & Time: Unload Me
            End If
        End If
    ElseIf Trim(CmbUserID.Text) = "Select User" Then
        MsgBox "Please Select The Valid UserId" & vbCrLf & "You have " & LogInTries & _
               " Tries Left.", vbInformation, "Error! User ID"
               lblTries = LogInTries: CmbUserID.SetFocus
        Exit Sub
    ElseIf Trim(Txt_Password.Text) = "" Then
        MsgBox "Please Enter Valid Password" & vbCrLf & "You have " & LogInTries & _
        " Tries Left.", vbInformation, " Error! Password"
        Txt_Password.SetFocus: lblTries = LogInTries: Exit Sub
    ElseIf Val(lblTries) <= 1 Then
        MsgBox "You are Un-Authorised User." & vbCrLf & _
               "Please contact your administrator." & vbCrLf & _
               "Session has been Terminate.", vbCritical, " Error! Invalid User . . .": End
    End If
End Sub
Private Sub Form_Load()
    frmLogin.Move (FrmMain.Width / 2), (FrmMain.Height / 4)
    Call Ctrl_PayRoll.Explode(frmLogin)  'Call 4 load form with Explode Property.
    Call Ctrl_PayRoll.Populate_Text_Clear(frmLogin) 'To Clear the textboxes of the Form.
    lblUserName = "Unknown": LogInTries = 5
    
    
    Rst_UID.Open "SELECT * FROM tblUsers", DB_Conect, adOpenStatic, adLockOptimistic
    If Rst_UID.RecordCount > 0 Then
        CmbUserID.AddItem "Select User"
        Do While Not Rst_UID.EOF = True
            CmbUserID.AddItem Trim(Rst_UID.Fields(0).Value)
            If Rst_UID.EOF = False Then Rst_UID.MoveNext
            If Rst_UID.EOF = True Then Exit Do
        Loop
        CmbUserID.Text = "Select User"
    ElseIf Rst_UID.RecordCount <= 0 Then
        Load frmUser_Create: frmUser_Create.Show
        Me.WindowState = 1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Rst_UID.Close
'    frmMain.lblUser = CmbUserId.Text
'    frmMain.lblDate = Date & "    " & Time
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub Txt_Password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If Txt_Password.Text <> "" Then
            Call CmdOk_Click 'For Next Processing
'        End If
    End If
End Sub
