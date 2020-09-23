VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Begin VB.Form frmLoanApprove 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Benifits Maintains"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Chk_Permission 
      BackColor       =   &H8000000E&
      Caption         =   "Special Permission"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   2600
      Width           =   2175
   End
   Begin VB.TextBox txtDesig 
      Height          =   390
      Left            =   2040
      TabIndex        =   12
      Text            =   "txtDesig"
      Top             =   1080
      Width           =   1845
   End
   Begin VB.TextBox txtEmp_ID 
      Height          =   390
      Left            =   2040
      TabIndex        =   11
      Text            =   "txtEmp_ID"
      Top             =   600
      Width           =   1245
   End
   Begin VB.TextBox txtAppLoanAmt 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2040
      TabIndex        =   6
      Text            =   "txtAppLoanAmt"
      Top             =   2160
      Width           =   1845
   End
   Begin VB.TextBox txtLoanLimit 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   2040
      TabIndex        =   5
      Text            =   "txtLoanLimit"
      Top             =   1680
      Width           =   1845
   End
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   450
      Left            =   3720
      TabIndex        =   3
      Top             =   3360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   794
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
      PictureUp       =   "frmLoanApprove.frx":0000
      PictureDown     =   "frmLoanApprove.frx":0623
      PictureHover    =   "frmLoanApprove.frx":0C29
      PictureFocus    =   "frmLoanApprove.frx":122F
      PictureDisabled =   "frmLoanApprove.frx":1852
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
   Begin StyleButtonX.StyleButton CmdCancel 
      Height          =   450
      Left            =   2520
      TabIndex        =   2
      Top             =   3360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   794
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
      Enabled         =   0   'False
      BackColorUp     =   -2147483634
      BackColorDown   =   -2147483634
      BackColorHover  =   -2147483634
      BackColorFocus  =   -2147483634
      BackColorDisabled=   -2147483634
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmLoanApprove.frx":1E75
      PictureDown     =   "frmLoanApprove.frx":24B7
      PictureHover    =   "frmLoanApprove.frx":2AF9
      PictureFocus    =   "frmLoanApprove.frx":312C
      PictureDisabled =   "frmLoanApprove.frx":376E
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBorderLevel1=   0   'False
      ShowBorderLevel2=   0   'False
   End
   Begin StyleButtonX.StyleButton CmdSubmit 
      Height          =   450
      Left            =   1320
      TabIndex        =   1
      Top             =   3360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   794
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
      Enabled         =   0   'False
      BackColorUp     =   -2147483634
      BackColorDown   =   -2147483634
      BackColorHover  =   -2147483634
      BackColorFocus  =   -2147483634
      BackColorDisabled=   -2147483634
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmLoanApprove.frx":3B5A
      PictureDown     =   "frmLoanApprove.frx":417C
      PictureHover    =   "frmLoanApprove.frx":479E
      PictureFocus    =   "frmLoanApprove.frx":4DE2
      PictureDisabled =   "frmLoanApprove.frx":5404
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBorderLevel1=   0   'False
      ShowBorderLevel2=   0   'False
   End
   Begin StyleButtonX.StyleButton CmdNew 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   794
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
      PictureUp       =   "frmLoanApprove.frx":5800
      PictureDown     =   "frmLoanApprove.frx":5E10
      PictureHover    =   "frmLoanApprove.frx":6420
      PictureFocus    =   "frmLoanApprove.frx":6A3E
      PictureDisabled =   "frmLoanApprove.frx":704E
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
   Begin VB.Label lblEmpName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "lblEmpName"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   1560
      TabIndex        =   14
      Top             =   2960
      Width           =   3090
   End
   Begin VB.Label lblDesig_ID 
      BackStyle       =   0  'Transparent
      Caption         =   "lblDesig_ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   390
      Left            =   3480
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee - ID : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   1725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Approved : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1845
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Range : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   4800
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   4800
      X2              =   4800
      Y1              =   480
      Y2              =   3240
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Designation : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Approve Loan Maintainance"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4965
   End
End
Attribute VB_Name = "frmLoanApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_EmpInfo As New ADODB.Recordset: Dim Rst_Desig As New ADODB.Recordset
Dim Rst_Loan_Set As New ADODB.Recordset: Dim Rst_Loan_App As New ADODB.Recordset
Dim Rst_Loan_Paid As New ADODB.Recordset
Dim LoanDes As String: Dim LoanAmt As Integer

Private Sub Pop_Save_Record()
    With Rst_Loan_App
        .AddNew
            .Fields(0).Value = txtEmp_ID.Text: .Fields(1).Value = txtLoanLimit.Text
            .Fields(2).Value = txtAppLoanAmt.Text: .Fields(3).Value = LoanDes 'For Special Permission.
            .Fields(4).Value = Date
        .Update
    End With: LoanDes = ""
    MsgBox "Record has been saved successfully.", vbCritical, "Record Save"
End Sub

Private Sub Chk_Permission_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdSubmit.Enabled = True Then CmdSubmit.SetFocus
        If CmdSubmit.Enabled = False Then Call Ctrl_PayRoll.msg_Consutruct
    End If
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmLoanApprove) 'Call 4 WahingOut Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmLoanApprove, False) 'To Not Allow Enteries.
    lblEmpName = "": Chk_Permission.Value = 0
    
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    CmdNew.Enabled = True: CmdNew.SetFocus
    
End Sub

Private Sub CmdExit_Click()
    Unload frmLoanApprove
End Sub

Private Sub CmdNew_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmLoanApprove) 'To Clearing the Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmLoanApprove, True) 'To Allow Enteries.
    
    CmdSubmit.Enabled = True: CmdCancel.Enabled = True
    CmdNew.Enabled = False: txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus
End Sub

Private Sub CmdSubmit_Click()
    If Chk_Permission.Value = 1 Then LoanDes = Chk_Permission.Caption
    If Chk_Permission.Value = 0 Then LoanDes = "Ordinary"
    
    If txtEmp_ID.Text = "" Then MsgBox "Please enter valid value in Employee ID.", vbCritical, "Error! Employee ID": _
                           txtEmp_ID.Text = "": txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus: Exit Sub
                           
    If txtAppLoanAmt.Text = "" Then MsgBox "Please enter valid value in Applied Loan Section.", vbCritical, "Error! Applied Loan Entry": _
                           txtAppLoanAmt.Text = "": txtAppLoanAmt.Text = "EMP-": SendKeys "{Home}+{End}": txtAppLoanAmt.SetFocus: Exit Sub

    Call Pop_Save_Record 'For Saving the Record in Database Table.
    Call CmdCancel_Click 'To Initiate the Form for next entry.
End Sub

Private Sub Form_Load()
    frmLoanApprove.Move FrmMain.Width / 3, FrmMain.Height / 8
    Call Ctrl_PayRoll.Populate_Text_Clear(frmLoanApprove) 'To Clearing the Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmLoanApprove, False) 'Not Allow entry
    lblEmpName = "": Chk_Permission.Value = 0
    
    Rst_EmpInfo.Open "SELECT * FROM tblEmployee_Info", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Desig.Open "SELECT * FROM tblDesignation", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Loan_Set.Open "SELECT * FROM tblLoan_Set", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Loan_Paid.Open "SELECT * FROM tblLoan_Paid", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Loan_App.Open "SELECT * FROM tblLoan_Approved", DB_Conect, adOpenStatic, adLockOptimistic
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rst_EmpInfo.Close: Rst_Desig.Close
    Rst_Loan_Set.Close: Rst_Loan_App.Close
    Rst_Loan_Paid.Close
End Sub

Private Sub txtAppLoanAmt_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Only Numeric values.
    If KeyAscii = 13 Then
        If Val(txtAppLoanAmt.Text) <= Val(txtLoanLimit.Text) Then
                CmdSubmit.SetFocus
        ElseIf Val(txtAppLoanAmt.Text) > Val(txtLoanLimit.Text) Then
            If Chk_Permission.Value = 1 Then
                CmdSubmit.SetFocus
            ElseIf Chk_Permission.Value = 0 Then
                MsgBox "This Loan is exceed from upper-limit." & vbCrLf & _
                       "Is the Special Permission allowed. If then Please chek it.", vbCritical, "Error! In Loan Approved"
                       SendKeys "{Home}+{End}": txtAppLoanAmt.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtDesig_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtEmp_ID_Change()
    If CmdNew.Enabled = False Then
        Rst_EmpInfo.Close: Rst_EmpInfo.Open "SELECT * FROM tblEmployee_Info WHERE Emp_ID='" & txtEmp_ID.Text & "'"
        If Rst_EmpInfo.RecordCount > 0 Then
            lblEmpName.Caption = Rst_EmpInfo.Fields(5).Value & " " & Rst_EmpInfo.Fields(6).Value
            lblDesig_ID = Rst_EmpInfo.Fields(2).Value
            
            Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation WHERE Desig_ID='" & lblDesig_ID & "'"
            txtDesig.Text = Rst_Desig.Fields(1).Value
            
            Rst_Loan_Set.Close: Rst_Loan_Set.Open "SELECT * FROM tblLoan_Set WHERE Desig_ID='" & lblDesig_ID & "'"
            txtLoanLimit.Text = Rst_Loan_Set.Fields(1).Value
'============================================================= If Already Applied Loan ===============
           Rst_Loan_App.Close: Rst_Loan_App.Open "SELECT * FROM tblLoan_Approved WHERE Emp_ID='" & txtEmp_ID.Text & "'"
           If Rst_Loan_App.RecordCount > 0 Then
                txtAppLoanAmt.Text = Rst_Loan_App.Fields(2).Value
'=====================================================================================================
'                Rst_Loan_Paid.Close: Rst_Loan_Paid.Open "SELECT * FROM tblLoan_Paid WHERE Emp_ID='" & txtEmp_ID.Text & "'"
'                If Rst_Loan_Paid.RecordCount > 0 Then
'                    For IntI = 1 To Rst_Loan_Paid.RecordCount
'                        LoanAmt = LoanAmt + .Fields(1).Value
'                    Next
'                    If LoanAmt = Val(txtAppLoanAmt.Text) Then
'                    If LoanAmt < Val(txtAppLoanAmt.Text) Then
'
'                End If
           ElseIf Rst_Loan_App.RecordCount <= 0 Then
           End If
        ElseIf Rst_EmpInfo.RecordCount <= 0 Then
            lblEmpName = "": lblDesig_ID = "": txtDesig.Text = "": txtLoanLimit.Text = ""
        End If
    End If
End Sub

Private Sub txtEmp_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEmp_ID.Text <> "" Then
        
            If Rst_EmpInfo.RecordCount <= 0 Then
                lblEmpName = "Record Not Exist In Database." 'Due to Active Change Event this Message didn't Display."
                txtEmp_ID.Text = "": txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus
            ElseIf Rst_EmpInfo.RecordCount > 0 Then
                txtAppLoanAmt.SetFocus
            End If
            
        ElseIf txtEmp_ID.Text = "" Then
            txtEmp_ID.Text = "": txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus
        End If
        Rst_EmpInfo.Close: Rst_EmpInfo.Open "SELECT * FROM tblEmployee_Info"
        Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation"
        Rst_Loan_Set.Close: Rst_Loan_Set.Open "SELECT * FROM tblLoan_Set"
        Rst_Loan_App.Close: Rst_Loan_App.Open "SELECT * FROM tblLoan_Approved"
    End If
End Sub

Private Sub txtLoanLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
