VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Begin VB.Form frmWrkDays 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Working Days Setting . . . . . . . "
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
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
   ScaleHeight     =   3450
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_ReHoli_OvrTime 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4230
      TabIndex        =   17
      Text            =   "txt_ReHoli_OvrTime"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txt_ReDay_OvrTime 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4230
      TabIndex        =   16
      Text            =   "txt_ReDay_OvrTime"
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txt_ReHoli_Time 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1350
      TabIndex        =   11
      Text            =   "txt_ReHoli_Time"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txt_ReDay_Time 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1350
      TabIndex        =   10
      Text            =   "txt_ReDay_Time"
      Top             =   1800
      Width           =   1095
   End
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   450
      Left            =   3405
      TabIndex        =   4
      Top             =   3000
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
      PictureUp       =   "frmWrkDays.frx":0000
      PictureDown     =   "frmWrkDays.frx":0623
      PictureHover    =   "frmWrkDays.frx":0C29
      PictureFocus    =   "frmWrkDays.frx":122F
      PictureDisabled =   "frmWrkDays.frx":1852
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
      Left            =   1995
      TabIndex        =   3
      Top             =   3000
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
      PictureUp       =   "frmWrkDays.frx":1E75
      PictureDown     =   "frmWrkDays.frx":24B7
      PictureHover    =   "frmWrkDays.frx":2AF9
      PictureFocus    =   "frmWrkDays.frx":312C
      PictureDisabled =   "frmWrkDays.frx":376E
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
      Left            =   540
      TabIndex        =   2
      Top             =   3000
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
      PictureUp       =   "frmWrkDays.frx":3B5A
      PictureDown     =   "frmWrkDays.frx":417C
      PictureHover    =   "frmWrkDays.frx":479E
      PictureFocus    =   "frmWrkDays.frx":4DE2
      PictureDisabled =   "frmWrkDays.frx":5404
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
   Begin VB.TextBox txtReWrkDay 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      TabIndex        =   0
      Text            =   "txtReWrkDay"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtReHolDay 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      TabIndex        =   1
      Text            =   "txtReHolDay"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   2520
      X2              =   2520
      Y1              =   1440
      Y2              =   2880
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Holidays : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   2880
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   2880
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Regular Working Over-Time"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   2640
      TabIndex        =   13
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Regular Working Time"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   2445
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Holidays : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   5400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   5400
      X2              =   5400
      Y1              =   480
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Holidays : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   720
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Working Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   720
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Number of Working Days per Month"
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
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5565
   End
End
Attribute VB_Name = "frmWrkDays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_WrkDays As New ADODB.Recordset

Private Sub Populate_Display_Record()
    If Rst_WrkDays.RecordCount > 0 Then
        txtReWrkDay.Text = Rst_WrkDays.Fields(0).Value
        txtReHolDay.Text = Rst_WrkDays.Fields(1).Value
        txt_ReDay_Time.Text = Rst_WrkDays.Fields(2).Value
        txt_ReHoli_Time.Text = Rst_WrkDays.Fields(3).Value
        txt_ReDay_OvrTime.Text = Rst_WrkDays.Fields(4).Value
        txt_ReHoli_OvrTime.Text = Rst_WrkDays.Fields(5).Value
        
    ElseIf Rst_WrkDays.RecordCount <= 0 Then
        txtReWrkDay.Text = "0": txtReHolDay.Text = "0"
        txt_ReDay_Time.Text = "0": txt_ReHoli_Time.Text = "0"
        txt_ReDay_OvrTime.Text = "0": txt_ReHoli_OvrTime.Text = "0"
    End If
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmWrkDays) 'Call 4 WahingOut Text Boxes.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    Call Populate_Display_Record 'To Display Record In Text Boxes.
    txtReWrkDay.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload frmWrkDays
End Sub

Private Sub CmdSubmit_Click()
'    Call Ctrl_PayRoll.msg_Consutruct
    If txtReWrkDay.Text = "" Then MsgBox "Please enter valid value for working days", vbCritical, "Error! Invalid Value": _
                    SendKeys "{Home}+{End}": txtReWrkDay.SetFocus: Exit Sub
    If txtReHolDay.Text = "" Then MsgBox "Please enter valid value for Holidays", vbCritical, "Error! Invalid Value": _
                    SendKeys "{Home}+{End}": txtReHolDay.SetFocus: Exit Sub
    
    If Val(txtReHolDay.Text) = 4 Then
        If Val(txtReWrkDay.Text) + Val(txtReHolDay.Text) <> 30 Then
            MsgBox "4-Holidays then the Regular working days within " & vbCrLf & _
                   "the Combination of 30 days. Please correct it.", vbCritical, "Error! InDays"
                   SendKeys "{Home}+{End}": txtReWrkDay.SetFocus: Exit Sub
        End If
    ElseIf Val(txtReHolDay.Text) = 5 Then
        If Val(txtReWrkDay.Text) + Val(txtReHolDay.Text) <> 31 Then
            MsgBox "5-Holidays then the Regular working days within " & vbCrLf & _
                   "the Combination of 31 days. Please correct it.", vbCritical, "Error! InDays"
                   SendKeys "{Home}+{End}": txtReWrkDay.SetFocus: Exit Sub
        End If
    End If
    With Rst_WrkDays
        If .RecordCount <= 0 Then
            .AddNew
                .Fields(0).Value = Val(txtReWrkDay.Text): .Fields(1).Value = Val(txtReHolDay.Text)
                .Fields(2).Value = Val(txt_ReDay_Time.Text): .Fields(3).Value = Val(txt_ReHoli_Time.Text)
                .Fields(4).Value = Val(txt_ReDay_OvrTime.Text): .Fields(5).Value = Val(txt_ReHoli_OvrTime.Text)
            .Update
        ElseIf .RecordCount > 0 Then
                .Fields(0).Value = Val(txtReWrkDay.Text): .Fields(1).Value = Val(txtReHolDay.Text)
                .Fields(2).Value = Val(txt_ReDay_Time.Text): .Fields(3).Value = Val(txt_ReHoli_Time.Text)
                .Fields(4).Value = Val(txt_ReDay_OvrTime.Text): .Fields(5).Value = Val(txt_ReHoli_OvrTime.Text)
            .Update
        End If
    End With: MsgBox "Record has been saved successfully", vbInformation, "Record Saved"
    Call CmdCancel_Click 'Call For Initialize the Form.
End Sub

Private Sub Form_Load()
    frmWrkDays.Move FrmMain.Width / 3, FrmMain.Height / 5
    Call Ctrl_PayRoll.Populate_Text_Clear(frmWrkDays) 'Call 4 WahingOut Text Boxes.
    Rst_WrkDays.Open "SELECT * FROM tblWorkDay_Set", DB_Conect, adOpenStatic, adLockOptimistic
    Call Populate_Display_Record 'To Display Record In Text Boxes.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rst_WrkDays.Close
End Sub

Private Sub txt_ReDay_OvrTime_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txt_ReDay_OvrTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt_ReDay_OvrTime.Text <> "0" Then
            If txt_ReDay_OvrTime.Text > 5 Then
                Msg_Responce = MsgBox("Please enter the valid Regular Day working Over-Time." & vbCrLf & _
                       "Keep in mind Labour Law." & vbCrLf & _
                       "Otherwise you want to countinue.", vbCritical + vbYesNo, "Error! In Days working Over-Time")
                       If Msg_Responce = vbYes Then
                            SendKeys "{Home}+{End}": txt_ReHoli_OvrTime.SetFocus
                       ElseIf Msg_Responce = vbNo Then
                            txt_ReDay_OvrTime.Text = "0": SendKeys "{Home}+{End}": txt_ReDay_OvrTime.SetFocus
                       End If
            ElseIf txt_ReDay_OvrTime.Text <= 5 Then
                txt_ReHoli_OvrTime.SetFocus
            End If

        ElseIf txt_ReDay_OvrTime.Text = "0" Then
            MsgBox "Please enter the Regular Day working Over-Time." & vbCrLf & _
                   "It must be greater than Zero (0)." & vbCrLf & _
                   "Standard/Default is Five(5) Hours per day.", vbCritical, "Error! In Days working Over-Time"
                   SendKeys "{Home}+{End}": txt_ReDay_OvrTime.SetFocus
        End If
    End If
End Sub

Private Sub txt_ReDay_Time_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txt_ReDay_Time_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt_ReDay_Time.Text <> "0" Then
            If txt_ReDay_Time.Text > 8 Then
                Msg_Responce = MsgBox("Please enter the valid Regular Day working Time." & vbCrLf & _
                       "Keep in mind Labour Law." & vbCrLf & _
                       "Otherwise you want to countinue.", vbCritical + vbYesNo, "Error! In Days working Time")
                       If Msg_Responce = vbYes Then
                            SendKeys "{Home}+{End}": txt_ReHoli_Time.SetFocus
                       ElseIf Msg_Responce = vbNo Then
                            SendKeys "{Home}+{End}": txt_ReDay_Time.SetFocus
                       End If
            ElseIf txt_ReDay_Time.Text <= 8 Then
                txt_ReHoli_Time.SetFocus
            End If
        ElseIf txt_ReDay_Time.Text = "0" Then
            MsgBox "Please enter the Regular Day working Time." & vbCrLf & _
                   "It must be greater than Zero (0)." & vbCrLf & _
                   "Standard/Default is Eight(8) Hours per day.", vbCritical, "Error! In Days working Time"
                   SendKeys "{Home}+{End}": txt_ReDay_Time.SetFocus
        End If
    End If
End Sub

Private Sub txt_ReHoli_OvrTime_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txt_ReHoli_OvrTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt_ReHoli_OvrTime.Text <> "0" Then
            If txt_ReHoli_OvrTime.Text > 5 Then
                Msg_Responce = MsgBox("Please enter the valid Regular Holiday working Over-Time." & vbCrLf & _
                       "Keep in mind Labour Law." & vbCrLf & _
                       "Otherwise you want to countinue.", vbCritical + vbYesNo, "Error! In Holiday working Over-Time")
                       If Msg_Responce = vbYes Then
                            SendKeys "{Home}+{End}": CmdSubmit.SetFocus
                       ElseIf Msg_Responce = vbNo Then
                            txt_ReHoli_OvrTime.Text = "0": SendKeys "{Home}+{End}": txt_ReHoli_OvrTime.SetFocus
                       End If
            ElseIf txt_ReHoli_OvrTime.Text <= 8 Then
                CmdSubmit.SetFocus
            End If

        ElseIf txt_ReHoli_OvrTime.Text = "0" Then
            MsgBox "Please enter the Regular Holiday working Over-Time." & vbCrLf & _
                   "It must be greater than Zero (0)." & vbCrLf & _
                   "Standard/Default is Five(5) Hours per day.", vbCritical, "Error! In Holiday working Over-Time"
                   SendKeys "{Home}+{End}": txt_ReHoli_OvrTime.SetFocus
        End If
    End If
End Sub

Private Sub txt_ReHoli_Time_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txt_ReHoli_Time_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt_ReHoli_Time.Text <> "0" Then
            If txt_ReHoli_Time.Text > 8 Then
                Msg_Responce = MsgBox("Please enter the valid Regular Holidays working Time." & vbCrLf & _
                       "Keep in mind Labour Law." & vbCrLf & _
                       "Otherwise you want to countinue.", vbCritical + vbYesNo, "Error! In Holidays working Time")
                       If Msg_Responce = vbYes Then
                            SendKeys "{Home}+{End}": txt_ReDay_OvrTime.SetFocus
                       ElseIf Msg_Responce = vbNo Then
                            txt_ReHoli_Time.Text = "0": SendKeys "{Home}+{End}": txt_ReHoli_Time.SetFocus
                       End If
            ElseIf txt_ReHoli_Time.Text <= 8 Then
                txt_ReDay_OvrTime.SetFocus
            End If
        ElseIf txt_ReHoli_Time.Text = "0" Then
            MsgBox "Please enter the Regular Holidays working Time." & vbCrLf & _
                   "It must be greater than Zero (0)." & vbCrLf & _
                   "Standard/Default is Eight(8) Hours per day.", vbCritical, "Error! In Holidays working Time"
                   SendKeys "{Home}+{End}": txt_ReHoli_Time.SetFocus
        End If
    End If
End Sub

Private Sub txtReHolDay_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Enter Only Numeric Values.
    If KeyAscii = 13 Then
        If txtReHolDay.Text <> "0" Then
            If txtReHolDay.Text > 10 Then
                MsgBox "Please verify the Holidays." & vbCrLf & _
                       "Not possible more than 10 Holidays", vbCritical, "Error! Monthly Holidays"
                       SendKeys "{Home}+{End}": txtReHolDay.SetFocus
            ElseIf txtReHolDay.Text <= 10 Then
                If (Val(txtReWrkDay.Text) + Val(txtReHolDay.Text) > 31) Then
                    MsgBox "Please verify the total days of Month." & vbCrLf & _
                           "Combination of Working & Holidays not more than 31", vbCritical, "Error! Days Of Month"
                            Call Populate_Display_Record 'To Dispay the Record in Text Boxes.
                            CmdSubmit.Enabled = False: CmdCancel.Enabled = False
                            SendKeys "{Home}+{End}": txtReWrkDay.SetFocus
                ElseIf (Val(txtReWrkDay.Text) + Val(txtReHolDay.Text) <= 31) Then
                    txt_ReDay_Time.SetFocus
                End If
            End If
        End If
    End If
End Sub

Private Sub txtReWrkDay_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtReWrkDay_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Enter Only Numeric Values.
    
    If KeyAscii = 13 Then
        If txtReWrkDay.Text <> "0" Then
            If txtReWrkDay.Text <= "31" Then
                CmdSubmit.Enabled = True: CmdCancel.Enabled = True
                SendKeys "{Home}+{End}": txtReHolDay.SetFocus
            ElseIf txtReWrkDay.Text > "31" Then
                MsgBox "Please Verify the Days of Month", vbCritical, "Error! Days Of Month"
                SendKeys "{Home}+{End}": txtReWrkDay.SetFocus
            End If
        End If
    End If
End Sub
