VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Begin VB.Form frmTaxes 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Funds and Taxes Maintains"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6180
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
   ScaleHeight     =   3975
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLoanRange 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   4800
      TabIndex        =   27
      Text            =   "txtLoanRange"
      Top             =   2880
      Width           =   885
   End
   Begin VB.TextBox txtHRent 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   4800
      TabIndex        =   18
      Text            =   "txtHRent"
      Top             =   1920
      Width           =   885
   End
   Begin VB.TextBox txtMedical 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   4800
      TabIndex        =   17
      Text            =   "txtMedical"
      Top             =   2400
      Width           =   885
   End
   Begin VB.TextBox txtWHTax 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1440
      TabIndex        =   12
      Text            =   "txtWHTax"
      Top             =   2880
      Width           =   1000
   End
   Begin VB.TextBox txtInTax 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1440
      TabIndex        =   11
      Text            =   "txtInTax"
      Top             =   2400
      Width           =   1000
   End
   Begin VB.TextBox txtBFund 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   1440
      TabIndex        =   10
      Text            =   "txtBFund"
      Top             =   1920
      Width           =   1000
   End
   Begin VB.ComboBox CmbDesig 
      Height          =   390
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   600
      Width           =   2200
   End
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   450
      Left            =   4800
      TabIndex        =   3
      Top             =   3480
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
      PictureUp       =   "frmTaxes.frx":0000
      PictureDown     =   "frmTaxes.frx":0623
      PictureHover    =   "frmTaxes.frx":0C29
      PictureFocus    =   "frmTaxes.frx":122F
      PictureDisabled =   "frmTaxes.frx":1852
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
      Left            =   3240
      TabIndex        =   2
      Top             =   3480
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
      PictureUp       =   "frmTaxes.frx":1E75
      PictureDown     =   "frmTaxes.frx":24B7
      PictureHover    =   "frmTaxes.frx":2AF9
      PictureFocus    =   "frmTaxes.frx":312C
      PictureDisabled =   "frmTaxes.frx":376E
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
   Begin StyleButtonX.StyleButton CmdSubmit 
      Height          =   450
      Left            =   1680
      TabIndex        =   1
      Top             =   3480
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
      PictureUp       =   "frmTaxes.frx":3B5A
      PictureDown     =   "frmTaxes.frx":417C
      PictureHover    =   "frmTaxes.frx":479E
      PictureFocus    =   "frmTaxes.frx":4DE2
      PictureDisabled =   "frmTaxes.frx":5404
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
   Begin StyleButtonX.StyleButton CmdNew 
      Height          =   450
      Left            =   120
      TabIndex        =   0
      Top             =   3480
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
      PictureUp       =   "frmTaxes.frx":5800
      PictureDown     =   "frmTaxes.frx":5E10
      PictureHover    =   "frmTaxes.frx":6420
      PictureFocus    =   "frmTaxes.frx":6A3E
      PictureDisabled =   "frmTaxes.frx":704E
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
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Limit  : "
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
      Left            =   2880
      TabIndex        =   26
      Top             =   2880
      Width           =   1845
   End
   Begin VB.Label lblDesig_ID 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "lblDesig_ID"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "% Of Benifits Applied"
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
      Left            =   2880
      TabIndex        =   24
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "% Funds && Taxes Applied"
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
      TabIndex        =   23
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Medcal Allowance : "
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
      Left            =   2880
      TabIndex        =   22
      Top             =   2400
      Width           =   1845
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   5640
      TabIndex        =   21
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   5640
      TabIndex        =   20
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Rent : "
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
      Left            =   2880
      TabIndex        =   19
      Top             =   1920
      Width           =   1845
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   6000
      X2              =   6000
      Y1              =   1800
      Y2              =   3360
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter value in %age without %age sign"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   360
      TabIndex        =   16
      Top             =   1080
      Width           =   3720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   120
      Y1              =   1800
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   6000
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   2760
      X2              =   2760
      Y1              =   1800
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   6000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   2400
      TabIndex        =   15
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   2400
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   2400
      TabIndex        =   13
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "W/H Tax : "
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
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Income Tax : "
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
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B - Funds : "
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
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
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
      Left            =   960
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Funds, Taxes && Loans Maintains"
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
      TabIndex        =   4
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmTaxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Desig As New ADODB.Recordset: Dim Rst_Taxes As New ADODB.Recordset
Dim Rst_Loan As New ADODB.Recordset: Dim Rst_Allow As New ADODB.Recordset

Private Sub Populate_SaveRecord()
    If Opt_Flag = "Add" Then
        With Rst_Taxes
            .AddNew 'For Taxes / Funds Records.
                .Fields(0).Value = lblDesig_ID: .Fields(1).Value = Val(txtBFund.Text)
                .Fields(2).Value = Val(txtInTax.Text): .Fields(3).Value = Val(txtWHTax.Text)
            .Update
        End With
        With Rst_Loan 'For Loan Limits Records.
            .AddNew
                .Fields(0).Value = lblDesig_ID: .Fields(1).Value = Val(txtLoanRange.Text)
            .Update
        End With
    
        With Rst_Allow 'For Benifits / Allowances Records.
            .AddNew
                .Fields(0).Value = lblDesig_ID: .Fields(1).Value = Val(txtHRent.Text)
                .Fields(2).Value = Val(txtMedical.Text)
            .Update
        End With: MsgBox "Record has been saved succeccfully", vbInformation, "Record Saving"
'==================================================================================================================
    ElseIf Opt_Flag = "Edit" Then
        Rst_Taxes.Close: Rst_Taxes.Open "SELECT * FROM tblTaxes_Funds_Set WHERE Desig_ID='" & lblDesig_ID & "'"
        Rst_Loan.Close: Rst_Loan.Open "SELECT * FROM tblLoan_Set WHERE Desig_ID='" & lblDesig_ID & "'"
        Rst_Allow.Close: Rst_Allow.Open "SELECT * FROM tblAllowances_Set WHERE Desig_ID='" & lblDesig_ID & "'"

        With Rst_Taxes 'For Taxes / Funds Records.
                .Fields(1).Value = Val(txtBFund.Text)
                .Fields(2).Value = Val(txtInTax.Text): .Fields(3).Value = Val(txtWHTax.Text)
            .Update
        End With
        
        With Rst_Loan 'For Loan Limits Records.
                .Fields(1).Value = Val(txtLoanRange.Text)
            .Update
        End With
    
        With Rst_Allow 'For Benifits / Allowances Records.
                .Fields(1).Value = Val(txtHRent.Text): .Fields(2).Value = Val(txtMedical.Text)
            .Update
        End With: MsgBox "Record has been Modified Succeccfully", vbInformation, "Record Modify"
    End If: Opt_Flag = ""
    Rst_Taxes.Close: Rst_Taxes.Open "SELECT * FROM tblTaxes_Funds_Set"
    Rst_Loan.Close: Rst_Loan.Open "SELECT * FROM tblLoan_Set": Rst_Allow.Close: Rst_Allow.Open "SELECT * FROM tblAllowances_Set"
End Sub

Private Sub CmbDesig_Click()
    Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation WHERE Desig_Name='" & CmbDesig.Text & "'"
    If Rst_Desig.RecordCount > 0 Then lblDesig_ID = Rst_Desig.Fields(0).Value
    If Rst_Desig.RecordCount <= 0 Then lblDesig_ID = ""
End Sub

Private Sub CmbDesig_GotFocus()
    Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation"
    Call Ctrl_PayRoll.Populate_Init_Cmb(Rst_Desig, 1, CmbDesig) 'To Initialize the Combo Boxes Enteries.
End Sub

Private Sub CmbDesig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbDesig.Text = "Choose" Then
            MsgBox "Please choose the correct selection." & vbCrLf & _
                   "It is not valid selection.", vbCritical, "Error! Department Selection"
                   CmbDesig.SetFocus
                   
        ElseIf CmbDesig.Text = "" Then
            MsgBox "Please choose the correct selection." & vbCrLf & _
                   "It is not valid selection.", vbCritical, "Error! Department Selection"
                   CmbDesig.SetFocus: Exit Sub
                   
        ElseIf CmbDesig.Text <> "Choose" Then
            Rst_Taxes.Close: Rst_Taxes.Open "SELECT * FROM tblTaxes_Funds_Set WHERE Desig_ID='" & lblDesig_ID & "'"
            Rst_Loan.Close: Rst_Loan.Open "SELECT * FROM tblLoan_Set WHERE Desig_ID='" & lblDesig_ID & "'"
            Rst_Allow.Close: Rst_Allow.Open "SELECT * FROM tblAllowances_Set WHERE Desig_ID='" & lblDesig_ID & "'"
            
            If Rst_Taxes.RecordCount > 0 Then
                Msg_Responce = MsgBox(CmbDesig.Text & "'s Taxes and Funds %age already exit." & vbCrLf & _
                              "Do you want to modify them.", vbYesNo + vbCritical, "Taxes & Funds Existance")
                              
                               txtBFund.Text = Rst_Taxes.Fields(1).Value: txtInTax.Text = Rst_Taxes.Fields(2).Value
                               txtWHTax.Text = Rst_Taxes.Fields(3).Value
                               txtHRent.Text = Rst_Allow.Fields(1).Value: txtMedical.Text = Rst_Allow.Fields(2).Value
                               txtLoanRange.Text = Rst_Loan.Fields(1).Value
                                    
                    If Msg_Responce = vbYes Then 'Modify the Records.
                        SendKeys "{Home}+{End}": txtBFund.SetFocus: Opt_Flag = "Edit"
                    ElseIf Msg_Responce = vbNo Then 'If didn't want modify.
                        Call Ctrl_PayRoll.Populate_Text_Clear(frmTaxes)
                        Call CmbDesig_GotFocus: CmbDesig.SetFocus
                    End If
            ElseIf Rst_Taxes.RecordCount <= 0 Then 'Then Enter New Record.
                SendKeys "{Home}+{End}": txtBFund.SetFocus
            End If
        End If
    End If
    Rst_Taxes.Close: Rst_Taxes.Open "SELECT * FROM tblTaxes_Funds_Set"
    Rst_Loan.Close: Rst_Loan.Open "SELECT * FROM tblLoan_Set"
    Rst_Allow.Close: Rst_Allow.Open "SELECT * FROM tblAllowances_Set"
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmTaxes) 'To Clearing the Text Boxes.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    CmdNew.Enabled = True: CmdNew.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload frmTaxes
End Sub

Private Sub CmdNew_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmTaxes) 'To Clearing the Text Boxes.
    CmdSubmit.Enabled = True: CmdCancel.Enabled = True
    CmdNew.Enabled = False: CmbDesig.SetFocus: Opt_Flag = "Add"
End Sub

Private Sub CmdSubmit_Click()
    Call Populate_SaveRecord 'For Saving Record in the Data Table.
    Call Ctrl_PayRoll.Populate_Text_Clear(frmTaxes) 'To Clearing the Text Boxes.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    CmdNew.Enabled = True: CmdNew.SetFocus
End Sub

Private Sub Form_Load()
    frmTaxes.Move FrmMain.Width / 3, FrmMain.Height / 8
    Call Ctrl_PayRoll.Populate_Text_Clear(frmTaxes) 'To Clea the Text Boxes.
    Opt_Flag = ""
    Rst_Desig.Open "SELECT * FROM tblDesignation", DB_Conect, adOpenStatic, adLockOptimistic 'For Designations.
    Rst_Taxes.Open "SELECT * FROM tblTaxes_Funds_Set", DB_Conect, adOpenStatic, adLockOptimistic 'For Taxes and Funds.
    Rst_Loan.Open "SELECT * FROM tblLoan_Set", DB_Conect, adOpenStatic, adLockOptimistic 'For Loan.
    Rst_Allow.Open "SELECT * FROM tblAllowances_Set", DB_Conect, adOpenStatic, adLockOptimistic 'For Allowances Benifits.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rst_Desig.Close 'To Close Database Table.
    Rst_Taxes.Close 'To Close Database Table.
    Rst_Loan.Close 'To Close Database Table.
    Rst_Allow.Close 'To Close Database Table.
End Sub

Private Sub txtBFund_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Only Numeric Values.
    If KeyAscii = 13 Then
        If txtBFund.Text <> "" Then
            SendKeys "{Home}+{End}": txtInTax.SetFocus
            
        ElseIf txtBFund.Text = "" Then
            MsgBox "Please verify! the Bonavolent Fund." & vbCrLf & _
                   "Are you Leave it Empty", vbCritical, "Error! Bonavolent Fund"
                   txtBFund.SetFocus
        End If
    End If
End Sub

Private Sub txtHRent_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Only Numeric Values.
    If KeyAscii = 13 Then
        If txtHRent.Text <> "" Then
            SendKeys "{Home}+{End}": txtMedical.SetFocus
            
        ElseIf txtHRent.Text = "" Then
            MsgBox "Please verify! The House Rent Value." & vbCrLf & _
                   "Are you Leave it Empty", vbCritical, "Error! House Rent"
                   txtHRent.SetFocus
        End If
    End If
End Sub

Private Sub txtInTax_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Only Numeric Values.
    If KeyAscii = 13 Then
        If txtInTax.Text <> "" Then
            SendKeys "{Home}+{End}": txtWHTax.SetFocus
            
        ElseIf txtInTax.Text = "" Then
            MsgBox "Please verify! The Income Taxe Value." & vbCrLf & _
                   "Are you Leave it Empty", vbCritical, "Error! Income Tax"
                   txtInTax.SetFocus
        End If
    End If
End Sub

Private Sub txtLoanRange_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Only Numeric Values.
    If KeyAscii = 13 Then
        If txtLoanRange.Text <> "" Then
            SendKeys "{Home}+{End}": CmdSubmit.SetFocus
            
        ElseIf txtLoanRange.Text = "" Then
            MsgBox "Please verify! The Loan Limit Value." & vbCrLf & _
                   "Are you Leave it Empty", vbCritical, "Error! Loan Limit"
                   txtLoanRange.SetFocus
        End If
    End If
End Sub

Private Sub txtMedical_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Only Numeric Values.
    If KeyAscii = 13 Then
        If txtMedical.Text <> "" Then
            SendKeys "{Home}+{End}": txtLoanRange.SetFocus
            
        ElseIf txtMedical.Text = "" Then
            MsgBox "Please verify! The Medical Allowance Value." & vbCrLf & _
                   "Are you Leave it Empty", vbCritical, "Error! Medical Allowance"
                   txtMedical.SetFocus
        End If
    End If
End Sub

Private Sub txtWHTax_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Only Numeric Values.
    If KeyAscii = 13 Then
        If txtWHTax.Text <> "" Then
            SendKeys "{Home}+{End}": txtHRent.SetFocus
            
        ElseIf txtWHTax.Text = "" Then
            MsgBox "Please verify! The Whealth Taxe Value." & vbCrLf & _
                   "Are you Leave it Empty", vbCritical, "Error! Whealth Tax"
                   txtWHTax.SetFocus
        End If
    End If
End Sub
