VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesignation 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Designation Information"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   5190
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   450
      Left            =   3480
      TabIndex        =   3
      Top             =   4680
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
      PictureUp       =   "frmDesignation.frx":0000
      PictureDown     =   "frmDesignation.frx":0623
      PictureHover    =   "frmDesignation.frx":0C29
      PictureFocus    =   "frmDesignation.frx":122F
      PictureDisabled =   "frmDesignation.frx":1852
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
      Left            =   2355
      TabIndex        =   2
      Top             =   4680
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
      PictureUp       =   "frmDesignation.frx":1E75
      PictureDown     =   "frmDesignation.frx":24B7
      PictureHover    =   "frmDesignation.frx":2AF9
      PictureFocus    =   "frmDesignation.frx":312C
      PictureDisabled =   "frmDesignation.frx":376E
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
      Left            =   1245
      TabIndex        =   1
      Top             =   4680
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
      PictureUp       =   "frmDesignation.frx":3B5A
      PictureDown     =   "frmDesignation.frx":417C
      PictureHover    =   "frmDesignation.frx":479E
      PictureFocus    =   "frmDesignation.frx":4DE2
      PictureDisabled =   "frmDesignation.frx":5404
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
      Top             =   4680
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
      PictureUp       =   "frmDesignation.frx":5800
      PictureDown     =   "frmDesignation.frx":5E10
      PictureHover    =   "frmDesignation.frx":6420
      PictureFocus    =   "frmDesignation.frx":6A3E
      PictureDisabled =   "frmDesignation.frx":704E
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
   Begin MSComctlLib.ListView LstDesig 
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr.#"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   3881
      EndProperty
   End
   Begin VB.TextBox txtDesigName 
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
      Left            =   1680
      TabIndex        =   7
      Text            =   "txtDesigName"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtDesigID 
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
      Left            =   1680
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "txtDesigID"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   4560
      X2              =   4560
      Y1              =   480
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   4560
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Existing Saved List of Designation"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   10
      Top             =   1520
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Designation Settings"
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   9
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Designation ID : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Desig As New ADODB.Recordset
Dim LItem As ListItem: Dim Responce

Private Sub Pop_Record()
    With Rst_Desig
        LstDesig.ListItems.Clear ' To Clear the List.
        If .RecordCount > 0 Then .MoveFirst
        For IntI = 1 To .RecordCount
            Set LItem = LstDesig.ListItems.Add(IntI, , IntI)
                LItem.SubItems(1) = .Fields(0).Value
                LItem.SubItems(2) = .Fields(1).Value
                If .EOF = True Then Exit For
                If .EOF = False Then .MoveNext
        Next
    End With
End Sub

Private Sub Pop_SaveRecord()
    If txtDesigID.Text = "" Then MsgBox "Please Enter the Designation ID." & vbCrLf & _
        "Not a valid Designation ID", vbCritical, "Error! Designation ID": Call CmdCancel_Click: Exit Sub
    If txtDesigName.Text = "" Then MsgBox "Please Enter the Designation Name." & vbCrLf & _
        "Not a valid Designation Name", vbCritical, "Error! Designation Name": Call CmdCancel_Click: Exit Sub
        
    With Rst_Desig
        .AddNew
            .Fields(0).Value = txtDesigID.Text: .Fields(1).Value = txtDesigName.Text
        .Update
        Call Pop_Record 'To Display the Record in the List.
        MsgBox "Record has been Saved. Successfully! ", vbInformation, "Saving Record"
    End With
End Sub

Private Sub Pop_Auto_ID()
    With Rst_Desig
        txtDesigID.Text = "DES-" & .RecordCount + 101
    End With
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmDesignation) 'To Clearing the Text Boxes.
    CmdCancel.Enabled = False: CmdSubmit.Enabled = False
    CmdNew.Enabled = True: CmdNew.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload frmDesignation
End Sub

Private Sub CmdNew_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmDesignation) 'To Clearing the Text Boxes.
    CmdSubmit.Enabled = True: CmdCancel.Enabled = True
    CmdNew.Enabled = False: txtDesigName.SetFocus
End Sub

Private Sub CmdSubmit_Click()
    Call Pop_SaveRecord 'Call For Saving Record In Database.
    Responce = MsgBox("Do you want to enter more records", vbInformation + vbYesNo, "Record Entery")
    If Responce = vbYes Then Call Pop_Auto_ID: txtDesigName.Text = "": txtDesigName.SetFocus
    If Responce = vbNo Then Call CmdCancel_Click: CmdNew.SetFocus
End Sub

Private Sub Form_Load()
    frmDesignation.Move FrmMain.Width / 3, FrmMain.Height / 8
    Call Ctrl_PayRoll.Populate_Text_Clear(frmDesignation) 'To Clearing the Text Boxes.
    
    Rst_Desig.Open "SELECT * FROM tblDesignation", DB_Conect, adOpenStatic, adLockOptimistic
    Call Pop_Record 'To Initilize the Display List.
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rst_Desig.Close 'For Closing Database Table to Secure.
End Sub

Private Sub txtDesigName_GotFocus()
    Call Pop_Auto_ID 'Call for Departmental ID Number.
End Sub

Private Sub txtDesigName_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_CharOnly(KeyAscii) 'Call for Only Character Values.
    Call Ctrl_PayRoll.Populate_Alpha_Char(KeyAscii, frmDesignation, txtDesigName) 'Call First Capital Letter.
    If KeyAscii = 13 Then
        If txtDesigName.Text <> "" Then
            Call CmdSubmit_Click 'For Saving Process.
            
        ElseIf txtDesigName.Text = "" Then
            MsgBox "Please verify! the Designation Name." & vbCrLf & _
                   "It is not a valid Designation Name", vbCritical, "Error! Designation Name"
                   txtDesigName.SetFocus
        End If
    End If
End Sub


