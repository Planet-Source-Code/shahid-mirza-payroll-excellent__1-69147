VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmNewEmployee 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " New Employee Information"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
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
   ScaleHeight     =   7290
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   Begin MSMask.MaskEdBox txtEnd_Date 
      Height          =   390
      Left            =   5280
      TabIndex        =   43
      Top             =   3120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtStart_Date 
      Height          =   390
      Left            =   5280
      TabIndex        =   42
      Top             =   2640
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   688
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtEmp_DOB 
      Height          =   390
      Left            =   1800
      TabIndex        =   41
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin LVbuttons.LaVolpeButton CmdEmp_Picture 
      Height          =   375
      Left            =   7600
      TabIndex        =   40
      Top             =   3450
      Width           =   1720
      _ExtentX        =   3043
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Browse 4 Picture"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14215660
      FCOL            =   16576
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmNewEmployee.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Deduction"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1155
      Left            =   5640
      TabIndex        =   39
      Top             =   4380
      Width           =   5895
      Begin LVbuttons.LaVolpeButton Cmd_Not_Deduct 
         Height          =   330
         Left            =   3480
         TabIndex        =   53
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Un Check All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   15591915
         FCOL            =   12582912
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmNewEmployee.frx":001C
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton Cmd_All_Deduct 
         Height          =   330
         Left            =   1800
         TabIndex        =   52
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Check All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   15591915
         FCOL            =   12582912
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmNewEmployee.frx":0038
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton Cmd_All_Exp_Loan 
         Height          =   330
         Left            =   120
         TabIndex        =   51
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "All Expect Loan"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   15591915
         FCOL            =   12582912
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmNewEmployee.frx":0054
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.CheckBox Chk_Loan 
         BackColor       =   &H8000000E&
         Caption         =   "Loan"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4080
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox Chk_In_WH_Tax 
         BackColor       =   &H8000000E&
         Caption         =   "Income && W/H Tax"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   49
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox Chk_BFund 
         BackColor       =   &H8000000E&
         Caption         =   "B-Funds"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Facilitate (Y/N)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   2145
      Left            =   9600
      TabIndex        =   36
      Top             =   1680
      Width           =   1935
      Begin LVbuttons.LaVolpeButton Cmd_Both_Benifits 
         Height          =   330
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Check All"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   15591915
         FCOL            =   12582912
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "frmNewEmployee.frx":0070
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   0
         IconSize        =   2
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin VB.CheckBox Chk_Medical 
         BackColor       =   &H8000000E&
         Caption         =   "Medical Only"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox Chk_HRent 
         BackColor       =   &H8000000E&
         Caption         =   "H/Rent Only"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Opt_Not_Benifits 
         BackColor       =   &H8000000E&
         Caption         =   "Not Allow"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.TextBox txtRe_Ovr_Day 
      Alignment       =   1  'Right Justify
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
      Left            =   4560
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "txtRe_Ovr_Day"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtRe_Ovr_Holi 
      Alignment       =   1  'Right Justify
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
      Left            =   4560
      TabIndex        =   28
      TabStop         =   0   'False
      Text            =   "txtRe_Ovr_Holi"
      Top             =   4920
      Width           =   855
   End
   Begin VB.ComboBox CmbDepart 
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
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox CmbDesig 
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
      Left            =   5280
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2160
      Width           =   2055
   End
   Begin StyleButtonX.StyleButton CmdSearch 
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
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
      PictureUp       =   "frmNewEmployee.frx":008C
      PictureDown     =   "frmNewEmployee.frx":0736
      PictureHover    =   "frmNewEmployee.frx":0DE0
      PictureFocus    =   "frmNewEmployee.frx":1488
      PictureDisabled =   "frmNewEmployee.frx":1B32
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBorderLevel1=   0   'False
      ShowBorderLevel2=   0   'False
   End
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
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
      PictureUp       =   "frmNewEmployee.frx":21DC
      PictureDown     =   "frmNewEmployee.frx":28AB
      PictureHover    =   "frmNewEmployee.frx":2F7A
      PictureFocus    =   "frmNewEmployee.frx":3631
      PictureDisabled =   "frmNewEmployee.frx":3D00
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
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
      Enabled         =   0   'False
      BackColorUp     =   -2147483634
      BackColorDown   =   -2147483634
      BackColorHover  =   -2147483634
      BackColorFocus  =   -2147483634
      BackColorDisabled=   -2147483634
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmNewEmployee.frx":43CF
      PictureDown     =   "frmNewEmployee.frx":4A8F
      PictureHover    =   "frmNewEmployee.frx":514F
      PictureFocus    =   "frmNewEmployee.frx":57FE
      PictureDisabled =   "frmNewEmployee.frx":5EBE
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
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
      Enabled         =   0   'False
      BackColorUp     =   -2147483634
      BackColorDown   =   -2147483634
      BackColorHover  =   -2147483634
      BackColorFocus  =   -2147483634
      BackColorDisabled=   -2147483634
      DotsInCornerColor=   16777215
      ForeColorDisabled=   12632256
      PictureUp       =   "frmNewEmployee.frx":657E
      PictureDown     =   "frmNewEmployee.frx":6C53
      PictureHover    =   "frmNewEmployee.frx":7328
      PictureFocus    =   "frmNewEmployee.frx":7A07
      PictureDisabled =   "frmNewEmployee.frx":80DC
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   6720
      Width           =   1335
      _ExtentX        =   2355
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
      PictureUp       =   "frmNewEmployee.frx":87B1
      PictureDown     =   "frmNewEmployee.frx":8E70
      PictureHover    =   "frmNewEmployee.frx":952F
      PictureFocus    =   "frmNewEmployee.frx":9BC2
      PictureDisabled =   "frmNewEmployee.frx":A281
      BeginProperty FontUp {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFocus {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowBorderLevel1=   0   'False
      ShowBorderLevel2=   0   'False
   End
   Begin VB.TextBox txtRe_Nor_Holi 
      Alignment       =   1  'Right Justify
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
      Left            =   1795
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "txtRe_Nor_Holi"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtRe_Nor_Day 
      Alignment       =   1  'Right Justify
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
      Left            =   1795
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "txtRe_Nor_Day"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtEmp_ID 
      Alignment       =   2  'Center
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
      Left            =   5280
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "txtEmp_ID - (Auto)"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtEmp_Address 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "frmNewEmployee.frx":A940
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtLastName 
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
      Left            =   1800
      TabIndex        =   18
      Text            =   "txtLastName"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox txtFirstName 
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
      Left            =   1800
      TabIndex        =   17
      Text            =   "txtFirstName"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label lblDepart_ID 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblDepart_ID"
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
      Left            =   9360
      TabIndex        =   55
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblPicPath 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblPicPath"
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
      Left            =   7680
      TabIndex        =   54
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblDesig_ID 
      Alignment       =   2  'Center
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
      Height          =   390
      Left            =   7320
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      X1              =   0
      X2              =   11760
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   11640
      X2              =   0
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funds/Taxes && Loans Info - (Deductions)"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5640
      TabIndex        =   38
      Top             =   4080
      Width           =   5895
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   11640
      X2              =   11640
      Y1              =   1320
      Y2              =   5640
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Benifits Info"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9600
      TabIndex        =   35
      Top             =   1365
      Width           =   1935
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   5520
      X2              =   5520
      Y1              =   3960
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   2760
      X2              =   2760
      Y1              =   3960
      Y2              =   5640
   End
   Begin VB.Image Img_Emp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Employee - Snap"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7560
      TabIndex        =   34
      Top             =   1365
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   2880
      TabIndex        =   33
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holidays : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   2880
      TabIndex        =   32
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Employement Information"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3960
      TabIndex        =   31
      Top             =   1365
      Width           =   3375
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hourly Rate 4 Overtime"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   30
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Department : "
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
      Left            =   3960
      TabIndex        =   26
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   9480
      X2              =   9480
      Y1              =   1320
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   0
      X2              =   11640
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   3840
      X2              =   3840
      Y1              =   1320
      Y2              =   3960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   7440
      X2              =   7440
      Y1              =   1320
      Y2              =   3960
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hourly Rate Normal Days"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Employee Personal Information"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   23
      Top             =   1365
      Width           =   3615
   End
   Begin VB.Label lblHelp_Bar 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewEmployee.frx":A951
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   11535
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User's Instruction"
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
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   2890
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Special Holidays : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "End Contract : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   3960
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Start Contract : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   3960
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Designation : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   3960
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   3720
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
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
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth (DOB) : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Left            =   0
      Picture         =   "frmNewEmployee.frx":A9DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11865
   End
End
Attribute VB_Name = "frmNewEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Emp_Info As New ADODB.Recordset: Dim Rst_Emp_Dtl As New ADODB.Recordset
Dim Rst_Depart As New ADODB.Recordset: Dim Rst_Desig As New ADODB.Recordset
Dim Rst_Rates As New ADODB.Recordset

'===============================================================================================
Private Sub Pop_SaveRecord()
    With Rst_Emp_Info
        .AddNew
            .Fields(0).Value = txtEmp_ID.Text: .Fields(1).Value = lblDepart_ID
            .Fields(2).Value = lblDesig_ID: .Fields(3).Value = Format(txtStart_Date.Text, "DD/MMM/YY")
            .Fields(4).Value = txtEnd_Date.Text: .Fields(5).Value = txtFirstName.Text
            .Fields(6).Value = txtLastName.Text: .Fields(7).Value = txtEmp_DOB.Text
            .Fields(8).Value = txtEmp_Address.Text: '.Fields(9).Value = Img_Emp.Picture
        .Update
    End With
    With Rst_Emp_Dtl
        .AddNew
            .Fields(0).Value = txtEmp_ID.Text
            .Fields(1).Value = Chk_HRent.Value: .Fields(2).Value = Chk_Medical.Value
            .Fields(3).Value = Chk_BFund.Value: .Fields(4).Value = Chk_In_WH_Tax.Value
            .Fields(5).Value = Chk_Loan.Value
        .Update
    End With
    MsgBox "Record has been Saved. Successfully! ", vbInformation, "Saving Record"
End Sub

Private Sub Pop_Auto_ID()
    With Rst_Emp_Info
        txtEmp_ID.Text = "EMP-" & .RecordCount + 101
    End With
End Sub
'===============================================================================================

Private Sub Chk_HRent_Click()
    If ((Chk_HRent.Value = 0) And (Chk_Medical.Value = 1)) Then
        Opt_Not_Benifits.Value = False: Cmd_Both_Benifits.Enabled = True
    ElseIf ((Chk_HRent.Value = 1) And (Chk_Medical.Value = 1)) Then
        Opt_Not_Benifits.Value = False
        Cmd_Both_Benifits.Enabled = False
    End If
End Sub

Private Sub Chk_Medical_Click()
    If ((Chk_HRent.Value = 1) And (Chk_Medical.Value = 0)) Then
        Opt_Not_Benifits.Value = False: Cmd_Both_Benifits.Enabled = True
    ElseIf ((Chk_HRent.Value = 1) And (Chk_Medical.Value = 1)) Then
        Opt_Not_Benifits.Value = False
        Cmd_Both_Benifits.Enabled = False
    End If
End Sub

Private Sub CmbDepart_Click()
    Rst_Depart.Close: Rst_Depart.Open "SELECT * FROM tblDepartment WHERE Depart_Name='" & CmbDepart.Text & "'"
    If Rst_Depart.RecordCount > 0 Then lblDepart_ID = Rst_Depart.Fields(0).Value
    If Rst_Depart.RecordCount <= 0 Then lblDepart_ID = ""
End Sub

Private Sub CmbDepart_GotFocus()
    Rst_Depart.Close: Rst_Depart.Open "SELECT * FROM tblDepartment"
    Call Ctrl_PayRoll.Populate_Init_Cmb(Rst_Depart, 1, CmbDepart) 'To Initialize the Combo Boxes Enteries.
    lblHelp_Bar = "Choose the Employee's Depatment Name and Press enter key, Proceed to Employee's Designation you can't enter the Department Other than List Provided."
    lblHelp_Bar.ForeColor = vbBlack
End Sub

Private Sub CmbDepart_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbDepart.Text <> "Choose" Then
            CmbDesig.SetFocus
        ElseIf CmbDepart.Text = "Choose" Then
            MsgBox "Please choose the correct selection." & vbCrLf & _
                   "It is not valid selection.", vbCritical, "Error! Department Selection"
                   CmbDepart.SetFocus
        End If
    End If
End Sub

Private Sub CmbDesig_Click()
    Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation WHERE Desig_Name='" & CmbDesig.Text & "'"
    If Rst_Desig.RecordCount > 0 Then lblDesig_ID = Rst_Desig.Fields(0).Value
    If Rst_Desig.RecordCount <= 0 Then lblDesig_ID = ""
    
    Rst_Rates.Close: Rst_Rates.Open "SELECT * FROM tblHourly_Rates_Set WHERE Desig_ID='" & lblDesig_ID & "'"
    If Rst_Rates.RecordCount > 0 Then 'If Found Fill TextBoxes.
       With Rst_Rates
           txtRe_Nor_Day.Text = .Fields(1).Value: txtRe_Nor_Holi.Text = .Fields(2).Value
           txtRe_Ovr_Day.Text = .Fields(3).Value: txtRe_Ovr_Holi.Text = .Fields(4).Value
       End With
    ElseIf Rst_Rates.RecordCount <= 0 Then 'Then Empty the TextBoxes.
       txtRe_Nor_Day.Text = "0": txtRe_Nor_Holi.Text = "0"
       txtRe_Ovr_Day.Text = "0": txtRe_Ovr_Holi.Text = "0"
    End If
End Sub

Private Sub CmbDesig_GotFocus()
    Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation"
    Call Ctrl_PayRoll.Populate_Init_Cmb(Rst_Desig, 1, CmbDesig) 'To Initialize the Combo Boxes Enteries.
    
    lblHelp_Bar = "Choose the Employee's Dsignation and Press enter key, Proceed to Starting Date of Employee's Contract you can't enter the Designation Other than List Provided."
    lblHelp_Bar.ForeColor = vbRed
End Sub

Private Sub CmbDesig_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmbDesig.Text <> "Choose" Then
            txtStart_Date.SetFocus
        ElseIf CmbDesig.Text = "Choose" Then
            MsgBox "Please choose the correct selection." & vbCrLf & _
                   "It is not valid selection.", vbCritical, "Error! Department Selection"
                   CmbDesig.SetFocus
        End If
    End If
    Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation"
    Rst_Rates.Close: Rst_Rates.Open "SELECT * FROM tblHourly_Rates_Set"
End Sub

Private Sub Cmd_All_Deduct_Click()
    Chk_BFund.Value = 1: Chk_In_WH_Tax.Value = 1: Chk_Loan.Value = 1
    Cmd_All_Deduct.Enabled = False: Cmd_All_Exp_Loan.Enabled = True: Cmd_Not_Deduct.Enabled = True
End Sub

Private Sub Cmd_All_Exp_Loan_Click()
    Chk_BFund.Value = 1: Chk_In_WH_Tax.Value = 1: Chk_Loan.Value = 0
    Cmd_All_Deduct.Enabled = True: Cmd_All_Exp_Loan.Enabled = False: Cmd_Not_Deduct.Enabled = True
End Sub


Private Sub Cmd_Not_Deduct_Click()
    Chk_BFund.Value = 0: Chk_In_WH_Tax.Value = 0: Chk_Loan.Value = 0
    Cmd_All_Deduct.Enabled = True: Cmd_All_Exp_Loan.Enabled = True: Cmd_Not_Deduct.Enabled = False
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmNewEmployee) 'To Clearing the Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmNewEmployee, False) 'Disable the From from Entry.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    CmdNew.Enabled = True: CmdNew.SetFocus: txtEmp_ID.Text = "(Auto Number)"
End Sub

Private Sub CmdEmp_Picture_Click()
    Call Ctrl_PayRoll.msg_Consutruct 'Waiting for Employee Picture Code.
End Sub

Private Sub CmdExit_Click()
    Unload frmNewEmployee
End Sub

Private Sub CmdNew_Click()
    Call Ctrl_PayRoll.Populate_Entery(frmNewEmployee, True) 'Disable the From from Entry.
    Call Ctrl_PayRoll.Populate_Text_Clear(frmNewEmployee) 'To Clearing the TextBoxes.
    CmdNew.Enabled = False: CmdSubmit.Enabled = True
    CmdCancel.Enabled = True: CmdSearch.Enabled = True
    txtFirstName.SetFocus: CtrlName = txtFirstName.Text
End Sub

Private Sub CmdNew_GotFocus()
    CtrlName = "New Button"
    lblHelp_Bar = "For new employee click New Button. To search for the employee you want to edit click Search button."
    lblHelp_Bar.ForeColor = vbRed
End Sub

Private Sub CmdSearch_Click()
    Call Ctrl_PayRoll.msg_Consutruct
End Sub

Private Sub CmdSubmit_Click()
    Call Pop_SaveRecord 'For Saving Record in the Data Table.
    Call Ctrl_PayRoll.Populate_Text_Clear(frmNewEmployee) 'To Clearing the Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmNewEmployee, False) 'Disable the From from Entry.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False
    Chk_BFund.Value = 0: Chk_In_WH_Tax.Value = 0: Chk_Loan.Value = 0
    Chk_HRent.Value = 0: Chk_Medical.Value = 0
    
    CmdNew.Enabled = True: CmdNew.SetFocus
End Sub

Private Sub Form_Load()
    frmNewEmployee.Move FrmMain.Width / 4.8, FrmMain.Height / 12
    Call Ctrl_PayRoll.Populate_Text_Clear(frmNewEmployee) 'To Clearing the TextBoxes.
    txtEmp_ID.Text = "(Auto Number)": Cmd_Not_Deduct.Enabled = False
    Call Ctrl_PayRoll.Populate_Entery(frmNewEmployee, False) 'Disable the From from Entry.
    
    Rst_Emp_Info.Open "SELECT * FROM tblEmployee_Info", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Emp_Dtl.Open "SELECT * FROM tblEmployee_Detail", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Depart.Open "SELECT * FROM tblDepartment", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Desig.Open "SELECT * FROM tblDesignation", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Rates.Open "SELECT * FROM tblHourly_Rates_Set", DB_Conect, adOpenStatic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rst_Emp_Info.Close 'To Close Database Table.
    Rst_Emp_Dtl.Close 'To Close Database Table.
    Rst_Depart.Close 'To Close Database Table.
    Rst_Desig.Close 'To Close Database Table.
    Rst_Rates.Close 'To Close Database Table.
End Sub

Private Sub Cmd_Both_Benifits_Click()
    If Cmd_Both_Benifits.Enabled = True Then
        Chk_HRent.Value = 1: Chk_Medical.Value = 1
        Cmd_Both_Benifits.Enabled = False
        Opt_Not_Benifits.Value = False
    End If
End Sub

Private Sub Opt_Not_Benifits_Click()
    If Opt_Not_Benifits.Value = True Then
        Chk_HRent.Value = 0: Chk_Medical.Value = 0
    End If
End Sub

Private Sub txtEmp_Address_GotFocus()
    lblHelp_Bar = "Enter the Employee's Residential Address, Please Don't Use Shift Key and Press enter key, Proceed to Employee's Department and Entries must be in Alpha Numeric Character. Maximum Characters 100."
    lblHelp_Bar.ForeColor = vbBlue
End Sub

Private Sub txtEmp_Address_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_Alpha_Char(KeyAscii, frmNewEmployee, txtEmp_Address) 'Call for First Character in Capital.
    If KeyAscii = 13 Then
        If txtEmp_Address.Text = "" Then
            MsgBox "Please enter the Employee's Residential Address." & vbCrLf & "It is not valid entery.", vbCritical, "Error! Residential Address"
            SendKeys "{Home}+{End}": txtEmp_Address.SetFocus
        ElseIf txtEmp_Address.Text <> "" Then
            CmbDepart.SetFocus: KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEmp_DOB_GotFocus()
    lblHelp_Bar = "Enter the Employee's Date Of Birth and Press enter key, Proceed to Employee's Address and Entries must be in Numeric Character. Seperator '/' can be used from Numeric PAD."
    lblHelp_Bar.ForeColor = vbRed
End Sub

Private Sub txtEmp_DOB_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Numeric Values Only.
    If KeyAscii = 13 Then
        If txtEmp_DOB.Text = "__/__/____" Then
            MsgBox "Please enter the Employee's Date Of Birth." & vbCrLf & "It is not valid entery.", vbCritical, "Error! Date OF Birth"
            SendKeys "{Home}+{End}": txtEmp_DOB.SetFocus
        ElseIf txtEmp_DOB.Text <> "" Then
            txtEmp_Address.SetFocus
        End If
    End If
End Sub

Private Sub txtEnd_Date_GotFocus()
    lblHelp_Bar = "Enter the End Date of Employee's Job Contract and Press enter key, Proceed to the Hourly Rate about Normal Days, Entries must be in Numeric Character. Seperator '/' is Automatic Please don't Use it."
    lblHelp_Bar.ForeColor = vbBlue
End Sub

Private Sub txtEnd_Date_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Numeric Values Only.
    If KeyAscii = 13 Then
        If txtEnd_Date.Text = "__/__/____" Then
            MsgBox "Please enter the Employee's End Contract Date." & vbCrLf & "It is not valid entery.", vbCritical, "Error! Contract End Date"
            SendKeys "{Home}+{End}": txtEnd_Date.SetFocus
        ElseIf txtEnd_Date.Text <> "" Then
            CmdEmp_Picture.SetFocus
        End If
    End If
End Sub

Private Sub txtFirstName_GotFocus()
    Call Pop_Auto_ID 'To Enter Emploee ID Auto.
    lblHelp_Bar = "Enter Employee's First Name and Press Enter Key Move Forward to the Employee's Last Name Please Don't Use Shift Key, Entries must be in Alphabatic Character. Length of Characters 45."
    lblHelp_Bar.ForeColor = vbBlue
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_CharOnly(KeyAscii) 'Allow Only Alphabatic Characters.
    Call Ctrl_PayRoll.Populate_Alpha_Char(KeyAscii, frmNewEmployee, txtFirstName) 'Call for First Character in Capital.
    If KeyAscii = 13 Then
        If txtFirstName.Text = "" Then
            MsgBox "Please enter the Employee's First Name." & vbCrLf & "It is not valid entery.", vbCritical, "Error! First Name"
            SendKeys "{Home}+{End}": txtFirstName.SetFocus
        ElseIf txtFirstName.Text <> "" Then
            txtLastName.SetFocus
        End If
    End If
End Sub

Private Sub txtLastName_GotFocus()
    lblHelp_Bar = "Enter the Employee's Last name [Please Don't Use Shift Key] and Press enter key, Proceed to Employee's Date Of Birth and Entries must be in Alphabatic Character. Length of Characters 45."
    lblHelp_Bar.ForeColor = vbBlack
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_CharOnly(KeyAscii) 'Allow Only Alphabatic Characters.
    Call Ctrl_PayRoll.Populate_Alpha_Char(KeyAscii, frmNewEmployee, txtLastName) 'Call for First Character in Capital.
    If KeyAscii = 13 Then
        If txtLastName.Text = "" Then
            MsgBox "Please enter the Employee's Last Name." & vbCrLf & "It is not valid entery.", vbCritical, "Error! Last Name"
            SendKeys "{Home}+{End}": txtLastName.SetFocus
        ElseIf txtLastName.Text <> "" Then
            txtEmp_DOB.SetFocus
        End If
    End If
End Sub

Private Sub txtRe_Nor_Day_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtRe_Nor_Holi_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtRe_Ovr_Day_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtRe_Ovr_Holi_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtStart_Date_GotFocus()
    lblHelp_Bar = "Enter the Start Date of Employee's Job Contract and Press enter key, Proceed to the Ending Date of Employee's Job Contract, Entries must be in Numeric Character. Seperator '/' is Automatic Please don't Use it."
    lblHelp_Bar.ForeColor = vbRed
End Sub

Private Sub txtStart_Date_KeyPress(KeyAscii As Integer)
    Call Ctrl_PayRoll.Populate_NumOnly(KeyAscii) 'Call for Numeric Values Only.
    If KeyAscii = 13 Then
        If txtStart_Date.Text = "__/__/____" Then
            MsgBox "Please enter the Employee's Start Contract Date." & vbCrLf & "It is not valid entery.", vbCritical, "Error! Contract Start Date"
            SendKeys "{Home}+{End}": txtStart_Date.Text = Format(txtStart_Date.Text, "DD/MMM/YY"): txtStart_Date.SetFocus
        ElseIf txtStart_Date.Text <> "" Then
            txtEnd_Date.SetFocus
        End If
    End If
End Sub
