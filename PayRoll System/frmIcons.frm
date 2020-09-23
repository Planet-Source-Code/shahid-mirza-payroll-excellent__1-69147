VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmIcons 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3150
   ControlBox      =   0   'False
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
   ScaleHeight     =   9195
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdRpt 
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   6480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Click For Reports"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   15591915
      FCOL            =   12583104
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":0000
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":19AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":268A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":401C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":59AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":7340
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":8CD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":99AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":A686
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":B360
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":C03C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":CD18
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":D5F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":E2D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":EFAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":FC88
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1056C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":11248
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":11B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":12800
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":14194
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":15B28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   2280
      Top             =   7680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   36
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":16404
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":170DE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":17F30
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":18D82
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1965C
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":19F36
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1A810
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1B1DA
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1BAB4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1BDCE
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1C6A8
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1CF82
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1D85C
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1DB76
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1E450
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1ED2A
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1F604
            Key             =   "IMG17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":1FEDE
            Key             =   "IMG18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":207B8
            Key             =   "IMG19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":21092
            Key             =   "IMG20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":2196C
            Key             =   "IMG21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":22246
            Key             =   "IMG22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":22B20
            Key             =   "IMG23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":233FA
            Key             =   "IMG24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":23CD4
            Key             =   "IMG25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":245AE
            Key             =   "IMG26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":24E88
            Key             =   "IMG27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":25762
            Key             =   "IMG28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":2603C
            Key             =   "IMG29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":26916
            Key             =   "IMG30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":271CC
            Key             =   "IMG31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":27AA6
            Key             =   "IMG32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":27EF8
            Key             =   "IMG33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":2834A
            Key             =   "IMG34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":2AAFC
            Key             =   "IMG35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIcons.frx":2BD7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CmdShutDown 
      Height          =   495
      Left            =   240
      TabIndex        =   28
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15591915
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E530
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "36"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdPayingLoan 
      Height          =   615
      Index           =   35
      Left            =   1680
      TabIndex        =   25
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15591915
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E54C
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "35"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdLoanApprove 
      Height          =   615
      Left            =   1680
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E568
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "12"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdRLoan 
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E584
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "05"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdFundLoans 
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E5A0
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "06"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdChangePWD 
      Height          =   615
      Left            =   1680
      TabIndex        =   10
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E5BC
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "12"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdNewUser 
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E5D8
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "19"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdWorkDays 
      Height          =   615
      Left            =   1680
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E5F4
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "11"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdDesignation 
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E610
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "18"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdCreatePayRoll 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E62C
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "19"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdWorkRate 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E648
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "16"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdDepartment 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E664
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "21"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton CmdNewEmp 
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   14215660
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E680
      ALIGN           =   1
      IMGLST          =   "ImageList1"
      IMGICON         =   "1"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin Crystal.CrystalReport Rpt_EmpList 
      Left            =   600
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Rpt_PaySlip 
      Left            =   1080
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Rpt_LoanApproved 
      Left            =   1560
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Rpt_LoanReturn 
      Left            =   2040
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin LVbuttons.LaVolpeButton CmdAbout 
      Height          =   495
      Left            =   1680
      TabIndex        =   31
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      BCOL            =   15591915
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmIcons.frx":2E69C
      ALIGN           =   1
      IMGLST          =   "SmallImages"
      IMGICON         =   "32"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   4
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      Height          =   255
      Left            =   1680
      TabIndex        =   32
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shut Down"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Create Payroll"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1680
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblUserName 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF00FF&
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Paying Loan"
      Height          =   270
      Left            =   1680
      TabIndex        =   26
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3000
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Deduct"
      Height          =   270
      Left            =   240
      TabIndex        =   24
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Approve"
      Height          =   255
      Left            =   1680
      TabIndex        =   23
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loan && Funds"
      Height          =   270
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   1200
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3000
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   1560
      X2              =   1560
      Y1              =   7320
      Y2              =   9120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   3000
      X2              =   120
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change PWd"
      Height          =   270
      Left            =   1680
      TabIndex        =   21
      Top             =   7920
      Width           =   1200
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New User"
      Height          =   270
      Left            =   240
      TabIndex        =   20
      Top             =   7920
      Width           =   1200
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "MAIN MENU"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "User Settings"
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
      Left            =   240
      TabIndex        =   18
      Top             =   6960
      Width           =   2655
   End
   Begin VB.Line Line8 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3000
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line7 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   1560
      X2              =   1560
      Y1              =   600
      Y2              =   1680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   1560
      X2              =   1560
      Y1              =   6360
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3000
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   9120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   3000
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   3000
      X2              =   3000
      Y1              =   600
      Y2              =   9120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Rate"
      Height          =   270
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Settings Option"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
      Height          =   270
      Left            =   1680
      TabIndex        =   16
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   270
      Left            =   240
      TabIndex        =   14
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      Height          =   270
      Left            =   1680
      TabIndex        =   13
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Employee"
      Height          =   270
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1245
   End
   Begin VB.Menu MainRpt 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu Rpt_ToolBar 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Reports|Font:Trebuchet MS|BOLD|Fsize:12|Fcolor:16777215|Bcolor:255|Gradient}"
      End
      Begin VB.Menu Seprator0 
         Caption         =   "-Reports"
      End
      Begin VB.Menu MainRpt_EmpList 
         Caption         =   "{IMG:1}Employee List"
      End
      Begin VB.Menu MainRpt_LoanApproved 
         Caption         =   "{IMG:1}Loan Approved"
      End
      Begin VB.Menu MainRpt_LoanReturn 
         Caption         =   "{IMG:1}Loan Return"
      End
      Begin VB.Menu MainRpt_PaySlip 
         Caption         =   "{IMG:1}Employee Pay Slip"
      End
      Begin VB.Menu Sperator1 
         Caption         =   "-End"
      End
   End
End
Attribute VB_Name = "frmIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OptExit

Private Sub CmdAbout_Click()
    Load frmAbout: frmAbout.Show
End Sub

Private Sub CmdChangePWD_Click()
    Load frmChangePwd: frmChangePwd.Show
End Sub

Private Sub CmdLoanApprove_Click()
    Load frmLoanApprove: frmLoanApprove.Show
End Sub

Private Sub CmdNewUser_Click()
    Load frmUser_Create: frmUser_Create.Show
End Sub

Private Sub CmdPayingLoan_Click(Index As Integer)
    Load frmPaidLoan: frmPaidLoan.Show
End Sub

Private Sub CmdRLoan_Click()
    Load frmRLoan: frmRLoan.Show
End Sub

Private Sub CmdCreatePayRoll_Click()
    Load frmPaySlip: frmPaySlip.Show
End Sub

Private Sub CmdDepartment_Click()
    Load frmDepartment: frmDepartment.Show
End Sub

Private Sub CmdDesignation_Click()
    Load frmDesignation: frmDesignation.Show
End Sub

Private Sub CmdFundLoans_Click()
    Load frmTaxes: frmTaxes.Show
End Sub

Private Sub CmdNewEmp_Click()
    Load frmNewEmployee: frmNewEmployee.Show
End Sub

Private Sub CmdRpt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        PopupMenu MainRpt
    End If
End Sub

Private Sub CmdShutDown_Click()
    OptExit = MsgBox("Are you sure you want to exit " & vbCrLf & "from the Payroll system." & vbCrLf & _
                    "Please vertiy.", vbCritical + vbYesNo, "System Shut Down. . . . . ")
        If OptExit = vbYes Then
            End
        End If
End Sub

Private Sub CmdWorkDays_Click()
    Load frmWrkDays: frmWrkDays.Show
End Sub

Private Sub CmdWorkRate_Click()
    Load frmRate: frmRate.Show
End Sub

Private Sub Form_Load()
    frmIcons.Move 0, FrmMain.Height / 15
    Call SetMenus(hwnd, SmallImages)
End Sub


Private Sub MainRpt_EmpList_Click()
    Rpt_EmpList.ReportFileName = App.Path & "\Reports\Rpt_EmpList.rpt"
    Rpt_EmpList.WindowState = crptMaximized
    Rpt_EmpList.Action = 1
End Sub

Private Sub MainRpt_LoanApproved_Click()
    Rpt_LoanApproved.ReportFileName = App.Path & "\Reports\Rpt_Aproved_Loan.rpt"
    Rpt_LoanApproved.WindowState = crptMaximized
    Rpt_LoanApproved.Action = 1
End Sub

Private Sub MainRpt_LoanReturn_Click()
    Rpt_LoanReturn.ReportFileName = App.Path & "\Reports\Rpt_Loan_Return.rpt"
    Rpt_LoanReturn.WindowState = crptMaximized
    Rpt_LoanReturn.Action = 1
End Sub

Private Sub MainRpt_PaySlip_Click()
    Rpt_PaySlip.ReportFileName = App.Path & "\Reports\Rpt_PaySlip.rpt"
    Rpt_PaySlip.WindowState = crptMaximized
    Rpt_PaySlip.Action = 1
End Sub
