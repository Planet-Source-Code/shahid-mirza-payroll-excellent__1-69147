VERSION 5.00
Object = "{C9680CB9-8919-4ED0-A47D-8DC07382CB7B}#1.0#0"; "StyleButtonX.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7888C00A-4808-4D27-9AAE-BD36EC13D16F}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form frmPaySlip 
   BackColor       =   &H8000000E&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
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
   ScaleHeight     =   9090
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin LVbuttons.LaVolpeButton CmdHide_Detail 
      Height          =   405
      Left            =   9840
      TabIndex        =   94
      Top             =   8160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Hide All Details"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   12648447
      MPTR            =   0
      MICON           =   "frmPaySlip.frx":0000
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
   Begin LVbuttons.LaVolpeButton CmdCancelCurrent 
      Height          =   375
      Left            =   3720
      TabIndex        =   93
      Top             =   8730
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancl Action"
      ENAB            =   0   'False
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
      FCOL            =   12582912
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "frmPaySlip.frx":001C
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
   Begin VB.Frame Frame_Deduct_Dtl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   120
      TabIndex        =   88
      Top             =   3480
      Visible         =   0   'False
      Width           =   11775
      Begin MSComctlLib.ListView Lst_Deduct_Dtl 
         Height          =   3495
         Left            =   15
         TabIndex        =   89
         Top             =   480
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   6165
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp - ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "For Month"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Gross Pay"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "B - Fund"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "In - Tax"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "W - Tax"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "B - Fund"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "In - Tax"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "W - Tax"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Total - Tax"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Inst. Loan"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Total Deduct"
            Object.Width           =   2293
         EndProperty
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Values in Figures "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   6025
         TabIndex        =   92
         Top             =   195
         Width           =   2580
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Values in Percentage"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   3325
         TabIndex        =   91
         Top             =   195
         Width           =   2580
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Approved Deduction Detail"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   45
         TabIndex        =   90
         Top             =   195
         Width           =   3180
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   8650
         X2              =   8650
         Y1              =   120
         Y2              =   480
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   5965
         X2              =   5965
         Y1              =   120
         Y2              =   480
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   3265
         X2              =   3265
         Y1              =   120
         Y2              =   480
      End
   End
   Begin VB.Frame Frame_Benifits_Dtl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   120
      TabIndex        =   82
      Top             =   3000
      Visible         =   0   'False
      Width           =   8715
      Begin MSComctlLib.ListView Lst_Benifits_Dtl 
         Height          =   3495
         Left            =   15
         TabIndex        =   83
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6165
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp - ID"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "For Month"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Gross Pay"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "H.Rent"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Med. Allow"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "H.Rent"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Med.Allow"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   " Applied Benifits Details"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   86
         Top             =   195
         Width           =   3375
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Values in Figure "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5580
         TabIndex        =   85
         Top             =   200
         Width           =   1920
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Values in Percentage "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3560
         TabIndex        =   84
         Top             =   200
         Width           =   1920
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   5520
         X2              =   5520
         Y1              =   480
         Y2              =   120
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   3500
         X2              =   3500
         Y1              =   480
         Y2              =   120
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   7540
         X2              =   7540
         Y1              =   480
         Y2              =   120
      End
   End
   Begin VB.Frame Frame_Gross_Detail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   120
      TabIndex        =   78
      Top             =   2775
      Visible         =   0   'False
      Width           =   11000
      Begin MSComctlLib.ListView Lst_Gross_Detail 
         Height          =   3495
         Left            =   15
         TabIndex        =   79
         Top             =   480
         Width           =   10930
         _ExtentX        =   19288
         _ExtentY        =   6165
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp - ID"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "For Month"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Holi"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Holi"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Total"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Gross Pay Details"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   45
         TabIndex        =   87
         Top             =   210
         Width           =   2415
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   2500
         X2              =   2500
         Y1              =   480
         Y2              =   120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   6735
         X2              =   6735
         Y1              =   120
         Y2              =   495
      End
      Begin VB.Label lbl_Regular_Details 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Regular Days / Holidays Detail  "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2560
         TabIndex        =   81
         Top             =   210
         Width           =   4120
      End
      Begin VB.Label lbl_Overtime_Details 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   " Over-Time Days / Holidays Detail "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6800
         TabIndex        =   80
         Top             =   210
         Width           =   4100
      End
   End
   Begin MSComctlLib.ListView Lst_PaySlip 
      Height          =   3285
      Left            =   120
      TabIndex        =   21
      Top             =   5280
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5794
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp - ID"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Month Of"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Gross. Total"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Benifits Total"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Sub Total"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Deduct Total"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Net Total"
         Object.Width           =   2469
      EndProperty
   End
   Begin LVbuttons.LaVolpeButton CmdFundTax_Detail 
      Height          =   405
      Left            =   9000
      TabIndex        =   77
      Top             =   8700
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Fund/Tax Detail"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   12648447
      MPTR            =   0
      MICON           =   "frmPaySlip.frx":0038
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
   Begin LVbuttons.LaVolpeButton CmdBenifit_Detail 
      Height          =   405
      Left            =   7200
      TabIndex        =   76
      Top             =   8700
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Benifits Detail"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   12648447
      MPTR            =   0
      MICON           =   "frmPaySlip.frx":0054
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
   Begin LVbuttons.LaVolpeButton CmdGross_Detail 
      Height          =   405
      Left            =   5400
      TabIndex        =   75
      Top             =   8700
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      BTYPE           =   3
      TX              =   "Gross Pay Detail"
      ENAB            =   0   'False
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
      FCOL            =   8388608
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   12648447
      MPTR            =   0
      MICON           =   "frmPaySlip.frx":0070
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
   Begin VB.TextBox txtFund_Tax_Amt 
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
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10560
      TabIndex        =   74
      Text            =   "txtFund_Tax_Amt"
      Top             =   4500
      Width           =   1335
   End
   Begin VB.TextBox txtBenifitAmt 
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
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   10560
      TabIndex        =   73
      Text            =   "txtBenifitAmt"
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
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
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   7200
      TabIndex        =   70
      Text            =   "txtTotal"
      Top             =   3065
      Width           =   1215
   End
   Begin VB.TextBox txtRe_Nor_Day 
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
      Left            =   6360
      TabIndex        =   67
      Text            =   "txtRe_Nor_Day"
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame_DateRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   54
      Top             =   720
      Visible         =   0   'False
      Width           =   4215
      Begin MSComCtl2.DTPicker DTPick_To 
         Height          =   375
         Left            =   2680
         TabIndex        =   58
         Top             =   240
         Width           =   1480
         _ExtentX        =   2593
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62717953
         CurrentDate     =   39291
      End
      Begin MSComCtl2.DTPicker DTPick_From 
         Height          =   375
         Left            =   700
         TabIndex        =   57
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62717953
         CurrentDate     =   39291
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To : "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   56
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "From : "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   80
         TabIndex        =   55
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame_Month 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   4215
      Begin VB.Label lblCurrentMonth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblCurrentMonth"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.OptionButton Opt_DateRange 
      BackColor       =   &H8000000E&
      Caption         =   "Date Range"
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
      Left            =   2760
      TabIndex        =   51
      Top             =   480
      Width           =   1455
   End
   Begin VB.OptionButton Opt_CurMonth 
      BackColor       =   &H8000000E&
      Caption         =   "Current Month"
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
      TabIndex        =   50
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtEmpName 
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
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "txtEmpName"
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtEmp_ID 
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
      TabIndex        =   48
      Text            =   "txtEmp_ID"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtDesig 
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
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "txtDesig"
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Allowances Apply"
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
      Height          =   765
      Left            =   4920
      TabIndex        =   42
      Top             =   3765
      Width           =   3495
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
         TabIndex        =   44
         Top             =   360
         Width           =   1455
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
         Left            =   1800
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   765
      Left            =   120
      TabIndex        =   38
      Top             =   3765
      Width           =   4575
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
         Left            =   3600
         TabIndex        =   41
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
         Left            =   1440
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtRe_Ovr_Holi 
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
      Left            =   6360
      TabIndex        =   37
      Text            =   "txtRe_Ovr_Holi"
      Top             =   2580
      Width           =   975
   End
   Begin VB.TextBox txtInTax 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10560
      TabIndex        =   36
      Text            =   "txtInTax"
      Top             =   2580
      Width           =   1335
   End
   Begin VB.TextBox txtInst_Loan 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10560
      TabIndex        =   34
      Text            =   "txtInst_Loan"
      Top             =   3560
      Width           =   1335
   End
   Begin VB.TextBox txtWHTax 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10560
      TabIndex        =   33
      Text            =   "txtWHTax"
      Top             =   3015
      Width           =   1335
   End
   Begin VB.TextBox txtBFund 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10560
      TabIndex        =   32
      Text            =   "txtBFund"
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtMedAllow 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10560
      TabIndex        =   27
      Text            =   "txtMedAllow"
      Top             =   1380
      Width           =   1335
   End
   Begin VB.TextBox txtHRent 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10560
      TabIndex        =   26
      Text            =   "txtHRent"
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtRe_Ovr_Day 
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
      Left            =   6360
      TabIndex        =   17
      Text            =   "txtRe_Ovr_Day"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtRe_Nor_Holi 
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
      Left            =   6360
      TabIndex        =   11
      Text            =   "txtRe_Nor_Holi"
      Top             =   1380
      Width           =   975
   End
   Begin StyleButtonX.StyleButton CmdExit 
      Height          =   450
      Left            =   10920
      TabIndex        =   3
      Top             =   8655
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
      PictureUp       =   "frmPaySlip.frx":008C
      PictureDown     =   "frmPaySlip.frx":06AF
      PictureHover    =   "frmPaySlip.frx":0CB5
      PictureFocus    =   "frmPaySlip.frx":12BB
      PictureDisabled =   "frmPaySlip.frx":18DE
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
      Top             =   8655
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
      PictureUp       =   "frmPaySlip.frx":1F01
      PictureDown     =   "frmPaySlip.frx":2543
      PictureHover    =   "frmPaySlip.frx":2B85
      PictureFocus    =   "frmPaySlip.frx":31B8
      PictureDisabled =   "frmPaySlip.frx":37FA
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
      Left            =   1320
      TabIndex        =   1
      Top             =   8655
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
      PictureUp       =   "frmPaySlip.frx":3BE6
      PictureDown     =   "frmPaySlip.frx":4208
      PictureHover    =   "frmPaySlip.frx":482A
      PictureFocus    =   "frmPaySlip.frx":4E6E
      PictureDisabled =   "frmPaySlip.frx":5490
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
      Top             =   8655
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
      PictureUp       =   "frmPaySlip.frx":588C
      PictureDown     =   "frmPaySlip.frx":5E9C
      PictureHover    =   "frmPaySlip.frx":64AC
      PictureFocus    =   "frmPaySlip.frx":6ACA
      PictureDisabled =   "frmPaySlip.frx":70DA
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
   Begin VB.Label Label44 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SubTotal -- Deduct"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   7680
      TabIndex        =   105
      Top             =   4965
      Width           =   1500
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Gross + Benifits]"
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5145
      TabIndex        =   104
      Top             =   4965
      Width           =   1365
   End
   Begin VB.Label txtDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "txtDate"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9960
      TabIndex        =   47
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Loan Information"
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
      Height          =   270
      Left            =   9360
      TabIndex        =   103
      Top             =   4965
      Width           =   2535
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Approved : "
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
      Height          =   255
      Left            =   9440
      TabIndex        =   102
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label lblAppLoan 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblAppLoan"
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
      Left            =   10560
      TabIndex        =   101
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance : "
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
      Height          =   255
      Left            =   9440
      TabIndex        =   100
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Paid : "
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
      Height          =   255
      Left            =   9440
      TabIndex        =   99
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label lblLoan_Balance 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblLoan_Balance"
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
      Left            =   10560
      TabIndex        =   98
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label lblPaid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblPaid"
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
      Left            =   10560
      TabIndex        =   97
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label38 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   95
      Top             =   3960
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4440
      Picture         =   "frmPaySlip.frx":749B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label19 
      BackColor       =   &H00808080&
      Caption         =   "                                     Information of Payslips"
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
      Height          =   270
      Left            =   120
      TabIndex        =   20
      Top             =   4965
      Width           =   9135
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      X1              =   8520
      X2              =   120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Fund/Tax Amount : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   8640
      TabIndex        =   72
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Benifits Amount : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   8640
      TabIndex        =   71
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Total : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   5760
      TabIndex        =   69
      Top             =   3075
      Width           =   1335
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   12000
      X2              =   8520
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   4800
      X2              =   12000
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   8520
      X2              =   4800
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblRe_Nor_Day 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblRe_Nor_Day"
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
      Height          =   360
      Left            =   7395
      TabIndex        =   68
      Top             =   990
      Width           =   495
   End
   Begin VB.Label lblT_Ovr_Holi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblT_Ovr_Holi"
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
      Height          =   360
      Left            =   7995
      TabIndex        =   66
      Top             =   2610
      Width           =   495
   End
   Begin VB.Label lblT_Ovr_Day 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblT_Ovr_Day"
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
      Height          =   360
      Left            =   7995
      TabIndex        =   65
      Top             =   2190
      Width           =   495
   End
   Begin VB.Label lblT_Nor_Holi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblT_Nor_Holi"
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
      Height          =   360
      Left            =   7995
      TabIndex        =   64
      Top             =   1395
      Width           =   495
   End
   Begin VB.Label lblT_Nor_Day 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblT_Nor_Day"
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
      Height          =   360
      Left            =   7995
      TabIndex        =   63
      Top             =   990
      Width           =   495
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   7920
      X2              =   7920
      Y1              =   840
      Y2              =   3000
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   8000
      TabIndex        =   62
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   7400
      TabIndex        =   61
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rates/Day"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   7440
      TabIndex        =   60
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblDesigID 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3360
      TabIndex        =   59
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Designation : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   45
      Top             =   3000
      Width           =   1545
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Benifits Options - (Addition)"
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
      Left            =   4920
      TabIndex        =   35
      Top             =   3510
      Width           =   3495
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Inst Of Loans : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   31
      Top             =   3560
      Width           =   1815
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "W/Holding Tax : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   30
      Top             =   3015
      Width           =   1815
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Income Tax : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   29
      Top             =   2580
      Width           =   1815
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "B - Funds : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   28
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Medical Allownce : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   25
      Top             =   1380
      Width           =   1815
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "House Rent : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   24
      Top             =   960
      Width           =   1815
   End
   Begin VB.Line Line5 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   12000
      X2              =   12000
      Y1              =   720
      Y2              =   8400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   8520
      X2              =   8520
      Y1              =   720
      Y2              =   4920
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funds / Taxes && Loans   in %age"
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
      Left            =   8640
      TabIndex        =   23
      Top             =   1845
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      X1              =   4800
      X2              =   4800
      Y1              =   720
      Y2              =   4560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   12000
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderStyle     =   2  'Dash
      X1              =   12000
      X2              =   120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Benifits - (Addition)      in %age"
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
      Left            =   8640
      TabIndex        =   22
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label lblRe_Ovr_HoliDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblRe_Ovr_HoliDay"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7395
      TabIndex        =   19
      Top             =   2610
      Width           =   495
   End
   Begin VB.Label lblRe_Ovr_Day 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblRe_Ovr_Day"
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
      Height          =   360
      Left            =   7395
      TabIndex        =   18
      Top             =   2190
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   4920
      TabIndex        =   16
      Top             =   2595
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   15
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Date Range for Pay Slip"
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
      TabIndex        =   14
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblRe_Nor_Holiday 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblRe_Nor_Holiday"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7395
      TabIndex        =   13
      Top             =   1395
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funds/Taxes && Loans Option - (Deductions)"
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
      TabIndex        =   12
      Top             =   3510
      Width           =   4575
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   4920
      TabIndex        =   10
      Top             =   1395
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Regular Days : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reported Days-(Overtime)"
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
      Left            =   4920
      TabIndex        =   8
      Top             =   1845
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reported Days-(Regular)"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name : "
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
      TabIndex        =   6
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee - ID : "
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
      Left            =   300
      TabIndex        =   5
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Information Payroll Of"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lblDetail_Help 
      BackStyle       =   0  'Transparent
      Caption         =   "It is the                                 Click on Detail List to exit from it."
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
      Height          =   240
      Left            =   240
      TabIndex        =   96
      Top             =   4620
      Width           =   7935
   End
End
Attribute VB_Name = "frmPaySlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst_Emp_Info As New ADODB.Recordset: Dim Rst_Emp_Dtl As New ADODB.Recordset:
Dim Rst_Desig As New ADODB.Recordset
Dim Rst_HurRate As New ADODB.Recordset: Dim Rst_WrkDay As New ADODB.Recordset
Dim Rst_FundTax As New ADODB.Recordset: Dim Rst_Allowce As New ADODB.Recordset
Dim Rst_LoanApp As New ADODB.Recordset: Dim Rst_LoanPaid As New ADODB.Recordset
Dim Rst_LoanInst As New ADODB.Recordset

Dim Rst_PaySlip_Info As New ADODB.Recordset: Dim Rst_PaySlip_Dtl As New ADODB.Recordset

Dim LItem_D As ListItem: Dim LItem_B As ListItem: Dim LItem_G As ListItem: Dim LItem_L As ListItem
Dim TotalPay As Integer: Dim PaidLoan As Double: Dim Opt_Paid As String


Private Sub Pop_SavingRecord()
'    Call Ctrl_PayRoll.msg_Consutruct
    For IntI = 1 To Lst_PaySlip.ListItems.Count
        With Rst_PaySlip_Info
            Set LItem = Lst_PaySlip.ListItems.Item(IntI)
            .AddNew
                .Fields(0).Value = LItem:
                .Fields(1).Value = txtDate: .Fields(2).Value = LItem.SubItems(1)
                .Fields(3).Value = LItem.SubItems(2): .Fields(4).Value = LItem.SubItems(3)
                .Fields(5).Value = LItem.SubItems(4): .Fields(6).Value = LItem.SubItems(5)
                .Fields(7).Value = LItem.SubItems(6)
            .Update
        End With
        
'=========================== Saving All Details of Employement Pay Slip ================================
        With Rst_PaySlip_Dtl
            'Following are the Gross Pay Details
            Set LItem_G = Lst_Gross_Detail.ListItems.Item(IntI)
            .AddNew
                .Fields(0).Value = LItem_G
                .Fields(1).Value = txtDate: .Fields(2).Value = LItem_G.SubItems(1)
                
                    'Following For Saving the Routine Time Calculation.
                .Fields(3).Value = LItem_G.SubItems(2): .Fields(4).Value = LItem_G.SubItems(3)
                .Fields(5).Value = LItem_G.SubItems(4): .Fields(6).Value = LItem_G.SubItems(5)
                .Fields(7).Value = LItem_G.SubItems(6) 'Total of Normal Days.
                
                    'Following For Saving the Over-Times Calculation.
                .Fields(8).Value = LItem_G.SubItems(7): .Fields(9).Value = LItem_G.SubItems(8)
                .Fields(10).Value = LItem_G.SubItems(9): .Fields(11).Value = LItem_G.SubItems(10)
                .Fields(12).Value = LItem_G.SubItems(11) 'Total of Over-Time Days.
                
            'Following are the Benifits Details
            Set LItem_B = Lst_Benifits_Dtl.ListItems.Item(IntI)
                .Fields(13).Value = LItem_B.SubItems(3): .Fields(14).Value = LItem_B.SubItems(4)
                .Fields(15).Value = LItem_B.SubItems(5): .Fields(16).Value = LItem_B.SubItems(6)
                .Fields(17).Value = LItem_B.SubItems(7)
                
            'Following are the Deduction Details
            Set LItem_D = Lst_Deduct_Dtl.ListItems.Item(IntI)
                .Fields(18).Value = LItem_D.SubItems(3): .Fields(19).Value = LItem_D.SubItems(4)
                .Fields(20).Value = LItem_D.SubItems(5): .Fields(21).Value = LItem_D.SubItems(6)
                .Fields(22).Value = LItem_D.SubItems(7): .Fields(23).Value = LItem_D.SubItems(8)
                .Fields(24).Value = LItem_D.SubItems(9) 'Total of Funds Amount
                 Opt_Paid = LItem_D.SubItems(10)
                .Fields(25).Value = LItem_D.SubItems(11) 'Total Decution Amount that Actually Deduct From Pay.
            .Update
            
            If Opt_Paid <> "Paid" Then
                If Opt_Paid <> "0" Then
                    With Rst_LoanPaid 'Installment of Loan Deduction.
                        .AddNew
                            .Fields(0).Value = LItem_D
                            .Fields(1).Value = LItem_D.SubItems(10)
                            
                            .Fields(2).Value = "Deduct From Pay"
                            .Fields(3).Value = Date
                        .Update
                    End With
                End If
            End If
            
        End With
    Next
    MsgBox "Record has been saved successfully", vbCritical, "Saving Record"
End Sub

Private Sub Pop_Reported_Days() 'For Initialize the Numeric TextBoxes.
    Chk_BFund.Value = 0: Chk_In_WH_Tax.Value = 0: Chk_Loan.Value = 0
    Chk_Medical.Value = 0: Chk_HRent.Value = 0
'======================================================
    txtRe_Nor_Day.Text = "0": txtRe_Nor_Holi.Text = "0"
    txtRe_Ovr_Day.Text = "0": txtRe_Ovr_Holi.Text = "0"
'======================================================
    lblT_Nor_Day = "": lblT_Nor_Holi = ""
    lblT_Ovr_Day = "": lblT_Ovr_Holi = ""
'======================================================
    lblRe_Nor_Day = "": lblRe_Nor_Holiday = ""
    lblRe_Ovr_Day = "": lblRe_Ovr_HoliDay = ""
'======================================================
    txtHRent.Text = "0": txtMedAllow.Text = "0"
'======================================================
    txtBFund.Text = "0": txtInTax.Text = "0"
    txtWHTax.Text = "0": txtInst_Loan.Text = "0"
    
End Sub

Private Sub CmdBenifit_Detail_Click()
    Frame_Gross_Detail.Visible = False: Frame_Deduct_Dtl.Visible = False
    Frame_Benifits_Dtl.Visible = True
    lblDetail_Help.Caption = "It is the Details of Benifits of Employee that has Approved. Click on Detail List to exit from it."
    
    CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = False: CmdFundTax_Detail.Enabled = True
    CmdHide_Detail.Enabled = True
End Sub

Private Sub CmdCancel_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmPaySlip) 'To Clearing the Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmPaySlip, False) 'Disable the From from Entry.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False:: CmdCancelCurrent.Enabled = False
    Lst_PaySlip.ListItems.Clear 'Clear the ListView.
    Lst_Benifits_Dtl.ListItems.Clear 'Clear the ListView.
    Lst_Deduct_Dtl.ListItems.Clear 'Clear the ListView.
    Lst_Gross_Detail.ListItems.Clear 'Clear the ListView.
    
    CmdNew.Enabled = True: CmdNew.SetFocus: txtEmp_ID.Text = "EMP-": txtDate = Date
    Call Pop_Reported_Days 'For Initialize the Numeric TextBoxes.
    Opt_CurMonth.Value = False: lblDesigID = "": lblAppLoan = "0": lblPaid = "0": lblLoan_Balance = "0"
    If Frame_Month.Visible = True Then Frame_Month.Visible = False
    If Frame_DateRange.Visible = True Then Frame_DateRange.Visible = False
End Sub

Private Sub CmdCancelCurrent_Click()
    Call Ctrl_PayRoll.Populate_Text_Clear(frmPaySlip) 'To Clearing the Text Boxes.
    Call Ctrl_PayRoll.Populate_Entery(frmPaySlip, False) 'Not Allow Enteries.
    CmdSubmit.Enabled = False: CmdCancel.Enabled = False: CmdCancelCurrent.Enabled = False
    CmdNew.Enabled = True: CmdNew.SetFocus: txtEmp_ID.Text = "EMP-": txtDate = Date
    Call Pop_Reported_Days 'For Initialize the Numeric TextBoxes.
    Opt_CurMonth.Value = False: lblDesigID = ""
    If Frame_Month.Visible = True Then Frame_Month.Visible = False
    If Frame_DateRange.Visible = True Then Frame_DateRange.Visible = False
    If Lst_PaySlip.ListItems.Count > 0 Then CmdSubmit.Enabled = True: CmdCancel.Enabled = True
End Sub

Private Sub CmdExit_Click()
    Unload frmPaySlip
End Sub

Private Sub CmdFundTax_Detail_Click()
    Frame_Gross_Detail.Visible = False: Frame_Benifits_Dtl.Visible = False
    Frame_Deduct_Dtl.Visible = True
    
    lblDetail_Help.Caption = "It is the Details of Approved Funds and Taxes. Click on Detail List to exit from it."
    
    CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = True: CmdFundTax_Detail.Enabled = False
    CmdHide_Detail.Enabled = True
End Sub

Private Sub CmdGross_Detail_Click()
    Frame_Deduct_Dtl.Visible = False: Frame_Benifits_Dtl.Visible = False
    Frame_Gross_Detail.Visible = True
    CmdGross_Detail.Enabled = False: CmdBenifit_Detail.Enabled = True: CmdFundTax_Detail.Enabled = True
    CmdHide_Detail.Enabled = True
    
    lblDetail_Help.Caption = "It is the Employee Payment Details after Processing. Click on Detail List to exit from it."
    
    lbl_Regular_Details.Caption = " Regular (" & Rst_WrkDay.Fields(2).Value & " X Days X Rate) / (" & Rst_WrkDay.Fields(3).Value & " X Holidays X Rate) Detail "
    lbl_Overtime_Details.Caption = " Over-Time (" & Rst_WrkDay.Fields(4).Value & " X Days X Rate) / (" & Rst_WrkDay.Fields(5).Value & " X Holidays X Rate) Detail "
End Sub

Private Sub CmdHide_Detail_Click()
    Frame_Gross_Detail.Visible = False: Frame_Deduct_Dtl.Visible = False
    Frame_Benifits_Dtl.Visible = False
    
    CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = True: CmdFundTax_Detail.Enabled = True
    CmdHide_Detail.Enabled = False

End Sub

Private Sub CmdNew_Click()
    Call Ctrl_PayRoll.Populate_Entery(frmPaySlip, True) 'Disable the From from Entry.
    Call Ctrl_PayRoll.Populate_Text_Clear(frmPaySlip) 'To Clearing the Text Boxes.
    Call Pop_Reported_Days 'For Initialize the Numeric TextBoxes.
    CmdSubmit.Enabled = True: CmdCancel.Enabled = True: CmdCancelCurrent.Enabled = True
    CmdNew.Enabled = False: txtEmp_ID.Text = "EMP-": txtDate = Date
    SendKeys "{End}": txtEmp_ID.SetFocus: Opt_Flag = "Add"
    Opt_CurMonth.Value = True
End Sub

Private Sub CmdSubmit_Click()
    Call Pop_SavingRecord
    Call CmdCancel_Click 'For Initiating the Pay Slip Form.
End Sub

Private Sub Form_Load()
    frmPaySlip.Move FrmMain.Width / 5, FrmMain.Height / 20
    Frame_Gross_Detail.Left = 120: Frame_Gross_Detail.Top = 480
    Frame_Benifits_Dtl.Left = 120: Frame_Benifits_Dtl.Top = 480
    Frame_Deduct_Dtl.Left = 120: Frame_Deduct_Dtl.Top = 480
    
    lblDetail_Help.Caption = "Welcome on Pay Slip System.": lblDetail_Help.AutoSize = True
    
    
    Call Ctrl_PayRoll.Populate_Text_Clear(frmPaySlip) 'To Clearing the Text Boxes.
    Opt_Flag = "": txtDate = Date: Call Pop_Reported_Days 'For Initialize the Numeric TextBoxes.
    Call Ctrl_PayRoll.Populate_Entery(frmPaySlip, False) 'Disable the From from Entry.
    lblDesigID = "": lblAppLoan = "0": lblPaid = "0": lblLoan_Balance = "0"
    
    Rst_Emp_Info.Open "SELECT * FROM tblEmployee_Info", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Emp_Dtl.Open "SELECT * FROM tblEmployee_Detail", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Desig.Open "SELECT * FROM tblDesignation", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_HurRate.Open "SELECT * FROM tblHourly_Rates_Set", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_WrkDay.Open "SELECT * FROM tblWorkDay_Set", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_FundTax.Open "SELECT * FROM tblTaxes_Funds_Set", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_Allowce.Open "SELECT * FROM tblAllowances_Set", DB_Conect, adOpenStatic, adLockOptimistic
    
    Rst_LoanApp.Open "SELECT * FROM tblLoan_Approved", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_LoanPaid.Open "SELECT * FROM tblLoan_Paid", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_LoanInst.Open "SELECT * FROM tblLoan_Installment", DB_Conect, adOpenStatic, adLockOptimistic
    
    Rst_PaySlip_Info.Open "SELECT * FROM tblPaySlip_Info", DB_Conect, adOpenStatic, adLockOptimistic
    Rst_PaySlip_Dtl.Open "SELECT * FROM tblPaySlip_Detail", DB_Conect, adOpenStatic, adLockOptimistic
       
    If Lst_PaySlip.ListItems.Count = 0 Then CmdGross_Detail.Enabled = False: CmdBenifit_Detail.Enabled = False: CmdFundTax_Detail.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Rst_Emp_Info.Close: Rst_Emp_Dtl.Close: Rst_Desig.Close
        
    Rst_HurRate.Close: Rst_WrkDay.Close
    Rst_FundTax.Close: Rst_Allowce.Close
    Rst_PaySlip_Info.Close: Rst_PaySlip_Dtl.Close
    
    Rst_LoanApp.Close: Rst_LoanPaid.Close: Rst_LoanInst.Close
End Sub

Private Sub lblAppLoan_Change()
    If Val(lblPaid) = Val(lblAppLoan) Then
        lblLoan_Balance = "Paid"
    ElseIf Val(lblPaid) < Val(lblAppLoan) Then
        lblLoan_Balance = Val(lblAppLoan) - Val(lblPaid)
    End If
End Sub

Private Sub lblDesigID_Change()
    If CmdNew.Enabled = False Then
        Rst_HurRate.Close: Rst_HurRate.Open "SELECT * FROM tblHourly_Rates_Set WHERE Desig_ID='" & lblDesigID & "'"
        If Rst_HurRate.RecordCount > 0 Then
            lblRe_Nor_Day = Rst_HurRate.Fields(1).Value
            lblRe_Nor_Holiday = Rst_HurRate.Fields(2).Value
            lblRe_Ovr_Day = Rst_HurRate.Fields(3).Value
            lblRe_Ovr_HoliDay = Rst_HurRate.Fields(4).Value
        ElseIf Rst_HurRate.RecordCount <= 0 Then
            lblRe_Nor_Day = "": lblRe_Nor_Holiday = ""
            lblRe_Ovr_Day = "": lblRe_Ovr_HoliDay = ""
        End If
    End If
End Sub

Private Sub lblPaid_Change()
    If Val(lblPaid) = Val(lblAppLoan) Then
        lblLoan_Balance = "Paid": txtInst_Loan.Text = "": txtInst_Loan.Text = "Paid"
    ElseIf Val(lblPaid) < Val(lblAppLoan) Then
        lblLoan_Balance = Val(lblAppLoan) - Val(lblPaid)
    End If
End Sub

Private Sub lblRe_Nor_Day_Change()
    If CmdNew.Enabled = False Then lblT_Nor_Day = Val(txtRe_Nor_Day.Text) * Val(lblRe_Nor_Day)
End Sub

Private Sub lblRe_Nor_Holiday_Change()
    If CmdNew.Enabled = False Then lblT_Nor_Holi = Val(txtRe_Nor_Holi.Text) * Val(lblRe_Nor_Holiday)
End Sub

Private Sub lblRe_Ovr_Day_Change()
    If CmdNew.Enabled = False Then lblT_Ovr_Day = Val(txtRe_Ovr_Day.Text) * Val(lblRe_Ovr_Day)
End Sub

Private Sub lblRe_Ovr_HoliDay_Change()
    If CmdNew.Enabled = False Then lblT_Ovr_Holi = Val(txtRe_Ovr_Day.Text) * Val(lblRe_Ovr_HoliDay)
End Sub

Private Sub lblT_Nor_Day_Change()
    If lblDesigID <> "" Then
'        txtTotal = Val(lblT_Nor_Day) + Val(lblT_Nor_Holi) + Val(lblT_Ovr_Day) + Val(lblT_Ovr_Holi)
        If CmdNew.Enabled = False Then txtTotal = (Rst_WrkDay.Fields(2).Value * Val(lblT_Nor_Day)) + (Rst_WrkDay.Fields(3).Value * Val(lblT_Nor_Holi)) + (Rst_WrkDay.Fields(4).Value * Val(lblT_Ovr_Day)) + (Rst_WrkDay.Fields(5).Value * Val(lblT_Ovr_Holi))
    End If
End Sub

Private Sub lblT_Nor_Holi_Change()
    If lblDesigID <> "" Then
'        txtTotal = Val(lblT_Nor_Day) + Val(lblT_Nor_Holi) + Val(lblT_Ovr_Day) + Val(lblT_Ovr_Holi)
        If CmdNew.Enabled = False Then txtTotal = (Rst_WrkDay.Fields(2).Value * Val(lblT_Nor_Day)) + (Rst_WrkDay.Fields(3).Value * Val(lblT_Nor_Holi)) + (Rst_WrkDay.Fields(4).Value * Val(lblT_Ovr_Day)) + (Rst_WrkDay.Fields(5).Value * Val(lblT_Ovr_Holi))
    End If
End Sub

Private Sub lblT_Ovr_Day_Change()
    If lblDesigID <> "" Then
'        txtTotal = Val(lblT_Nor_Day) + Val(lblT_Nor_Holi) + Val(lblT_Ovr_Day) + Val(lblT_Ovr_Holi)
        If CmdNew.Enabled = False Then txtTotal = (Rst_WrkDay.Fields(2).Value * Val(lblT_Nor_Day)) + (Rst_WrkDay.Fields(3).Value * Val(lblT_Nor_Holi)) + (Rst_WrkDay.Fields(4).Value * Val(lblT_Ovr_Day)) + (Rst_WrkDay.Fields(5).Value * Val(lblT_Ovr_Holi))
    End If
End Sub

Private Sub lblT_Ovr_Holi_Change()
    If lblDesigID <> "" Then
'        txtTotal = Val(lblT_Nor_Day) + Val(lblT_Nor_Holi) + Val(lblT_Ovr_Day) + Val(lblT_Ovr_Holi)
        If CmdNew.Enabled = False Then txtTotal = (Rst_WrkDay.Fields(2).Value * Val(lblT_Nor_Day)) + (Rst_WrkDay.Fields(3).Value * Val(lblT_Nor_Holi)) + (Rst_WrkDay.Fields(4).Value * Val(lblT_Ovr_Day)) + (Rst_WrkDay.Fields(5).Value * Val(lblT_Ovr_Holi))
    End If
End Sub

Private Sub Lst_Benifits_Dtl_Click()
    Frame_Benifits_Dtl.Visible = False
    CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = True
    CmdFundTax_Detail.Enabled = True: CmdHide_Detail.Enabled = True

    lblDetail_Help.Caption = "Welcome on Pay Slip System."
End Sub

Private Sub Lst_Deduct_Dtl_Click()
    Frame_Deduct_Dtl.Visible = False
    CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = True
    CmdFundTax_Detail.Enabled = True: CmdHide_Detail.Enabled = True
    
    lblDetail_Help.Caption = "Welcome on Pay Slip System."
End Sub

Private Sub Lst_Gross_Detail_Click()
    Frame_Gross_Detail.Visible = False
    CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = True
    CmdFundTax_Detail.Enabled = True: CmdHide_Detail.Enabled = True

    lblDetail_Help.Caption = "Welcome on Pay Slip System."
End Sub

Private Sub Opt_CurMonth_Click()
    If CmdNew.Enabled = False Then
        If Opt_CurMonth.Value = True Then
            Opt_DateRange.Value = False
            Frame_Month.Visible = True: Frame_DateRange.Visible = False
            lblCurrentMonth = ""
            lblCurrentMonth = MonthName(Month(Date), True): txtEmp_ID.SetFocus
        End If
    ElseIf CmdNew.Enabled = True Then
        MsgBox "Firt Press New Button for Creating New Slip.", vbCritical, "Error! New Entry"
        Opt_CurMonth.Value = False: CmdNew.SetFocus
    End If
End Sub

Private Sub Opt_DateRange_Click()
    If CmdNew.Enabled = False Then
        If Opt_DateRange.Value = True Then
            Opt_CurMonth.Value = False
            Frame_DateRange.Visible = True: Frame_Month.Visible = False
            DTPick_From.SetFocus
        End If
    ElseIf CmdNew.Enabled = True Then
        MsgBox "Firt Press New Button for Creating New Slip.", vbCritical, "Error! New Entry"
        Opt_DateRange.Value = False: CmdNew.SetFocus
    End If

End Sub

Private Sub txtBenifitAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtBFund_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtDesig_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtEmp_ID_Change()
    If CmdNew.Enabled = False Then
        Rst_Emp_Info.Close: Rst_Emp_Info.Open "SELECT * FROM tblEmployee_Info WHERE Emp_ID='" & txtEmp_ID.Text & "'"
        If Rst_Emp_Info.RecordCount > 0 Then
            txtEmpName.Text = Rst_Emp_Info.Fields(5).Value & " " & Rst_Emp_Info.Fields(6).Value
            lblDesigID = Rst_Emp_Info.Fields(2).Value
            
            Rst_Emp_Dtl.Close: Rst_Emp_Dtl.Open "SELECT * FROM tblEmployee_Detail WHERE Emp_ID='" & txtEmp_ID.Text & "'"
            Rst_Allowce.Close: Rst_Allowce.Open "SELECT * FROM tblAllowances_Set WHERE Desig_ID='" & lblDesigID & "'"
            Rst_FundTax.Close: Rst_FundTax.Open "SELECT * FROM tblTaxes_Funds_Set WHERE Desig_ID='" & lblDesigID & "'"

            If Rst_Emp_Dtl.Fields(1).Value = True Then Chk_HRent.Value = 1
                If Rst_Allowce.RecordCount > 0 Then txtHRent.Text = Rst_Allowce.Fields(1).Value
           
            If Rst_Emp_Dtl.Fields(2).Value = True Then Chk_Medical.Value = 1
                If Rst_Allowce.RecordCount > 0 Then txtMedAllow.Text = Rst_Allowce.Fields(2).Value

            
            If Rst_Emp_Dtl.Fields(3).Value = True Then Chk_BFund.Value = 1
                If Rst_FundTax.RecordCount > 0 Then txtBFund.Text = Rst_FundTax.Fields(1).Value
                
            If Rst_Emp_Dtl.Fields(4).Value = True Then Chk_In_WH_Tax = 1
                If Rst_FundTax.RecordCount > 0 Then txtInTax.Text = Rst_FundTax.Fields(2).Value: txtWHTax.Text = Rst_FundTax.Fields(3).Value
                
            If Rst_Emp_Dtl.Fields(5).Value = True Then Chk_Loan = 1
            
            If Rst_Emp_Dtl.Fields(1).Value = False Then Chk_HRent.Value = 0: txtHRent.Text = "0"
            If Rst_Emp_Dtl.Fields(2).Value = False Then Chk_Medical.Value = 0: txtMedAllow.Text = "0"
            If Rst_Emp_Dtl.Fields(3).Value = False Then Chk_BFund.Value = 0: txtBFund.Text = "0"
            If Rst_Emp_Dtl.Fields(4).Value = False Then Chk_In_WH_Tax = 0: txtInTax.Text = "0": txtWHTax.Text = "0"
            If Rst_Emp_Dtl.Fields(5).Value = False Then Chk_Loan = 0
            
            Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation WHERE Desig_ID='" & lblDesigID & "'"
            txtDesig.Text = Rst_Desig.Fields(1).Value
            
            Rst_LoanApp.Close: Rst_LoanApp.Open "SELECT * FROM tblLoan_Approved WHERE Emp_ID='" & txtEmp_ID.Text & "'"
            If Rst_LoanApp.RecordCount > 0 Then lblAppLoan = Rst_LoanApp.Fields(2).Value
            If Rst_LoanApp.RecordCount <= 0 Then lblAppLoan = "0"

            Rst_LoanPaid.Close: Rst_LoanPaid.Open "SELECT * FROM tblLoan_Paid WHERE Emp_ID='" & txtEmp_ID.Text & "'"
            
            PaidLoan = 0
            If Rst_LoanPaid.RecordCount > 0 Then
            For IntI = 1 To Rst_LoanPaid.RecordCount
                With Rst_LoanPaid
                    PaidLoan = .Fields(1).Value + PaidLoan
                    If Rst_LoanPaid.EOF = True Then Exit For
                    If Rst_LoanPaid.EOF = False Then Rst_LoanPaid.MoveNext
                End With
            Next
            End If: lblPaid = PaidLoan
            
            
            Rst_LoanInst.Close: Rst_LoanInst.Open "SELECT * FROM tblLoan_Installment WHERE Emp_ID='" & txtEmp_ID.Text & "'"
            If Rst_LoanInst.RecordCount > 0 Then
                If Val(lblPaid) = Val(lblAppLoan) Then lblLoan_Balance = "Paid": txtInst_Loan.Text = "": _
                                                       txtInst_Loan.Text = "Paid"
                If Val(lblPaid) < Val(lblAppLoan) Then txtInst_Loan.Text = Rst_LoanInst.Fields(1).Value: _
                                                       lblLoan_Balance = Val(lblAppLoan) - Val(lblPaid)
            ElseIf Rst_LoanInst.RecordCount <= 0 Then
                txtInst_Loan.Text = "0"
            End If
            If txtInst_Loan.Text = "Paid" Then Opt_Paid = txtInst_Loan.Text
            If txtInst_Loan.Text <> "Paid" Then Opt_Paid = txtInst_Loan.Text
            
        ElseIf Rst_Emp_Info.RecordCount <= 0 Then
            txtEmpName.Text = "Unknown Employee record": lblDesigID = "Unknown": txtDesig.Text = "Unknown"
        End If
    End If
End Sub

Private Sub txtEmp_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtEmpName.Text <> "" Then
            
            If Rst_Emp_Info.RecordCount > 0 Then
                txtRe_Nor_Day.SetFocus
            ElseIf Rst_Emp_Info.RecordCount <= 0 Then
                MsgBox "Plaese enter the Employee ID Number." & vbCrLf & "You can't leave it Empty.", _
                        vbCritical, "Error! Employee ID": txtEmp_ID.Text = ""
                        txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus
            End If
        ElseIf txtEmpName.Text = "" Then
            MsgBox "Plaese enter the Employee ID Number." & vbCrLf & "You can't leave it Empty.", _
                    vbCritical, "Error! Employee ID": txtEmp_ID.Text = ""
                    txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus
        End If
        
        Rst_Emp_Info.Close: Rst_Emp_Info.Open "SELECT * FROM tblEmployee_Info"
        Rst_Emp_Dtl.Close: Rst_Emp_Dtl.Open "SELECT * FROM tblEmployee_Detail"
        Rst_Desig.Close: Rst_Desig.Open "SELECT * FROM tblDesignation"
        Rst_HurRate.Close: Rst_HurRate.Open "SELECT * FROM tblHourly_Rates_Set"

        Rst_Allowce.Close: Rst_Allowce.Open "SELECT * FROM tblAllowances_Set"
        Rst_FundTax.Close: Rst_FundTax.Open "SELECT * FROM tblTaxes_Funds_Set"
        
        Rst_LoanInst.Close: Rst_LoanInst.Open "SELECT * FROM tblLoan_Installment"
    End If
End Sub

Private Sub txtEmpName_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtFund_Tax_Amt_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtHRent_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtInst_Loan_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtInst_Loan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Pop_Generate_PaySlip 'Generate Current Pay Slip & Values Moves In List View.
    Else
'        KeyAscii = 0
    End If
End Sub

Private Sub txtInTax_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtMedAllow_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtRe_Nor_Day_Change()
    If CmdNew.Enabled = False Then lblT_Nor_Day = Val(txtRe_Nor_Day.Text) * Val(lblRe_Nor_Day)
End Sub

Private Sub txtRe_Nor_Day_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRe_Nor_Day_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRe_Nor_Day.Text <> "0" Then
            If Val(txtRe_Nor_Day.Text) <= Rst_WrkDay.Fields(0).Value Then
                txtRe_Nor_Holi.SetFocus
            ElseIf Val(txtRe_Nor_Day.Text) > Rst_WrkDay.Fields(0).Value Then
                MsgBox "Days of month must be less or equal to " & Rst_WrkDay.Fields(0).Value & "." & vbCrLf & _
                       "Please verify it.", vbCritical, "Error! In Days Of Month"
                       SendKeys "{Home}+{End}": txtRe_Nor_Day.SetFocus
            End If
        ElseIf txtRe_Nor_Day.Text = "0" Then
            MsgBox "Days of month must be greater than 0(Zero) & less or equal to " & Rst_WrkDay.Fields(0).Value & "." & vbCrLf & _
                   "Please verify the Days.", vbCritical, "Error! In Days Of Month"
                   SendKeys "{Home}+{End}": txtRe_Nor_Day.SetFocus
        End If
    End If
End Sub

Private Sub txtRe_Nor_Holi_Change()
    If CmdNew.Enabled = False Then lblT_Nor_Holi = Val(txtRe_Nor_Holi.Text) * Val(lblRe_Nor_Holiday)
End Sub

Private Sub txtRe_Nor_Holi_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRe_Nor_Holi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRe_Nor_Holi.Text <> "0" Then
            If Val(txtRe_Nor_Holi.Text) <= Rst_WrkDay.Fields(1).Value Then
                txtRe_Ovr_Day.SetFocus
            ElseIf Val(txtRe_Nor_Holi.Text) > Rst_WrkDay.Fields(1).Value Then
                MsgBox "Holidays of month must be greater than 0(Zero) & less or equal to " & Rst_WrkDay.Fields(1).Value & "." & vbCrLf & _
                       "Please verify it.", vbCritical, "Error! In Holidays Of Month"
                       SendKeys "{Home}+{End}": txtRe_Nor_Holi.SetFocus
            End If
        ElseIf txtRe_Nor_Holi.Text = "0" Then
            MsgBox "Holidays of month must be greater than 0(Zero) & less or equal to " & Rst_WrkDay.Fields(1).Value & "." & vbCrLf & _
                   "Please verify the Holidays.", vbCritical, "Error! In Days Of Month for Overtime"
                   SendKeys "{Home}+{End}": txtRe_Nor_Holi.SetFocus
        End If
    End If
End Sub

Private Sub txtRe_Ovr_Day_Change()
    If CmdNew.Enabled = False Then lblT_Ovr_Day = Val(txtRe_Ovr_Day.Text) * Val(lblRe_Ovr_Day)
End Sub

Private Sub txtRe_Ovr_Day_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRe_Ovr_Day_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRe_Ovr_Day.Text <> "0" Then
            If Val(txtRe_Ovr_Day.Text) <= Rst_WrkDay.Fields(0).Value Then
                txtRe_Ovr_Holi.SetFocus
            ElseIf Val(txtRe_Ovr_Day.Text) > Rst_WrkDay.Fields(0).Value Then
                MsgBox "Days of month must be less or equal to " & Rst_WrkDay.Fields(0).Value & "." & vbCrLf & _
                       "Please verify it for Over-Time.", vbCritical, "Error! In Days Of Month for Overtime"
                       SendKeys "{Home}+{End}": txtRe_Ovr_Day.SetFocus
            End If
        ElseIf txtRe_Ovr_Day.Text = "0" Then
            MsgBox "Days of month must be greater than 0(Zero) & less or equal to " & Rst_WrkDay.Fields(0).Value & "." & vbCrLf & _
                   "Please verify the Days for Over-Time.", vbCritical, "Error! In Days Of Month for Overtime"
                   SendKeys "{Home}+{End}": txtRe_Ovr_Day.SetFocus
        End If
    End If
End Sub

Private Sub txtRe_Ovr_Holi_Change()
    If CmdNew.Enabled = False Then lblT_Ovr_Holi = Val(txtRe_Ovr_Holi.Text) * Val(lblRe_Ovr_HoliDay)
End Sub

Private Sub txtRe_Ovr_Holi_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtRe_Ovr_Holi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRe_Ovr_Holi.Text <> "0" Then
            If Val(txtRe_Ovr_Holi.Text) <= Val(txtRe_Nor_Holi.Text) Then
                txtInst_Loan.SetFocus
            ElseIf Val(txtRe_Ovr_Holi.Text) > Val(txtRe_Nor_Holi.Text) Then
                MsgBox "Holidays of month must be greater than 0(Zero) & less or equal to " & Val(txtRe_Nor_Holi.Text) & "." & vbCrLf & _
                       "Please verify it.", vbCritical, "Error! In Holidays Of Month"
                       SendKeys "{Home}+{End}": txtRe_Ovr_Holi.SetFocus
            End If
        ElseIf txtRe_Ovr_Holi.Text = "0" Then
            MsgBox "Holidays of month must be greater than 0(Zero) & less or equal to " & Val(txtRe_Nor_Holi.Text) & "." & vbCrLf & _
                   "Please verify the Holidays.", vbCritical, "Error! In Days Of Month for Overtime"
                   SendKeys "{Home}+{End}": txtRe_Ovr_Holi.SetFocus
        End If
    End If
End Sub

Private Sub txtTotal_Change()
    If CmdNew.Enabled = False Then txtBenifitAmt.Text = (Val(txtTotal.Text) * Val(txtHRent.Text)) / 100 + (Val(txtTotal.Text) * Val(txtMedAllow.Text)) / 100
    If CmdNew.Enabled = False Then txtFund_Tax_Amt.Text = (Val(txtTotal.Text) * Val(txtBFund.Text)) / 100 + (Val(txtTotal.Text) * Val(txtInTax.Text)) / 100 + (Val(txtTotal.Text) * Val(txtWHTax.Text)) / 100
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtWHTax_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Pop_Generate_PaySlip()
    
'    TotalPay = (Val(txtTotal.Text) + Val(txtBenifitAmt.Text)) - (Val(txtFund_Tax_Amt.Text) + Val(txtInst_Loan.Text))
    
    If Lst_PaySlip.ListItems.Count > 0 Then
        For IntI = 1 To Lst_PaySlip.ListItems.Count
            Set LItem = Lst_PaySlip.ListItems.Item(IntI)
                If ((LItem = txtEmp_ID.Text) And (LItem.SubItems(1) = lblCurrentMonth)) Then
                    MsgBox "Current Employee ID : " & txtEmp_ID.Text & " already exist." & vbCrLf & _
                           "Please check it out.", vbCritical, "Error! Exitance in List"
                           Call Ctrl_PayRoll.Populate_Text_Clear(frmPaySlip): Call Pop_Reported_Days 'For Clearing and Initiate to Zero in TextBoxes.
                           txtEmp_ID.Text = "EMP-": SendKeys "{End}": txtEmp_ID.SetFocus: Exit Sub
                End If
        Next
    End If
    
    
    Set LItem = Lst_PaySlip.ListItems.Add(1, , txtEmp_ID.Text)
        If Opt_CurMonth.Value = True Then LItem.SubItems(1) = lblCurrentMonth
        If Opt_DateRange.Value = True Then LItem.SubItems(1) = DTPick_From.Value & " - " & DTPick_To.Value
        LItem.SubItems(2) = txtTotal.Text
        LItem.SubItems(3) = txtBenifitAmt.Text
        LItem.SubItems(4) = Val(LItem.SubItems(2)) + Val(LItem.SubItems(3))
        LItem.SubItems(5) = Val(txtFund_Tax_Amt.Text) + Val(txtInst_Loan.Text)
        LItem.SubItems(6) = Val(LItem.SubItems(4)) - Val(LItem.SubItems(5))
    
        Call Pop_DetailEntries
'=================== For Initiate the next entry ========================================================
        Call Ctrl_PayRoll.Populate_Text_Clear(frmPaySlip) 'To Clearing the Text Boxes.
        Call Ctrl_PayRoll.Populate_Entery(frmPaySlip, False) 'Disable the From from Entry.
        
        CmdGross_Detail.Enabled = True: CmdBenifit_Detail.Enabled = True: CmdFundTax_Detail.Enabled = True
        CmdCancelCurrent.Enabled = False
        CmdNew.Enabled = True: CmdNew.SetFocus: txtEmp_ID.Text = "EMP-": txtDate = Date
        
        Call Pop_Reported_Days 'For Initialize the Numeric TextBoxes.
        
        Opt_CurMonth.Value = False: lblDesigID = ""
        
        If Frame_Month.Visible = True Then Frame_Month.Visible = False
        If Frame_DateRange.Visible = True Then Frame_DateRange.Visible = False
End Sub

Private Sub Pop_DetailEntries()
'====================Following code manage the Gross Detail List ==========================================
    Set LItem = Lst_Gross_Detail.ListItems.Add(1, , txtEmp_ID.Text)
        If Opt_CurMonth.Value = True Then LItem.SubItems(1) = lblCurrentMonth
        If Opt_DateRange.Value = True Then LItem.SubItems(1) = DTPick_From.Value & " - " & DTPick_To.Value
        LItem.SubItems(2) = txtRe_Nor_Day.Text '(Rst_WrkDay.Fields(2).Value * Val(lblT_Nor_Day))
        LItem.SubItems(3) = lblRe_Nor_Day 'Rates
        LItem.SubItems(4) = txtRe_Nor_Holi.Text '(Rst_WrkDay.Fields(3).Value * Val(lblT_Nor_Holi))
        LItem.SubItems(5) = lblRe_Nor_Holiday 'Rates
        LItem.SubItems(6) = (Rst_WrkDay.Fields(2).Value * Val(lblT_Nor_Day)) + (Rst_WrkDay.Fields(3).Value * Val(lblT_Nor_Holi))
        
        LItem.SubItems(7) = txtRe_Ovr_Day.Text '(Rst_WrkDay.Fields(4).Value * Val(lblT_Ovr_Day))
        LItem.SubItems(8) = lblRe_Ovr_Day 'Rates
        LItem.SubItems(9) = txtRe_Ovr_Holi.Text '(Rst_WrkDay.Fields(5).Value * Val(lblT_Ovr_Holi))
        LItem.SubItems(10) = lblRe_Ovr_HoliDay 'Rates
        LItem.SubItems(11) = (Rst_WrkDay.Fields(4).Value * Val(lblT_Ovr_Day)) + (Rst_WrkDay.Fields(5).Value * Val(lblT_Ovr_Holi))
        
'====================Following code manage the Benifits Detail List ========================================
    Set LItem = Lst_Benifits_Dtl.ListItems.Add(1, , txtEmp_ID.Text)
        If Opt_CurMonth.Value = True Then LItem.SubItems(1) = lblCurrentMonth
        If Opt_DateRange.Value = True Then LItem.SubItems(1) = DTPick_From.Value & " - " & DTPick_To.Value
        LItem.SubItems(2) = txtTotal.Text
        LItem.SubItems(3) = txtHRent.Text
        LItem.SubItems(4) = txtMedAllow.Text
        LItem.SubItems(5) = (Val(txtTotal.Text) * Val(txtHRent.Text)) / 100
        LItem.SubItems(6) = (Val(txtTotal.Text) * Val(txtMedAllow.Text)) / 100
        LItem.SubItems(7) = Val(LItem.SubItems(5)) + Val(LItem.SubItems(6))
        
'====================Following code manage the Deductions (Funds/Taxes/Loans) Detail List ========================================
    Set LItem = Lst_Deduct_Dtl.ListItems.Add(1, , txtEmp_ID.Text)
        If Opt_CurMonth.Value = True Then LItem.SubItems(1) = lblCurrentMonth
        If Opt_DateRange.Value = True Then LItem.SubItems(1) = DTPick_From.Value & " - " & DTPick_To.Value
        LItem.SubItems(2) = txtTotal.Text
        LItem.SubItems(3) = txtBFund.Text
        LItem.SubItems(4) = txtInTax.Text
        LItem.SubItems(5) = txtWHTax.Text
        LItem.SubItems(6) = (Val(txtTotal.Text) * Val(txtBFund.Text)) / 100
        LItem.SubItems(7) = (Val(txtTotal.Text) * Val(txtInTax.Text)) / 100
        LItem.SubItems(8) = (Val(txtTotal.Text) * Val(txtWHTax.Text)) / 100
        LItem.SubItems(9) = Val(LItem.SubItems(6)) + Val(LItem.SubItems(7)) + Val(LItem.SubItems(8))
        LItem.SubItems(10) = Val(txtInst_Loan.Text) 'Installment of Loan Entry.
        LItem.SubItems(11) = Val(LItem.SubItems(9)) + Val(LItem.SubItems(10)) 'Total Actual Deductions From Pay.
        
End Sub
