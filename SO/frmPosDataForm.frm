VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSO_Posform 
   Caption         =   "POINT OF SALE"
   ClientHeight    =   10170
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15120
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   Icon            =   "frmPosDataForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10170
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBillCopy 
      Height          =   405
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   10320
      Width           =   615
   End
   Begin VB.TextBox txtChipCode 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   10320
      Width           =   3615
   End
   Begin VB.TextBox txtDptTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   9720
      Width           =   735
   End
   Begin VB.TextBox txtCancelSale 
      BackColor       =   &H00C0E0FF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   9480
      MaxLength       =   255
      PasswordChar    =   "*"
      TabIndex        =   55
      Top             =   1080
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txtMoreAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   9720
      Width           =   1215
   End
   Begin VB.TextBox txtOfferAmnt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   9720
      Width           =   1215
   End
   Begin VB.TextBox txtMedicin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   9720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   9285
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   15060
      Begin VB.CommandButton cmdCancelSale 
         Caption         =   "&Cancel Sale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8160
         TabIndex        =   54
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtpackdisc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   10515
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8790
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6840
         TabIndex        =   42
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtqty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5760
         MaxLength       =   255
         TabIndex        =   40
         Text            =   "1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtarticleno 
         BackColor       =   &H00C0E0FF&
         Height          =   345
         Left            =   1605
         MaxLength       =   255
         TabIndex        =   1
         Top             =   1065
         Width           =   3480
      End
      Begin VB.TextBox txtstock 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   4095
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8790
         Width           =   1170
      End
      Begin VB.TextBox txtdiscamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   8220
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8775
         Width           =   990
      End
      Begin VB.TextBox txttotalqty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   6495
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8790
         Width           =   1035
      End
      Begin VB.TextBox txtitemdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8790
         Width           =   3180
      End
      Begin VB.Frame Frame4 
         Height          =   1005
         Left            =   15
         TabIndex        =   19
         Top             =   0
         Width           =   7920
         Begin VB.TextBox txtdumytranscode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00;(""$""#,##0.00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   3165
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   37
            TabStop         =   0   'False
            ToolTipText     =   "Total Issue Value"
            Top             =   165
            Width           =   1185
         End
         Begin VB.CommandButton cmdLookup 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2820
            Picture         =   "frmPosDataForm.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   150
            Width           =   315
         End
         Begin VB.TextBox txttransno 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1605
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "0000000000"
            Top             =   165
            Width           =   1185
         End
         Begin Crystal.CrystalReport rptVoucher 
            Left            =   0
            Top             =   0
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            Destination     =   1
            WindowBorderStyle=   3
            WindowControlBox=   0   'False
            WindowMaxButton =   0   'False
            WindowMinButton =   0   'False
            DiscardSavedData=   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowGroupTree=   -1  'True
            WindowShowCloseBtn=   -1  'True
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin MSCommLib.MSComm MSComm1 
            Left            =   0
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            CommPort        =   3
            DTREnable       =   -1  'True
         End
         Begin VB.Label dtptransdate 
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1620
            TabIndex        =   26
            Top             =   600
            Width           =   2220
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "INVOICE DATE :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   300
            TabIndex        =   21
            Top             =   615
            Width           =   1200
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "INVOICE NO :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   20
            Top             =   165
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1005
         Left            =   7875
         TabIndex        =   18
         Top             =   0
         Width           =   3480
         Begin VB.CommandButton Command3 
            Caption         =   "&Delete Dumy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   2340
            TabIndex        =   36
            Top             =   300
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Save Dumy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   180
            TabIndex        =   4
            Top             =   300
            Width           =   990
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Restore Dumy"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1305
            TabIndex        =   5
            Top             =   300
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1005
         Left            =   11370
         TabIndex        =   15
         Top             =   0
         Width           =   3615
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Cashier :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   17
            Top             =   420
            Width           =   765
         End
         Begin VB.Label lblcasherName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   915
            TabIndex        =   16
            Top             =   420
            Width           =   2205
         End
      End
      Begin VB.TextBox txtempDisc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3915
         MaxLength       =   5000
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   255
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtdiscamount1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2925
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txttotalamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   13365
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   8805
         Width           =   1590
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3510
         MaxLength       =   5000
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   225
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Frame Frame5 
         Height          =   7335
         Left            =   30
         TabIndex        =   22
         Top             =   1395
         Width           =   14970
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
            Height          =   6945
            Left            =   105
            TabIndex        =   0
            Top             =   150
            Width           =   14790
            _ExtentX        =   26088
            _ExtentY        =   12250
            _Version        =   393216
            BackColor       =   16777215
            RowHeightMin    =   475
            BackColorSel    =   16777215
            ForeColorSel    =   0
            GridColor       =   -2147483632
            WordWrap        =   -1  'True
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
      End
      Begin VB.Frame Frame6 
         Height          =   1005
         Left            =   7080
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   840
         Begin VB.CheckBox ChkEmpBill 
            Caption         =   "Employee Bill"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   75
            TabIndex        =   3
            Top             =   405
            Width           =   1395
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Pack Disc :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9225
         TabIndex        =   44
         Top             =   8835
         Width           =   1380
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "QTY :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   41
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "ARTICLE # :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   39
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Stock :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3285
         TabIndex        =   35
         Top             =   8835
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "Disc :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   7545
         TabIndex        =   33
         Top             =   8820
         Width           =   750
      End
      Begin VB.Label Label4 
         Caption         =   "Total Qty :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5295
         TabIndex        =   29
         Top             =   8835
         Width           =   1320
      End
      Begin VB.Label Label11 
         Caption         =   " Total Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   11640
         TabIndex        =   10
         Top             =   8835
         Width           =   1755
      End
      Begin VB.Label txtstatus 
         Height          =   165
         Left            =   4230
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame Frame7 
      Enabled         =   0   'False
      Height          =   585
      Left            =   120
      TabIndex        =   27
      Top             =   9120
      Width           =   15000
      Begin VB.TextBox txtNetAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$""#,##0.00;(""$""#,##0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   13320
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   120
         Width           =   1590
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   11640
         TabIndex        =   49
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   315
         Left            =   4800
         TabIndex        =   38
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label IblRateMsg 
         Alignment       =   1  'Right Justify
         Caption         =   "Rate Change Allowed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   7080
         TabIndex        =   31
         Top             =   240
         Width           =   4065
      End
      Begin VB.Label IblQtymsg 
         Caption         =   "Quantity Change Allowed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   195
         Width           =   4890
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Dept Tot."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   0
      TabIndex        =   57
      Top             =   9840
      Width           =   1095
   End
   Begin VB.Label Label16 
      Caption         =   "For Small Pizza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   11880
      TabIndex        =   53
      Top             =   9840
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Need More:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   9240
      TabIndex        =   52
      Top             =   9840
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "Pizza  Offer Amount:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5520
      TabIndex        =   47
      Top             =   9840
      Width           =   2535
   End
   Begin VB.Label Label12 
      Caption         =   "Non Offer Amount:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1920
      TabIndex        =   46
      Top             =   9840
      Width           =   2190
   End
   Begin VB.Label TXTINVTOTAL 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   16560
      TabIndex        =   6
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Begin VB.Menu New_Window 
         Caption         =   "New Window"
         Shortcut        =   ^W
      End
      Begin VB.Menu Change_Printer 
         Caption         =   "Change Printer"
      End
      Begin VB.Menu ComPort_Setting 
         Caption         =   "Com Port Setting"
      End
      Begin VB.Menu MNUEXIT 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu edit_menu 
      Caption         =   "EDIT"
      Begin VB.Menu Copy_date 
         Caption         =   "COPY"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste_data 
         Caption         =   "PASTE"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu MNUDATA 
      Caption         =   "DATA"
      Begin VB.Menu RESET 
         Caption         =   "RESET"
         Shortcut        =   ^N
      End
      Begin VB.Menu MNUNEWINVOICE 
         Caption         =   "NEW BILL"
         Shortcut        =   ^P
      End
      Begin VB.Menu EDIT_EXISTING_BILL 
         Caption         =   "EDIT EXISTING BILL"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu INSERT_ROW 
         Caption         =   "INSERT ROW"
         Shortcut        =   {F2}
      End
      Begin VB.Menu delete_row 
         Caption         =   "DELETE ROW"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Goto_Article 
         Caption         =   "GOTO ARTICLE"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Save_Record 
         Caption         =   "SAVE RECORD"
         Shortcut        =   {F12}
      End
      Begin VB.Menu GOTO_RATE 
         Caption         =   "GOTO RATE"
         Shortcut        =   {F7}
      End
      Begin VB.Menu goto_qty 
         Caption         =   "GOTO QTY"
         Shortcut        =   {F8}
      End
      Begin VB.Menu Goto_Grid 
         Caption         =   "Goto Grid"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Re_Print 
         Caption         =   "RE-PRINT"
         Enabled         =   0   'False
         Shortcut        =   ^R
      End
      Begin VB.Menu DInvoice 
         Caption         =   "Duplicate Invoice"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu Item_Sale_History 
         Caption         =   "Item Sale History"
         Enabled         =   0   'False
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmSO_Posform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Dim pr_dumy As New Recordset
Dim PR_IcItem As New Recordset
Dim ls_sql As String
Dim ls_Dispname As String
Dim ln_strlen As Integer
Public PO_CODE As Object
Public PO_DESC As Object
Dim LN_OLDQTY, ln_CQty
Dim LN_WGo As Integer
Dim LN_TabGo As Integer
Dim ln_cnt As Integer
Dim Pr_SysDate As New Recordset
Dim ls_directprint As Boolean
Dim pr_Stock As New Recordset
Dim ln_rowid As Integer
Dim NetAmt As Integer
Dim NoToken As String
Dim SalePws As String
Dim InpotBoxValue As String



Private Sub InitializeGrid()

    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Custom Code|<Description|<QTY|<Sale Price|<Amount|<Category|<U.O.M|<Discperc|<Disc Amount|<Itemcode|<EmpDiscAmount|<ChangeStatus|<ChangeQStatus|<AvgRate|<Discamount|<Discper|<AStock|<Qty in Pack|<Pack Disc|<PQty|<PDisc|<QtyANAllow|<Scan Cnt|<AvgRate1|<CatCode|<LPRate"
        .ColWidth(1) = 1700
        .ColWidth(2) = 3700
        
        .ColWidth(3) = 1000
        .ColAlignment(3) = 7

        .ColWidth(4) = 1000
        .ColAlignment(4) = 7
        
        .ColWidth(5) = 1200
        .ColAlignment(5) = 7

        .ColWidth(6) = 2400
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 1000
        .ColAlignment(9) = 7
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 700
        .ColWidth(19) = 700
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        
        .ColWidth(22) = 0
        .ColWidth(23) = 600
        .ColWidth(24) = 0
        .ColWidth(25) = 0
        
        .ColAlignment(18) = 7
        .ColAlignment(19) = 7
        .ColAlignment(23) = 7
        .ColAlignment(24) = 7
        .ColAlignment(25) = 7
        .ColAlignment(26) = 7
        
        .Redraw = True
        .Row = 1
        
        .CellBackColor = vbHighlight
    End With
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer, ls_bcode As String)
 'On Error GoTo LocalErr
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler


If KeyAscii = 13 Then 'if Enter Key then...
  
  With argFlexGrid
        ' .SelectionMode = flexSelectionByRow
        .Row = .RowSel
    
    
    ln_rowid = SearchInGridSale(GrdGRN, ls_bcode)
    
    If ln_rowid > 0 And ls_bcode <> "" And Val(.TextMatrix(ln_rowid, 10)) = 0 Then
        .Row = ln_rowid
        .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row, 3)) + Val(txtqty)
        .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 4)) * Val(.TextMatrix(.Row, 3))
        .TextMatrix(.Row, 23) = Val(.TextMatrix(.Row, 23)) + 1
        
        If Val(.TextMatrix(.Row, 3)) > 0 And Val(.TextMatrix(.Row, 20)) > 0 Then
         If Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20))) >= 1 Then
              .TextMatrix(.Row, 18) = Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20)))
              .TextMatrix(.Row, 19) = Val(.TextMatrix(.Row, 18)) * Val(.TextMatrix(.Row, 21))
         End If
        End If
        
        
       Call SaveDumyRecord(.Row)
        
        TotalAmount
        Exit Sub
    Else
  '  .Row = .Rows - 1
        If ls_bcode <> "" Then
         .TextMatrix(.Row, 1) = ls_bcode
       End If
     
    End If
    
    If .Col = 1 Then
         LN_WGo = 0
        .CellBackColor = vbWindowBackground
        ln_CQty = 1
        
       If Trim(.TextMatrix(.Row, 1)) <> "" Then
       
       
         If Len(Trim(.TextMatrix(.Row, 1))) = 13 And Left(Trim(.TextMatrix(.Row, 1)), 2) = 99 Then
          
          If PR_IcItem.State = 1 Then PR_IcItem.Close
          ls_sql = "SELECT IC_Item.CustomCode from ic_item"
          ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "' and Itemcode  = '" & "0" & Mid(.TextMatrix(.Row, 1), 3, 5) & "'"
          
          PR_IcItem.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
           If Not PR_IcItem.EOF Then
                ln_CQty = Val(Mid(.TextMatrix(.Row, 1), 8, 2) & "." & Mid(.TextMatrix(.Row, 1), 10, 3))
               .TextMatrix(.Row, 1) = Trim(PR_IcItem("Customcode") & "")
               
           End If
               PR_IcItem.Close
          
          End If
 
 ' *************** Strat No RamZan Package ********************
       
          If PR_IcItem.State = 1 Then PR_IcItem.Close
          ls_sql = "SELECT IC_Item.CustomCode,IC_Item.QtyANAllow,IC_Item.Packqty,IC_Item.PackDisc,IC_Item.avgrate,IC_Item.avgrate1,IC_Item.ItemCode, IC_Item.Description, IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_Item.discper, IC_Item.discamt,IC_Item.CatCode, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc,IC_ItemCategory.EmpDiscPer,IC_Item.PriceDescCStatus,IC_Item.QtyCStatus,IC_Item.PurchaseCost"
          ls_sql = ls_sql & " FROM IC_Item INNER JOIN IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
          ls_sql = ls_sql & " where IC_Item.compcode = '" & Gs_compcode & "' and IC_Item.Customcode  = '" & Trim(.TextMatrix(.Row, 1)) & "'"
                
      PR_IcItem.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
      
      If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
         If GrdGRN.TextMatrix(GrdGRN.Row, 2) <> "" Then
            .Col = 1
            Exit Sub
         End If
      End If
       
       
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox("Item Code not found !!! ", vbCritical)
              .TextMatrix(.Row, 1) = ""
              .Col = 1
              Exit Sub
              '.SetFocus
           '***********************************************************
            Else
               .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(PR_IcItem("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
               .TextMatrix(.Row, 20) = Val(0 & PR_IcItem("Packqty"))
               .TextMatrix(.Row, 21) = Val(0 & PR_IcItem("PackDisc"))
               .TextMatrix(.Row, 22) = Val(0 & PR_IcItem("qtyanAllow"))
               .TextMatrix(.Row, 23) = 1
                
                txtitemdesc = .TextMatrix(.Row, 2)
                If Trim(.TextMatrix(.Row, 3)) = "" Then
               .TextMatrix(.Row, 3) = ln_CQty
                End If
                
               If Trim(.TextMatrix(.Row, 4)) = "" Then
               .TextMatrix(.Row, 4) = Val(0 & PR_IcItem("Salecost"))
               End If
               
               If Trim(.TextMatrix(.Row, 10)) <> Trim(PR_IcItem("Itemcode") & "") Then
               .TextMatrix(.Row, 4) = Val(0 & PR_IcItem("Salecost"))
               End If
               
               
               .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 4)) * Val(.TextMatrix(.Row, 3))
               .TextMatrix(.Row, 6) = Trim(PR_IcItem("CatDesc") & "")
               
                              
               .TextMatrix(.Row, 7) = Trim(PR_IcItem("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & PR_IcItem("SaleDiscPerc"))
               
            
               
               If Val(0 & PR_IcItem("DiscPer")) <> 0 Then
               .TextMatrix(.Row, 9) = Round(Val(.TextMatrix(.Row, 5)) * Val(0 & PR_IcItem("discper")) / 100, 0)
               .TextMatrix(.Row, 16) = Val(0 & PR_IcItem("DiscPer"))
               End If
               
               If Val(0 & PR_IcItem("Discamt")) <> 0 Then
               .TextMatrix(.Row, 9) = Val(0 & PR_IcItem("discamt"))
               .TextMatrix(.Row, 15) = Val(0 & PR_IcItem("discamt"))
               
               End If
               
               .TextMatrix(.Row, 10) = Trim(PR_IcItem("Itemcode") & "")
               .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_IcItem("EmpDiscPer")) / 100
               .TextMatrix(.Row, 12) = Val(0 & PR_IcItem("PriceDescCStatus"))
               .TextMatrix(.Row, 13) = Val(0 & PR_IcItem("QtyCStatus"))
               .TextMatrix(.Row, 14) = Val(0 & PR_IcItem("AvgRate"))
               .TextMatrix(.Row, 24) = Val(0 & PR_IcItem("PurchaseCost"))
               .TextMatrix(.Row, 25) = PR_IcItem("CatCode")
               .TextMatrix(.Row, 26) = Val(0 & PR_IcItem("AvgRate1"))
               
               
                If Val(.TextMatrix(.Row, 3)) > 0 And Val(.TextMatrix(.Row, 20)) > 0 Then
                 If Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20))) >= 1 Then
                      .TextMatrix(.Row, 18) = Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20)))
                      .TextMatrix(.Row, 19) = Val(.TextMatrix(.Row, 18)) * Val(.TextMatrix(.Row, 21))
                 End If
                End If
               
                     
               Call SaveDumyRecord(.Row)
                
                
                
                
'               If pr_Stock.State = 1 Then pr_Stock.Close
'               pr_Stock.Open "Select isnull(qty,0) as Qty from StockSummary  where siteid =2 and itemcode = '" & .TextMatrix(.Row, 10) & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'               If Not pr_Stock.EOF Then
'               txtstock = Val(pr_Stock("Qty"))
'              .TextMatrix(.Row, 17) = Val(pr_Stock("Qty"))
'               End If
'               pr_Stock.Close
'
'               If Val(txtstock) < 0 Then
'                    Call MsgBox("Stock less then zero !!!", vbCritical)
'               ElseIf Val(txtstock) = 0 Then
'                    Call MsgBox("Stock is zero !!!", vbCritical)
'               End If
               
               
                LN_OLDQTY = 1
                LN_TabGo = 0
               IblQtymsg.Caption = ""
               IblRateMsg.Caption = ""
               If (.TextMatrix(.Row, 12)) = 1 Then
                    IblRateMsg.Caption = "Rate Change Allowed !!!"
               Else
                    IblRateMsg.Caption = "Rate Change Not Allowed !!!"
               End If
               
               If (.TextMatrix(.Row, 13)) = 1 Then
                    IblQtymsg.Caption = "Qty Change Allowed !!!"
               Else
                    IblQtymsg.Caption = "Qty Change Not Allowed !!!"
               End If
               ' Insert_row_Click
                 If .Rows = 6 Or .Rows = 11 Or .Rows = 16 Or .Rows = 21 Or .Rows = 26 Or .Rows > 26 Then
                    Command1_Click
                 End If
               
                 TotalAmount
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                End If
                    .Col = 1
                    .LeftCol = 1
                    .Row = .Row + 1
                    .SetFocus
          
                If .RowSel > 10 Then
               .TopRow = .Rows - 1 'To Move the Scrollbar
                End If
               .CellBackColor = vbHighlight
               
               If Trim(.TextMatrix(.Row, 10)) = Trim(PR_IcItem("Itemcode") & "") Then
               .CellBackColor = vbWindowBackground
               .Row = .Row - 1
               .Col = 3
               .CellBackColor = vbHighlight
               End If
               
              ' If Val(.Rows) >= 5 Then
          
              ' End If
               
                 ' Buffer to hold input string
               On Error GoTo LocalErr
               If .Rows > 2 Then
            
              If gn_comportset > 0 Then
               
               MSComm1.CommPort = gn_comportset
               MSComm1.Settings = "9600,N,8,1"
               MSComm1.InputLen = 0

               If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
                On Error GoTo LocalErr
               MSComm1.PortOpen = True
                   'MSComm1.Output = txtitemdesc & Chr$(13)
                  ls_Dispname = Trim(.TextMatrix(.Row - 1, 2))
                  ln_strlen = 20 - Len(ls_Dispname)
                  If ln_strlen > 0 Then
                  ls_Dispname = ls_Dispname + Space(ln_strlen)
                  Else
                  ls_Dispname = Left(ls_Dispname, 20)
                  End If
                  MSComm1.Output = Space(40) + Chr$(13)
                  MSComm1.Output = ls_Dispname & "Qty :" + str(Val(.TextMatrix(.Row - 1, 3))) & "Rate :" + str(Val(.TextMatrix(.Row - 1, 4))) & Chr$(13) & Chr$(10) ' Ensure that

               MSComm1.PortOpen = False
              End If
              End If
               
               
          End If
         If PR_IcItem.State = 1 Then PR_IcItem.Close
        
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
       
 ElseIf .Col = 3 Then
       .CellBackColor = vbWindowBackground
        LN_TabGo = 2
        Insert_row_Click
    
       
       If LN_OLDQTY <> Val(.TextMatrix(.Row, 3)) Then
               
               On Error GoTo LocalErr
             If gn_comportset > 0 Then
               
               MSComm1.CommPort = gn_comportset
               MSComm1.Settings = "9600,N,8,1"
               MSComm1.InputLen = 0
             
            
               If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
             
               MSComm1.PortOpen = True
                   'MSComm1.Output = txtitemdesc & Chr$(13)
                  ls_Dispname = Trim(.TextMatrix(.Row, 2))
                  ln_strlen = 20 - Len(ls_Dispname)
                  If ln_strlen > 0 Then
                  ls_Dispname = ls_Dispname + Space(ln_strlen)
                  Else
                  ls_Dispname = Left(ls_Dispname, 20)
                  End If
                  MSComm1.Output = Space(40) + Chr$(13)
                  MSComm1.Output = ls_Dispname & "Qty :" + str(Val(.TextMatrix(.Row, 3))) & "Rate :" + str(Val(.TextMatrix(.Row, 4))) & Chr$(13) & Chr$(10)  ' Ensure that
                 
                MSComm1.PortOpen = False
                
             End If
                
                
            
       End If
       
       If Val(.TextMatrix(.Row, 12)) = 1 And LN_WGo = 0 Then
              .Col = .Col + 1
               LN_WGo = 0
       .CellBackColor = vbHighlight
       Else
       
        If Trim(.TextMatrix(.Row, 3)) <> "" Then
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
'          .SetFocus
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
      End If
   
     
    ElseIf .Col = 4 Then
       .CellBackColor = vbWindowBackground
        LN_TabGo = 2
        Insert_row_Click
        
      
    End If
      GrdGRN.CellBackColor = vbHighlight
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If .Col = 1 Or .Col = 3 Or .Col = 4 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
   If Trim(.TextMatrix(.Row, 3)) <> "" Or Trim(.TextMatrix(.Row, 4)) <> "" Then
    .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 3)) * Val(.TextMatrix(.Row, 4))
   End If
   
   If Val(.TextMatrix(.Row, 16)) > 0 Then
     .TextMatrix(.Row, 9) = Round(Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 16)) / 100, 0)
   End If
   
   
   If Val(.TextMatrix(.Row, 15)) > 0 Then
        .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 3)) * Val(.TextMatrix(.Row, 15))
   End If
        
   If Val(.TextMatrix(.Row, 3)) > 0 And Val(.TextMatrix(.Row, 20)) > 0 Then
    If Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20))) >= 1 Then
         .TextMatrix(.Row, 18) = Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20)))
         .TextMatrix(.Row, 19) = Val(.TextMatrix(.Row, 18)) * Val(.TextMatrix(.Row, 21))
    End If
   End If
        
        
   Call SaveDumyRecord(.Row)
        
   
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      If .Col = 1 Then
        If .CellBackColor = vbHighlight Then
         .Text = "": .CellBackColor = vbWindowBackground
        End If
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
      
      'barcode reading for constant values
      If Left(.TextMatrix(.Row, 1), 4) = "MEDI" Then
      .TextMatrix(.Row, 4) = Val(Mid(.TextMatrix(.Row, 1), 5, 10))
      .TextMatrix(.Row, 5) = Val(Mid(.TextMatrix(.Row, 1), 5, 10))
      .TextMatrix(.Row, 1) = "MEDI"
      .TextMatrix(.Row, 3) = 1
      End If
      
      If Left(.TextMatrix(.Row, 1), 4) = "COMS" Then
      .TextMatrix(.Row, 4) = Val(Mid(.TextMatrix(.Row, 1), 5, 10))
      .TextMatrix(.Row, 5) = Val(Mid(.TextMatrix(.Row, 1), 5, 10))
      .TextMatrix(.Row, 1) = "COMS"
      .TextMatrix(.Row, 3) = 1
      End If
      
      'end bar code
      
      ElseIf .Col = 3 Then
        If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
        End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
           Exit Sub
          End If
      ElseIf .Col = 4 And Val(.TextMatrix(.Row, 12)) = 1 Then
        If .CellBackColor = vbHighlight Then
                .Text = "": .CellBackColor = vbWindowBackground
        End If
         .Text = .Text & Chr(KeyAscii)
          If Not IsNumeric(.Text) Then
          .Text = ""
           Call MsgBox("Enter Numeric entry !!!", vbCritical)
           Exit Sub
          End If
      
      End If
        
        If Trim(.TextMatrix(.Row, 3)) <> "" Or Trim(.TextMatrix(.Row, 4)) <> "" Then
        .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 3)) * Val(.TextMatrix(.Row, 4))
        End If
        
        If Val(.TextMatrix(.Row, 16)) > 0 Then
          .TextMatrix(.Row, 9) = Round(Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 16)) / 100, 0)
        End If
             
        
        If Val(.TextMatrix(.Row, 15)) > 0 Then
        .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 3)) * Val(.TextMatrix(.Row, 15))
        End If
        
        If Val(.TextMatrix(.Row, 3)) > 0 And Val(.TextMatrix(.Row, 20)) > 0 Then
         If Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20))) >= 1 Then
              .TextMatrix(.Row, 18) = Int(Val(.TextMatrix(.Row, 3)) / Val(.TextMatrix(.Row, 20)))
              .TextMatrix(.Row, 19) = Val(.TextMatrix(.Row, 18)) * Val(.TextMatrix(.Row, 21))
         End If
        End If
             
        
        
        Call SaveDumyRecord(.Row)
        
        TotalAmount
       
    End With
  End If
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub
Private Sub SaveDumyRecord(ln_rowno As Integer)
Dim mRemarks As String
 mRemarks = "SALE"

With GrdGRN

ls_sql = "delete  from  SO_Trans where srno = " & ln_rowno & "  and  userid = " & Gn_UserCode & ""
ls_servername.Execute ls_sql

ls_sql = "INSERT into SO_Trans(Compcode, TransCode,customcode, ItemCode, Quantity,Itemrate,Amount,discamount,AvgRate,srno,userid,Description,CatDesc,Remarks)"
ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Trim(txtdumytranscode) & "','" & Trim(.TextMatrix(ln_rowno, 1)) & "','" & Trim(.TextMatrix(ln_rowno, 10)) & "'," & (Val(0 & .TextMatrix(ln_rowno, 3))) & "," & Val(.TextMatrix(ln_rowno, 4)) & "," & Val(.TextMatrix(ln_rowno, 5)) & "," & Val(.TextMatrix(ln_rowno, 9)) & "," & Val(.TextMatrix(ln_rowno, 14)) & "," & Val(ln_rowno) & "," & Gn_UserCode & ",'" & RepApp(Trim(.TextMatrix(ln_rowno, 2))) & "','" & RepApp(Trim(.TextMatrix(ln_rowno, 6))) & "','" & mRemarks & "')"
ls_servername.Execute ls_sql
End With
End Sub
Private Sub Change_Printer_Click()
If ln_changeprinter = 1 Then
    ln_changeprinter = 0
Else
    ln_changeprinter = 1
End If
End Sub

Private Sub ChkEmpBill_GotFocus()
 On Error GoTo 0
 Call GetKeysAdd(GrdGRN, 13, txtarticleno)
If LN_TabGo = 1 Then
 GrdGRN.Col = 4
 GrdGRN.SetFocus
 LN_TabGo = 2
ElseIf LN_TabGo = 2 Then
 GrdGRN.Col = 1
 GrdGRN.SetFocus
 LN_TabGo = 0
Else
 goto_qty_Click
 LN_TabGo = 1
End If
 LN_WGo = 0
End Sub

Private Sub cmdCancelSale_Click()
 
 If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then

   If txtCancelSale.Visible = True Then
      txtCancelSale.Visible = False
   ElseIf txtCancelSale.Visible = False Then
      txtCancelSale.Visible = True
      txtCancelSale.SetFocus
   End If
     OptDell = ""
     txtCancelSale.Text = ""
  Else
        
    Unload Me
     
  End If
End Sub



Private Sub txtCancelSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  
  txtCancelSale.Text = ""
  txtCancelSale.Visible = False

ElseIf KeyCode = 13 Then
 
If pr_dumy.State = 1 Then pr_dumy.Close
   Set pr_dumy = gc_dbcon.Execute("Select * from SyUsers  where SalePws = '" & UCase(Trim(txtCancelSale.Text)) & "'")
  If pr_dumy.RecordCount <= 0 Then
     MsgBox ("Rerocd Not Found ...."), vbCritical
     txtCancelSale.SetFocus
  Else
  
  If OptDell = "DELLROW" Then
   If UCase(Trim(txtCancelSale.Text)) = Trim(pr_dumy("SalePws")) Then
      With GrdGRN
        If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
           .RemoveItem .Row
              If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
                txtitemdesc = ""
                txttotalamount = ""
                txttotalqty = ""
                 End If
         End With
        TotalAmount
        txtCancelSale.Text = ""
        txtCancelSale.Visible = False
        OptDell = ""
        gc_dbcon.Execute "Insert into Sys_passlog(CompCode, Userid, Remarks, Uoption, Adddate) Values ('001','" & Gc_UserId & "','" & "DELETE ROW :" + Trim(Gs_ComputerName) & "',1,'" & Format(Gd_SysDate, "YYYY/MM/DD HH:MM:SS") & "')"
    Else
    MsgBox ("Wrong PwsWord ..."), vbCritical
    txtCancelSale.SetFocus
    End If
 
   Else
  
    If UCase(Trim(txtCancelSale.Text)) = Trim(pr_dumy("SalePws")) Then
       gc_dbcon.Execute "Update SyUsers  set activestatus = (activestatus - 1) where userid = '" & Gc_UserId & "'"
       txtCancelSale.Text = ""
       txtCancelSale.Visible = False
       gc_dbcon.Execute "Insert into Sys_passlog(CompCode, Userid, Remarks, Uoption, Adddate) Values ('001','" & Gc_UserId & "','" & "SALE CLOSE :" + Trim(Gs_ComputerName) & "',1,'" & Format(Gd_SysDate, "YYYY/MM/DD HH:MM:SS") & "')"
    End
   Else
     MsgBox ("Wrong PwsWord  "), vbCritical
     txtCancelSale.SetFocus
   End If
  End If
  End If
End If
 
End Sub


Private Sub cmdLookup_Click()
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txttransno
        Set PO_DESC = Text1
        
        
        Gs_SQL = "SELECT Invoices.TransCode 'InvoiceNo', TransDate 'Invoicedate', Customer.Description as 'Customer.Description', Items.Description 'Items.Description', Invoices.NetAmount 'Invoices.NetAmount'"
        Gs_SQL = Gs_SQL & " FROM SO_TransMaster Invoices INNER JOIN   SO_Trans ON Invoices.Compcode = SO_Trans.Compcode AND Invoices.TransCode = SO_Trans.TransCode INNER JOIN"
        Gs_SQL = Gs_SQL & " IC_Item Items ON SO_Trans.Compcode = Items.Compcode AND SO_Trans.ItemCode = Items.ItemCode INNER JOIN   IC_Clients Customer ON Invoices.Compcode = Customer.Compcode AND Invoices.AccountCode = Customer.ClientCode"
        
        Gs_OrderBy = "ORDER BY Invoices.TransCode Desc"
                
        Gs_OtherPara = " Where Invoices.compcode = '" & Gs_compcode & "'"
                
        frmSosearchRecords.Caption = "Invoices"
        frmSosearchRecords.Show 1
        
        If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
ls_servername.Execute "Delete from  SO_Trans  where Compcode ='" & Gs_compcode & "' and userid = " & Gn_UserCode & " "
Call MsgBox("Dumy Records Successfully Deleted !!!", vbInformation)
 'GrdGRN.SetFocus
End Sub



Private Sub ComPort_Setting_Click()
frmcomportsetting.Show
End Sub

Private Sub Copy_date_Click()
With GrdGRN
Clipboard.Clear
Clipboard.SetText .TextMatrix(.Row, .Col)
End With
End Sub

Private Sub Delete_row_Click()
On Error GoTo LocalErr
  With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            Dim ms As String
            ms = InputBox("Enter Pwasword ....")
            If ms = "Ok777" Then
            
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
                txtitemdesc = ""
                txttotalamount = ""
                txttotalqty = ""
                TotalAmount
            End If
            Else
            MsgBox ("Access Denied ....")
            End If
    End With
  TotalAmount
Exit Sub
LocalErr:
End Sub

Private Sub DInvoice_Click()
frmSOInvoice.Show
frmSOInvoice.Caption = "Duplicate Invoice"
End Sub

Private Sub EDIT_EXISTING_BILL_Click()
txttransno.Enabled = True
cmdLookup.Enabled = True
txttransno = ""
cmdLookup_Click
NoToken = "NOTOKEN"
End Sub


Private Sub Form_Unload(Cancel As Integer)
 gc_dbcon.Execute "Update SyUsers  set activestatus = 0 where userid = '" & Gc_UserId & "'"

' gc_dbcon.Execute "Update SyUsers  set activestatus = 0,compname = '" & Trim(ls_CompName1) & "', logintime = '" & Time & "',logouttime = 'Still Active' where userid = '" & Gc_UserId & "'"
End Sub

Private Sub Goto_Article_Click()
txtarticleno.SetFocus
End Sub

Private Sub Goto_Grid_Click()
GrdGRN.SetFocus
End Sub

Private Sub goto_qty_Click()
With GrdGRN
.CellBackColor = vbWindowBackground
.Col = 3
 LN_WGo = 1
If .TextMatrix(.Row, .Col) = "" Then
.Col = 3
.Row = .Row - 1
ElseIf .TextMatrix(.Row, .Col) <> "" And .Row > 1 Then
.Row = .Row - 1
End If
.CellBackColor = vbHighlight
.SetFocus
End With
End Sub

Private Sub GOTO_RATE_Click()
With GrdGRN
.CellBackColor = vbWindowBackground
.Col = 4
If .TextMatrix(.Row, .Col) = "" Then
.Col = 4
.Row = .Row - 1
ElseIf .TextMatrix(.Row, .Col) <> "" And .Row > 1 Then
.Row = .Row - 1
End If
.CellBackColor = vbHighlight
End With

End Sub

Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
On Error GoTo ErrHandler
 If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
    If Trim(GrdGRN.TextMatrix(GrdGRN.Row, 2)) = "" Then
       Call GetKeysAdd(GrdGRN, KeyAscii, txtarticleno)
     ElseIf GrdGRN.Col <> 1 Then
       Call GetKeysAdd(GrdGRN, KeyAscii, txtarticleno)
    End If
  Else
     Call GetKeysAdd(GrdGRN, KeyAscii, txtarticleno)
  End If
Exit Sub

ErrHandler:
MsgBox ("An Error has Occured In The MSFlexgrid1_KeyPress() Procedure") & vbCr & "Report This Error To Latifjat@hotmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
End Sub

Private Sub GrdGRN_Click()
With GrdGRN
     .CellBackColor = vbHighlight
      txtitemdesc = .TextMatrix(.Row, 2)
      txtstock = .TextMatrix(.Row, 17)
 End With
End Sub
Private Sub GrdGRN_DblClick()
    GrdGRN.SelectionMode = flexSelectionFree
End Sub

Private Sub GrdGRN_EnterCell()
'With GrdGRN
'If .Col <> 1 Then
'.CellBackColor = vbButtonFace
'
'Else
'.CellBackColor = vbWhite
'End If
'End With
With GrdGRN
    GrdGRN.CellBackColor = vbHighlight
    If .Col = 1 Then
     LN_TabGo = 0
    ElseIf .Col = 3 Then
     LN_TabGo = 1
    ElseIf .Col = 4 Then
     LN_TabGo = 2
    End If
  txtstock = .TextMatrix(.Row, 17)
   
End With
End Sub


Private Sub GrdGRN_LeaveCell()
With GrdGRN
 .CellBackColor = vbWindowBackground
End With
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then

 If Me.WindowState = 0 Then
     Me.WindowState = 1
 ElseIf Me.WindowState = 1 Then
     Me.WindowState = 0
 End If
End If

If KeyCode = 13 Then ' F1 Enter key pressed n
 
 If GrdGRN.TextMatrix(GrdGRN.Row, 1) = "" Then
    DoEvents
    Text1 = ""
    Text2 = ""
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description,SaleCost as SalePrice,(((SaleCost/100)*DiscPer)+DiscAmt) as DiscAmt,(SaleCost-(((SaleCost/100)*DiscPer)+DiscAmt)) as AfterDisc FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    Gs_ExtraPara = " And Customcode = '1'"
    MyLookupSaleOLDB.Caption = "Items"
    MyLookupSaleOLDB.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
  End If
 
 'txtChipCode = "6921836900637"
 'txtChipCode = txtChipCode + ","
 'txtChipCode = txtChipCode + "024131280148"
 
 If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
   'Remarks   Trim(txtCancelSale.Text) = "024131280148" Or Trim(txtCancelSale.Text) = "6930603700288" Or Trim(txtCancelSale.Text) = "075821505126" Or Trim(txtCancelSale.Text) = "075821505041"

 If Trim(GrdGRN.TextMatrix(GrdGRN.Row, 1)) = "6921836900637" Or GrdGRN.TextMatrix(GrdGRN.Row, 1) = "024131280148" Or GrdGRN.TextMatrix(GrdGRN.Row, 1) = "075821505041" Then
    
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = " "
'
'      With GrdGRN
'        If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
'           .RemoveItem .Row
'              If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
'                InitializeGrid
'                txtitemdesc = ""
'                txttotalamount = ""
'                txttotalqty = ""
'                 End If
'         End With
'        TotalAmount
'
'        gc_dbcon.Execute "Insert into Sys_passlog(CompCode, Userid, Remarks, Uoption, Adddate) Values ('001','" & Gc_UserId & "','" & "DELETE ROW :" + Trim(Gs_ComputerName) & "',1,'" & Format(Gd_SysDate, "YYYY/MM/DD HH:MM:SS") & "')"
'
   End If
End If

Call GrdGRN_KeyPress(13)

ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed

If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
   txtCancelSale.Visible = True
   txtCancelSale.SetFocus
   OptDell = "DELLROW"
Else
   With GrdGRN
    If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
       .RemoveItem .Row
     If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
        InitializeGrid
        txtitemdesc = ""
        txttotalamount = ""
        txttotalqty = ""
     End If
     End With
     TotalAmount
End If
     
 
 ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then  'up down Key Pressed
    With GrdGRN
    txtitemdesc = .TextMatrix(.Row, 2)
    
               If Val(.TextMatrix(.Row, 12)) = 1 Then
                    IblRateMsg.Caption = "Rate Change Allowed !!!"
               Else
                    IblRateMsg.Caption = "Rate Change Not Allowed !!!"
               End If
               
               If Val(.TextMatrix(.Row, 13)) = 1 Then
                    IblQtymsg.Caption = "Qty Change Allowed !!!"
               Else
                    IblQtymsg.Caption = "Qty Change Not Allowed !!!"
               End If
               
    End With
 End If

    
End Sub
Private Sub Command1_Click()
If GrdGRN.Rows > 2 Then
'Dim ls_transcodehold As String
'ls_transcodehold = txttransno

ls_sql = "Delete from So_TransHold where Transcode = '" & txttransno & "' and compcode = '" & Gs_compcode & "' and userid  = " & Gn_UserCode & ""
gc_dbcon.Execute ls_sql
 
 
 With GrdGRN
       For ln_cnt = 1 To .Rows - 1
       If .TextMatrix(ln_cnt, 1) <> "" Then
        ls_sql = "INSERT into SO_TransHold(Compcode, TransCode,transdate,customcode, ItemCode, Quantity,Itemrate,Amount,userid)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & txttransno & "', '" & Format(dtptransdate, "YYYY/MM/DD HH:MM:SS") & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 10)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 3))) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & "," & Gn_UserCode & ")"
        gc_dbcon.Execute ls_sql
      End If
      Next
  End With
  GrdGRN.SetFocus
  GrdGRN.Row = GrdGRN.Rows - 1
  GrdGRN.Col = 1

'dtptransdate = Now
'txttransno = maxtranscode
'InitializeGrid
'GrdGRN.Row = 1
'GrdGRN.SetFocus
'Else
'    Call MsgBox("Nothing for Hold!!!", vbCritical)
End If
End Sub
Private Sub Command2_Click()
On Error GoTo LocalErr
  If pr_dumy.State = 1 Then pr_dumy.Close
  ls_sql = "Select * from so_Transhold where compcode = '" & Gs_compcode & "'"
  pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
  If pr_dumy.EOF Then
  Call MsgBox("Nothing for restore !!!", vbCritical)
  pr_dumy.Close
  Exit Sub
  End If
  
  pr_dumy.Close
  Gs_ExtraPara = ""
  Set PO_AnyForm = Nothing
  Set PO_AnyForm = Me
  Set PO_CODE = Text1
  Set PO_DESC = Text2
  Gs_SQL = "Select TransCode, TransDate from So_Transhold "
  Gs_FindFld = "Transcode"
  Gs_OrderBy = "Order by Transdate"
  Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' and userid = " & Gn_UserCode & " group by Transcode,Transdate "
  MyLookupOLDB.Caption = "Hold Trans"
  MyLookupOLDB.Show 1
  
  InitializeGrid
  GrdGRN.CellBackColor = vbWindowBackground
  If Text1 <> "" Then
   ls_sql = "SELECT IC_Item.CustomCode, IC_Item.ItemCode, IC_Item.Description, SO_TransHold.Quantity, SO_TransHold.ItemRate, SO_TransHold.Amount,"
    ls_sql = ls_sql & " IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc,IC_ItemCategory.empdiscper,IC_Item.PriceDescCStatus,IC_Item.QtyCStatus,(((IC_Item.DiscAmt/100)*IC_Item.DiscPer)+IC_Item.DiscAmt) as DiscAmt FROM IC_Item INNER JOIN"
    ls_sql = ls_sql & " IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN"
    ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode INNER JOIN"
    ls_sql = ls_sql & " SO_TransHold ON IC_Item.Compcode = SO_TransHold.Compcode AND IC_Item.ItemCode = SO_TransHold.ItemCode"
    ls_sql = ls_sql & " WHERE SO_TransHold.Compcode = '" & Gs_compcode & "' AND SO_TransHold.TransCode = '" & Text1 & "'"
  End If

With GrdGRN
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
            Do While Not pr_dumy.EOF
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(pr_dumy("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(pr_dumy("Description") & "")
               .TextMatrix(.Row, 3) = Val(0 & pr_dumy("Quantity"))
               .TextMatrix(.Row, 4) = Val(0 & pr_dumy("ItemRate"))
               .TextMatrix(.Row, 5) = Val(0 & pr_dumy("Amount"))
               .TextMatrix(.Row, 6) = Trim(pr_dumy("CatDesc") & "")
               .TextMatrix(.Row, 7) = Trim(pr_dumy("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & pr_dumy("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(0 & pr_dumy("DiscAmt"))
              ' .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 5)) * Val(0 & pr_dumy("SaleDiscPerc")) / 100
               .TextMatrix(.Row, 10) = Trim(pr_dumy("Itemcode") & "")
               .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(0 & pr_dumy("EmpDiscPer")) / 100
               .TextMatrix(.Row, 12) = Val(0 & pr_dumy("PriceDescCStatus"))
               .TextMatrix(.Row, 13) = Val(0 & pr_dumy("QtyCStatus"))
        If .Row = .Rows - 1 Then
        .Col = 1
        .Row = .Rows - 1
        .Rows = .Rows + 1
        .Row = .Rows - 1
        GrdGRN.CellBackColor = vbHighlight
       ' .SetFocus
        End If
      pr_dumy.MoveNext
      Loop
      .Row = .Rows - 1
      .TopRow = .Rows - 1
      .SetFocus
      TotalAmount

Else
    MsgBox "Hold Transaction not found", vbExclamation, "Error"
    InitializeGrid

End If
End With
pr_dumy.Close

ls_sql = "Delete from So_TransHold where Transcode = '" & Text1 & "' and compcode = '" & Gs_compcode & "' and userid = " & Gn_UserCode & " "
gc_dbcon.Execute ls_sql
dtptransdate = Now
txttransno = maxtranscode

Exit Sub
LocalErr:
Call MsgBox(Err.Description, vbCritical)

End Sub

Private Sub Form_Click()
'InitializeGrid
'dtptransdate = Date
'TXTBARCODE.Visible = False

End Sub

Private Sub Insert_row_Click()
On Error GoTo LocalErr
With GrdGRN
If .TextMatrix(.Row, 1) <> "" Then
          .CellBackColor = vbWindowBackground
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
          .Row = .Rows - 1
          .SetFocus
        Else
         'Call MsgBox("Enter/Select Item Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
           GrdGRN.CellBackColor = vbHighlight
         '.CellBackColor = vbWindowBackground
          .SetFocus
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
End With
Exit Sub
LocalErr:
Call MsgBox(Err.Description, vbCritical)

End Sub

Private Sub Item_Sale_History_Click()
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txttransno
        Set PO_DESC = Text1
        
        Gs_SQL = " SELECT  Invoice.TransCode AS 'Invoice.Transcode',Invoice.NetAmount,Invoice.TransDate as 'Invoice.TransDate',  IC_Item.Description as 'IC_Item.Description',IC_Item.CustomCode as 'IC_Item.CustomCode', SO_Trans.Quantity,   SO_Trans.ItemRate, SO_Trans.Amount, SyUsers.UserName"
        Gs_SQL = Gs_SQL & " FROM  SO_TransMaster Invoice INNER JOIN   SO_Trans ON Invoice.TransCode = SO_Trans.TransCode INNER JOIN     IC_Item ON SO_Trans.ItemCode = IC_Item.ItemCode INNER JOIN  SyUsers ON Invoice.UserCode = SyUsers.UserCode"

        
        Gs_OrderBy = "ORDER BY Invoice.TransCode Desc"
                
        Gs_OtherPara = " Where Invoice.compcode = '" & Gs_compcode & "' and SO_Trans.customcode  = '" & GrdGRN.TextMatrix(GrdGRN.Row, 1) & "'"
        
        frmSosearchRecords.Caption = "Invoices"
        frmSosearchRecords.Show 1
        If txttransno <> "" Then
            Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
            ls_directprint = True
        ElseIf txttransno = "" Then
            ls_directprint = False
        End If

End Sub

Private Sub MNUEXIT_Click()
If ln_posaccess = 1 Or ln_posaccess = 2 Or ln_posaccess = 3 Then
   gc_dbcon.Execute "Update SyUsers  set activestatus = 0 where userid = '" & Gc_UserId & "'"
Else
   gc_dbcon.Execute "Update SyUsers  set activestatus = 0 where userid = '" & Gc_UserId & "'"
   Unload Me
End If

End Sub

Private Sub mnuNewinvoice_Click()
On Error GoTo LocalErr
If ls_directprint = True Then
    Printinvoice
    ls_directprint = False
    txttransno.Enabled = False
    cmdLookup.Enabled = False
    getnewdate
    dtptransdate = Gd_SysDate
    txttransno = maxtranscode
    InitializeGrid
    Exit Sub
End If

 With GrdGRN
       For ln_cnt = 1 To .Rows - 1
           If .TextMatrix(ln_cnt, 1) <> "" Then
                    If Val(.TextMatrix(ln_cnt, 3)) = 0 Then
                        Call MsgBox("Qty not enter for some Item", vbCritical)
                        Exit Sub
                    ElseIf Val(.TextMatrix(ln_cnt, 4)) = 0 Then
                        Call MsgBox("Rate not enter for some Item", vbCritical)
                        Exit Sub
                    ElseIf Val(.TextMatrix(ln_cnt, 5)) = 0 Then
                        Call MsgBox("Amount not enter for some Item", vbCritical)
                        Exit Sub
                    End If
                End If
           Next
    End With
Label7.Caption = "Calculation Amount And Discount in Grid..."
DoEvents
CalcBeforeSave
Label7.Caption = "Calculation Total Amount..."
DoEvents

TotalAmount
Label7.Caption = ""
DoEvents
Me.Refresh

If Val(txttotalamount) > 0 Then
    'print on display
               On Error GoTo LocalErr
             
             If gn_comportset > 0 Then
               If MSComm1.PortOpen Then MSComm1.PortOpen = False
               MSComm1.CommPort = gn_comportset
               MSComm1.Settings = "9600,N,8,1"
               MSComm1.InputLen = 0
        

               If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
               
               MSComm1.PortOpen = True
                   'MSComm1.Output = txtitemdesc & Chr$(13)
                  ls_Dispname = "Total:"
                  MSComm1.Output = Space(40) + Chr$(13)
                  MSComm1.Output = ls_Dispname & str(Val(txttotalamount) - Val(txtdiscamount)) & Chr$(13) & Chr$(10)  ' Ensure that

               MSComm1.PortOpen = False
            
              End If

    txtstatus = ""
    frmSOPaidAmtform.txtitemDiscounts = Round(Val(txtdiscamount) + Val(txtpackdisc), 0)
    frmSOPaidAmtform.txttotalamount = Round(txttotalamount, 0)
    frmSOPaidAmtform.txtNetAmount = Round(txttotalamount, 0) - Round(Val(txtdiscamount) + Val(txtpackdisc), 0)
   
    frmSOPaidAmtform.Show 1
    DoEvents
        
    If txtstatus = "OK" Then
      Dim BC As Integer
      BC = 0
       For BC = 0 To Val(txtBillCopy)
           Call Printinvoice
      Next
        ls_sql = "Delete from So_TransHold where userid = " & Gn_UserCode & ""
        gc_dbcon.Execute ls_sql
        
        ls_sql = "delete  from  SO_Trans where  userid = " & Gn_UserCode & ""
        ls_servername.Execute ls_sql

    End If
End If

If txtstatus <> "Cancel" Then
txttransno.Enabled = False
cmdLookup.Enabled = False
getnewdate
dtptransdate = Gd_SysDate
txttransno = maxtranscode
InitializeGrid
txttotalamount = ""
Me.Refresh
End If

Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub getnewdate()
If Pr_SysDate.State = 1 Then Pr_SysDate.Close
Pr_SysDate.Open "Select getdate()", gc_dbcon, adOpenStatic, adLockOptimistic
Gd_SysDate = Pr_SysDate(0)
Pr_SysDate.Close
End Sub
Private Sub Printinvoice()
'On Error GoTo LocalErr

   
   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleInvoiceN.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(2) = "Reportname = 'Good Receive Note'"
         .Formulas(1) = "BarCodeStatus = " & Gn_posaccess & ""
        .Formulas(2) = "NonOfferAmt = '" & txtMedicin & "'"
        
        Dim dsc As Double
        
        If Val(txtOfferAmnt) > 0 Then
           'dsc = Val(txtOfferAmnt) - Val((txtOfferAmnt) / 100 * 10)
           dsc = Val(txtOfferAmnt)
        Else
          dsc = 0
        End If
        
        .Formulas(3) = "OfferAmt = '" & dsc & "'"
        
            
        
        .SQLQuery = "SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.DiscAmount, SO_TransMaster.RecAmount, SO_TransMaster.BalAmount, "
        .SQLQuery = .SQLQuery & " SO_TransMaster.CompName ,SO_Trans.PackQty,SO_Trans.PackDisc, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, SyUsers.UserName, IC_Item.Description,IC_Clients.Description"
        .SQLQuery = .SQLQuery & " FROM SO_TransMaster SO_TransMaster LEFT OUTER JOIN SyUsers SyUsers ON SO_TransMaster.Compcode = SyUsers.CompCode AND SO_TransMaster.UserCode = SyUsers.UserCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_Trans SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Item IC_Item ON SO_Trans.Compcode = IC_Item.Compcode AND SO_Trans.ItemCode = IC_Item.ItemCode  "
        .SQLQuery = .SQLQuery & " LEFT OUTER JOIN IC_Clients IC_Clients ON SO_TransMaster.Compcode = IC_Clients.Compcode AND SO_TransMaster.AccountCode = IC_Clients.ClientCode"
        .SQLQuery = .SQLQuery & " where SO_TransMaster.compcode = '" & Gs_compcode & "' and  SO_TransMaster.transcode = '" & Trim(txttransno) & "'"
        .SQLQuery = .SQLQuery & " ORDER BY SO_TransMaster.TransCode "
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With

Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub

Private Sub CalcBeforeSave()
  With GrdGRN
    For ln_cnt = 1 To .Rows - 1
        'check the qty before save
         If .TextMatrix(ln_cnt, 1) <> "" Then
             If Val(.TextMatrix(ln_cnt, 3)) <= 0 Then
                 Call MsgBox("Qty not enter for some Item", vbCritical)
                  Exit Sub
             ElseIf Val(.TextMatrix(ln_cnt, 4)) <= 0 Then
                 Call MsgBox("Rate not enter for some Item", vbCritical)
                  Exit Sub
             ElseIf Val(.TextMatrix(ln_cnt, 5)) <= 0 Then
                  Call MsgBox("Amount not enter for some Item", vbCritical)
                   Exit Sub
             End If
         End If
        
        
        
        If Trim(.TextMatrix(ln_cnt, 3)) <> "" Or Trim(.TextMatrix(ln_cnt, 4)) <> "" Then
         .TextMatrix(ln_cnt, 5) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 4))
        End If
        
        If Val(.TextMatrix(ln_cnt, 16)) > 0 Then
          .TextMatrix(ln_cnt, 9) = Round(Val(.TextMatrix(ln_cnt, 5)) * Val(.TextMatrix(ln_cnt, 16)) / 100, 0)
        End If
        
        If Val(.TextMatrix(ln_cnt, 15)) > 0 Then
             .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 15))
        End If

    Next
   
  End With
  
End Sub

Private Sub TotalAmount()

On Error GoTo LocalErr
ln_cnt = 0
txttotalamount = ""
txttotalqty = ""
txtdiscamount = ""
txtempDisc = ""
txtpackdisc = ""
txtNetAmount = ""

txtMedicin = ""
txtOfferAmnt = ""

 
 
 With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Format(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
            txttotalqty = Format(Val(txttotalqty) + Val(.TextMatrix(ln_cnt, 3)), "######0.00")
            txtdiscamount = Format(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
            txtempDisc = Format(Val(txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
            txtpackdisc = Format(Val(txtpackdisc) + Val(.TextMatrix(ln_cnt, 19)), "######0.00")
            txtNetAmount = Format(Val(txttotalamount) - Val(txtdiscamount), "######0.00")

    If Trim(.TextMatrix(ln_cnt, 10)) = "044780" Or Trim(.TextMatrix(ln_cnt, 10)) = "026551" Or Trim(.TextMatrix(ln_cnt, 10)) = "001006" Or Trim(.TextMatrix(ln_cnt, 10)) = "054425" Or Trim(.TextMatrix(ln_cnt, 10)) = "001005" Or Trim(.TextMatrix(ln_cnt, 10)) = "012171" Or Trim(.TextMatrix(ln_cnt, 10)) = "012631" Or Trim(.TextMatrix(ln_cnt, 10)) = "022999" Or Trim(.TextMatrix(ln_cnt, 10)) = "012632" Or Trim(.TextMatrix(ln_cnt, 10)) = "052724" Or Trim(.TextMatrix(ln_cnt, 10)) = "000849" Or Trim(.TextMatrix(ln_cnt, 10)) = "000850" Or Trim(.TextMatrix(ln_cnt, 10)) = "000851" Or Trim(.TextMatrix(ln_cnt, 10)) = "002453" Or Trim(.TextMatrix(ln_cnt, 10)) = "002454" Or Trim(.TextMatrix(ln_cnt, 10)) = "002455" Or Trim(.TextMatrix(ln_cnt, 10)) = "002466" Or Trim(.TextMatrix(ln_cnt, 10)) = "002919" Or Trim(.TextMatrix(ln_cnt, 10)) = "002920" Or Trim(.TextMatrix(ln_cnt, 10)) = "003540" Or Trim(.TextMatrix(ln_cnt, 10)) = "003598" Or Trim(.TextMatrix(ln_cnt, 10)) = "003622" Then
        
           txtMedicin = Format(Val(txtMedicin) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
           txtMedicin = Format(Val(txtMedicin) - Val(.TextMatrix(ln_cnt, 9)), "######0.00")
      
    ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "004663" Or Trim(.TextMatrix(ln_cnt, 10)) = "004664" Or Trim(.TextMatrix(ln_cnt, 10)) = "004703" Or Trim(.TextMatrix(ln_cnt, 10)) = "004764" Or Trim(.TextMatrix(ln_cnt, 10)) = "005396" Or Trim(.TextMatrix(ln_cnt, 10)) = "005398" Or Trim(.TextMatrix(ln_cnt, 10)) = "012275" Or Trim(.TextMatrix(ln_cnt, 10)) = "012454" Or Trim(.TextMatrix(ln_cnt, 10)) = "013311" Or Trim(.TextMatrix(ln_cnt, 10)) = "013677" Or Trim(.TextMatrix(ln_cnt, 10)) = "015830" Or Trim(.TextMatrix(ln_cnt, 10)) = "015832" Or Trim(.TextMatrix(ln_cnt, 10)) = "017437" Or Trim(.TextMatrix(ln_cnt, 10)) = "018301" Or Trim(.TextMatrix(ln_cnt, 10)) = "018305" Or Trim(.TextMatrix(ln_cnt, 10)) = "018710" Or Trim(.TextMatrix(ln_cnt, 10)) = "019571" Or Trim(.TextMatrix(ln_cnt, 10)) = "020128" Or Trim(.TextMatrix(ln_cnt, 10)) = "021755" Or Trim(.TextMatrix(ln_cnt, 10)) = "023762" Or Trim(.TextMatrix(ln_cnt, 10)) = "024421" Or Trim(.TextMatrix(ln_cnt, 10)) = "028955" Then
           txtMedicin = Format(Val(txtMedicin) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
           txtMedicin = Format(Val(txtMedicin) - Val(.TextMatrix(ln_cnt, 9)), "######0.00")
          
        ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "031337" Or Trim(.TextMatrix(ln_cnt, 10)) = "031956" Or Trim(.TextMatrix(ln_cnt, 10)) = "034370" Or Trim(.TextMatrix(ln_cnt, 10)) = "037467" Or Trim(.TextMatrix(ln_cnt, 10)) = "039536" Or Trim(.TextMatrix(ln_cnt, 10)) = "039831" Or Trim(.TextMatrix(ln_cnt, 10)) = "039833" Or Trim(.TextMatrix(ln_cnt, 10)) = "040987" Or Trim(.TextMatrix(ln_cnt, 10)) = "042491" Or Trim(.TextMatrix(ln_cnt, 10)) = "042493" Or Trim(.TextMatrix(ln_cnt, 10)) = "042494" Or Trim(.TextMatrix(ln_cnt, 10)) = "043975" Or Trim(.TextMatrix(ln_cnt, 10)) = "044238" Or Trim(.TextMatrix(ln_cnt, 10)) = "044304" Or Trim(.TextMatrix(ln_cnt, 10)) = "044472" Or Trim(.TextMatrix(ln_cnt, 10)) = "044473" Then
 
           txtMedicin = Format(Val(txtMedicin) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
           txtMedicin = Format(Val(txtMedicin) - Val(.TextMatrix(ln_cnt, 9)), "######0.00")
           
    ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "044886" Or Trim(.TextMatrix(ln_cnt, 10)) = "046018" Or Trim(.TextMatrix(ln_cnt, 10)) = "046579" Or Trim(.TextMatrix(ln_cnt, 10)) = "046844" Or Trim(.TextMatrix(ln_cnt, 10)) = "047478" Or Trim(.TextMatrix(ln_cnt, 10)) = "048732" Or Trim(.TextMatrix(ln_cnt, 10)) = "048853" Or Trim(.TextMatrix(ln_cnt, 10)) = "048854" Or Trim(.TextMatrix(ln_cnt, 10)) = "049491" Or Trim(.TextMatrix(ln_cnt, 10)) = "049851" Or Trim(.TextMatrix(ln_cnt, 10)) = "050494" Or Trim(.TextMatrix(ln_cnt, 10)) = "050655" Or Trim(.TextMatrix(ln_cnt, 10)) = "051936" Or Trim(.TextMatrix(ln_cnt, 10)) = "054096" Or Trim(.TextMatrix(ln_cnt, 10)) = "054323" Or Trim(.TextMatrix(ln_cnt, 10)) = "054324" Or Trim(.TextMatrix(ln_cnt, 10)) = "054325" Or Trim(.TextMatrix(ln_cnt, 10)) = "055479" Then
           
           txtMedicin = Format(Val(txtMedicin) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
           txtMedicin = Format(Val(txtMedicin) - Val(.TextMatrix(ln_cnt, 9)), "######0.00")
           
                
    ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "055745" Or Trim(.TextMatrix(ln_cnt, 10)) = "055780" Or Trim(.TextMatrix(ln_cnt, 10)) = "056491" Or Trim(.TextMatrix(ln_cnt, 10)) = "056662" Or Trim(.TextMatrix(ln_cnt, 10)) = "056733" Or Trim(.TextMatrix(ln_cnt, 10)) = "056734" Or Trim(.TextMatrix(ln_cnt, 10)) = "056911" Then
        
           txtMedicin = Format(Val(txtMedicin) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
           txtMedicin = Format(Val(txtMedicin) - Val(.TextMatrix(ln_cnt, 9)), "######0.00")
         ' txtOfferAmnt = Format(Val(txtNetAmount) - Val(txtMedicin), "######0.00")
         ' Trim (PR_IcItem("CatDesc") & "")
     
    End If
        
    
Next
     
          txtOfferAmnt = Format(Val(txtNetAmount) - Val(txtMedicin), "######0.00")
          
                    
          If Val(txtOfferAmnt) <= 10000 Then
               txtMoreAmt = 10000 - Val(txtOfferAmnt)
               Label16.Caption = "FOR SMALL PIZZA FREE"
          ElseIf Val(txtOfferAmnt) <= 20000 Then
               txtMoreAmt = 20000 - Val(txtOfferAmnt)
               Label16.Caption = "FOR MEDIUM PIZZA FREE"
          ElseIf Val(txtOfferAmnt) <= 30000 Then
                txtMoreAmt = 30000 - Val(txtOfferAmnt)
                Label16.Caption = "FOR LARGE PIZZA FREE"
          ElseIf Val(txtOfferAmnt) >= 30000 Then
                Label16.Caption = "GET LARGE PIZZA FREE"
          End If
     
     
   'If Val(txtMedicin) > 0 Then
   '  frmSO_Posform.Height = 11535
   'Else
   ' frmSO_Posform.Height = 11565
  
 '  End If
   
    End With
    'txttotalamount = Val(txttotalamount) - Val(txtdiscamount)
    
 'Pamper Discound Allow or Not Allow ..............
' ln_cnt = 0
' With GrdGRN
'
'   For ln_cnt = 1 To .Rows - 1
'
'     If Trim(.TextMatrix(ln_cnt, 10)) = "044780" Or Trim(.TextMatrix(ln_cnt, 10)) = "026551" Or Trim(.TextMatrix(ln_cnt, 10)) = "001006" Or Trim(.TextMatrix(ln_cnt, 10)) = "054425" Or Trim(.TextMatrix(ln_cnt, 10)) = "001005" Or Trim(.TextMatrix(ln_cnt, 10)) = "012171" Or Trim(.TextMatrix(ln_cnt, 10)) = "012631" Or Trim(.TextMatrix(ln_cnt, 10)) = "022999" Or Trim(.TextMatrix(ln_cnt, 10)) = "012632" Or Trim(.TextMatrix(ln_cnt, 10)) = "052724" Or Trim(.TextMatrix(ln_cnt, 10)) = "000849" Or Trim(.TextMatrix(ln_cnt, 10)) = "000850" Or Trim(.TextMatrix(ln_cnt, 10)) = "000851" Or Trim(.TextMatrix(ln_cnt, 10)) = "002453" Or Trim(.TextMatrix(ln_cnt, 10)) = "002454" Or Trim(.TextMatrix(ln_cnt, 10)) = "002455" Or Trim(.TextMatrix(ln_cnt, 10)) = "002466" Or Trim(.TextMatrix(ln_cnt, 10)) = "002919" Or Trim(.TextMatrix(ln_cnt, 10)) = "002920" Or Trim(.TextMatrix(ln_cnt, 10)) = "003540" Or Trim(.TextMatrix(ln_cnt, 10)) = "003598" Or Trim(.TextMatrix(ln_cnt, 10)) = "003622" Then
'
'           'If Val(txtOfferAmnt) <= 3000 Then
'          If Val(txtNetAmount) <= 3000 Or Val(.TextMatrix(ln_cnt, 3)) > 2 Then
'             .TextMatrix(ln_cnt, 9) = 0
'          Else
'              .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 15))
'          End If
'
'     ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "004663" Or Trim(.TextMatrix(ln_cnt, 10)) = "004664" Or Trim(.TextMatrix(ln_cnt, 10)) = "004703" Or Trim(.TextMatrix(ln_cnt, 10)) = "004764" Or Trim(.TextMatrix(ln_cnt, 10)) = "005396" Or Trim(.TextMatrix(ln_cnt, 10)) = "005398" Or Trim(.TextMatrix(ln_cnt, 10)) = "012275" Or Trim(.TextMatrix(ln_cnt, 10)) = "012454" Or Trim(.TextMatrix(ln_cnt, 10)) = "013311" Or Trim(.TextMatrix(ln_cnt, 10)) = "013677" Or Trim(.TextMatrix(ln_cnt, 10)) = "015830" Or Trim(.TextMatrix(ln_cnt, 10)) = "015832" Or Trim(.TextMatrix(ln_cnt, 10)) = "017437" Or Trim(.TextMatrix(ln_cnt, 10)) = "018301" Or Trim(.TextMatrix(ln_cnt, 10)) = "018305" Or Trim(.TextMatrix(ln_cnt, 10)) = "018710" Or Trim(.TextMatrix(ln_cnt, 10)) = "019571" Or Trim(.TextMatrix(ln_cnt, 10)) = "020128" Or Trim(.TextMatrix(ln_cnt, 10)) = "021755" Or Trim(.TextMatrix(ln_cnt, 10)) = "023762" Or Trim(.TextMatrix(ln_cnt, 10)) = "024421" Or Trim(.TextMatrix(ln_cnt, 10)) = "028955" Then
'
'
'          'If Val(txtOfferAmnt) <= 3000 Then
'          If Val(txtNetAmount) <= 3000 Or Val(.TextMatrix(ln_cnt, 3)) > 2 Then
'             .TextMatrix(ln_cnt, 9) = 0
'               Else
'              .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 15))
'          End If
'
'     ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "031337" Or Trim(.TextMatrix(ln_cnt, 10)) = "031956" Or Trim(.TextMatrix(ln_cnt, 10)) = "034370" Or Trim(.TextMatrix(ln_cnt, 10)) = "037467" Or Trim(.TextMatrix(ln_cnt, 10)) = "039536" Or Trim(.TextMatrix(ln_cnt, 10)) = "039831" Or Trim(.TextMatrix(ln_cnt, 10)) = "039833" Or Trim(.TextMatrix(ln_cnt, 10)) = "040987" Or Trim(.TextMatrix(ln_cnt, 10)) = "042491" Or Trim(.TextMatrix(ln_cnt, 10)) = "042493" Or Trim(.TextMatrix(ln_cnt, 10)) = "042494" Or Trim(.TextMatrix(ln_cnt, 10)) = "043975" Or Trim(.TextMatrix(ln_cnt, 10)) = "044238" Or Trim(.TextMatrix(ln_cnt, 10)) = "044304" Or Trim(.TextMatrix(ln_cnt, 10)) = "044472" Or Trim(.TextMatrix(ln_cnt, 10)) = "044473" Then
'
'
'           'If Val(txtOfferAmnt) <= 3000 Then
'          If Val(txtNetAmount) <= 3000 Or Val(.TextMatrix(ln_cnt, 3)) > 2 Then
'             .TextMatrix(ln_cnt, 9) = 0
'            Else
'              .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 15))
'          End If
'     ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "044886" Or Trim(.TextMatrix(ln_cnt, 10)) = "046018" Or Trim(.TextMatrix(ln_cnt, 10)) = "046579" Or Trim(.TextMatrix(ln_cnt, 10)) = "046844" Or Trim(.TextMatrix(ln_cnt, 10)) = "047478" Or Trim(.TextMatrix(ln_cnt, 10)) = "048732" Or Trim(.TextMatrix(ln_cnt, 10)) = "048853" Or Trim(.TextMatrix(ln_cnt, 10)) = "048854" Or Trim(.TextMatrix(ln_cnt, 10)) = "049491" Or Trim(.TextMatrix(ln_cnt, 10)) = "049851" Or Trim(.TextMatrix(ln_cnt, 10)) = "050494" Or Trim(.TextMatrix(ln_cnt, 10)) = "050655" Or Trim(.TextMatrix(ln_cnt, 10)) = "051936" Or Trim(.TextMatrix(ln_cnt, 10)) = "054096" Or Trim(.TextMatrix(ln_cnt, 10)) = "054323" Or Trim(.TextMatrix(ln_cnt, 10)) = "054324" Or Trim(.TextMatrix(ln_cnt, 10)) = "054325" Or Trim(.TextMatrix(ln_cnt, 10)) = "055479" Then
'
'
'           'If Val(txtOfferAmnt) <= 3000 Then
'          If Val(txtNetAmount) <= 3000 Or Val(.TextMatrix(ln_cnt, 3)) > 2 Then
'             .TextMatrix(ln_cnt, 9) = 0
'            Else
'              .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 15))
'          End If
'
'     ElseIf Trim(.TextMatrix(ln_cnt, 10)) = "055745" Or Trim(.TextMatrix(ln_cnt, 10)) = "055780" Or Trim(.TextMatrix(ln_cnt, 10)) = "056491" Or Trim(.TextMatrix(ln_cnt, 10)) = "056662" Or Trim(.TextMatrix(ln_cnt, 10)) = "056733" Or Trim(.TextMatrix(ln_cnt, 10)) = "056734" Or Trim(.TextMatrix(ln_cnt, 10)) = "056911" Then
'
'          'If Val(txtOfferAmnt) <= 3000 Then
'          If Val(txtNetAmount) <= 3000 Or Val(.TextMatrix(ln_cnt, 3)) > 2 Then
'             .TextMatrix(ln_cnt, 9) = 0
'            Else
'              .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 3)) * Val(.TextMatrix(ln_cnt, 15))
'          End If
'     End If
'End Pamper Discound Allow or Not Allow ..............
'  Next
'End With
    

txttotalamount = ""
txttotalqty = ""
txtdiscamount = ""
txtempDisc = ""
txtpackdisc = ""
txtNetAmount = ""
'txtMedicin = ""
'txtOfferAmnt = ""
 ln_cnt = 0

 With GrdGRN
 'Final Total  Allow or Not Allow ..............
   For ln_cnt = 1 To .Rows - 1
            txttotalamount = Format(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
            txttotalqty = Format(Val(txttotalqty) + Val(.TextMatrix(ln_cnt, 3)), "######0.00")
            txtdiscamount = Format(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
            txtempDisc = Format(Val(txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
            txtpackdisc = Format(Val(txtpackdisc) + Val(.TextMatrix(ln_cnt, 19)), "######0.00")
            txtNetAmount = Format(Val(txttotalamount) - Val(txtdiscamount), "######0.00")
 'Final Total Allow or Not Allow ..............
  Next
End With
    
    
    
Exit Sub
LocalErr:
Call MsgBox(Err.Description, vbCritical)
  
End Sub
Private Sub FinalTotalAmount()

txttotalamount = ""
txtdiscamount = ""
txtempDisc = ""
txttotalqty = ""
txtpackdisc = ""
'txtMedicin = ""
'txtOfferAmnt = ""
 

 With GrdGRN
 'Pamper Discound Allow or Not Allow ..............
   For ln_cnt = 1 To .Rows - 1
            txttotalamount = Format(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 5)), "######0.00")
            txttotalqty = Format(Val(txttotalqty) + Val(.TextMatrix(ln_cnt, 3)), "######0.00")
            txtdiscamount = Format(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
            txtempDisc = Format(Val(txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
            txtpackdisc = Format(Val(txtpackdisc) + Val(.TextMatrix(ln_cnt, 19)), "######0.00")
            txtNetAmount = Format(Val(txttotalamount) - Val(txtdiscamount), "######0.00")
'End Pamper Discound Allow or Not Allow ..............
  Next
End With

End Sub




Private Sub Form_Load()

If Trim(Gc_UserName) = "Administrator" Then
   EDIT_EXISTING_BILL.Enabled = True
   Re_Print.Enabled = True
   DInvoice.Enabled = True
   Item_Sale_History.Enabled = True
End If
InitializeGrid
dtptransdate = Gd_SysDate
lblcasherName = Gc_UserName
ls_directprint = False
txttransno = maxtranscode
'Call mnuNewinvoice_Click
 NoToken = ""
Call CheckDumyRecord
'Me.WindowState = 2

End Sub
Private Sub CheckDumyRecord()
Dim PR_Dumyrecord As New Recordset
Dim PR_Dumyrecord1 As New Recordset
Dim res
ls_sql = "SELECT * from so_trans where userid = " & Gn_UserCode & " order by srno"

If PR_Dumyrecord.State = 1 Then PR_Dumyrecord.Close
PR_Dumyrecord.Open ls_sql, ls_servername, adOpenStatic, adLockReadOnly, 1

If Not PR_Dumyrecord.EOF Then
res = MsgBox("Due to Close of software Suddenly Do you want to load save local data", vbYesNoCancel + vbInformation)
If res = vbYes Then
    If PR_Dumyrecord1.State = 1 Then PR_Dumyrecord1.Close
    PR_Dumyrecord1.Open "SELECT Transcode,Userid FROM So_Trans group by  Transcode,Userid ", ls_servername, adOpenStatic, adLockReadOnly, 1

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtdumytranscode
    Set PO_DESC = Text2
    GoTop PR_Dumyrecord1
    MyLookup.Caption = "Local Save Records ID"
    MyLookup.FillGrid PR_Dumyrecord1, "Transcode", "Transcode", 5
    MyLookup.Show 1
    PR_Dumyrecord1.Close
    
  ls_sql = "SELECT * from so_trans where userid = " & Gn_UserCode & " and transcode = '" & txtdumytranscode & "' order by srno"

  If PR_Dumyrecord.State = 1 Then PR_Dumyrecord.Close
  PR_Dumyrecord.Open ls_sql, ls_servername, adOpenStatic, adLockReadOnly, 1
    
    
    getnewdate
    dtptransdate = Gd_SysDate
    txttransno = maxtranscode
    InitializeGrid

With GrdGRN
If Not PR_Dumyrecord.EOF Then
            Do While Not PR_Dumyrecord.EOF
                .CellBackColor = vbWindowBackground
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(PR_Dumyrecord("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(PR_Dumyrecord("Description") & "")
               .TextMatrix(.Row, 3) = Val(0 & PR_Dumyrecord("Quantity"))
               .TextMatrix(.Row, 4) = Val(0 & PR_Dumyrecord("ItemRate"))
               .TextMatrix(.Row, 5) = Val(0 & PR_Dumyrecord("Amount"))
               .TextMatrix(.Row, 6) = Trim(PR_Dumyrecord("CatDesc") & "")
           '    .TextMatrix(.Row, 7) = Trim(PR_Dumy("UOM") & "")
            '   .TextMatrix(.Row, 8) = Val(0 & PR_Dumy("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(0 & PR_Dumyrecord("DiscAmount"))
               .TextMatrix(.Row, 10) = Trim(PR_Dumyrecord("Itemcode") & "")
              If .Row = .Rows - 1 Then
              .Col = 1
              .Row = .Rows - 1
              .Rows = .Rows + 1
              .Row = .Rows - 1
              End If
            PR_Dumyrecord.MoveNext
          Loop
       TotalAmount
       .Row = .Rows - 1
       .TopRow = .Rows - 1
     
      
End If
End With
ElseIf res = vbCancel Then
    Command3_Click
End If

End If
If PR_Dumyrecord.State = 1 Then PR_Dumyrecord.Close


End Sub
Public Function maxtranscode() As String
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open "select max(transcode) as transcode from SO_TransMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
txtdumytranscode = Right(maxtranscode, 3) + "T" + Trim(str(Hour(Now))) + Trim(str(Minute(Now)) + Trim(str(Second(Now))))
End Function
Public Function maxtranscodehold() As String
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open "select max(transcode) as transcode from SO_Transhold where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscodehold = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscodehold = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function

Private Sub New_Window_Click()
Shell App.Path & "/ecounts.exe"
End Sub

Private Sub Paste_data_Click()
With GrdGRN
.TextMatrix(.Row, .Col) = Clipboard.GetText
End With
End Sub

Private Sub RE_PRINT_Click()
        Dim ld_ldate As Date
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txttransno
        Set PO_DESC = Text1
        
        ld_ldate = DateAdd("D", -5, Date)
        
        Gs_SQL = "SELECT Invoices.TransCode AS InvoiceNo, Invoices.TransDate AS Invoicedate, Customer.Description AS 'Customer.Description',"
        Gs_SQL = Gs_SQL & " Invoices.NetAmount AS 'Invoices.NetAmount', SyUsers.UserName fROM SO_TransMaster Invoices INNER JOIN"
        Gs_SQL = Gs_SQL & " IC_Clients Customer ON Invoices.Compcode = Customer.Compcode AND Invoices.AccountCode = Customer.ClientCode LEFT OUTER JOIN"
        Gs_SQL = Gs_SQL & " SyUsers ON Invoices.Compcode = SyUsers.CompCode AND Invoices.UserCode = SyUsers.UserCode"
        
        Gs_OrderBy = "ORDER BY Invoices.TransCode Desc"
        
        If chkRights1("SYSTEM99SM") Then
            Gs_OtherPara = " Where Invoices.compcode = '" & Gs_compcode & "'"
        Else
            Gs_OtherPara = " Where Invoices.compcode = '" & Gs_compcode & "' and  SyUsers.UserCode = " & Gn_UserCode & " and Invoices.TransDate >= '" & Format(ld_ldate, "YYYY/MM/DD") & "'"
        
        End If
        
        frmSosearchRecords.Caption = "Invoices"
        frmSosearchRecords.Show 1
        If txttransno <> "" Then
            Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)
            ls_directprint = True
        ElseIf txttransno = "" Then
            ls_directprint = False
        End If
NoToken = "NOTOKEN"
End Sub

Private Sub RESET_Click()
ls_directprint = False
txttransno.Enabled = False
cmdLookup.Enabled = False
getnewdate
dtptransdate = Gd_SysDate
txttransno = maxtranscode
InitializeGrid
Exit Sub
End Sub

Private Sub Save_Record_Click()
mnuNewinvoice_Click
End Sub

Private Sub txtarticleno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtarticleno <> "" Then
 With GrdGRN
 Call GetKeysAdd(GrdGRN, vbKeyReturn, txtarticleno)
 txtarticleno = ""
 txtarticleno.SetFocus
End With
ElseIf KeyCode = vbKeyReturn And txtarticleno = "" Then
 DoEvents
   txtarticleno = ""
    Text2 = ""
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtarticleno
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description,SaleCost as SalePrice FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    Gs_ExtraPara = " And Customcode = '1'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
     Call GetKeysAdd(GrdGRN, vbKeyReturn, txtarticleno)
     txtarticleno = ""
    txtarticleno.SetFocus


End If
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txttransno.Text) > 0 Then
         If pr_dumy.State = 1 Then pr_dumy.Close
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         pr_dumy.Open "select * from SO_TransMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If pr_dumy.EOF Then
                   Call MsgBox("Record not found !!!", vbCritical)
                   If txttransno.Enabled Then txttransno.SetFocus
         Else
                   dtptransdate = pr_dumy("Transdate")
                   InitializeGrid
                   GrdGRN.CellBackColor = vbWindowBackground
                   LoadGRNTrans
         End If
End If
End Sub
Private Sub LoadGRNTrans()

ls_sql = "SELECT IC_Item.CustomCode, IC_Item.ItemCode, IC_Item.Description, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount,"
ls_sql = ls_sql & " IC_Item.SaleDiscPerc, IC_Item.SaleCost,SO_Trans.PackQty,SO_Trans.PackDisc, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc FROM IC_Item INNER JOIN"
ls_sql = ls_sql & " IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN"
ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode INNER JOIN"
ls_sql = ls_sql & " SO_Trans ON IC_Item.Compcode = SO_Trans.Compcode AND IC_Item.ItemCode = SO_Trans.ItemCode"
ls_sql = ls_sql & " WHERE (SO_Trans.Compcode = '" & Gs_compcode & "') AND (SO_Trans.TransCode = '" & txttransno & "')"

With GrdGRN
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
            Do While Not pr_dumy.EOF
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(pr_dumy("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(pr_dumy("Description") & "")
               .TextMatrix(.Row, 3) = Val(0 & pr_dumy("Quantity"))
               .TextMatrix(.Row, 4) = Val(0 & pr_dumy("ItemRate"))
               .TextMatrix(.Row, 5) = Val(0 & pr_dumy("Amount"))
               .TextMatrix(.Row, 6) = Trim(pr_dumy("CatDesc") & "")
               .TextMatrix(.Row, 7) = Trim(pr_dumy("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & pr_dumy("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 5)) * Val(0 & pr_dumy("SaleDiscPerc")) / 100
               .TextMatrix(.Row, 10) = Trim(pr_dumy("Itemcode") & "")
               .TextMatrix(.Row, 18) = Val(0 & pr_dumy("PackQty"))
               .TextMatrix(.Row, 19) = Val(0 & pr_dumy("PackDisc"))
            
        
        
        If .Row = .Rows - 1 Then
        .Col = 1
        .Row = .Rows - 1
        .Rows = .Rows + 1
        .Row = .Rows - 1
         'TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
         'TXTBARCODE.Text = .TextMatrix(.Row, 1)
         'TXTBARCODE.Visible = True
        ElseIf .Row < .Rows - 1 Then
            .Row = .Rows - 1
            ' TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
            ' TXTBARCODE.Text = .TextMatrix(.Row, 1)
        End If
        
      pr_dumy.MoveNext
      Loop
       TotalAmount
       .Row = .Rows - 1
       .TopRow = .Rows - 1
       .SetFocus
      
Else
    MsgBox "Transaction not found", vbExclamation, "Error"
    InitializeGrid
    'TXTBARCODE.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' .CellHeight - CZ
End If
End With
pr_dumy.Close

'TXTBARCODE.SetFocus
End Sub

