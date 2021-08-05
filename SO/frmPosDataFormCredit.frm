VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmSO_PosformCredit 
   Caption         =   "Credit Sale"
   ClientHeight    =   7950
   ClientLeft      =   -1980
   ClientTop       =   795
   ClientWidth     =   11580
   Icon            =   "frmPosDataFormCredit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   7395
      Left            =   75
      TabIndex        =   6
      Top             =   0
      Width           =   11460
      Begin VB.TextBox txtmiscper 
         Alignment       =   1  'Right Justify
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   9555
         MaxLength       =   11
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6555
         Width           =   435
      End
      Begin VB.CheckBox chkprate1 
         Caption         =   "Purchase Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1200
         TabIndex        =   42
         Top             =   6690
         Width           =   1830
      End
      Begin VB.TextBox txtmiscamount 
         Alignment       =   1  'Right Justify
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   10005
         MaxLength       =   11
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6555
         Width           =   1185
      End
      Begin VB.CheckBox chkPRate 
         Caption         =   "Avg Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   75
         TabIndex        =   39
         Top             =   6675
         Value           =   1  'Checked
         Width           =   1830
      End
      Begin VB.TextBox txtnetamount 
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   9570
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6990
         Width           =   1635
      End
      Begin VB.TextBox txtdiscamt 
         Alignment       =   1  'Right Justify
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   9555
         MaxLength       =   11
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6135
         Width           =   1635
      End
      Begin VB.TextBox txtdiscper 
         Alignment       =   1  'Right Justify
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   9555
         MaxLength       =   11
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5715
         Width           =   1140
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2040
         Picture         =   "frmPosDataFormCredit.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   5730
         Width           =   360
      End
      Begin VB.TextBox txtCreditDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2430
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   5745
         Width           =   5625
      End
      Begin VB.TextBox txtCreditCode 
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
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   28
         Top             =   5745
         Width           =   900
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5310
         Width           =   7995
      End
      Begin VB.Frame Frame6 
         Height          =   1005
         Left            =   3705
         TabIndex        =   22
         Top             =   0
         Width           =   2145
         Begin VB.CommandButton Command3 
            Caption         =   "&Save"
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
            Left            =   90
            TabIndex        =   27
            Top             =   360
            Width           =   990
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Cancel"
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
            Left            =   1140
            TabIndex        =   26
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1005
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   3735
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Left            =   3675
            TabIndex        =   1
            Top             =   705
            Width           =   120
         End
         Begin VB.CommandButton cmdLookup 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2985
            Picture         =   "frmPosDataFormCredit.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   150
            Width           =   315
         End
         Begin VB.TextBox txttransno 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1770
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "0000000000"
            Top             =   165
            Width           =   1185
         End
         Begin Crystal.CrystalReport rptVoucher 
            Left            =   -120
            Top             =   825
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
         Begin VB.Label dtptransdate 
            Caption         =   "Date"
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
            Left            =   1785
            TabIndex        =   25
            Top             =   615
            Width           =   2025
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
            Left            =   240
            TabIndex        =   20
            Top             =   615
            Width           =   1425
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
            Left            =   465
            TabIndex        =   19
            Top             =   165
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1005
         Left            =   5760
         TabIndex        =   17
         Top             =   0
         Width           =   2340
         Begin VB.CommandButton Command1 
            Caption         =   "&Hold"
            Enabled         =   0   'False
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
            Left            =   180
            TabIndex        =   3
            Top             =   360
            Width           =   990
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Restore"
            Enabled         =   0   'False
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
            Left            =   1230
            TabIndex        =   4
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1005
         Left            =   8070
         TabIndex        =   14
         Top             =   0
         Width           =   3390
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Casher :"
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
            Left            =   15
            TabIndex        =   16
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
            Left            =   840
            TabIndex        =   15
            Top             =   420
            Width           =   2280
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
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5745
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3915
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   255
         Visible         =   0   'False
         Width           =   420
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
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5745
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
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   9555
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5310
         Width           =   1635
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3510
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   225
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Frame Frame5 
         Height          =   4275
         Left            =   15
         TabIndex        =   21
         Top             =   975
         Width           =   11385
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
            Height          =   3915
            Left            =   90
            TabIndex        =   0
            Top             =   195
            Width           =   11190
            _ExtentX        =   19738
            _ExtentY        =   6906
            _Version        =   393216
            BackColor       =   16777215
            RowHeightMin    =   300
            BackColorSel    =   16777215
            ForeColorSel    =   0
            GridColor       =   -2147483632
            FocusRect       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
      Begin VB.Label Label5 
         Caption         =   "Misc Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8250
         TabIndex        =   41
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "&Disc % :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8775
         TabIndex        =   38
         Top             =   5790
         Width           =   960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "&Discount Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   2670
      End
      Begin VB.Label Label7 
         Caption         =   " Total Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8160
         TabIndex        =   36
         Top             =   7065
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Disc &Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8250
         TabIndex        =   34
         Top             =   6195
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   60
         TabIndex        =   31
         Top             =   5805
         Width           =   1035
      End
      Begin VB.Label Label11 
         Caption         =   " Total Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8175
         TabIndex        =   9
         Top             =   5340
         Width           =   1455
      End
      Begin VB.Label txtstatus 
         Height          =   165
         Left            =   4230
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   705
      End
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
      TabIndex        =   5
      Top             =   9960
      Width           =   1935
   End
   Begin VB.Menu MNUFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNUEXIT 
         Caption         =   "EXIT"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu EDIT_MENU 
      Caption         =   "EDIT"
      Begin VB.Menu COPY_MENU 
         Caption         =   "COPY"
         Shortcut        =   ^C
      End
      Begin VB.Menu PASTE_MENU 
         Caption         =   "PASTE"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu MNUDATA 
      Caption         =   "DATA"
      Begin VB.Menu MNUNEWINVOICE 
         Caption         =   "NEW BILL"
         Shortcut        =   ^P
      End
      Begin VB.Menu EDIT_EXISTING_BILL 
         Caption         =   "EDIT EXISTING BILL"
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu INSERT_ROW 
         Caption         =   "INSERT ROW"
         Shortcut        =   {F2}
      End
      Begin VB.Menu DELETE_ROW 
         Caption         =   "DELETE ROW"
         Shortcut        =   {F3}
      End
      Begin VB.Menu SAVE_RECORD 
         Caption         =   "SAVE RECORD"
         Shortcut        =   {F12}
      End
      Begin VB.Menu GOTO_QTY 
         Caption         =   "GOTO QTY"
         Shortcut        =   {F7}
      End
      Begin VB.Menu GOTO_RATE 
         Caption         =   "GOTO RATE"
         Shortcut        =   {F8}
      End
      Begin VB.Menu re_print 
         Caption         =   "RE-PRINT"
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmSO_PosformCredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pr_dumy As New Recordset
Dim PR_IcItem As New Recordset
Dim ls_sql As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim ln_gotostatus As Integer
Dim LN_TabGo As Integer
Dim ln_dicamount As Double
Dim ln_salestatus As Integer
Dim ls_directprint As Boolean
Public ln_printerCopy As Integer


Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Custom Code|<Description|<QTY|<Sale Price|<Amount|<Category|<U.O.M|<Discperc|<DiscAmount|<Itemcode|<EmpDiscAmount|<SaleRate|<SaleAmount"
        .ColWidth(1) = 2000
        .ColWidth(2) = 3300
        .ColWidth(3) = 900
        .ColAlignment(3) = 7

        .ColWidth(4) = 1200
        .ColAlignment(4) = 7
        
        .ColWidth(5) = 1500
        .ColAlignment(5) = 7

        .ColWidth(6) = 2000
        .ColWidth(7) = 1500
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .Redraw = True
        .Row = 1
        
        .CellBackColor = vbHighlight
    End With
End Sub
Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  
  With argFlexGrid
        ' .SelectionMode = flexSelectionByRow
        .Row = .RowSel
    If .Col = 1 Then
        .CellBackColor = vbWindowBackground
       If .TextMatrix(.Row, 1) <> "" Then
          If PR_IcItem.State = 1 Then PR_IcItem.Close
          ls_sql = "SELECT IC_Item.CustomCode,IC_Item.ItemCode,IC_Item.avgrate1,IC_Item.purchasecost, IC_Item.Description, IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc,IC_ItemCategory.EmpDiscPer"
          ls_sql = ls_sql & " FROM IC_Item INNER JOIN IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode"
          ls_sql = ls_sql & " where IC_Item.compcode = '" & Gs_compcode & "' and IC_Item.Customcode  = '" & .TextMatrix(.Row, 1) & "'"

          PR_IcItem.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox("Item Code not found !!!", vbCritical)
             .TextMatrix(.Row, 1) = ""
             
          Else
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 0) = .Row
               .TextMatrix(.Row, 1) = Trim(PR_IcItem("Customcode") & "")
               .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
                txtitemdesc = .TextMatrix(.Row, 2)
               .TextMatrix(.Row, 3) = 1
               If chkPRate.Value = 1 Then
               .TextMatrix(.Row, 4) = Val(0 & PR_IcItem("Avgrate1"))
               .TextMatrix(.Row, 12) = Val(0 & PR_IcItem("Salecost"))
               
               ElseIf chkprate1.Value = 1 Then
               .TextMatrix(.Row, 4) = Val(0 & PR_IcItem("Purchasecost"))
               .TextMatrix(.Row, 12) = Val(0 & PR_IcItem("Salecost"))
               Else
               .TextMatrix(.Row, 4) = Val(0 & PR_IcItem("Salecost"))
               .TextMatrix(.Row, 12) = Val(0 & PR_IcItem("Salecost"))
               End If
               
               .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 4)) * .TextMatrix(.Row, 3)
               .TextMatrix(.Row, 6) = Trim(PR_IcItem("CatDesc") & "")
               .TextMatrix(.Row, 7) = Trim(PR_IcItem("UOM") & "")
               .TextMatrix(.Row, 8) = Val(0 & PR_IcItem("SaleDiscPerc"))
               .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_IcItem("SaleDiscPerc")) / 100
               .TextMatrix(.Row, 10) = Trim(PR_IcItem("Itemcode") & "")
               .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(0 & PR_IcItem("EmpDiscPer")) / 100
               
               .TextMatrix(.Row, 13) = Val(.TextMatrix(.Row, 12)) * .TextMatrix(.Row, 3)
               
                .Col = 3
                .CellBackColor = vbHighlight
                TotalAmount
                
          End If
         PR_IcItem.Close
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
       
       ElseIf .Col = 3 Then
       .CellBackColor = vbWindowBackground
        If .TextMatrix(.Row, 3) <> "" Then
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
            
          
        Else
         Call MsgBox("Enter Qty!!!", vbCritical)
         .Row = .Row
         .Col = 3
         
        End If
      ElseIf .Col = 4 Then
       .CellBackColor = vbWindowBackground
        LN_TabGo = 0
        Insert_row_Click
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
            
   End If
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If .Col = 1 Or .Col = 3 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
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
      ElseIf .Col = 3 Or .Col = 4 Then
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
        .TextMatrix(.Row, 13) = Val(.TextMatrix(.Row, 12)) * .TextMatrix(.Row, 3)
        End If
        
        
        TotalAmount
       
    End With
  End If
End Sub

Private Sub Check1_GotFocus()
 On Error GoTo 0
 Call GetKeysAdd(GrdGRN, 13)
If LN_TabGo = 1 Then
 GrdGRN.Col = 4
 GrdGRN.Row = GrdGRN.Row - 1
 GrdGRN.CellBackColor = vbHighlight
 
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
 

End Sub

Private Sub chkPRate_Click()
If chkPRate.Value = 1 Then
chkprate1.Value = 0
End If
End Sub
Private Sub chkPRate1_Click()
If chkprate1.Value = 1 Then
chkPRate.Value = 1
End If
End Sub


Private Sub cmdLookup_Click()
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txtTransNo
        Set PO_DESC = Text1
        
        
        Gs_SQL = "SELECT SO_CreditSaleMaster.TransCode 'SO_CreditSaleMaster.TransCode', SO_CreditSaleMaster.TransDate, IC_Clients.Description 'IC_Clients.Description', SO_TransReturn.CustomCode 'SO_TransReturn.CustomCode', IC_Item.Description 'IC_Item.Description', SO_TransReturnMaster.NetAmount 'SO_TransReturnMaster.NetAmount'"
        Gs_SQL = Gs_SQL & " FROM SO_CreditSaleMaster INNER JOIN   SO_CreditSaleTrans ON SO_CreditSaleMaster.Compcode = SO_CreditSaleTrans.Compcode AND SO_CreditSaleMaster.TransCode = SO_CreditSaleTrans.TransCode INNER JOIN"
        Gs_SQL = Gs_SQL & " IC_Item ON SO_CreditSaleTrans.Compcode = IC_Item.Compcode AND SO_CreditSaleTrans.ItemCode = IC_Item.ItemCode INNER JOIN   IC_Clients ON SO_CreditSaleMaster.Compcode = IC_Clients.Compcode AND SO_CreditSaleMaster.AccountCode = IC_Clients.ClientCode"
        
        Gs_OrderBy = "ORDER BY SO_CreditSaleMaster.TransCode"
        
        
        Gs_OtherPara = " Where SO_CreditSaleMaster.compcode = '" & Gs_compcode & "' "
        
        frmSosearchRecords.Caption = "Invoices"
        frmSosearchRecords.Show 1
        
        If txtTransNo <> "" Then Call txttransno_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command3_Click()
Dim ls_clientcode As String

Dim ln_cnt
If Val(txttotalamount) > 0 Then
    ls_clientcode = txtCreditCode
    If ls_clientcode = "000019" Then
        ln_salestatus = 0
    Else
        ln_salestatus = 1
    End If


If ls_directprint = True Then
txtTransNo = txtTransNo
gc_dbcon.Execute "delete from SO_CreditSaleMaster where compcode = '" & Gs_compcode & "' and transcode = '" & txtTransNo & "' "
gc_dbcon.Execute "delete from SO_CreditSaleTrans where compcode = '" & Gs_compcode & "' and transcode = '" & txtTransNo & "' "
ls_directprint = False
Else
txtTransNo = maxtranscode
End If
ls_sql = "Insert into SO_CreditSaleMaster(Compcode, TransCode, TransDate, AccountCode, Remarks, TotalAmount, DiscPer, DiscAmount, NetAmount, RecAmount, BalAmount,usercode,compname,SaleStatus,estatus,miscamount,miscper)"
ls_sql = ls_sql & " Values ('" & Gs_compcode & "' , '" & txtTransNo & "', '" & Format(dtptransdate, "YYYY/MM/DD HH:MM:SS") & "', '" & ls_clientcode & "' , 'Credit Sale' , " & Val(txttotalamount) & ", " & Val(txtdiscper) & ", " & Val(txtdiscamt) & ", " & Val(txtnetamount) & ", " & Val(txtnetamount) & ", 0," & Gn_UserCode & ",'" & Gs_ComputerName & "'," & ln_salestatus & ",1," & Val(txtmiscamount) & "," & Val(txtmiscper) & ")"
gc_dbcon.Execute ls_sql

 With GrdGRN
       For ln_cnt = 1 To .Rows - 1
       If .TextMatrix(ln_cnt, 1) <> "" Then
        ln_dicamount = ((Val(.TextMatrix(ln_cnt, 5)) / Val(txttotalamount)) * Val(txtdiscamt))
        ls_sql = "INSERT into SO_CreditSaleTrans(Compcode, TransCode,customcode, ItemCode, Quantity,Itemrate,Amount,discamount,salerate,saleamount)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 10)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 3))) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & "," & Val(ln_dicamount) & "," & Val(.TextMatrix(ln_cnt, 12)) & "," & Val(.TextMatrix(ln_cnt, 13)) & ")"
        gc_dbcon.Execute ls_sql
      End If
      Next
  End With
  frmSoPrinterCopy.Show 1
  If ln_printerCopy > 0 Then
  Printinvoice
  End If
  GrdGRN.SetFocus
  GrdGRN.Col = 1
  GrdGRN.Row = 1
  
Else
Call MsgBox("Nothing for save !!!", vbExclamation)
 
End If
dtptransdate = Gd_SysDate
txtTransNo = maxtranscode
InitializeGrid
txtdiscamt = ""
txtnetamount = ""
txtdiscper = ""
txtmiscamount = ""
txtCreditCode = "000019"
Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)
GrdGRN.SetFocus
GrdGRN.Row = 1
txttotalamount = ""
End Sub
Private Sub Printinvoice()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "SaleInvoiceCredit.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        '.Formulas(2) = "Reportname = 'Good Receive Note'"
        .CopiesToPrinter = ln_printerCopy
        
        
        .SQLQuery = "SELECT SO_TransMaster.TransCode, SO_TransMaster.TransDate, SO_TransMaster.DiscAmount, SO_TransMaster.CompName, SO_TransMaster.MiscAmount,"
        .SQLQuery = .SQLQuery & " SO_Trans.ItemCode, SO_Trans.SRNo, SO_Trans.Quantity, SO_Trans.ItemRate, SO_Trans.Amount, IC_Clients.Description, SyUsers.UserName,"
        .SQLQuery = .SQLQuery & " IC_Item.Description AS Expr1, IC_Item.PriceDescCStatus FROM SO_CreditSaleMaster SO_TransMaster LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " IC_Clients IC_Clients ON SO_TransMaster.Compcode = IC_Clients.Compcode AND"
        .SQLQuery = .SQLQuery & " SO_TransMaster.AccountCode = IC_Clients.ClientCode LEFT OUTER JOIN  SyUsers SyUsers ON SO_TransMaster.Compcode = SyUsers.CompCode AND SO_TransMaster.UserCode = SyUsers.UserCode LEFT OUTER JOIN"
        .SQLQuery = .SQLQuery & " SO_CreditSaleTrans SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND"
        .SQLQuery = .SQLQuery & "  SO_TransMaster.TransCode = SO_Trans.TransCode LEFT OUTER JOIN IC_Item IC_Item ON SO_Trans.Compcode = IC_Item.Compcode AND SO_Trans.ItemCode = IC_Item.ItemCode"
        .SQLQuery = .SQLQuery & " where SO_TransMaster.compcode = '" & Gs_compcode & "' "
        If Trim(txtTransNo) <> "" Then
            .SQLQuery = .SQLQuery & " and  SO_TransMaster.transcode = '" & Trim(txtTransNo) & "'"
        End If
        .SQLQuery = .SQLQuery & "  ORDER BY SO_TransMaster.TransCode"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
End Sub

Private Sub Command4_Click()
dtptransdate = Gd_SysDate
txtTransNo = maxtranscode
InitializeGrid
txtdiscamt = ""
txtnetamount = ""
txtdiscper = ""
txtmiscamount = ""
txtCreditCode = "000019"
Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)
GrdGRN.SetFocus
GrdGRN.Row = 1
txttotalamount = ""
End Sub
Private Sub COPY_MENU_Click()
With GrdGRN
Clipboard.Clear
Clipboard.SetText .TextMatrix(.Row, .Col)
End With
End Sub
Private Sub Delete_row_Click()
  With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
    End With
End Sub

Private Sub EDIT_EXISTING_BILL_Click()
txtTransNo.Enabled = True
cmdLookup.Enabled = True
txtTransNo = ""
txtTransNo.SetFocus
cmdLookup_Click
End Sub

Private Sub goto_qty_Click()
With GrdGRN
.CellBackColor = vbWindowBackground
.Col = 3
If .TextMatrix(.Row, .Col) = "" Then
.Col = 3
.Row = .Row - 1
ElseIf .TextMatrix(.Row, .Col) <> "" And .Row > 1 Then
.Row = .Row - 1
End If
.CellBackColor = vbHighlight
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
'On Error GoTo ErrHandler
'If ls_directprint = True Then
'    Call MsgBox("Can enter data when it is in print mode [Press CTRL+P for printing ]", vbCritical)
'    Exit Sub
'End If
Call GetKeysAdd(GrdGRN, KeyAscii)
Exit Sub

'ErrHandler:
'MsgBox ("An Error has Occured In The MSFlexgrid1_KeyPress() Procedure") & vbCr & "Report This Error To Latifjat@hotmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
End Sub

Private Sub GrdGRN_Click()
With GrdGRN
.CellBackColor = vbHighlight
 txtitemdesc = .TextMatrix(.Row, 2)
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
GrdGRN.CellBackColor = vbHighlight
End Sub


Private Sub GrdGRN_LeaveCell()
With GrdGRN
 .CellBackColor = vbWindowBackground
End With
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And GrdGRN.Col = 1 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    If GrdGRN.TextMatrix(GrdGRN.Row, 1) <> "" Then
        Call GrdGRN_KeyPress(13)
    End If
 ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                InitializeGrid
            End If
    End With
 ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then  'Delete Key Pressed
    With GrdGRN
    txtitemdesc = .TextMatrix(.Row, 2)
    End With
 End If

    
End Sub


Private Sub Command1_Click()
If GrdGRN.Rows > 2 Then
Dim ls_transcodehold As String
Dim ln_cnt As Integer
ls_transcodehold = maxtranscodehold
 With GrdGRN
       For ln_cnt = 1 To .Rows - 1
       If .TextMatrix(ln_cnt, 1) <> "" Then
        ls_sql = "INSERT into SO_TransHold(Compcode, TransCode,transdate,customcode, ItemCode, Quantity,Itemrate,Amount)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & ls_transcodehold & "', '" & Format(dtptransdate, "YYYY/MM/DD HH:MM:SS") & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 10)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 3))) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & ")"
        gc_dbcon.Execute ls_sql
      End If
      Next
  End With

dtptransdate = Now
txtTransNo = maxtranscode
InitializeGrid
GrdGRN.Row = 1
Else
    Call MsgBox("Nothing for Hold!!!", vbCritical)
End If
End Sub

Private Sub Command2_Click()
  ls_sql = "Select * from so_Transhold where compcode = '" & Gs_compcode & "'"
  pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
  If pr_dumy.EOF Then
  Call MsgBox("Nothing for restore !!!", vbCritical)
  pr_dumy.Close
  Exit Sub
  End If
  
  pr_dumy.Close
  
  Set PO_AnyForm = Nothing
  Set PO_AnyForm = Me
  Set PO_CODE = Text1
  Set PO_DESC = Text2
  Gs_SQL = "Select TransCode, TransDate from So_Transhold "
  Gs_FindFld = "Transcode"
  Gs_OrderBy = "Order by Transdate"
  Gs_OtherPara = " where Compcode ='" & Gs_compcode & "' group by Transcode,Transdate "
  MyLookupOLDB.Caption = "Hold Trans"
  MyLookupOLDB.Show 1
  InitializeGrid
  
  If Text1 <> "" Then
   ls_sql = "SELECT IC_Item.CustomCode, IC_Item.ItemCode, IC_Item.Description, SO_TransHold.Quantity, SO_TransHold.ItemRate, SO_TransHold.Amount,"
    ls_sql = ls_sql & " IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc,IC_ItemCategory.empdiscper FROM IC_Item INNER JOIN"
    ls_sql = ls_sql & " IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN"
    ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode INNER JOIN"
    ls_sql = ls_sql & " SO_TransHold ON IC_Item.Compcode = SO_TransHold.Compcode AND IC_Item.ItemCode = SO_TransHold.ItemCode"
    ls_sql = ls_sql & " WHERE (SO_TransHold.Compcode = '" & Gs_compcode & "') AND (SO_TransHold.TransCode = '" & Text1 & "')"
  End If

With GrdGRN
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
               .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(0 & pr_dumy("EmpDiscPer")) / 100
        If .Row = .Rows - 1 Then
        .Col = 1
        .Row = .Rows - 1
        .Rows = .Rows + 1
        .Row = .Rows - 1
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

ls_sql = "Delete from So_TransHold where Transcode = '" & Text1 & "' and compcode = '" & Gs_compcode & "'"
gc_dbcon.Execute ls_sql
dtptransdate = Now
txtTransNo = maxtranscode

End Sub

Private Sub Form_Click()
'InitializeGrid
'dtptransdate = Date
'TXTBARCODE.Visible = False
End Sub

Private Sub Insert_row_Click()
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
         Call MsgBox("Enter/Select Item Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
End With
End Sub

Private Sub mnuNewinvoice_Click()
If Val(txttotalamount) > 0 Then
    Command3_Click
End If
If txtstatus <> "Cancel" Then
txtTransNo.Enabled = False
cmdLookup.Enabled = False
dtptransdate = Gd_SysDate
txtTransNo = maxtranscode
InitializeGrid
txttotalamount = ""

End If
End Sub


Private Sub TotalAmount()
    Dim ln_cnt As Integer
      txttotalamount = ""
      txtdiscamount = ""
      txtempDisc = ""
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Round(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 5)))
            'txtdiscamount = Format(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), "######0.00")
            'txtempDisc = Format(Val(txtempDisc) + Val(.TextMatrix(ln_cnt, 11)), "######0.00")
        Next
    End With
    txtnetamount = (Val(txttotalamount) + Val(txtmiscamount)) - Val(txtdiscamt)
    
End Sub

Private Sub Form_Load()
InitializeGrid
dtptransdate = Gd_SysDate
lblcasherName = Gc_UserName
Call mnuNewinvoice_Click
txtCreditCode = "000019"
Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)
LN_TabGo = 0
ls_directprint = False
End Sub
Public Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from  SO_CreditSaleMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function
Public Function maxtranscodehold() As String
pr_dumy.Open "select max(transcode) as transcode from SO_Transhold where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscodehold = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscodehold = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function

Private Sub PASTE_MENU_Click()
With GrdGRN
.TextMatrix(.Row, .Col) = Clipboard.GetText
End With
End Sub

Private Sub RE_PRINT_Click()
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txtTransNo
        Set PO_DESC = Text1
        
        
        Gs_SQL = "SELECT Invoices.TransCode AS InvoiceNo, Invoices.TransDate AS Invoicedate, Customer.Description AS 'Customer.Description',"
        Gs_SQL = Gs_SQL & " Invoices.NetAmount AS 'Invoices.NetAmount', SyUsers.UserName fROM SO_CreditSaleMaster Invoices INNER JOIN"
        Gs_SQL = Gs_SQL & " IC_Clients Customer ON Invoices.Compcode = Customer.Compcode AND Invoices.AccountCode = Customer.ClientCode LEFT OUTER JOIN"
        Gs_SQL = Gs_SQL & " SyUsers ON Invoices.Compcode = SyUsers.CompCode AND Invoices.UserCode = SyUsers.UserCode"
        
        Gs_OrderBy = "ORDER BY Invoices.TransCode Desc"
                
        Gs_OtherPara = " Where Invoices.compcode = '" & Gs_compcode & "'"
        
        frmSosearchRecords.Caption = "Invoices"
        frmSosearchRecords.Show 1
        ls_directprint = True
        If txtTransNo <> "" Then Call txttransno_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Save_Record_Click()
Command3_Click
End Sub
Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCreditCode
    Set PO_DESC = txtCreditDesc
    Gs_SQL = "Select Clientcode, Description from IC_clients "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Credit Clients"
    MyLookupOLDB.Show 1
    
    If txtCreditCode <> "" Then Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub


Private Sub txtmiscamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtCreditCode.SetFocus
End If
End Sub

Private Sub txtCreditCode_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo LocalErr
If Trim(txtCreditCode) <> "" And KeyCode = vbKeyReturn Then
        txtCreditCode = DoPad(txtCreditCode, 6)
        pr_dumy.Open "Select * from IC_clients where Compcode  = '" & Gs_compcode & "' and clientcode= '" & txtCreditCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Client Code not found !!!", vbCritical)
            txtCreditCode = ""
            txtCreditCode = ""
            txtCreditCode.SetFocus
        Else
            txtCreditDesc = pr_dumy("Description")
        End If
        pr_dumy.Close

ElseIf Trim(txtCreditCode) = "" And KeyCode = vbKeyReturn Then
        txtCreditCode = ""
        txtCreditDesc = ""
        Call Command5_Click
End If

Exit Sub
LocalErr:
Call MsgBox(Err.Description, vbCritical)

End Sub

Private Sub txtdiscamt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtmiscper.SetFocus
End If
End Sub

Private Sub txtdiscamt_LostFocus()
If txtdiscamt <> "" Then
    txtdiscamount = Val(txttotalamount) * Val(txtdiscper) / 100
    txtnetamount = Val(txttotalamount) - Val(txtdiscamt)
End If
End Sub

Private Sub txtdiscper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtdiscamt.SetFocus
End If
End Sub

Private Sub txtdiscper_LostFocus()
If txtdiscper <> "" Then
   txtdiscamt = Val(txttotalamount) * Val(txtdiscper) / 100
   txtnetamount = Val(txttotalamount) - Val(txtdiscamt)
End If
End Sub

Private Sub txtmiscamount_LostFocus()
If txtmiscamount <> "" Then
txtnetamount = Val(txttotalamount) + Val(txtmiscamount)
End If
End Sub

Private Sub txtmiscper_Change()
If txtmiscper <> "" Then
txtmiscamount = Round(((Val(txttotalamount) * Val(txtmiscper)) / 100), 0)
txtnetamount = Val(txttotalamount) + Val(txtmiscamount)
End If
End Sub

Private Sub txtmiscper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtmiscamount.SetFocus
End If
End Sub

Private Sub txttransno_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txtTransNo.Text) > 0 Then
         If pr_dumy.State = 1 Then pr_dumy.Close
         txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
         pr_dumy.Open "select * from SO_CreditSaleMaster where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         If pr_dumy.EOF Then
                   Call MsgBox("Record not found !!!", vbExclamation)
                   If txtTransNo.Enabled Then txtTransNo.SetFocus
         Else
                   dtptransdate = pr_dumy("Transdate")
                   txtdiscper = Val(0 & pr_dumy("Discper"))
                   txtdiscamt = Val(0 & pr_dumy("DiscAmount"))
                   txtmiscamount = Val(0 & pr_dumy("MiscAmount"))
                   txtmiscper = Val(0 & pr_dumy("Miscper"))
                   InitializeGrid
                   GrdGRN.CellBackColor = vbWindowBackground
                   LoadGRNTrans
                   
         End If
End If
End Sub
Private Sub LoadGRNTrans()
ls_sql = "Select * from  SO_CreditSaleMaster where compcode = '" & Gs_compcode & "' and transcode = '" & txtTransNo & "'"
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
txtCreditCode = Trim(pr_dumy("AccountCode") & "")

txtdiscamt = Val(0 & pr_dumy("Discamount"))
txtdiscper = Val(0 & pr_dumy("Discper"))
End If
pr_dumy.Close
Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)


ls_sql = "SELECT IC_Item.CustomCode, IC_Item.ItemCode, IC_Item.Description, SO_CreditSaleTrans.salerate,SO_CreditSaleTrans.saleamount,SO_CreditSaleTrans.Quantity, SO_CreditSaleTrans.ItemRate, SO_CreditSaleTrans.Amount,"
ls_sql = ls_sql & " IC_Item.SaleDiscPerc, IC_Item.SaleCost, IC_ItemUM.Description AS UOM, IC_ItemCategory.Description AS CatDesc FROM IC_Item INNER JOIN"
ls_sql = ls_sql & " IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode INNER JOIN"
ls_sql = ls_sql & " IC_ItemCategory ON IC_Item.Compcode = IC_ItemCategory.Compcode AND IC_Item.CatCode = IC_ItemCategory.CatCode INNER JOIN"
ls_sql = ls_sql & " SO_CreditSaleTrans ON IC_Item.Compcode = SO_CreditSaleTrans.Compcode AND IC_Item.ItemCode = SO_CreditSaleTrans.ItemCode"
ls_sql = ls_sql & " WHERE (SO_CreditSaleTrans.Compcode = '" & Gs_compcode & "') AND (SO_CreditSaleTrans.TransCode = '" & txtTransNo & "')"

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
               .TextMatrix(.Row, 12) = Val(0 & pr_dumy("SaleRate"))
               .TextMatrix(.Row, 13) = Val(0 & pr_dumy("SaleAmount"))
               
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
       'txtnetamount = Val(txttotalamount) - Val(txtdiscamt)
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

