VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmadjustment 
   Caption         =   "Sale Return (Adjustment)"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   15
      TabIndex        =   1
      Top             =   570
      Width           =   7530
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   6135
         TabIndex        =   21
         Top             =   165
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Format          =   56295425
         CurrentDate     =   37580
      End
      Begin VB.TextBox txtpartydesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2430
         MaxLength       =   64
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   4920
      End
      Begin VB.TextBox txtpartycode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   540
         Width           =   660
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   2100
         Picture         =   "frmAdjustment.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3060
         MaxLength       =   50
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   180
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   510
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   900
         Width           =   5925
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2520
         Picture         =   "frmAdjustment.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtTransNo 
         BackColor       =   &H00FFFF00&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   180
         Width           =   1095
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   7140
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Party Code :"
         Height          =   255
         Left            =   105
         TabIndex        =   32
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   900
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference #  :"
         Height          =   255
         Left            =   105
         TabIndex        =   13
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label label2 
         Caption         =   "Adjustment Date :"
         Height          =   255
         Left            =   4785
         TabIndex        =   12
         ToolTipText     =   "Enter Value Date"
         Top             =   195
         Width           =   1410
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New"
            Description     =   "Add"
            Object.ToolTipText     =   "Add new record"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Edit"
            Description     =   "Edit"
            Object.ToolTipText     =   "Edit an existing record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Delete"
            Description     =   "Remove "
            Object.ToolTipText     =   "Remove an existing record."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save"
            Description     =   "Save a new Record"
            Object.ToolTipText     =   "Save on disk"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Slip"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            Description     =   "Find a Record."
            Object.ToolTipText     =   "Find a record."
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancel"
            Description     =   "Cancel Operation"
            Object.ToolTipText     =   "Cancel operation mode"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   14
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4920
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAdjustment.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3225
      Left            =   0
      TabIndex        =   2
      Top             =   1980
      Width           =   7530
      Begin VB.TextBox txtmtype 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1185
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   405
         Width           =   1155
      End
      Begin VB.TextBox txttaxreceived 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   6135
         MaxLength       =   11
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   2820
         Width           =   1260
      End
      Begin VB.TextBox txtamountreceived 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   3420
         MaxLength       =   11
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   2820
         Width           =   1320
      End
      Begin VB.TextBox txtnoofitems 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Frame Frame3 
         Height          =   1785
         Left            =   120
         TabIndex        =   36
         Top             =   4125
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox txtvchrno 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6090
            MaxLength       =   64
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   165
            Width           =   1275
         End
         Begin VB.CommandButton Command6 
            Height          =   315
            Left            =   3045
            Picture         =   "frmAdjustment.frx":273A
            Style           =   1  'Graphical
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox txtpartyaccount 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   13
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1365
            Width           =   1590
         End
         Begin VB.TextBox txtpartyactdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3375
            MaxLength       =   64
            TabIndex        =   49
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1350
            Width           =   3990
         End
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   3045
            Picture         =   "frmAdjustment.frx":28AC
            Style           =   1  'Graphical
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   975
            Width           =   315
         End
         Begin VB.TextBox txttaxaccount 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   13
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   990
            Width           =   1590
         End
         Begin VB.TextBox txttaxactdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3375
            MaxLength       =   64
            TabIndex        =   45
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   975
            Width           =   3990
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   3045
            Picture         =   "frmAdjustment.frx":2A1E
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   600
            Width           =   315
         End
         Begin VB.TextBox txtsaleaccount 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   13
            TabIndex        =   42
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   615
            Width           =   1590
         End
         Begin VB.TextBox txtsaleactdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3375
            MaxLength       =   64
            TabIndex        =   41
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   600
            Width           =   3990
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2025
            Picture         =   "frmAdjustment.frx":2B90
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox txtvchrtype 
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "XXX"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   38
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   195
            Width           =   660
         End
         Begin VB.TextBox txtVchrDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2355
            MaxLength       =   64
            TabIndex        =   37
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Top             =   180
            Width           =   2355
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Voucher No :"
            Height          =   255
            Left            =   4785
            TabIndex        =   53
            Top             =   210
            Width           =   1275
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Party Account # :"
            Height          =   255
            Left            =   60
            TabIndex        =   52
            Top             =   1410
            Width           =   1305
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Tax Account # :"
            Height          =   255
            Left            =   90
            TabIndex        =   48
            Top             =   1005
            Width           =   1275
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Sale Account # :"
            Height          =   255
            Left            =   90
            TabIndex        =   44
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Voucher Type :"
            Height          =   255
            Left            =   90
            TabIndex        =   40
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.TextBox txttotaltaxamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   6135
         MaxLength       =   11
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   2445
         Width           =   1260
      End
      Begin VB.CheckBox ChkPaidAmt 
         Alignment       =   1  'Right Justify
         Caption         =   "Post Gl Voucher :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   2460
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox txttaxamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   6405
         MaxLength       =   11
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   405
         Width           =   1020
      End
      Begin VB.TextBox txtamount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4455
         MaxLength       =   11
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   405
         Width           =   1155
      End
      Begin VB.TextBox txttax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   5640
         MaxLength       =   11
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   405
         Width           =   540
      End
      Begin VB.TextBox txtunitprice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   3345
         MaxLength       =   11
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   405
         Width           =   1080
      End
      Begin VB.TextBox txtItemDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   150
         MaxLength       =   64
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   765
         Width           =   7275
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
         Height          =   315
         Left            =   3420
         MaxLength       =   11
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   2445
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   840
         Picture         =   "frmAdjustment.frx":2D02
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   405
         Width           =   315
      End
      Begin VB.TextBox TxtItemCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   150
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Item Code"
         Top             =   405
         Width           =   675
      End
      Begin VB.TextBox txtqty 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0;(#,##0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   2385
         MaxLength       =   11
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Issue Quantity"
         Top             =   405
         Width           =   900
      End
      Begin MSFlexGridLib.MSFlexGrid GrdGRN 
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   1095
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax Paid :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3555
         TabIndex        =   60
         Top             =   2865
         Width           =   2490
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount Paid  Party :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   840
         TabIndex        =   58
         Top             =   2865
         Width           =   2490
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Units :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1155
         TabIndex        =   55
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Sale Tax Total :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4845
         TabIndex        =   34
         Top             =   2490
         Width           =   1290
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Tax :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5415
         TabIndex        =   31
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Amount :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4290
         TabIndex        =   30
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Quantity :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2295
         TabIndex        =   27
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Item Code :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "%"
         Height          =   315
         Left            =   6225
         TabIndex        =   25
         Top             =   435
         Width           =   150
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Tax Amount :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6270
         TabIndex        =   24
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label Label11 
         Caption         =   " Total :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2880
         TabIndex        =   15
         Top             =   2490
         Width           =   885
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Unit Price :"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3360
         TabIndex        =   11
         Top             =   180
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmadjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object


Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String


Dim pr_dumy As New Recordset


Dim Pr_ICParty As New Recordset
Dim PR_ICISSU As New Recordset
Dim PR_IcItem As New Recordset
Dim PR_VchCntr As New Recordset
Dim PR_VchType As New Recordset

Dim PR_Branch As New Recordset

Public ls_VchType As String
Public ls_VchNo   As String
Public ls_VchDesc As String
Dim ls_Invoiceno As String
Dim enterkeystatus As Boolean

Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from ic_trans where compcode = '" & Gs_compcode & "' and transtype = 'A' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransNo
    Set PO_DESC = Text1
    GoTop PR_ICISSU
    MyLookup.Caption = "Issue References "
    MyLookup.FillGrid PR_ICISSU, "Transcode", "SaleDate", 10
    MyLookup.Show 1
    
    If Len(txtTransNo) > 0 Then txtTransNo_KeyDown vbKeyReturn, vbKeyShift
End Sub



'Private Sub Command2_Click()
'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txtVchrType
'    Set PO_DESC = txtVchrDesc
'    GoTop PR_VchType
'    PR_VchType.Filter = "BranchCode = '" & Gs_BranchCode & "'"
'    MyLookup.Caption = "Voucher Types"
'    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
'    MyLookup.Show 1
'    PR_VchType.Filter = adFilterNone
'
'    If Len(txtVchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
'End Sub

'Private Sub Command3_Click()
'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txtsaleaccount
'    Set PO_DESC = txtsaleactdesc
'    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description' from gl_Detail"
'    Gs_FindFld = "Acct_Desc"
'    Gs_Subon = True
'    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
'    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
'    MyLookupOLDB.Caption = "Account Nos."
'    MyLookupOLDB.Show 1
'    If Len(txtsaleaccount) > 0 Then txtsaleaccount_KeyDown vbKeyReturn, vbKeyShift
'End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtItemCode
    Set PO_DESC = txtItemDesc
    GoTop PR_IcItem

    If PR_IcItem.EOF Then
       Call SetErr("No Item has been found.", vbCritical)
       Exit Sub
    End If
    
    MyLookup.Caption = "Items. "
    MyLookup.FillGrid PR_IcItem, "Itemcode", "Description", 8
    MyLookup.Show 1
    

    If Len(TxtItemCode) > 0 Then txtItemcode_KeyDown vbKeyReturn, vbKeyShift
End Sub


'Private Sub Command4_Click()
'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txttaxaccount
'    Set PO_DESC = txttaxactdesc
'    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description' from gl_Detail"
'    Gs_FindFld = "Acct_Desc"
'    Gs_Subon = True
'    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
'    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
'    MyLookupOLDB.Caption = "Account Nos."
'    MyLookupOLDB.Show 1
'    If Len(txttaxaccount) > 0 Then txttaxaccount_KeyDown vbKeyReturn, vbKeyShift
'
'End Sub

Private Sub Command5_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpartycode
    Set PO_DESC = txtpartydesc
    Gs_SQL = "Select  ClientCode 'Code' ,Description from IC_Clients"
    Gs_FindFld = "Description"
    Gs_Subon = False
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Description,ClientCode"
    MyLookupOLDB.Caption = "Customer"
    MyLookupOLDB.Show 1
    
    If Len(txtpartycode) > 0 Then txtpartycode_KeyDown vbKeyReturn, vbKeyShift
End Sub



'Private Sub Command6_Click()
'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txtpartyaccount
'    Set PO_DESC = txtpartydesc
'    Gs_SQL = "Select Accountno 'Account No', Acct_Desc  'Description' from gl_Detail"
'    Gs_FindFld = "Acct_Desc"
'    Gs_Subon = True
'    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
'    Gs_OrderBy = "Order by Acct_Desc,AccountNo"
'    MyLookupOLDB.Caption = "Account Nos."
'    MyLookupOLDB.Show 1
'    If Len(txtpartyaccount) > 0 Then txtpartyaccount_KeyDown vbKeyReturn, vbKeyShift
'
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_ICISSU, Me, txtTransNo, txtvaluedate, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()

 
  SetToolBar(1) = chkRights("ICISUSTP01")
  SetToolBar(2) = chkRights("ICISUSTP02")
  SetToolBar(3) = chkRights("ICISUSTP03")
  SetToolBar(4) = chkRights("ICISUSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
'  frmGroupCompanies.PR_SysComp.Filter = "Compcode ='" & Gs_compcode & "'"
'  txtsaleaccount = Trim(frmGroupCompanies.PR_SysComp("Saleaccount"))
'  txttaxaccount = Trim(frmGroupCompanies.PR_SysComp("taxaccount"))
'  frmGroupCompanies.PR_SysComp.Filter = adFilterNone
'
  PR_VchCntr.Open "SELECT * FROM Gl_VchrCntrs WHERE CompCode = '" & Gs_compcode & "' And VchrYear = " & Year(Gs_Fnperiod) & " ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_VchType.Open "SELECT *,BranchCode+VchrType As FindFld FROM GlVchrType WHERE CompCode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    
  
  Pr_ICParty.Open "Select * from Ic_clients where compcode = '" & Gs_compcode & "'  order by ClientCode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  
  PR_Branch.Open "Select * from SysBranch Where Compcode = '" & Gs_compcode & "' order by Branchcode", gc_dbcon, adOpenStatic, adLockReadOnly, 1

  PR_IcItem.Open "Select * from Ic_Item where compcode ='" & Gs_compcode & "' ", gc_dbcon, adOpenDynamic, adLockPessimistic, 1
  PR_ICISSU.Open "Select * from Ic_Trans where compcode ='" & Gs_compcode & "' and transtype = 'A'  order by Transcode", gc_dbcon, adOpenDynamic, adLockOptimistic
  
  PB_BlnkGRN = IIf(PR_ICISSU.EOF, True, False)
  txtvaluedate.Value = Date
  enterkeystatus = False
 
'  Call txtsaleaccount_KeyDown(vbKeyReturn, vbKeyShift)
'  Call txttaxaccount_KeyDown(vbKeyReturn, vbKeyShift)
  InitializeGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Pr_ICParty.Close
    PR_ICISSU.Close
    PR_IcItem.Close
    PR_Branch.Close
    PR_VchCntr.Close
    PR_VchType.Close
  
End Sub

Private Sub txtItemcode_Validate(Cancel As Boolean)
If Trim(TxtItemCode.Text) <> "" Then
         TxtItemCode.Text = DoPad(Trim(TxtItemCode.Text), 4)
         lb_found = MySeek(TxtItemCode.Text, "itemcode", PR_IcItem)
         
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtItemCode.SetFocus
             Cancel = True
         Else
             txtItemDesc = PR_IcItem("Description")
             txtmtype = PR_IcItem("mdesc")
             If Val(PR_IcItem("NoofItems")) > 1 Then
                txtunitprice = PR_IcItem("TotalPurchaseprice")
             Else
                txtunitprice = PR_IcItem("Purchaseprice")
             End If
             If txtqty.Enabled Then txtqty.SetFocus
         End If
 End If
End Sub



Private Sub txtpartycode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And txtpartycode.Text <> "" Then
         txtpartycode.Text = DoPad(txtpartycode.Text, txtpartycode.MaxLength)
         
        
         If Not MySeek(txtpartycode.Text, "ClientCode", Pr_ICParty) Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtpartycode.SetFocus
             txtpartydesc.Text = ""
         Else
             txtpartydesc.Text = Pr_ICParty("Description")
             'txtpartyaccount = Pr_ICParty("GLAccountNo")
             'enterkeystatus = False
             'Call txtpartyaccount_KeyDown(vbKeyReturn, vbKeyShift)
             If TxtRemarks.Enabled Then TxtRemarks.SetFocus
         End If
         
 ElseIf KeyCode = vbKeyF12 Then
    Command5_Click
 ElseIf KeyCode = vbKeyPageUp Then
    
 End If
End Sub
Private Sub txtpartycode_Change()
txtpartydesc = ""
End Sub

Private Sub txtItemcode_Change()
txtItemDesc = ""
End Sub




Private Sub txtpartycode_Validate(Cancel As Boolean)
 If Trim(txtpartycode.Text) <> "" Then
         txtpartycode.Text = DoPad(txtpartycode.Text, txtpartycode.MaxLength)
         
        
         If Not MySeek(txtpartycode.Text, "ClientCode", Pr_ICParty) Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtpartycode.SetFocus
             txtpartydesc.Text = ""
             Cancel = True
         Else
             txtpartydesc.Text = Pr_ICParty("Description")
             If TxtRemarks.Enabled Then TxtRemarks.SetFocus
         End If
End If
End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
        If txtitemtype = "Box" Then
         txtnoofitems = Val(txtqty) * PR_IcItem("noofitems")
        Else
         txtnoofitems = Val(txtqty) * 1
        End If
        txtunitprice.SetFocus
   End If
   If KeyCode = vbKeyPageUp Then TxtItemCode.SetFocus
End Sub

Private Sub txtqty_LostFocus()
If Val(txtqty) > 0 Then
         txtnoofitems = Val(txtqty) * PR_IcItem("noofitems")
        
        If Val(txtunitprice) > 0 Then
            txtamount = Val(txtqty) * Val(txtunitprice)
        Else
            txtunitprice.SetFocus
        End If
End If
End Sub


Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then TxtItemCode.SetFocus
If KeyCode = vbKeyPageUp Then txtpartycode.SetFocus
End Sub

'Private Sub txtsaleaccount_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    pr_Dumy.Open "select * from gl_detail where compcode = '" & Gs_compcode & "' and accountno = '" & txtsaleaccount & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'        If Not pr_Dumy.EOF Then
'            txtsaleactdesc = pr_Dumy("acct_desc")
'
'            If enterkeystatus = True Then
'                txttaxaccount.SetFocus
'
'            End If
'
'        Else
'            Call MsgBox("Record not found", vbCritical)
'            txtsaleaccount.SetFocus
'        End If
'        pr_Dumy.Close
'End If
'End Sub
'
'Private Sub txttaxaccount_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    pr_Dumy.Open "select * from gl_detail where compcode = '" & Gs_compcode & "' and accountno = '" & txttaxaccount & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'        If Not pr_Dumy.EOF Then
'            txttaxactdesc = pr_Dumy("acct_desc")
'
'            If enterkeystatus = True Then
'                txtpartyaccount.SetFocus
'
'            End If
'
'        Else
'            Call MsgBox("Record not found", vbCritical)
'            txttaxaccount.SetFocus
'        End If
'        pr_Dumy.Close
'End If
'End Sub
'
'Private Sub txtpartyaccount_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'    pr_Dumy.Open "select * from gl_detail where compcode = '" & Gs_compcode & "' and accountno = '" & txtpartyaccount & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
'        If Not pr_Dumy.EOF Then
'            txtpartyactdesc = pr_Dumy("acct_desc")
'
'            If enterkeystatus = True Then
'                txtpartyaccount.SetFocus
'
'            End If
'
'        Else
'            Call MsgBox("Record not found", vbCritical)
'            txtpartyaccount.SetFocus
'        End If
'        pr_Dumy.Close
'End If
'End Sub
'
Private Sub txttax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txttaxamount = Val(txtamount) * Val(txttax) / 100
    If checkvalidate Then AddToGrid
End If
End Sub
Function checkvalidate() As Boolean
If Trim(TxtItemCode) = "" Then
    Call MsgBox("Please enter Item Code", vbCritical)
    TxtItemCode.SetFocus
    checkvalidate = False
ElseIf Val(txtqty) = 0 Then
    Call MsgBox("Please enter qty", vbCritical)
    txtqty.SetFocus
    checkvalidate = False
ElseIf Val(txtunitprice) = 0 Then
    Call MsgBox("Please enter unit price", vbCritical)
    txtunitprice.SetFocus
    checkvalidate = False
Else
    checkvalidate = True
End If
End Function
Private Sub txttotalamount_Change()
txtamountreceived = txttotalamount
End Sub

Private Sub txttotaltaxamount_Change()
txttaxreceived = txttotaltaxamount
End Sub

Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtTransNo.Text) > 0 Then
         
         txtTransNo.Text = IIf(IsNumeric(LTrim(str(txtTransNo.Text))), DoPad(UCase(txtTransNo.Text), 10), UCase(txtTransNo.Text))
         lb_found = MySeek(txtTransNo.Text, "Transcode", PR_ICISSU)
         
         
         InitializeGrid
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   Cancel = True
                   Call Me.ClearVal
                  If txtTransNo.Enabled Then txtTransNo.SetFocus
                Else
                   txtvaluedate.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   Cancel = True
                   Call Me.ClearVal
                   txtTransNo.SetFocus
                Else
                   Call SetVal
                   LoadGRNTrans
                   If Mode <> "D" Then
                      txtTransNo.SetFocus
                   End If
                End If
            End Select
     ElseIf KeyCode = vbKeyF12 Then
            cmdLookup_Click
     End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup.Enabled = False
       InitializeGrid
    Else
       txtTransNo.SetFocus
       cmdLookup.Enabled = True
    End If
    If Button.Index = 7 Then InitializeGrid
    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICISSU, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
     '  txtVchrType = "JVS"
     '  Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
       txtTransNo = maxtranscode
       txtpartycode.SetFocus
    End If
End Sub


Public Sub SaveValues()
On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String

ls_Invoiceno = txtTransNo

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
              gc_dbcon.Execute "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "' and transtype = 'A'"
              
              gc_dbcon.Execute "DELETE FROM IC_Payments WHERE CompCode = '" & Gs_compcode & "' AND Referencecode = '" & Trim(txtTransNo) & "' and codeid = 'C'"
              
           Case Else
                If Mode = "E" Then
                    gc_dbcon.Execute "DELETE FROM IC_Trans WHERE CompCode = '" & Gs_compcode & "' AND Transcode = '" & Trim(txtTransNo) & "' and transtype = 'A'"
                    gc_dbcon.Execute "DELETE FROM IC_Payments WHERE CompCode = '" & Gs_compcode & "' AND Referencecode = '" & Trim(txtTransNo) & "'  and codeid = 'C'"
                End If
                If Mode = "A" Then
                
                    txtTransNo = maxtranscode
                End If
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      ls_sql = "INSERT into IC_Trans(Compcode, TransCode, SaleDate, PartyCode, ItemCode, Quantity, ItemRate, Amount,costamount, STaxRate, SaleTaxAmount, Remarks, SaleAccount, TaxAccount, "
                      ls_sql = ls_sql & " PartyAccount, VchrType, VoucherNo,itemtype,noofitems,Transtype)"
                      ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Trim(txtTransNo) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Trim(txtpartycode) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 2))) & "," & Val(0 & .TextMatrix(ln_cnt, 3)) & "," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 4)) & "," & Val(0 & .TextMatrix(ln_cnt, 5)) & "," & Val(0 & .TextMatrix(ln_cnt, 6)) & ",'" & Trim(TxtRemarks) & "','" & Trim(txtsaleaccount) & "','" & Trim(txttaxaccount) & "','" & Trim(txtpartyaccount) & "','" & Trim(txtvchrtype) & "','" & Trim(txtvchrno) & "','" & Trim(.TextMatrix(ln_cnt, 7)) & "'," & Trim(.TextMatrix(ln_cnt, 8)) & ",'A')"
                      gc_dbcon.Execute ls_sql
                    Next
                  '  Call SetVoucher
                 End With
                 
                 
                 'payments receible
                    If Val(txttotalamount) > 0 Then
                     pr_dumy.Open "select max(transcode) as transcode from ic_payments where compcode = '" & Gs_compcode & "'  and codeid = 'C'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                     If Not pr_dumy.EOF Then
                     ls_transcode = DoPad(Trim(str(Val(0 & pr_dumy("transcode")) + 1)), 10)
                     Else
                     ls_transcode = DoPad(Trim(str(1)), 10)
                     End If
                     pr_dumy.Close
                    
                        ls_sql = "Insert into IC_Payments( CompCode, Partycode,Codeid,TransCode, TransDate, Amount,Taxamount, Remarks,referencecode) "
                        ls_sql = ls_sql & " Values ('" & Gs_compcode & "','" & txtpartycode & "','C','" & ls_transcode & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "'," & Val(0 & txttotalamount) & "," & Val(txttotaltaxamount) & ",'" & "Amount payable to " + txtpartydesc & "','" & txtTransNo & "'  )"
                         gc_dbcon.Execute ls_sql
                    End If
                 'payments paid
                 If Val(txtamountreceived) > 0 Then
                   pr_dumy.Open "select max(transcode) as transcode from ic_payments where compcode = '" & Gs_compcode & "'  and codeid = 'C'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                     If Not pr_dumy.EOF Then
                     ls_transcode = DoPad(Trim(str(Val(0 & pr_dumy("transcode")) + 1)), 10)
                     Else
                     ls_transcode = DoPad(Trim(str(1)), 10)
                     End If
                     pr_dumy.Close
                    
                    ls_sql = "Insert into IC_Payments( CompCode, Partycode,Codeid,TransCode, TransDate, Amount,Taxamount, Remarks,referencecode) "
                    ls_sql = ls_sql & " Values ('" & Gs_compcode & "','" & txtpartycode & "','C','" & ls_transcode & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "'," & (-1) * Val(0 & txtamountreceived) & "," & (-1) * Val(txttaxreceived) & ",'" & "Amount paid to " + txtpartydesc & "' ,'" & txtTransNo & "')"
                    
                    gc_dbcon.Execute ls_sql
                    
                End If
                 
                 
                 
     End Select
gc_dbcon.CommitTrans
PR_ICISSU.Requery
If Mode = "A" Then
    txtTransNo = maxtranscode
End If

    'If Mode = "A" Or Mode = "E" And txtvchrno <> "" Then
     '   ls_opt = SetErr("Print Voucher ?.", vbYesNo)
     '   If ls_opt = vbYes Then Call setprint
    'End If
    If Mode = "A" Or Mode = "E" Then
        ls_opt = SetErr("Print Adjustment Advice?.", vbYesNo)
        If ls_opt = vbYes Then Call Printinvoice
    End If
InitializeGrid
Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub SetVoucher()
Dim ls_sql As String
Dim ln_TmpVchNo
Dim ln_OrgVchNo
        ln_OrgVchNo = txtvchrno
        ' Save reference of Voucher
        ls_sql = "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks,InstrumentNo,CrncyCode,ExchgRate,userid,adddate,addtime) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Format(Date, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrtype & "','" & TxtRemarks & "','" & Txtinstrument & "','PKR'," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
        gc_dbcon.Execute ls_sql
                 
        ' Save Details of Voucher
                
        'debit the party account
        If Val(txttotalamount) + Val(txttotaltaxamount) > 0 Then
            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) "
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtpartyaccount & "'," & 1 & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrtype & "'," & Val(txttotalamount) + Val(txttotaltaxamount) & "," & 0 & "," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(TxtRemarks) & "','" & Trim(TxtRemarks) & "')"
            gc_dbcon.Execute ls_sql
        End If
              
        'credit the sale account
        If Val(txttotalamount) > 0 Then
            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) "
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtsaleaccount & "'," & 2 & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrtype & "',0," & Val(txttotalamount) & "," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(TxtRemarks) & "','" & Trim(TxtRemarks) & "')"
            gc_dbcon.Execute ls_sql
        End If
        'credit the tax account
        If Val(txttotaltaxamount) > 0 Then
            ls_sql = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, ExchgRate,userid,adddate,addtime,Acct_Nirration,AcctName) "
            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txttaxaccount & "'," & 3 & ",'" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtvchrno & "','" & txtvchrtype & "',0," & Val(txttotaltaxamount) & "," & Val(0) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(TxtRemarks) & "','" & Trim(TxtRemarks) & "')"
            gc_dbcon.Execute ls_sql
        End If
        
        
        'update voucher reference
        If Mode = "A" Then
                    PR_VchCntr.Requery
                    PR_VchCntr.Filter = "Branchcode = '" & Gs_BranchCode & "' And  VchrType = '" & txtvchrtype & "'"
  
                    If PR_VchType.Fields("VchrFrequency") = "1" Then
                       ln_TmpVchNo = Val(0 & PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value))))) + 1
                    Else
                       ln_TmpVchNo = Val(0 & PR_VchCntr.Fields("VchrCount")) + 1
                    End If
 
                    If ln_TmpVchNo > ln_OrgVchNo And ln_OrgVchNo = Val(txtvchrno) Then
                        'lb_Vstat = True
                        txtvchrno = DoPad(Trim(str(ln_TmpVchNo)), 10)
                        
                        If PR_VchType.Fields("VchrFrequency") = "1" Then
                          PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value)))) = ln_TmpVchNo
                        Else
                          PR_VchCntr.Fields("VchrCount") = ln_TmpVchNo
                        End If
                        PR_VchCntr.Update
                    Else
                        If PR_VchType.Fields("VchrFrequency") = "1" Then
                          PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value)))) = PR_VchCntr.Fields("VchrMonth" & Trim(str(Month(txtvaluedate.Value)))) + 1
                        Else
                          PR_VchCntr.Fields("VchrCount") = PR_VchCntr.Fields("VchrCount") + 1
                        End If
                        PR_VchCntr.Update
                    End If
                End If
                
                
End Sub

Public Sub ClearVal()
End Sub
Private Sub setprint()
On Error GoTo LocalErr
Dim ls_BranchName As String
 If txtvchrno <> "" Then
         If MySeek(Gs_BranchCode, "BranchCode", PR_Branch) Then ls_BranchName = PR_Branch("BranchDesc")
   With rptVoucher
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.RPT"
        .Destination = crptToWindow
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & ls_VchDesc & "'"
        .Formulas(5) = "BranchName = '" & Gs_BranchCode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(txtvchrno) & "' and {Gl_Trans.BranchCode} = '" & Gs_BranchCode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = '" & Trim(txtvchrtype) & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(txtvaluedate) & "," & Month(txtvaluedate) & "," & Day(txtvaluedate) & ")"
        .Formulas(2) = "Sig1 = '" & Gc_UserName & "'"
        .Formulas(3) = "Sig2 = '" & Gs_Sign2 & "'"
        .Formulas(4) = "Sig3 = '" & Gs_Sign3 & "'"
        .Action = 1
   End With
 End If
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub
Private Sub Printinvoice()
On Error GoTo LocalErr

   With CrystalReport1
        .WindowTitle = Me.Caption
        '.Destination = crptToPrinter
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "Invoice.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Invoice'"
        .SelectionFormula = "{Ic_Trans.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.Transtype}= 'A'"
        .SelectionFormula = .SelectionFormula & "  and {Ic_Trans.transcode} = '" & Trim(ls_Invoiceno) & "'"
        
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub

Private Sub SetVal()
     txtpartycode = PR_ICISSU("partyCode") & ""
     Call txtpartycode_KeyDown(vbKeyReturn, vbKeyShift)
     txtvaluedate = PR_ICISSU("Saledate")
     TxtRemarks = PR_ICISSU("Remarks")
    ' txtsaleaccount = PR_ICISSU("Saleaccount")
    ' Call txtsaleaccount_KeyDown(vbKeyReturn, vbKeyShift)
    ' txttaxaccount = PR_ICISSU("Taxaccount")
    ' Call txttaxaccount_KeyDown(vbKeyReturn, vbKeyShift)
    ' txtpartyaccount = PR_ICISSU("partyaccount")
    ' Call txtpartyaccount_KeyDown(vbKeyReturn, vbKeyShift)
    ' txtVchrType = PR_ICISSU("vchrtype")
    ' Call txtVchrType_KeyDown(vbKeyReturn, vbKeyShift)
    ' txtvchrno = PR_ICISSU("voucherno")
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtTransNo.Text) = txtTransNo.MaxLength And Len(txtpartycode) = txtpartycode.MaxLength And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
    Pr_ICParty.Requery
    PR_ICISSU.Requery
    PR_IcItem.Requery
    PR_Branch.Requery
    PR_VchCntr.Requery
    PR_VchType.Requery
End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim ln_cnt As Integer
 
 If KeyCode = vbKeyReturn And Len(TxtItemCode.Text) > 0 Then

         TxtItemCode.Text = DoPad(Trim(TxtItemCode.Text), 4)
         lb_found = MySeek(TxtItemCode.Text, "itemcode", PR_IcItem)
         
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtItemCode.SetFocus
         Else
             txtItemDesc = PR_IcItem("Description")
             txtmtype = PR_IcItem("mdesc")
             If Val(PR_IcItem("NoofItems")) > 1 Then
                txtunitprice = PR_IcItem("TotalPurchaseprice")
             Else
                txtunitprice = PR_IcItem("Purchaseprice")
             End If
             If txtqty.Enabled Then txtqty.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
        Command1_Click
 ElseIf KeyCode = vbKeyReturn And Len(TxtItemCode.Text) = 0 Then
'        txtVchrType.SetFocus
 ElseIf KeyCode = vbKeyPageUp Then TxtRemarks.SetFocus
 End If
End Sub
Private Sub AddToGrid()
Dim ln_cnt As Integer
            If (Val(txtqty) > 0 And Val(txtunitprice) > 0) Then
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                    If TxtItemCode.Text <> "" Then
                        With GrdGRN
                            If PS_RowClicked = "" Then
                                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                    Else
                                        .Row = PI_CurRow
                                    End If
                                    If PS_RowClicked = "" Then
                                        .TextMatrix(.Row, 0) = PI_SrNo
                                    Else
                                        .TextMatrix(.Row, 0) = PI_CurRow
                                    End If
                               If MySeek(TxtItemCode.Text, "ItemFind", PR_IcItem) Then
                                                .TextMatrix(.Row, 1) = Trim(TxtItemCode)
                                                .TextMatrix(.Row, 2) = Val(txtqty)
                                                .TextMatrix(.Row, 3) = Val(txtunitprice)
                                                .TextMatrix(.Row, 4) = Val(txtamount)
                                                .TextMatrix(.Row, 5) = Val(txttax)
                                                .TextMatrix(.Row, 6) = Val(txttaxamount)
                                                .TextMatrix(.Row, 7) = txtmtype
                                                .TextMatrix(.Row, 8) = Val(txtnoofitems)
                            Else
                                Call SetErr("Item Code Not Found.", vbCritical)
                                TxtItemCode.SetFocus
                            End If
                                TxtItemCode.Text = ""
                                txtqty = ""
                                txtunitprice = ""
                                txttax = ""
                                txttaxamount = ""
                                txtamount = ""
                                txtItemDesc = ""
                                
                                PS_RowClicked = ""
                        End With
                    End If
                        TotalGRN
                        TxtItemCode.SetFocus
                   
        Else
            Call SetErr("Please Enter Qty./Unit Price", vbCritical)
            txtqty.SetFocus
       End If
      

End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Item Code|<Quantity|<Unit Price|<Amount|<Sale Tax|<Tax Amount|<ItemType|<No of Items"
        .ColWidth(1) = 900
        .ColWidth(2) = 900
        .ColWidth(3) = 1000
        .ColAlignment(3) = 7
        .ColWidth(4) = 1400
        .ColWidth(5) = 900
        .ColWidth(6) = 1400
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
    PS_RowClicked = ""
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
    With GrdGRN
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            TotalGRN
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub
Private Sub txtUnitPrice_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtamount = Val(txtqty) * Val(txtunitprice)
    txttax.SetFocus
End If
End Sub
Private Sub txtunitprice_LostFocus()
If Val(txtunitprice) > 0 Then
    txtamount = Val(txtqty) * Val(txtunitprice)
Else
    txtunitprice.SetFocus
End If
End Sub

Private Sub TotalGRN()
    Dim ln_cnt As Integer
      txttotalamount = ""
      txttotaltaxamount = ""
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            .TextMatrix(ln_cnt, 0) = ln_cnt
            txttotalamount = Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 4))
            txttotaltaxamount = Val(txttotaltaxamount) + Val(.TextMatrix(ln_cnt, 6))
            
            PI_SrNo = ln_cnt
        Next
    End With
End Sub

Private Sub LoadGRNTrans()
On Error GoTo LocalErr
Dim lb_found As Boolean
Dim ln_cnt   As Integer
Dim temp As String
    
InitializeGrid
    
    lb_found = MySeek(txtTransNo, "transcode", PR_ICISSU)
   
    If lb_found Then
        With GrdGRN
            Do While Trim(txtTransNo) = PR_ICISSU("Transcode")
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(PR_ICISSU("ItemCode") & "")
                .TextMatrix(.Row, 2) = PR_ICISSU("Quantity")
                .TextMatrix(.Row, 3) = Val(0 & PR_ICISSU("Itemrate"))
                .TextMatrix(.Row, 4) = Val(0 & PR_ICISSU("amount"))
                .TextMatrix(.Row, 5) = Val(0 & PR_ICISSU("STaxRate"))
                .TextMatrix(.Row, 6) = Val(0 & PR_ICISSU("SaleTaxAmount"))
                .TextMatrix(.Row, 7) = Trim(PR_ICISSU("ItemType") & "")
                .TextMatrix(.Row, 8) = Val(0 & PR_ICISSU("Noofitems"))
                .Rows = .Rows + 1
                PR_ICISSU.MoveNext
                If PR_ICISSU.EOF Then Exit Do
             Loop
            
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalGRN
    Else
        Call SetErr("Transaction Issue not found.", vbCritical)
        
    End If
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub

Private Sub GrdGRN_DblClick()
    With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
        
        TxtItemCode = .TextMatrix(.Row, 1)
        txtqty = .TextMatrix(.Row, 2)
        txtunitprice = Val(.TextMatrix(.Row, 3))
        txtamount = Val(.TextMatrix(.Row, 4))
        txttax = Val(.TextMatrix(.Row, 5))
        txttaxamount = Val(.TextMatrix(.Row, 6))
        txtitemtype = .TextMatrix(.Row, 7)
        txtnoofitems = Val(.TextMatrix(.Row, 8))
        
        PS_RowClicked = "Y"
    End With
End Sub

Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtpartycode.SetFocus
    End If
End Sub

Public Sub SetFrmEnv(ls_mode As String)
    txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
    txtpartycode.Enabled = IIf(ls_mode <> "D", True, False)
    TxtRemarks.Enabled = IIf(ls_mode <> "D", True, False)
    Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub

'Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim lb_found As Boolean
'
' If KeyCode = vbKeyReturn Then
'         txtVchrType = UCase(txtVchrType)
'         PR_VchType.Filter = "BranchCode = '" & Gs_BranchCode & "'"
'         lb_found = MySeek(Gs_BranchCode + txtVchrType.Text, "FindFld", PR_VchType)
'                If Not lb_found Then
'                   Call SetErr(Gs_RecNFMsg, vbCritical)
'                   txtvchrno = ""
'                   txtVchrType.SetFocus
'                Else
'                    txtVchrDesc = PR_VchType.Fields("VchrDescrip")
'
'                   If Mode = "A" Then
'                     'PR_VchCntr.Filter = adFilterNone
'                      PR_VchCntr.Requery
'                      PR_VchCntr.Filter = "Branchcode = '" & Gs_BranchCode & "' And  VchrType = '" & txtVchrType & "'"
'
'                      If PR_VchCntr.EOF Then
'                         Call SetErr("Voucher type not found in <GL_VchCntr> Table.", vbCritical)
'                         txtVchrType = ""
'                         txtVchrType.SetFocus
'                         Exit Sub
'                      End If
'
'                      If txtVchrType.Text = "0OB" And PR_VchCntr.Fields("VchrCount") = 1 Then
'                         Call SetErr("Cannot be more than one voucher in a year.", vbCritical)
'                         txtVchrType.SetFocus
'                         Exit Sub
'                      Else
'                         If PR_VchType.Fields("VchrFrequency") = "1" Then
'                            txtvchrno = DoPad((Trim(Str(Val(0 & PR_VchCntr.Fields("VchrMonth" & Trim(Str(Month(txtvaluedate.Value))))) + 1))), 10)
'                         Else
'                            txtvchrno = DoPad((Trim(Str(Val(0 & PR_VchCntr.Fields("VchrCount")) + 1))), 10)
'                         End If
'                         ln_OrgVchNo = Val(txtvchrno)
'                      '   txtsaleaccount.SetFocus
'
'                      End If
'                   Else
'                   If txtvchrno.Enabled Then txtvchrno.SetFocus
'                   End If
'                End If
'                PR_VchType.Filter = adFilterNone
'  ElseIf KeyCode = vbKeyF12 Then
'   Call Command2_Click
'  End If
'  End Sub
'
'Private Sub txtvchrtype_LostFocus()
'enterkeystatus = True
'End Sub
Private Sub txtvchrno_Change()

End Sub
