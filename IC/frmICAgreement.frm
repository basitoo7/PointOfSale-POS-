VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmICAgreement 
   Caption         =   "Purchase Agreement"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmICAgreement.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Tag             =   "SKIP"
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   1005
      ButtonWidth     =   1376
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
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
            Caption         =   "&Listing"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Re&fresh"
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
               Picture         =   "frmICAgreement.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmICAgreement.frx":075E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmICAgreement.frx":0BB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmICAgreement.frx":1006
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmICAgreement.frx":145A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmICAgreement.frx":18AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmICAgreement.frx":2002
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6645
      Left            =   30
      TabIndex        =   0
      Top             =   525
      Width           =   7080
      Begin VB.TextBox txtdesc 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1215
         MaxLength       =   255
         TabIndex        =   55
         Top             =   1305
         Width           =   5805
      End
      Begin VB.CheckBox chkPPVchr 
         Caption         =   "Post Payment Voucher"
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   5400
         TabIndex        =   54
         Tag             =   "SKIPN"
         Top             =   2730
         Width           =   1620
      End
      Begin VB.Frame Frame3 
         Caption         =   "Payment Schedule"
         Height          =   1080
         Left            =   30
         TabIndex        =   40
         Top             =   3555
         Width           =   7020
         Begin VB.TextBox txtbankname1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2130
            MaxLength       =   50
            TabIndex        =   48
            Top             =   630
            Width           =   1785
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1800
            Picture         =   "frmICAgreement.frx":2456
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "SKIP"
            Top             =   630
            Width           =   315
         End
         Begin VB.TextBox txtbankcode 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   1095
            MaxLength       =   3
            TabIndex        =   46
            Top             =   645
            Width           =   690
         End
         Begin VB.TextBox txtbankinstr 
            BackColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   4965
            MaxLength       =   20
            TabIndex        =   45
            Top             =   630
            Width           =   1950
         End
         Begin MSComCtl2.DTPicker DTPAccural 
            Height          =   330
            Left            =   1095
            TabIndex        =   41
            Tag             =   "SKIPN"
            Top             =   255
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
            _Version        =   393216
            Format          =   19726337
            CurrentDate     =   37696
         End
         Begin MSMask.MaskEdBox txtamount 
            Height          =   315
            Left            =   5520
            TabIndex        =   43
            Top             =   255
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Bank :"
            Height          =   210
            Left            =   615
            TabIndex        =   50
            Top             =   675
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Instrument # :"
            Height          =   210
            Left            =   3975
            TabIndex        =   49
            Top             =   675
            Width           =   1110
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount :"
            Height          =   240
            Left            =   4395
            TabIndex        =   44
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Accural Date :"
            Height          =   210
            Left            =   45
            TabIndex        =   42
            Top             =   285
            Width           =   1035
         End
      End
      Begin VB.CheckBox chkglaccount 
         Caption         =   "Open GL Account"
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   5400
         TabIndex        =   39
         Tag             =   "SKIPN"
         Top             =   2445
         Value           =   1  'Checked
         Width           =   1620
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1935
         Picture         =   "frmICAgreement.frx":25C8
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "SKIP"
         Top             =   930
         Width           =   315
      End
      Begin VB.TextBox txtptypedesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "SKIP"
         Top             =   930
         Width           =   4770
      End
      Begin VB.TextBox txtptype 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   28
         Tag             =   "SKIPN"
         Top             =   930
         Width           =   690
      End
      Begin VB.CommandButton cmdlookup2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Picture         =   "frmICAgreement.frx":273A
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "SKIP"
         Top             =   2430
         Width           =   315
      End
      Begin VB.TextBox txtpmode 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1215
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2445
         Width           =   690
      End
      Begin VB.TextBox txtbankname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2250
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "SKIP"
         Top             =   2790
         Width           =   2160
      End
      Begin VB.CommandButton CmdLookUp8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1905
         Picture         =   "frmICAgreement.frx":28AC
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "SKIP"
         Top             =   2805
         Width           =   315
      End
      Begin VB.TextBox txtbank 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   14
         Top             =   2820
         Width           =   690
      End
      Begin VB.TextBox txtaccountno 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   13
         Top             =   3195
         Width           =   3180
      End
      Begin VB.TextBox txtpmdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2235
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "SKIP"
         Top             =   2430
         Width           =   2175
      End
      Begin VB.TextBox txttranscode 
         BackColor       =   &H00FFFF80&
         Height          =   315
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   10
         Tag             =   "SKIPN"
         Top             =   195
         Width           =   1170
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2415
         Picture         =   "frmICAgreement.frx":2A1E
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1935
         Picture         =   "frmICAgreement.frx":2B90
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "SKIP"
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2265
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "SKIP"
         Top             =   555
         Width           =   4770
      End
      Begin VB.TextBox txtSupplier 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1215
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "SKIPN"
         Top             =   555
         Width           =   690
      End
      Begin VB.CheckBox chkstatus 
         Caption         =   "Status"
         Height          =   300
         Left            =   6225
         TabIndex        =   4
         Top             =   165
         Width           =   810
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000080&
         Height          =   2085
         Left            =   30
         TabIndex        =   1
         Top             =   4530
         Width           =   7020
         Begin VB.TextBox Text1 
            Height          =   315
            Left            =   7125
            MaxLength       =   100
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   2715
            Visible         =   0   'False
            Width           =   195
         End
         Begin Crystal.CrystalReport rptVoucher 
            Left            =   6450
            Top             =   885
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
         Begin MSFlexGridLib.MSFlexGrid grdVoucher 
            Height          =   1545
            Left            =   75
            TabIndex        =   51
            Top             =   165
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   2725
            _Version        =   393216
            Rows            =   1
            BackColorFixed  =   -2147483637
         End
         Begin MSMask.MaskEdBox txtTotal 
            Height          =   315
            Left            =   2190
            TabIndex        =   52
            Top             =   1725
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Total :"
            Height          =   240
            Left            =   1065
            TabIndex        =   53
            Top             =   1770
            Width           =   1125
         End
      End
      Begin MSComCtl2.DTPicker DTPAgr 
         Height          =   330
         Left            =   5790
         TabIndex        =   19
         Tag             =   "SKIPN"
         Top             =   2085
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         Format          =   19726337
         CurrentDate     =   37696
      End
      Begin MSMask.MaskEdBox txtqty 
         Height          =   315
         Left            =   1215
         TabIndex        =   20
         Top             =   1725
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtrate 
         Height          =   315
         Left            =   3525
         TabIndex        =   21
         Top             =   1725
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txttotalamount 
         Height          =   315
         Left            =   5775
         TabIndex        =   33
         Top             =   1725
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtamountpaid 
         Height          =   315
         Left            =   1215
         TabIndex        =   35
         Top             =   2085
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtbal 
         Height          =   315
         Left            =   3540
         TabIndex        =   37
         Top             =   2085
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport Crrpt 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         Destination     =   1
         CopiesToPrinter =   2
         DiscardSavedData=   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   210
         Left            =   360
         TabIndex        =   56
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance:"
         Height          =   240
         Left            =   2595
         TabIndex        =   38
         Top             =   2100
         Width           =   930
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount Paid:"
         Height          =   240
         Left            =   90
         TabIndex        =   36
         Top             =   2085
         Width           =   1125
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount :"
         Height          =   240
         Left            =   4650
         TabIndex        =   34
         Top             =   1755
         Width           =   1125
      End
      Begin VB.Label lblcaption 
         AutoSize        =   -1  'True
         Caption         =   "Caption"
         Height          =   210
         Left            =   2130
         TabIndex        =   32
         Top             =   1755
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Type:"
         Height          =   210
         Left            =   90
         TabIndex        =   31
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Agr. Date :"
         Height          =   210
         Left            =   4980
         TabIndex        =   27
         Top             =   2100
         Width           =   780
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Qty :"
         Height          =   240
         Left            =   435
         TabIndex        =   26
         Top             =   1740
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   " Rate :"
         Height          =   240
         Left            =   2655
         TabIndex        =   25
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pament Mode :"
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   2490
         Width           =   1050
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Bank :"
         Height          =   210
         Left            =   720
         TabIndex        =   23
         Top             =   2850
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Instrument # :"
         Height          =   210
         Left            =   195
         TabIndex        =   22
         Top             =   3225
         Width           =   1110
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Ref :"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Code :"
         Height          =   210
         Left            =   105
         TabIndex        =   8
         Top             =   585
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmICAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_Supplier As New Recordset
Public Pr_Agreement As New Recordset
Dim PR_Currency As New Recordset
Dim PR_GbaIds As New Recordset
Dim PR_Bank As New Recordset
Dim PR_Group As New Recordset
Dim PR_Type As New Recordset
Dim PR_GlDetl As New Recordset

Dim PR_VchrCntr As New Recordset
Dim PR_VchrType As New Recordset
Dim PR_Branch As New Recordset
Dim PR_Para As New Recordset

Dim ls_VchrType As String
Dim ls_Vchrno As String

Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String


Private Sub Check1_Click()

End Sub

Private Sub cmdlookup2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpmode
    Set PO_DESC = Text1
    PR_GbaIds.Filter = "Recid = 'PAM'"
    GoTop PR_GbaIds
    MyLookup.Caption = "Payment Mode"
    MyLookup.FillGrid PR_GbaIds, "IdCode", "IdDescrip", 2
    MyLookup.Show 1
    PR_GbaIds.Filter = adFilterNone
    If Len(txtpmode) > 0 Then txtpmode_KeyDown vbKeyReturn, vbKeyShift
End Sub
Public Sub InitializeGrid()
    With grdVoucher
        .Redraw = False
        .Clear
        .Rows = 2
        .Cols = 2
        .FormatString = "Sr# |<Accural Date|<Amount|<Bank Code|<Instrument#"
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 5000
        .Redraw = True
    End With
    PI_SrNo = 0
End Sub
Private Sub TotalDrCr()
    Dim ln_cnt As Integer
    txtTotal = 0
    With grdVoucher
        For ln_cnt = 1 To .Rows - 1
            .TextMatrix(ln_cnt, 0) = ln_cnt
            txtTotal = txtTotal + Val(.TextMatrix(ln_cnt, 2))
            PI_SrNo = ln_cnt
        Next
    End With
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttranscode
    Set PO_DESC = Text1
    Gs_SQL = "Select PTransCode,AgrDate from Ic_PAgreement "
    Gs_FindFld = "PTransCode"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by PTransCode,AgrDate"
    
    MyLookupOLDB.Caption = "Purchase Reference"
    MyLookupOLDB.Show 1
    
    If Trim(txttranscode) <> "" Then Call txttranscode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtptype
    Set PO_DESC = txtptypedesc
    Gs_SQL = "Select Typecode,Typedesc from Ic_Ptype "
    Gs_FindFld = "Typdesc"
    Gs_OrderBy = "Order by TypeCode,TypeDesc"
    
    MyLookupOLDB.Caption = "Purchase Type"
    MyLookupOLDB.Show 1
    
    If Trim(txtptype) <> "" Then Call txtptype_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbankcode
    Set PO_DESC = txtbankname1

    GoTop PR_Bank
    MyLookup.Caption = "Banks"
    MyLookup.FillGrid PR_Bank, "BankCode", "BankName", txtbank.MaxLength
    MyLookup.Show 1

    If Len(txtbankcode) > 0 Then txtbankcode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub DTPAccural_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtamount.SetFocus
End Sub

Private Sub txtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then DTPAccural.SetFocus
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtqty.SetFocus
End Sub

Private Sub txttranscode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
         txttranscode.Text = DoPad(txttranscode.Text, txttranscode.MaxLength)
         If Not MySeek(txttranscode.Text, "PtransCode", Pr_Agreement) Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtbank.SetFocus
         Else
          Call SetVal
           
        End If
 End If
End Sub

Private Sub txtbank_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
         txtbank.Text = DoPad(txtbank.Text, txtbank.MaxLength)
         If Not MySeek(txtbank.Text, "BankCode", PR_Bank) Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtbank.SetFocus
         Else
            txtbankname = PR_Bank("bankname")
            If txtaccountno.Enabled Then txtaccountno.SetFocus
        End If
 ElseIf KeyCode = vbKeyF12 Then
       Call CmdLookUp8_Click
 End If
End Sub

Private Sub txtbankcode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
         txtbankcode.Text = DoPad(txtbankcode.Text, txtbankcode.MaxLength)
         If Not MySeek(txtbankcode.Text, "BankCode", PR_Bank) Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtbankcode.SetFocus
         Else
            txtbankname1 = PR_Bank("bankname")
            If txtbankinstr.Enabled Then txtbankinstr.SetFocus
        End If
 ElseIf KeyCode = vbKeyF12 Then
       Call CmdLookUp8_Click
 ElseIf KeyCode = vbKeyPageUp Then
       txtother2.SetFocus
 End If
End Sub



Private Sub AddGrid()
If txtamount.Text <> "" Then

        If PS_RowClicked = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
        
            With grdVoucher
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

                .TextMatrix(.Row, 1) = DTPAccural
                .TextMatrix(.Row, 2) = Val(txtamount)
                .TextMatrix(.Row, 3) = txtbankcode
                .TextMatrix(.Row, 4) = txtbankinstr
                
                 PS_RowClicked = ""
                 txtamount = ""
                 txtbankcode = ""
                 txtbankinstr = ""
                 txtbankname1 = ""
                 DTPAccural.SetFocus
              End With
End If
End Sub
Private Sub txtbankinstr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If (Val(txtTotal) + Val(txtamount)) > txtbal Then
        Call MsgBox("Balance Amount Less then Schedule Amount", vbInformation)
    Else
        AddGrid
        TotalDrCr
    End If
End If

End Sub
Private Sub grdVoucher_DblClick()
    With grdVoucher
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
        
        DTPAccural = .TextMatrix(.Row, 1)
        txtamount = .TextMatrix(.Row, 2)
        txtbankcode = Val(.TextMatrix(.Row, 3))
        If txtbankcode <> "" Then Call txtbankcode_KeyDown(vbKeyReturn, vbKeyShift)
        txtbankinstr = Val(.TextMatrix(.Row, 4))
        
        PS_RowClicked = "Y"
        DTPAccural.SetFocus
    End With
End Sub

Private Sub grdVoucher_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdVoucher
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            TotalDrCr
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        End If
    End With
End Sub

Private Sub txtpmode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then

         txtpmode.Text = UCase(txtpmode.Text)
         PR_GbaIds.Filter = "Recid = 'PAM'"
         If Not MySeek(txtpmode.Text, "idCode", PR_GbaIds) Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            txtpmode.SetFocus
         Else
            txtpmdesc = PR_GbaIds("IdDescrip")
            If txtbank.Enabled Then txtbank.SetFocus
        End If
        PR_GbaIds.Filter = adFilterNone
 ElseIf KeyCode = vbKeyF12 Then
       Call cmdlookup2_Click
 End If
End Sub

Private Sub CmdLookUp8_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbank
    Set PO_DESC = txtbankname

    GoTop PR_Bank
    MyLookup.Caption = "Banks"
    MyLookup.FillGrid PR_Bank, "BankCode", "BankName", txtbank.MaxLength
    MyLookup.Show 1

    If Len(txtbank) > 0 Then txtbank_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub DTPAgr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpmode.SetFocus
End Sub
Private Sub txtamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtbankcode.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, Pr_Agreement, Me, txttranscode, txtSupplier, PR_Para, "ICPACnt", 10, " CustomerNo", "CustomerName", 0, False, Toolbar1)
End Sub
Private Sub Form_Load()
  SetToolBar(1) = chkRights("CUSTOMREG1")
  SetToolBar(2) = chkRights("CUSTOMREG2")
  SetToolBar(3) = chkRights("CUSTOMREG3")
  SetToolBar(4) = chkRights("CUSTOMREG4")

  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)


  PR_Supplier.Open "Select Ic_Supplier.* from Ic_Supplier where compcode = '" & Gs_compcode & "' order by customerCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  Pr_Agreement.Open "Select IC_PAgreement.* from IC_PAgreement where compcode = '" & Gs_compcode & "' order by SupplierCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Bank.Open "Select *  from Pa_banks  where compcode = '" & Gs_compcode & "' order by BankCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Type.Open "Select *  from IC_PType order by TypeCode ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GbaIds.Open "Select * from FCM_IDs order by idcode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GlDetl.Open "Select * from Gl_Detail where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  PR_VchrCntr.Open "SELECT * FROM Gl_VchrCntrs WHERE CompCode = '" & Gs_compcode & "' And VchrYear = " & Year(Gs_Fnperiod) & " ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_VchrType.Open "SELECT *,BranchCode+VchrType As FindFld FROM GlVchrType WHERE CompCode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Branch.Open "Select * from SysBranch Where Compcode = '" & Gs_compcode & "' order by Branchcode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Para.Open "select ICPACnt from Syscomp where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  
  Call InitializeGrid
  End Sub
Private Sub Form_Unload(Cancel As Integer)
  PR_Supplier.Close
  Pr_Agreement.Close
  PR_Bank.Close
  PR_GbaIds.Close
  PR_Type.Close
  PR_GlDetl.Close
  PR_VchrCntr.Close
  PR_VchrType.Close
  PR_Branch.Close
  PR_Para.Close
End Sub
Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSupplier
    Set PO_DESC = txtName
    Gs_SQL = "Select CustomerCode 'Customer Code', CustomerName  'Customer Name' from Ic_Supplier "
    Gs_FindFld = "CustomerName"
    Gs_OrderBy = "Order by CustomerCode,CustomerName"
    Gs_OtherPara = " Where compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Customers"
    MyLookupOLDB.Show 1
    If Trim(txtSupplier) <> "" Then Call txtSupplier_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtptype_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtptype.Text) > 0 Then
         txtptype.Text = DoPad(txtptype, txtptype.MaxLength)
         lb_found = MySeek(txtptype.Text, "CustomerCode", PR_Type)
          If Not lb_found Then
               Call SetErr(Gs_RecNFMsg, vbCritical)
               txtptype.SetFocus
          Else
               txtptypedesc = Trim(PR_Type("TypeDEsc") & "")
               lblcaption = Trim(PR_Type("Units") & "")
              If txtDesc.Enabled Then txtDesc.SetFocus
          End If
 End If
 End Sub

Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtrate.SetFocus
End Sub
Private Sub txtrate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txttotalamount = txtqty * txtrate
txtamountpaid.SetFocus
End If
End Sub
Private Sub txttotalamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtamountpaid.SetFocus
End Sub
Private Sub txtamountpaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtbal = txttotalamount - txtamountpaid
    DTPAgr.SetFocus
End If
End Sub

Private Sub txtSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtSupplier.Text) > 0 Then
         txtSupplier.Text = DoPad(txtSupplier, txtSupplier.MaxLength)
         lb_found = MySeek(txtSupplier.Text, "CustomerCode", PR_Supplier)
          
          If Not lb_found Then
               Call SetErr(Gs_RecNFMsg, vbCritical)
               txtSupplier.SetFocus
          Else
               txtName = Trim(PR_Supplier("Customername") & "")
                txtptype.SetFocus
          End If
 End If
 End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtName.Text <> "" Then txtAddress.SetFocus
    If KeyCode = vbKeyPageUp And TxtCustNo.Enabled Then TxtCustNo.SetFocus
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If PB_BlnkSupp And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found. ", vbCritical)
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, Pr_Agreement, Me, txttranscode, txtSupplier, PR_Para, "ICPACnt", 10, " CustomerNo", "CustomerName", 0, False, Toolbar1)
    End If
End Sub
Public Sub SaveValues()
Dim ln_cnt As Integer
Dim ls_CodeID As String
PB_BlnkSupp = False
Dim ls_sql As String
Dim ls_Accountno As String

gc_dbcon.BeginTrans

     Select Case Mode
           Case "D"
                gc_dbcon.Execute "DELETE FROM IC_PAgreement WHERE CompCode = '" & Gs_compcode & "' and PTransCode = '" & txttranscode & "' "
                gc_dbcon.Execute "DELETE FROM IC_PSchedule WHERE CompCode = '" & Gs_compcode & "' and PTransCode = '" & txttranscode & "' "
           Case Else
                 If Mode = "E" Then
                     gc_dbcon.Execute "DELETE FROM IC_PAgreement WHERE CompCode = '" & Gs_compcode & "' and PTransCode = '" & txttranscode & "' "
                     gc_dbcon.Execute "DELETE FROM IC_PSchedule WHERE CompCode = '" & Gs_compcode & "' and PTransCode = '" & txttranscode & "' "
                 End If
             If Mode = "A" And chkglaccount.Value = 1 Then
                    ls_Accountno = "002002001"
                      If Not MySeek(ls_Accountno + Right(txtSupplier, 4), "Accountno", PR_GlDetl) Then
                       gc_dbcon.Execute "INSERT into Gl_detail(compcode,Acct_sub,Acct_Detail,AccountNo,Acct_desc,crncy_code,Acct_Base,Acct_Type,Acct_Status,userid,adddate,addtime,Bs_DrLineNo,Bs_CrLineNo) VALUES ('" & Gs_compcode & "','" & ls_Accountno & "', '" & Right(txtSupplier, 4) & "' , '" & ls_Accountno + Right(txtSupplier, 4) & "','" & txtName & "','PKR','B','G','D','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','0000005500','0000005500')"
                     End If
             End If

                    
                    ls_sql = "Insert into   IC_PAgreement(CompCode, PTransCode, SupplierCode, PType, Amount, DAmount, Rate, AgrDate, PMode, BankCode, BankAccount, GLAccountNo, Transdate,Status,qty,pdesc) "
                    ls_sql = ls_sql + " VALUES  ('" & Gs_compcode & "','" & txttranscode & "','" & Trim(txtSupplier.Text) & "','" & Trim(txtptype) & "'," & Val(txttotalamount) & "," & Val(txtamountpaid) & "," & txtrate & ", '" & Format(DTPAgr, "YYYY/MM/DD") & "','" & Trim(txtpmode) & "' ,'" & Trim(txtbank) & "'  ,'" & txtaccountno & "','" & ls_Accountno & "','" & Format(Date, "YYYY/MM/DD") & "',1," & Val(txtqty) & ",'" & (txtDesc) & "' )"
                    gc_dbcon.Execute ls_sql
                    
                    With grdVoucher
                       For ln_cnt = 1 To .Rows - 1
                          If Len(Trim(.TextMatrix(ln_cnt, 1))) > 0 Then
                           gc_dbcon.Execute "INSERT into IC_PSchedule(CompCode, PTransCode, AccrualDate, DueAmount,BankCode,Instrno,Instno) VALUES ('" & Gs_compcode & "','" & txttranscode & "','" & Format(.TextMatrix(ln_cnt, 1), "YYYY/MM/DD") & "'," & .TextMatrix(ln_cnt, 2) & ",'" & .TextMatrix(ln_cnt, 3) & "','" & .TextMatrix(ln_cnt, 4) & "','" & DoPad(Trim(Str(.TextMatrix(ln_cnt, 0))), 3) & "')"
                          End If
                       Next
                     End With


            
                Call Voucher1
                res = MsgBox("Print Voucher", vbYesNo + vbExclamation)
                If res = vbYes Then Call setprint
     End Select

gc_dbcon.CommitTrans
Pr_Agreement.Requery
            If Mode = "A" Or Mode = "E" Then
             res = MsgBox("Print Payment Schedule", vbYesNo + vbExclamation)
               If res = vbYes Then Call PrintSchedule
            End If
End Sub
Private Sub PrintSchedule()
          With Crrpt
            .WindowTitle = Me.Caption
            .SelectionFormula = "{IC_Pagreement.PtransCode} ='" & txttranscode & "'"
            'If txtCustNo <> "" Then .SelectionFormula = .SelectionFormula & " and  {Customer.CustomerCode} = '" & txtCustNo & "'"
            'If txtCustNo <> "" And txtReferenceNo <> "" Then .SelectionFormula = .SelectionFormula + " And {PA_Agreement.referenceno} = '" & txtReferenceNo & "'"
            .ReportFileName = App.Path & Gs_ICRepoPath & "\PmtpSchedule.rpt"
            .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
            .Formulas(2) = "Reportname = 'Payment Schedule'"
            .Action = 1
         End With
End Sub



Private Sub Voucher1()
Dim ls_Remarks As String
Dim ls_Narration, ls_acctname, ls_accounts1, ls_accounts As String
             ls_VchrType = "JVS"
             PR_VchrCntr.Requery
             PR_GlDetl.Requery
             PR_VchrCntr.Filter = "VchrType = '" & ls_VchrType & "' and BranchCode ='" & Gs_BranchCode & "' And VchrYear = " & Year(DateValue(Gs_Fnperiod)) & ""
             PR_VchrType.Filter = "VchrType = '" & ls_VchrType & "' and BranchCode ='" & Gs_BranchCode & "'"

             If PR_VchrType.Fields("VchrFrequency") = "1" Then
                 ls_Vchrno = DoPad((Trim(Str(Val(0 & PR_VchrCntr.Fields("VchrMonth" & Trim(Str(Month(DTPAgr.Value))))) + 1))), 10)
             Else
                 ls_Vchrno = DoPad((Trim(Str(Val(0 & PR_VchrCntr.Fields("VchrCount")) + 1))), 10)
             End If

            ' Increment in Voucher Counter
             If PR_VchrType.Fields("VchrFrequency") = "1" Then
                PR_VchrCntr.Fields("VchrMonth" & Trim(Str(Month(DTPAgr.Value)))) = Val(0 & PR_VchrCntr.Fields("VchrMonth" & Trim(Str(Month(DTPAgr.Value))))) + 1
             Else
                PR_VchrCntr.Fields("VchrCount") = Val(0 & PR_VchrCntr.Fields("VchrCount")) + 1
             End If
             PR_VchrCntr.Update

              ls_Remarks = txtptypedesc + " Purchased from " + Trim(txtName)

             ' Save References of Voucher
               gc_dbcon.Execute "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks,ExchgRate,CrncyCode,userid,adddate,addtime,InstrumentNo) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(DTPAgr, "YYYY/MM/DD") & "','" & Format(Date, "YYYY/MM/DD") & "','" & ls_Vchrno & "','" & ls_VchrType & "','" & ls_Remarks & "',0 ,'PKR','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(txtaccountno) & "')"

            
               ls_accounts1 = "0040010010001"
            

               If MySeek(ls_accounts1, "Accountno", PR_GlDetl) Then
               ls_acctname = Trim(PR_GlDetl("Acct_desc"))
               ls_Narration = ls_Remarks
               End If

               gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_accounts1 & "',1,'" & Format(DTPAgr, "YYYY/MM/DD") & "','" & ls_Vchrno & "','" & ls_VchrType & "'," & Val(txttotalamount) & ",0 ,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_Narration & "','" & ls_acctname & "')"

                
                ls_Accountno = "002002001" + Right(txtSupplier, 4)
                

               'ls_accounts = ls_Accountno + Right(TxtCustNo, 4)
               If MySeek(ls_accounts, "Accountno", PR_GlDetl) Then
               ls_Narration = ls_Remarks
               ls_acctname = Trim(PR_GlDetl("Acct_desc"))
               End If

               gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_accounts & "',2,'" & Format(DTPAgr, "YYYY/MM/DD") & "','" & ls_Vchrno & "','" & ls_VchrType & "',0," & Val(txttotalamount) & " ,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_Narration & "','" & ls_acctname & "')"

              PR_VchrCntr.Filter = adFilterNone
              PR_VchrType.Filter = adFilterNone

              'Call setprint
End Sub

Private Sub Voucher2()
Dim ls_Remarks As String
Dim ls_Narration As String
            
             ls_VchrType = PR_Bank("Vchrtype")
             PR_VchrCntr.Requery
             PR_VchrCntr.Filter = "VchrType = '" & ls_VchrType & "' and BranchCode ='" & Gs_BranchCode & "' And VchrYear = " & Year(DateValue(Gs_Fnperiod)) & ""
             PR_VchrType.Filter = "VchrType = '" & ls_VchrType & "' and BranchCode ='" & Gs_BranchCode & "'"
             
             If PR_VchrType.Fields("VchrFrequency") = "1" Then
                 ls_Vchrno = DoPad((Trim(Str(Val(0 & PR_VchrCntr.Fields("VchrMonth" & Trim(Str(Month(DTPAgr.Value))))) + 1))), 10)
             Else
                 ls_Vchrno = DoPad((Trim(Str(Val(0 & PR_VchrCntr.Fields("VchrCount")) + 1))), 10)
             End If
            
            ' Increment in Voucher Counter
             If PR_VchrType.Fields("VchrFrequency") = "1" Then
                PR_VchrCntr.Fields("VchrMonth" & Trim(Str(Month(DTPAgr.Value)))) = Val(0 & PR_VchrCntr.Fields("VchrMonth" & Trim(Str(Month(DTPAgr.Value))))) + 1
             Else
                PR_VchrCntr.Fields("VchrCount") = Val(0 & PR_VchrCntr.Fields("VchrCount")) + 1
             End If
             PR_VchrCntr.Update
             
              ls_Remarks = txtpmdesc + " Received from " + Trim(txtName) + " For Issuing Plot No " + txtplotno + " Bank(" + Text1 + ")"
             ' Save References of Voucher
               gc_dbcon.Execute "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks,ExchgRate,CrncyCode,userid,adddate,addtime,InstrumentNo) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(DTPAgr, "YYYY/MM/DD") & "','" & Format(Date, "YYYY/MM/DD") & "','" & ls_Vchrno & "','" & ls_VchrType & "','" & ls_Remarks & "',0 ,'PKR','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & Trim(txtaccountno) & "')"
              
               
               ls_accounts1 = "001002002" + Right(TxtCustNo, 4)
               
               If MySeek(ls_accounts1, "Accountno", PR_GlDetl) Then
               ls_acctname = PR_GlDetl("Acct_desc")
               ls_Narration = ls_Remarks
               End If
               
               
               If MySeek(txtbank, "Bankcode", PR_Bank) Then
               ls_accounts = Trim(PR_Bank("Accountno") + "")
               End If
                 
               gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_accounts & "',1,'" & Format(DTPAgr, "YYYY/MM/DD") & "','" & ls_Vchrno & "','" & ls_VchrType & "'," & Val(txtconamount) & ",0 ,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_Narration & "','" & ls_acctname & "')"
   
              
               If MySeek(ls_accounts, "Accountno", PR_GlDetl) Then
               ls_acctname = PR_GlDetl("Acct_desc")
               ls_Narration = ls_Remarks
               End If
               
               gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_accounts1 & "',2,'" & Format(DTPAgr, "YYYY/MM/DD") & "','" & ls_Vchrno & "','" & ls_VchrType & "',0," & Val(txtconamount) & " ,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_Narration & "','" & ls_acctname & "')"
               
                
               
              PR_VchrCntr.Filter = adFilterNone
              PR_VchrType.Filter = adFilterNone
              
              Call setprint
End Sub




Private Sub setprint()
'On Error GoTo LocalErr
Dim ls_BranchName As String
Dim ls_VchDesc As String
 If ls_Vchrno <> "" Then
         If MySeek(Gs_BranchCode, "BranchCode", PR_Branch) Then ls_BranchName = PR_Branch("BranchDesc")
         If MySeek(ls_VchrType, "VchrType", PR_VchrType) Then ls_VchDesc = PR_VchrType("VchrDescrip")
   With rptVoucher
        .ReportFileName = App.Path & Gs_GlRepoPath & "\Vchr_Print.RPT"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & ls_VchDesc & "'"
        .Formulas(5) = "BranchName = '" & Gs_BranchCode + "-" + ls_BranchName & "'"
        .SelectionFormula = "{Gl_Trans.Voucher_No} = '" & Trim(ls_Vchrno) & "' and {Gl_Trans.BranchCode} = '" & Gs_BranchCode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.VchrType} = '" & Trim(ls_VchrType) & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.CompCode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & " and {Gl_Trans.Value_Date} = Date(" & Year(DTPAgr) & "," & Month(DTPAgr) & "," & Day(DTPAgr) & ")"
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
Private Sub SetVal()
     DTPAgr = Pr_Agreement("AgrDate")
     txttotalamount = Val(0 & Pr_Agreement("Amount"))
     txtDesc = Trim(Pr_Agreement("Pdesc") & "")
     txtqty = Val(0 & Pr_Agreement("Qty"))
     txtrate = Val(0 & Pr_Agreement("Rate"))
     txtamountpaid = Val(0 & Pr_Agreement("DAmount"))
     txtbal = txttotalamount - txtamountpaid
     txtptype = Trim(Pr_Agreement("Ptype") & "")
     If txtptype <> "" Then Call txtptype_KeyDown(vbKeyReturn, vbKeyShift)
     txtSupplier = Trim(Pr_Agreement("SupplierCode") & "")
     If txtSupplier <> "" Then Call txtSupplier_KeyDown(vbKeyReturn, vbKeyShift)
     txtpmode = Trim(Pr_Agreement("PMode") & "")
     If txtpmode <> "" Then Call txtpmode_KeyDown(vbKeyReturn, vbKeyShift)
     txtbank = Trim(Pr_Agreement("Bankcode") & "")
     If txtbank <> "" Then Call txtbank_KeyDown(vbKeyReturn, vbKeyShift)
     txtaccountno = Trim(Pr_Agreement("BankAccount") & "")
     chkstatus = Pr_Agreement("Status")
     Call LoadSchedule
End Sub
Private Sub LoadSchedule()
Dim lb_found As Boolean
Dim ln_cnt   As Integer
Dim temp As String
Dim pr_dumy As New Recordset
pr_dumy.Open "Select * from ic_PSchedule where ptranscode = '" & txttranscode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
InitializeGrid

 If Not pr_dumy.EOF Then
        ln_cnt = 1
        With grdVoucher
            Do While Not pr_dumy.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = ln_cnt
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = pr_dumy("AccrualDate")
                .TextMatrix(.Row, 2) = pr_dumy("DueAmount")
                .TextMatrix(.Row, 3) = pr_dumy("BankCode")
                .TextMatrix(.Row, 4) = pr_dumy("InstrNo")
                .Rows = .Rows + 1
                pr_dumy.MoveNext
                If pr_dumy.EOF Or pr_dumy.BOF Then Exit Do
             Loop
            
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalDrCr
  Else
        Call SetErr("Transaction not found.", vbCritical)
        txttranscode.SetFocus
 End If
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtSupplier.Text) = txtSupplier.MaxLength And Trim(txtptype) <> "" And txttotalamount <> "" And Trim(txtrate) <> "" And Trim(txtqty) <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Public Sub FrmRefresh()
  PR_Supplier.Requery
  Pr_Agreement.Requery
  PR_Bank.Requery
  PR_Branch.Requery
  PR_GbaIds.Requery
  PR_Branch.Requery
  PR_Currency.Requery
End Sub
