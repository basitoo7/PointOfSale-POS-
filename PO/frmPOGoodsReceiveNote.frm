VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOGoodsReceiveNote 
   Caption         =   "Goods Receive Note"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOGoodsReceiveNote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1680
      Left            =   45
      TabIndex        =   18
      Top             =   570
      Width           =   15150
      Begin VB.ComboBox txtpurin 
         Height          =   330
         ItemData        =   "frmPOGoodsReceiveNote.frx":030A
         Left            =   7935
         List            =   "frmPOGoodsReceiveNote.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   510
         Width           =   1290
      End
      Begin VB.TextBox txtwhtaxrate 
         Height          =   315
         Left            =   10635
         MaxLength       =   100
         TabIndex        =   69
         Tag             =   "SKIPN"
         Text            =   "3.5"
         Top             =   540
         Width           =   1305
      End
      Begin VB.CheckBox ChkWtax 
         Caption         =   "WhTax"
         Height          =   225
         Left            =   9600
         TabIndex        =   68
         Top             =   585
         Width           =   810
      End
      Begin VB.ComboBox txtType 
         Height          =   330
         ItemData        =   "frmPOGoodsReceiveNote.frx":0327
         Left            =   5160
         List            =   "frmPOGoodsReceiveNote.frx":0337
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   4050
      End
      Begin VB.TextBox txtmiscCharges 
         Height          =   315
         Left            =   10635
         MaxLength       =   100
         TabIndex        =   11
         Tag             =   "SKIPN"
         Text            =   "0"
         Top             =   870
         Width           =   1305
      End
      Begin VB.TextBox txtGRNNo 
         Height          =   315
         Left            =   13245
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1260
         Width           =   1725
      End
      Begin VB.TextBox txtvdocNo 
         Height          =   315
         Left            =   13245
         MaxLength       =   100
         TabIndex        =   5
         Top             =   870
         Width           =   1725
      End
      Begin VB.TextBox txtlabour 
         Height          =   315
         Left            =   7920
         MaxLength       =   100
         TabIndex        =   10
         Tag             =   "SKIPN"
         Text            =   "0"
         Top             =   870
         Width           =   1305
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1770
         Picture         =   "frmPOGoodsReceiveNote.frx":0374
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1665
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtacode 
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
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   12
         Tag             =   "SKIPN"
         Top             =   1665
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtaname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2115
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   1650
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   7200
         Picture         =   "frmPOGoodsReceiveNote.frx":04E6
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1665
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtacode1 
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
         Left            =   6780
         MaxLength       =   3
         TabIndex        =   13
         Tag             =   "SKIPN"
         Top             =   1680
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtaname1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7545
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3510
      End
      Begin VB.TextBox txtloading 
         Height          =   315
         Left            =   4800
         MaxLength       =   100
         TabIndex        =   9
         Tag             =   "SKIPN"
         Text            =   "0"
         Top             =   885
         Width           =   1305
      End
      Begin VB.TextBox txtfreight 
         Height          =   315
         Left            =   1335
         MaxLength       =   100
         TabIndex        =   8
         Tag             =   "SKIPN"
         Text            =   "0"
         Top             =   900
         Width           =   1305
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
         Left            =   1350
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2475
         Picture         =   "frmPOGoodsReceiveNote.frx":0658
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   150
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5760
         MaxLength       =   50
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   -120
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5910
         MaxLength       =   50
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   -150
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   360
         Left            =   1335
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1275
         Width           =   10605
      End
      Begin VB.TextBox txtVendorCode 
         Height          =   315
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   3
         Top             =   540
         Width           =   645
      End
      Begin VB.TextBox txtVendordesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2325
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   4545
      End
      Begin VB.TextBox TxtsiteDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   10800
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   23
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   -60
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxtSiteID 
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
         Left            =   10275
         MaxLength       =   3
         TabIndex        =   15
         Tag             =   "SKIPN"
         Text            =   "001"
         Top             =   -45
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   10590
         Picture         =   "frmPOGoodsReceiveNote.frx":07CA
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   -60
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox TxtBinDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   13815
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1065
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.TextBox txtbinID 
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
         Left            =   11385
         MaxLength       =   3
         TabIndex        =   16
         Tag             =   "SKIPN"
         Text            =   "001"
         Top             =   -60
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   11115
         Picture         =   "frmPOGoodsReceiveNote.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   -30
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   1995
         Picture         =   "frmPOGoodsReceiveNote.frx":0AAE
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   315
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   13245
         TabIndex        =   2
         Top             =   165
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   142934017
         CurrentDate     =   37580
      End
      Begin Crystal.CrystalReport rptVoucher 
         Left            =   3120
         Top             =   105
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
         Left            =   3480
         Top             =   15
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
      Begin MSComCtl2.DTPicker DTPSupdate 
         Height          =   315
         Left            =   13245
         TabIndex        =   4
         Top             =   510
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Format          =   142934017
         CurrentDate     =   37580
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Purchase In :"
         Height          =   255
         Left            =   6885
         TabIndex        =   72
         Top             =   555
         Width           =   990
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Type #  :"
         Height          =   255
         Left            =   3885
         TabIndex        =   67
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Supp Inv. Date :"
         Height          =   255
         Left            =   11940
         TabIndex        =   66
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Misc Charges:"
         Height          =   255
         Left            =   9330
         TabIndex        =   65
         Top             =   915
         Width           =   1275
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "GRN # :"
         Height          =   255
         Left            =   11940
         TabIndex        =   64
         Top             =   1275
         Width           =   1260
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Supp Inv. # :"
         Height          =   255
         Left            =   11925
         TabIndex        =   43
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Verified By :"
         Height          =   255
         Left            =   360
         TabIndex        =   42
         Top             =   1680
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved By :"
         Height          =   255
         Left            =   5700
         TabIndex        =   41
         Top             =   1680
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Labour :"
         Height          =   255
         Left            =   7095
         TabIndex        =   40
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Loading :"
         Height          =   255
         Left            =   4125
         TabIndex        =   39
         Top             =   945
         Width           =   660
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Freight :"
         Height          =   255
         Left            =   60
         TabIndex        =   38
         Top             =   945
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice #  :"
         Height          =   255
         Left            =   75
         TabIndex        =   37
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   525
         TabIndex        =   36
         Top             =   1275
         Width           =   780
      End
      Begin VB.Label label2 
         Caption         =   "Invoice Date :"
         Height          =   255
         Left            =   12195
         TabIndex        =   35
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendor :"
         Height          =   255
         Left            =   360
         TabIndex        =   34
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Site ID :"
         Height          =   255
         Left            =   9540
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Bin ID :"
         Height          =   255
         Left            =   11280
         TabIndex        =   32
         Top             =   -15
         Visible         =   0   'False
         Width           =   630
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
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
         Left            =   5250
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
               Picture         =   "frmPOGoodsReceiveNote.frx":0C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOGoodsReceiveNote.frx":1074
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOGoodsReceiveNote.frx":14C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOGoodsReceiveNote.frx":191C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOGoodsReceiveNote.frx":1D70
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOGoodsReceiveNote.frx":21C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOGoodsReceiveNote.frx":2918
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Save"
         Height          =   270
         Left            =   3180
         TabIndex        =   59
         Top             =   720
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6435
      Left            =   15
      TabIndex        =   44
      Top             =   2190
      Width           =   15180
      Begin MSComCtl2.DTPicker dtpexpdate 
         Height          =   360
         Left            =   11865
         TabIndex        =   71
         Top             =   990
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Format          =   142999553
         CurrentDate     =   41972
      End
      Begin VB.CheckBox chkdiscpamt 
         Caption         =   "Discount/Tax on Purchase Amount"
         Height          =   375
         Left            =   105
         TabIndex        =   70
         Top             =   5985
         Width           =   1950
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
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6015
         Width           =   1470
      End
      Begin VB.TextBox txtflatdisc 
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
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   4560
         MaxLength       =   11
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5610
         Width           =   1500
      End
      Begin VB.TextBox txtbonusamount 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10485
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   57
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5610
         Width           =   1515
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
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   13605
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   51
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6015
         Width           =   1455
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
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   10485
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6015
         Width           =   1515
      End
      Begin VB.TextBox txtsedamount 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7410
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   6015
         Width           =   1485
      End
      Begin VB.TextBox txtgstamount 
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7410
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5610
         Width           =   1485
      End
      Begin VB.TextBox txtitemname 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   47
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   5610
         Width           =   3555
      End
      Begin VB.TextBox txtnoofitems 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   0
         MaxLength       =   50
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   195
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
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   13605
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5610
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   5340
         Left            =   105
         TabIndex        =   14
         Top             =   210
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   9419
         _Version        =   393216
         BackColor       =   16777215
         RowHeightMin    =   300
         BackColorSel    =   16777215
         ForeColorSel    =   0
         GridColor       =   -2147483632
         FocusRect       =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.Label Label5 
         Caption         =   "Stock:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4080
         TabIndex        =   63
         Top             =   6045
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Flat Disc :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3855
         TabIndex        =   61
         Top             =   5640
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Sale Cost :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9645
         TabIndex        =   58
         Top             =   5625
         Width           =   885
      End
      Begin VB.Label Label20 
         Caption         =   " Net Amount :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   12615
         TabIndex        =   56
         Top             =   6045
         Width           =   1395
      End
      Begin VB.Label Label19 
         Caption         =   "Disc Amount :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9465
         TabIndex        =   55
         Top             =   6045
         Width           =   1395
      End
      Begin VB.Label Label18 
         Caption         =   "SED Amount:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6465
         TabIndex        =   54
         Top             =   6045
         Width           =   1050
      End
      Begin VB.Label Label17 
         Caption         =   "GST Amount:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6435
         TabIndex        =   53
         Top             =   5625
         Width           =   1485
      End
      Begin VB.Label Label11 
         Caption         =   " Total :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   13095
         TabIndex        =   52
         Top             =   5625
         Width           =   1020
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      NegotiatePosition=   1  'Left
      Begin VB.Menu NewRecord 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu EditRecord 
         Caption         =   "Edit"
         Shortcut        =   ^E
      End
      Begin VB.Menu DeleteRecord 
         Caption         =   "Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu SaveRecord 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu Item_Setup 
         Caption         =   "Item Setup"
         Shortcut        =   ^T
      End
      Begin VB.Menu Post_Purchase 
         Caption         =   "Post Purchase"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu editmain 
      Caption         =   "Edit"
      Begin VB.Menu Copy_data 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu Paste_data 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu add_row 
         Caption         =   "Add Row"
         Shortcut        =   ^I
      End
      Begin VB.Menu Delete_Row 
         Caption         =   "Delete Row"
         Shortcut        =   ^R
      End
      Begin VB.Menu Find_Existing_Record 
         Caption         =   "Find Existing Record"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "frmPOGoodsReceiveNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object

Dim Po_Status  As Integer
Dim Ls_ItemName  As String
Dim ln_qty, LN_EnterQty
Dim Opt

Dim ls_sql As String

Dim pr_dumy As New Recordset

Dim PR_UOM As New Recordset

Dim PR_ICIssue As New Recordset
Dim PR_INote As New Recordset
Dim PR_IcItem As New Recordset
Dim PR_Branch As New Recordset
Dim LeftOrRight$, FirstPass%
Dim ls_siteopt As Integer
Dim ln_cnt As Integer
Dim CX, CY
Dim ClickRow
Dim TboxCols
Dim ln_netRate As Double

Dim ln_totalamount, ln_TotalQTY, ln_totalnetrate As Double



Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from PO_POGRN where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
'txtCustRef = ClientCoderef("005") + Right(maxtranscode, 6)
End Function

Private Sub Add_Row_Click()
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

Private Sub chkdiscpamt_Click()
'If chkdiscpamt.Value = 1 Then
CalTax
'End If
End Sub
Private Sub CalTax()
    
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
           If Val(.TextMatrix(ln_cnt, 13)) > 0 Then
            If chkdiscpamt.Value = 1 Then
            
                .TextMatrix(ln_cnt, 15) = Val(.TextMatrix(ln_cnt, 11)) * Val(.TextMatrix(ln_cnt, 13)) / 100
            Else
                .TextMatrix(ln_cnt, 15) = Val(.TextMatrix(ln_cnt, 12)) * Val(.TextMatrix(ln_cnt, 13)) / 100
            End If
        End If
        
        If Val(.TextMatrix(ln_cnt, 14)) > 0 Then
            If chkdiscpamt.Value = 1 Then
              .TextMatrix(ln_cnt, 16) = Val(.TextMatrix(ln_cnt, 11)) * Val(.TextMatrix(ln_cnt, 14)) / 100
            Else
              .TextMatrix(ln_cnt, 16) = Val(.TextMatrix(ln_cnt, 12)) * Val(.TextMatrix(ln_cnt, 14)) / 100
            
            End If
        End If
       
        Next
    End With
    
    TotalAmount
    

End Sub
Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVendorCode
    Set PO_DESC = txtVendordesc
    Gs_SQL = "Select SupplierCode, Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1
    
    If txtVendorCode <> "" Then Call txtVendorCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command5_Click()
'Call MsgBox("Save Record")
End Sub

Private Sub Copy_data_Click()
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
            ResetRowSRNO
            TotalAmount
    End With
End Sub

Private Sub dtpexpdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
Call dtpexpdate_LostFocus
 With GrdGRN
    If .TextMatrix(.Row, 1) <> "" Then
     .CellBackColor = vbWindowBackground
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
           If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
            End If
            
          .SetFocus
    End If
    End With
End If
End Sub

Private Sub dtpexpdate_LostFocus()
With GrdGRN
If ClickRow <> "" Then
.TextMatrix(ClickRow, 23) = dtpexpdate
 
 ClickRow = ""
 End If
 .CellBackColor = vbWindowBackground
  dtpexpdate.Visible = False
  If TboxCol = True Then
  .Col = 23
  .CellBackColor = vbHighlight
  .SetFocus

  TboxCol = False
  Else
 .SetFocus
 .Col = 1
 .CellBackColor = vbHighlight
 End If
 
 
End With
End Sub

Private Sub DTPSupdate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvdocNo.SetFocus
End Sub

Private Sub Find_Existing_Record_Click()
Command1_Click
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

Private Sub GrdGRN_KeyUp(KeyCode As Integer, Shift As Integer)
LeftOrRight = "Right"
End Sub

Private Sub GrdGRN_LeaveCell()
With GrdGRN
 .CellBackColor = vbWindowBackground
End With
End Sub


Private Sub Item_Setup_Click()

frmItemstp.Show
frmItemstp.New_Record_Click
End Sub

Public Sub NewRecord_Click()
Mode = DentMode(Mode, 1, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
Command1.Enabled = False
InitializeGrid
txtTransNo = maxtranscode
txtTransNo.Enabled = False
CheckLogTrans
End Sub

Private Sub Paste_data_Click()
With GrdGRN
.TextMatrix(.Row, .Col) = Clipboard.GetText
End With
End Sub

Private Sub Post_Purchase_Click()
'Dim ls_transcodepost As String
'If Mode = "A" Then
'    txttransno = maxtranscode
'    ls_transcodepost = txttransno
'    Mode = DentMode(Mode, 4, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
'Else
' ls_transcodepost = txttransno
'End If
'
''Dim res
''res = MsgBox("Are you sure to post the voucher !!!", vbYesNo + vbInformation)
''If res = vbYes Then
''If PostPurchaseVoucher(Gs_compcode, ls_transcodepost) Then
''   Call MsgBox("Voucher Successfully Posted !!!", vbInformation)
''End If
'
'Call NewRecord_Click
'End If
End Sub

Private Sub saverecord_Click()
Mode = DentMode(Mode, 4, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
End Sub
Private Sub editrecord_Click()
Mode = DentMode(Mode, 2, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
Command1.Enabled = True
txtTransNo.Enabled = True
txtTransNo.SetFocus
End Sub
Private Sub Deleterecord_Click()
Mode = DentMode(Mode, 3, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
If txtVendorCode.Enabled Then txtVendorCode.SetFocus
Command1.Enabled = True
txtTransNo.Enabled = True
End Sub





Private Sub txtacode_LostFocus()
'    If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtacode1_LostFocus()
  '  If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
End Sub



Private Sub txtbinID_LostFocus()
If txtbinID <> "" Then Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtFlatDisc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If txtVendorCode.Enabled Then txtVendorCode.SetFocus
End If
End Sub
Private Sub CalcFlatDisc()
On Error GoTo LocalErr
  With GrdGRN
       For ln_cnt = 1 To .Rows - 1
         .TextMatrix(ln_cnt, 25) = Round((Val(.TextMatrix(ln_cnt, 17)) / (Val(txtnetamount) + Val(txtflatdisc)) * Val(txtflatdisc)), 0)
      Next
  End With
Exit Sub
LocalErr:
On Error GoTo 0
End Sub
Private Sub txtflatdisc_LostFocus()
If txtflatdisc <> "" Then
     CalcFlatDisc
End If
End Sub
Private Sub txtfreight_Change()
If txtfreight <> "" And Not IsNumeric(txtfreight) Then
    Call MsgBox("Enter Numeric Value!!!", vbCritical)
    txtfreight = ""
    txtfreight.SetFocus
 End If
End Sub

Private Sub txtfreight_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtloading.SetFocus
End Sub
Private Sub txtfreight_LostFocus()
If txtfreight <> "" Then
    TotalAmount
End If
End Sub

Private Sub txtGRNNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 txtacode = "001"
 If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
 txtacode1 = "001"
 If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
 GrdGRN.Col = 1
 GrdGRN.LeftCol = 1
 GrdGRN.SetFocus
End If
End Sub

Private Sub txtloading_LostFocus()
If txtloading <> "" Then
    TotalAmount
End If
End Sub
Private Sub txtlabour_LostFocus()
If txtlabour <> "" Then
    TotalAmount
End If
End Sub

Private Sub txtloading_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtlabour.SetFocus
End Sub
Private Sub txtloading_Change()
If txtloading <> "" And Not IsNumeric(txtloading) Then
    Call MsgBox("Enter Numeric Value!!!", vbCritical)
    txtloading = ""
    txtloading.SetFocus
 End If
End Sub

Private Sub txtlabour_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtmiscCharges.SetFocus
End Sub
Private Sub txtmisccharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvdocNo.SetFocus
End Sub
Private Sub txtlabour_Change()
If txtlabour <> "" And Not IsNumeric(txtlabour) Then
    Call MsgBox("Enter Numeric Value!!!", vbCritical)
    txtlabour = ""
    txtlabour.SetFocus
 End If
End Sub
Private Sub txtmisccharges_LostFocus()
If txtmiscCharges <> "" Then
    TotalAmount
End If
End Sub


Private Sub txtpurin_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
If DTPSupdate.Enabled Then DTPSupdate.SetFocus
End If
End Sub

Private Sub TxtSiteID_LostFocus()
    If TxtSiteID <> "" Then Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
End Sub




Private Sub txttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvaluedate.SetFocus
End Sub

Private Sub txtvdocNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then TxtRemarks.SetFocus
End Sub

Private Sub txtVendorCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtVendorCode) <> "" And KeyCode = vbKeyReturn Then
        txtVendorCode = DoPad(txtVendorCode, 6)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_Supplier where Compcode  = '" & Gs_compcode & "' and Suppliercode = '" & txtVendorCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Vendor Code not found !!!", vbCritical)
            txtVendorCode = ""
            txtVendordesc = ""
            txtVendorCode.SetFocus
            Exit Sub
        Else
            txtVendordesc = pr_dumy("Description")
            'TxtRemarks = "Goods received from " & Trim(txtVendordesc) & " GRN = " + txtTransNo
            If txtpurin.Enabled Then txtpurin.SetFocus
            
            
            
        End If
        pr_dumy.Close
           

ElseIf Trim(txtVendorCode) = "" And KeyCode = vbKeyReturn Then
        txtVendorCode = ""
        txtVendordesc = ""
        Call Command4_Click
End If

End Sub

Private Sub GrdGRN_Click()
GrdGRN.SelectionMode = flexSelectionFree
With GrdGRN
    txtitemname = LoadLastRate(.TextMatrix(.Row, 19))
    If .TextMatrix(.Row, 4) = "GODOWN" Then
        txtstock = Val(.TextMatrix(.Row, 21))
    ElseIf .TextMatrix(.Row, 4) = "SHOWROOM" Then
        txtstock = Val(.TextMatrix(.Row, 21))
    End If
    txtbonusamount = .TextMatrix(.Row, 20)
End With
GrdGRN.CellBackColor = vbHighlight
With GrdGRN
'If .Col = 4 Then
'txtsitetype.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
'txtsitetype.Visible = True
'ClickRow = .Row
'txtsitetype.SetFocus
'End If
End With
End Sub
Private Sub ResetRowSRNO()
With GrdGRN
   For ln_cnt = 1 To .Rows - 1
    .TextMatrix(ln_cnt, 0) = ln_cnt
   Next
End With
End Sub
Private Sub Command1_Click()
        On Error Resume Next
        Set PO_AnyForm = Nothing
        Set PO_AnyForm = Me
        Set PO_CODE = txtTransNo
        Set PO_DESC = Text1
        Gs_SQL = "SELECT GRN.TransCode AS ComputerCode, GRN.GRNCode AS GRNCode, Vendors.Description AS 'Vendors.Description', GRN.TransDate AS GRNDate,    GRN.NetAmount AS 'GRN.NetAmount' FROM         PO_POGRN GRN INNER JOIN         IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode"
        Gs_OrderBy = "ORDER BY GRN.TransCode desc"
        Gs_OtherPara = " Where GRN.compcode = '" & Gs_compcode & "' and GRN.glstatus = 0 "
        
        frmPosearchRecords.Caption = "GRN"
        frmPosearchRecords.Show 1
        
        If txtTransNo <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub Form_Load()
 Opt = ""
 
  SetToolBar(1) = chkRights("ICISUSTP01")
  SetToolBar(2) = chkRights1("PURCH00002")
  SetToolBar(3) = chkRights1("PURCH00003")
  SetToolBar(4) = chkRights("ICISUSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
'  PR_IcItem.Open "Select * from Ic_Item where compcode ='" & Gs_compcode & "' ", gc_dbcon, adOpenDynamic, adLockPessimistic, 1
  'PR_ICIssue.Open "Select * from Ic_TransMaster where compcode ='" & Gs_compcode & "' and transtype in ('D')  order by Transcode", gc_dbcon, adOpenDynamic, adLockOptimistic
  
  dtpexpdate.Value = Date
  txtvaluedate.Value = Date
  DTPSupdate.Value = Date
  txtacode = "001"
  txtacode1 = "001"
  txtpurin.ListIndex = 0
 
  InitializeGrid


  
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    PR_ICIssue.Close
   ' PR_IcItem.Close
   ' PR_Branch.Close
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtGRNNo.SetFocus
End If
End Sub
Function checkvalidate() As Boolean
If Trim(txtitemcode) = "" Then
    Call MsgBox("Enter Item Code !!!", vbCritical)
    txtitemcode.SetFocus
    checkvalidate = False
ElseIf Val(txtqty) = 0 Then
    Call MsgBox("Enter Quantity !!!", vbCritical)
    txtqty.SetFocus
    checkvalidate = False
ElseIf Val(txtunitprice) = 0 Then
    Call MsgBox("Enter unit price !!!", vbCritical)
    txtunitprice.SetFocus
    checkvalidate = False
Else
    checkvalidate = True
End If
End Function


Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Trim(txtTransNo.Text) <> "" Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txtTransNo.Text = DoPad(UCase(txtTransNo.Text), 10)
         PR_ICIssue.Open "select * from PO_POGRN where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       Select Case Mode
            Case "A"
                If Not PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   If txtTransNo.Enabled Then txtTransNo.SetFocus
                Else
                   txtvaluedate.SetFocus
                End If
            Case Else
                If PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   If txtTransNo.Enabled Then txtTransNo.SetFocus
                    txtTransNo.Enabled = True
                   txtVendorCode.Enabled = True
                   Command4.Enabled = True
                Else
                   Call SetVal
                   LoadGRNTrans
                   txtTransNo.Enabled = False
                   txtVendorCode.Enabled = False
                   Command4.Enabled = False
                End If
            End Select
 ElseIf KeyCode = vbKeyReturn And Trim(txtTransNo.Text) = "" Then
           Command1_Click
 End If
 End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      Command1.Enabled = False
      InitializeGrid
       
    Else
       If txtVendorCode.Enabled Then txtVendorCode.SetFocus
       Command1.Enabled = True
    End If
    If Button.Index = 7 Then
    InitializeGrid
    End If
    
    
    
    
    If PB_BlnkGRN And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
       Command1.Enabled = False
       txtTransNo = maxtranscode
       txtTransNo.Enabled = False
       If txtType.Enabled Then txtType.SetFocus
      CheckLogTrans
    Else
       txtTransNo.Enabled = True
       Command1.Enabled = True
        
    End If
End Sub


Private Sub TotalBeforeSave()
        
  With GrdGRN
       For ln_cnt = 1 To .Rows - 1
        If Trim(.TextMatrix(ln_cnt, 5)) <> "" Or Trim(.TextMatrix(ln_cnt, 7)) <> "" Then
        .TextMatrix(ln_cnt, 11) = Val(.TextMatrix(ln_cnt, 5)) * Val(.TextMatrix(ln_cnt, 7))
        End If
        
        If Val(.TextMatrix(ln_cnt, 8)) > 0 Then
        .TextMatrix(ln_cnt, 9) = Val(.TextMatrix(ln_cnt, 8)) * Val(.TextMatrix(ln_cnt, 11)) / 100
        End If
        
        If Trim(.TextMatrix(ln_cnt, 6)) <> "" Then
        .TextMatrix(ln_cnt, 10) = Val(.TextMatrix(ln_cnt, 6)) * Val(.TextMatrix(ln_cnt, 7))
        End If
        
        .TextMatrix(ln_cnt, 12) = Val(.TextMatrix(ln_cnt, 11)) - Val(.TextMatrix(ln_cnt, 9))

        If Val(.TextMatrix(ln_cnt, 13)) > 0 Then
                .TextMatrix(ln_cnt, 15) = Val(.TextMatrix(ln_cnt, 12)) * Val(.TextMatrix(ln_cnt, 13)) / 100
        End If
        
        If Trim(.TextMatrix(ln_cnt, 14)) <> "" Then
              .TextMatrix(ln_cnt, 16) = Val(.TextMatrix(ln_cnt, 12)) * Val(.TextMatrix(ln_cnt, 14)) / 100
        End If
        
        .TextMatrix(ln_cnt, 17) = Val(.TextMatrix(ln_cnt, 12)) + Val(.TextMatrix(ln_cnt, 15)) + Val(.TextMatrix(ln_cnt, 16))
        
      Next
  End With
  TotalAmount
End Sub


Public Sub SaveValues()
On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String
Dim ls_transtype As String



     Select Case Mode
           Case "D"
           
           
           
              gc_dbcon.Execute "DELETE FROM PO_POGRN WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "'"
              gc_dbcon.Execute "DELETE FROM PO_POGRNDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "'"

           
              
           Case Else
                If Mode = "E" Then
                          gc_dbcon.Execute "DELETE FROM PO_POGRN WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "'"
                          gc_dbcon.Execute "DELETE FROM PO_POGRNDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txtTransNo) & "'"
                End If
                If Mode = "A" Then
                    txtTransNo = maxtranscode
                    Me.Refresh
                End If
                      gc_dbcon.BeginTrans
                            ls_sql = "INSERT into PO_POGRN(Compcode,branchcode, TransCode, TransDate, AccountCode,Remarks,userid,adddate,addtime,Vcode,Acode,Freight,loading,labour,vdocno,Siteid,BinID,NetAmount,FlatDisc,GRNCode,MiscCharges,type,suppdate,Whstatus,whtaxrate,purin)"
                            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtVendorCode & "','" & RepApp(TxtRemarks) & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & txtacode & "','" & txtacode1 & "'," & Val(txtfreight) & "," & Val(txtloading) & "," & Val(txtlabour) & ",'" & txtvdocNo & "' ,'" & TxtSiteID & "','" & txtbinID & "', " & Val(txtnetamount) & ", " & Val(txtflatdisc) & " , '" & Trim(txtGRNNo) & "', " & Val(txtmiscCharges) & "," & txtType.ListIndex & " ,'" & Format(DTPSupdate, "YYYY/MM/DD") & "' ," & ChkWtax.Value & " ," & txtwhtaxrate & " ," & txtpurin.ListIndex & " )"
                            gc_dbcon.Execute ls_sql
                      gc_dbcon.CommitTrans
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      If .TextMatrix(ln_cnt, 1) <> "" Then
                    If .TextMatrix(ln_cnt, 4) = "GODOWN" Then
                        ls_siteopt = 1
                     Else
                        ls_siteopt = 2
                     End If
                        gc_dbcon.BeginTrans
                            ls_sql = "INSERT into PO_POGRNDetail(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,Siteid,expdays,expdate,FlatAmount)"
                            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 19)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 5))) & "," & Val(.TextMatrix(ln_cnt, 7)) & "," & Val(.TextMatrix(ln_cnt, 11)) & "," & Val(.TextMatrix(ln_cnt, 13)) & "," & Val(.TextMatrix(ln_cnt, 14)) & "," & Val(.TextMatrix(ln_cnt, 15)) & "," & Val(.TextMatrix(ln_cnt, 16)) & "," & Val(.TextMatrix(ln_cnt, 8)) & "," & Val(.TextMatrix(ln_cnt, 9)) + Val(.TextMatrix(ln_cnt, 25)) & ",'" & Trim(.TextMatrix(ln_cnt, 18)) & "'," & Val(.TextMatrix(ln_cnt, 6)) & "," & Val(.TextMatrix(ln_cnt, 10)) & "," & ls_siteopt & "," & Val(.TextMatrix(.Row, 22)) & ",'" & Format(.TextMatrix(.Row, 23), "YYYY/MM/DD") & "'," & Val(.TextMatrix(ln_cnt, 25)) & ")"
                            gc_dbcon.Execute ls_sql
                        gc_dbcon.CommitTrans
                        
                        gc_dbcon.BeginTrans
                            ls_sql = " update IC_Item set IC_Item.AvgRate = ItemAvgRate.AVGRate "
                            ls_sql = ls_sql & " FROM ItemAvgRate INNER JOIN  IC_Item ON ItemAvgRate.ItemCode = IC_Item.ItemCode"
                            ls_sql = ls_sql & " where ItemAvgRate.Itemcode = '" & Trim(.TextMatrix(ln_cnt, 19)) & "'"
                            gc_dbcon.Execute ls_sql
                        gc_dbcon.CommitTrans
                         
                         
                        gc_dbcon.BeginTrans
                        
                            If Val(.TextMatrix(ln_cnt, 7)) <> Val(.TextMatrix(ln_cnt, 24)) Then
                                ls_sql = " INSERT into IC_Itemsetuplog(Compcode, ItemCode, transdate,pcost,ccost,Ppcost,CPcost,AddUser) values('" & Gs_compcode & "' ,'" & Trim(.TextMatrix(ln_cnt, 19)) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "',0 ,0," & Val(.TextMatrix(ln_cnt, 24)) & "," & Val(.TextMatrix(ln_cnt, 7)) & ",'" & Gc_UserId + "-GRN" & "')"
                                gc_dbcon.Execute ls_sql
                              
                            End If
                            
                                ln_netRate = 0
                                ln_netRate = Round(((Val(.TextMatrix(ln_cnt, 11)) + Val(.TextMatrix(ln_cnt, 15))) - (Val(.TextMatrix(ln_cnt, 9)) + Val(.TextMatrix(ln_cnt, 25)))) / (Val(.TextMatrix(ln_cnt, 5)) + Val(.TextMatrix(ln_cnt, 6))), 2)
        
                                
                                ls_sql = " update IC_Item set avgrate1 = " & ln_netRate & " ,purchasecost = " & Val(.TextMatrix(ln_cnt, 7)) & ",ManuCode = '" & txtVendorCode & "' where Itemcode = '" & Trim(.TextMatrix(ln_cnt, 19)) & "'"
                                gc_dbcon.Execute ls_sql
                                
                                
                        gc_dbcon.CommitTrans
 
                         
                     End If
                    Next
                      ls_sql = "delete from PO_POGRNDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'"
                      gc_dbcon.Execute ls_sql
                 End With
              
              
              
                 
                 
     End Select
     

If Mode <> "D" Then
'         ls_opt = MsgBox("Post Account GL Voucher # ?.", vbYesNo)
'              If ls_opt = vbYes Then
'                Call PostPurchaseVoucher(Gs_compcode)
'                Call MsgBox("All Unpost GRN Vouchers Successfully Posted !!!", vbInformation)
'              Else
'                Call MsgBox("GL Voucher not Posted !!!", vbInformation)
'              End If
 End If

If Mode <> "D" Then
   ls_opt = MsgBox("Print GRN Note # ?.", vbYesNo)
   If ls_opt = vbYes Then Call PrintGRNnote
End If

If Mode = "A" Then
    txtTransNo = maxtranscode
End If
InitializeGrid
txtfreight = 0
txtloading = 0
txtlabour = 0
txtflatdisc = 0
txtmiscCharges = 0

If Mode = "E" Or Mode = "D" Then
txtTransNo = ""
End If

Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub PrintGRNnote()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POGRN.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Good Receive Note'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txtTransNo) & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
   End With
Exit Sub
LocalErr:
Call SetErr("Printer Not Ready", vbCritical)
End Sub
Private Sub setprint()
End Sub

Private Sub SetVal()
     txtvaluedate = PR_ICIssue("Transdate")
     DTPSupdate = IIf(IsNull(PR_ICIssue("Suppdate")), Date, PR_ICIssue("Suppdate"))
     txtType.ListIndex = PR_ICIssue("type")
     txtVendorCode = Trim(PR_ICIssue("AccountCode") & "")
     Call txtVendorCode_KeyDown(vbKeyReturn, vbKeyShift)
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtacode = Trim(PR_ICIssue("VCode") & "")
     Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
     txtacode1 = Trim(PR_ICIssue("ACode") & "")
     Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
     txtfreight = Val(PR_ICIssue("freight"))
     txtloading = Val(PR_ICIssue("Loading"))
     txtlabour = Val(PR_ICIssue("Labour"))
     txtvdocNo = Trim(PR_ICIssue("VdocNo") & "")
     txtflatdisc = Val(PR_ICIssue("flatDisc"))
     txtmiscCharges = Val(PR_ICIssue("MiscCharges"))
     txtGRNNo = Trim(PR_ICIssue("Grncode") & "")
     TxtSiteID = Trim(PR_ICIssue("SiteID") & "")
     Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
     txtbinID = Trim(PR_ICIssue("BinId") & "")
     Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
     ChkWtax.Value = Val(PR_ICIssue("whstatus"))
     txtwhtaxrate = Val(PR_ICIssue("whtaxrate"))
     txtpurin.ListIndex = Val(PR_ICIssue("Purin"))
     
End Sub
Private Function CheckPOQTY() As Boolean
Dim ls_sql As String
Dim ls_ItemCode As String
Dim ln_POQTY As Double
Dim ln_INQTY As Double
Dim ln_TotalQTY As Double

Dim Pr_dumyPOQty As New Recordset
    
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            ls_ItemCode = .TextMatrix(ln_cnt, 1)
            
            'check po qty
            ls_sql = "SELECT sum(IC_Trans.Quantity) AS QTY"
            ls_sql = ls_sql & " FROM IC_TransMaster INNER JOIN IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode"
            ls_sql = ls_sql & " where IC_Trans.ItemCode = '" & ls_ItemCode & "' and  IC_TransMaster.Transtype in('P','I','R') and IC_TransMaster.compcode = '" & Gs_compcode & "' "
            Pr_dumyPOQty.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
        
            If Not Pr_dumyPOQty.EOF Then
            ln_POQTY = Val(0 & Pr_dumyPOQty("QTY"))
            End If
            Pr_dumyPOQty.Close
        
            'check invoice qty
            ls_sql = "SELECT sum(IC_Trans.Quantity) AS QTY"
            ls_sql = ls_sql & " FROM IC_TransMaster INNER JOIN IC_Trans ON IC_TransMaster.Compcode = IC_Trans.Compcode AND IC_TransMaster.TransCode = IC_Trans.TransCode"
            ls_sql = ls_sql & " where IC_Trans.ItemCode = '" & ls_ItemCode & "' and  IC_TransMaster.Transtype in('S','O','D','B') and IC_TransMaster.compcode = '" & Gs_compcode & "' "
            Pr_dumyPOQty.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly
        
            If Not Pr_dumyPOQty.EOF Then
                ln_INQTY = Val(0 & Pr_dumyPOQty("QTY"))
            End If
            Pr_dumyPOQty.Close
            
            ln_TotalQTY = ln_POQTY - (ln_INQTY + Val(0 & .TextMatrix(ln_cnt, 2)))
            ln_qty = ln_POQTY - ln_INQTY
            LN_EnterQty = Val(0 & .TextMatrix(ln_cnt, 2))
            If ln_TotalQTY < 0 Then
                CheckPOQTY = False
                Ls_ItemName = ls_ItemCode
                Exit Function
            End If
        Next
     
     CheckPOQTY = True
    
    End With

End Function

Public Function ChkInputs() As Boolean
    Dim lb_opt As Boolean
    If Trim(txtTransNo) = "" Then
      Call MsgBox("Enter/Select Inovice !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtVendorCode) = "" Then
      Call MsgBox("Enter/Select Vendor Code !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtGRNNo) = "" Then
      Call MsgBox("Enter GRN !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtType) = "" Then
      Call MsgBox("Enter/Select  Type !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(TxtRemarks) = "" Then
      Call MsgBox("Enter Remarks !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtacode) = "" Then
      Call MsgBox("Enter/Select Verified Code !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtacode1) = "" Then
      Call MsgBox("Enter/Select Approved Code !!!", vbCritical)
      ChkInputs = False
    ElseIf GrdGRN.TextMatrix(1, 1) = "" Then
      Call MsgBox("Enter Items in grid !!!", vbCritical)
      ChkInputs = False
      GrdGRN.SetFocus
    Else
        With GrdGRN
          For ln_cnt = 1 To .Rows - 1
            If .TextMatrix(ln_cnt, 1) <> "" Then
            If Val(.TextMatrix(ln_cnt, 5)) = 0 Then
                Call MsgBox("QTY must be entered !!!", vbCritical)
                lb_opt = False
                Exit For
             Else
                lb_opt = True
             End If
            End If
          Next
        End With
       Call TotalBeforeSave
       If CheckSaleRate Then
       ChkInputs = lb_opt
       Else
       ChkInputs = False
       End If
    End If
End Function
Private Function CheckSaleRate() As Boolean
  
  With GrdGRN
       CheckSaleRate = True
       For ln_cnt = 1 To .Rows - 1
       ln_totalamount = (Val(.TextMatrix(ln_cnt, 12)) + Val(.TextMatrix(ln_cnt, 15)) + Val(.TextMatrix(ln_cnt, 16))) - (Val(.TextMatrix(ln_cnt, 10)) + Val(.TextMatrix(ln_cnt, 25)))
       ln_TotalQTY = Val(.TextMatrix(ln_cnt, 5)) + Val(.TextMatrix(ln_cnt, 6))
       ln_totalnetrate = Round(ln_totalamount / ln_TotalQTY, 2)
       If ln_totalnetrate > Val(.TextMatrix(ln_cnt, 20)) Then
       Call MsgBox("Net Rate of [" & .TextMatrix(ln_cnt, 2) & "] Greater then Sale Rate [Row #]" & str(ln_cnt) & " [Old Rate " & .TextMatrix(ln_cnt, 20) & " and New Rate " & str(ln_totalnetrate) & "]")
       CheckSaleRate = False
       Exit For
       End If
      Next
       
      
  End With
End Function
Public Sub FrmRefresh()
    Pr_ICParty.Requery
    PR_ICIssue.Requery
    PR_IcItem.Requery
    PR_Branch.Requery
    PR_VchCntr.Requery
    PR_VchType.Requery
End Sub


Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<Custom Code|<Item Name|<UOM|<Site|<Qty|<B-Qty|<Rate|<Disc%|<Disc Amount|<B-Amount|<Total|<Gross Amount|<GST%|<FED%|<GST Amount|<FED Amount|<Net Amount|<Batch|<ItemCode|<SaleCost|<Stock|<Exp Days|<Exp Date|<Purcost|<Flat Disc"
        .ColWidth(1) = 1500
        .ColWidth(2) = 2300
        .ColWidth(3) = 0
        .ColWidth(4) = 1500
        
        
        '5+6 =11+15-9+10+25
        
        
        .ColWidth(5) = 1000
        .ColAlignment(5) = 7
        
        .ColWidth(6) = 1000
        .ColAlignment(6) = 7
        
        .ColWidth(7) = 1000
        .ColAlignment(7) = 7
        
        .ColWidth(8) = 700
        .ColAlignment(8) = 7
        
        .ColWidth(9) = 1200
        .ColAlignment(9) = 7
        
        .ColWidth(10) = 1200
        .ColAlignment(10) = 7
        
        .ColWidth(11) = 1200
        .ColAlignment(11) = 7
        
        .ColWidth(12) = 1200
        .ColAlignment(12) = 7
        
        .ColWidth(13) = 700
        .ColAlignment(13) = 7
        .ColWidth(14) = 700
        .ColAlignment(14) = 7
        .ColWidth(15) = 1200
        .ColAlignment(15) = 7
        .ColWidth(16) = 1200
        .ColAlignment(16) = 7
        .ColWidth(17) = 1200
        .ColWidth(18) = 1500
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        .ColWidth(21) = 0
        .ColWidth(22) = 900
        .ColWidth(23) = 1200
        .ColWidth(24) = 0
        .ColWidth(25) = 900
        .Redraw = True
    End With
    

    
End Sub
Private Sub Command7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtSiteID
    Set PO_DESC = TxtsiteDesc
    Gs_SQL = "Select SiteCode, Description from IC_Sites "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Sites"
    MyLookupOLDB.Show 1
    If TxtSiteID <> "" Then Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command8_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbinID
    Set PO_DESC = TxtBinDesc
    Gs_SQL = "Select BinCode, Description from IC_SitesBins "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where sitecode ='" & TxtSiteID & "' and Compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Company Site Bins"
    MyLookupOLDB.Show 1
    
    If txtbinID <> "" Then Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtSiteID_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(TxtSiteID) <> "" And KeyCode = vbKeyReturn Then
        TxtSiteID = DoPad(TxtSiteID, 3)
        pr_dumy.Open "Select * from IC_Sites where Compcode  = '" & Gs_compcode & "' and Sitecode = '" & TxtSiteID & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Site Code not found !!!", vbCritical)
            TxtSiteID = ""
            TxtsiteDesc = ""
            'If TxtSiteID.Enabled Then TxtSiteID.SetFocus
        Else
            TxtsiteDesc = pr_dumy("Description")
           ' txtbinID.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(TxtSiteID) = "" And KeyCode = vbKeyReturn Then
        TxtSiteID = ""
        TxtsiteDesc = ""
        'Call Command7_Click
End If

End Sub

Private Sub txtBinID_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtbinID) <> "" And KeyCode = vbKeyReturn Then
        txtbinID = DoPad(txtbinID, 3)
        pr_dumy.Open "Select * from IC_SitesBins where  sitecode ='" & TxtSiteID & "' and compcode = '" & Gs_compcode & "'  and bincode = '" & txtbinID & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Bin Code not found !!!", vbCritical)
            txtbinID = ""
            TxtBinDesc = ""
            'txtbinID.SetFocus
        Else
            TxtBinDesc = pr_dumy("Description")
           ' If txtacode.Enabled Then txtacode.SetFocus
            
            
        End If
        pr_dumy.Close

ElseIf Trim(txtbinID) = "" And KeyCode = vbKeyReturn Then
        txtbinID = ""
        TxtBinDesc = ""
        'Call Command8_Click
End If

End Sub


Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacode
    Set PO_DESC = txtaname
    Gs_SQL = "Select ACode, Aname Description from PO_AuthorityPerson "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by Aname"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Authority Person"
    MyLookupOLDB.Show 1
    
    If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtACode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtacode) <> "" And KeyCode = vbKeyReturn Then
        txtacode = DoPad(txtacode, 3)
        pr_dumy.Open "Select * from PO_AuthorityPerson where Compcode  = '" & Gs_compcode & "' and Acode = '" & txtacode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Authority Code not found !!!", vbCritical)
            txtacode = ""
            txtaname = ""
            txtacode.SetFocus
        Else
            txtaname = pr_dumy("aname")
            'txtacode1.SetFocus
        End If
        pr_dumy.Close

ElseIf Trim(txtacode) = "" And KeyCode = vbKeyReturn Then
        txtacode = ""
        txtaname = ""
        Call Command2_Click
End If

End Sub


Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtacode1
    Set PO_DESC = txtaname1
    Gs_SQL = "Select ACode, Aname Description from PO_AuthorityPerson "
    Gs_FindFld = "Aname"
    Gs_OrderBy = "Order by AName"
    Gs_OtherPara = " where Compcode ='" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Authority Person"
    MyLookupOLDB.Show 1
    
    If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub txtACode1_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtacode1) <> "" And KeyCode = vbKeyReturn Then
        txtacode1 = DoPad(txtacode1, 3)
        pr_dumy.Open "Select * from PO_AuthorityPerson where Compcode  = '" & Gs_compcode & "' and Acode = '" & txtacode1 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Authority Code not found !!!", vbCritical)
            txtacode1 = ""
            txtaname1 = ""
            txtacode1.SetFocus
        Else
            txtaname1 = pr_dumy("aname")
            GrdGRN.SetFocus
        End If
        pr_dumy.Close

ElseIf Trim(txtacode1) = "" And KeyCode = vbKeyReturn Then
        txtacode1 = ""
        txtaname1 = ""
        Call Command3_Click
End If
End Sub


Private Sub TotalAmount()
      txttotalamount = ""
      txtgstamount = ""
      txtsedamount = ""
      txtdiscamount = ""
     ' txtbonusamount = ""
      txtnetamount = ""
     
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Round(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 11)), 2)
            'txtbonusamount = Val(txtbonusamount) + Val(.TextMatrix(ln_cnt, 10))
            txtgstamount = Round(Val(txtgstamount) + Val(.TextMatrix(ln_cnt, 15)), 2)
            txtsedamount = Round(Val(txtsedamount) + Val(.TextMatrix(ln_cnt, 16)), 2)
            txtdiscamount = Round(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), 2)
        Next
    End With
    txtnetamount = Round((Val(txttotalamount) + Val(txtgstamount) + Val(txtsedamount)) - (Val(txtdiscamount) + Val(txtflatdisc)), 0)
    txtnetamount = Round(Val(txtnetamount) + Val(txtfreight) + Val(txtloading) + Val(txtmiscCharges) + Val(txtlabour), 0)
 'If Val(txtflatdisc) > 0 Then CalcFlatDisc
End Sub

Private Sub LoadGRNTrans()
'On Error GoTo LocalErr

Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = "delete from PO_POGRNDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'"
gc_dbcon.Execute ls_sql
       

ls_sql = " SELECT PO_POGRNDetail.CustomCode,PO_POGRNDetail.ItemCode,PO_POGRNDetail.siteid, IC_Item.salecost,IC_Item.Description, PO_POGRNDetail.Quantity, PO_POGRNDetail.Rate, PO_POGRNDetail.Amount, PO_POGRNDetail.BonusQty, PO_POGRNDetail.BonusAmount ,PO_POGRNDetail.GSTper,PO_POGRNDetail.Sedper,PO_POGRNDetail.GSTAmount,PO_POGRNDetail.SedAmount,PO_POGRNDetail.Discper,PO_POGRNDetail.Discamount,PO_POGRNDetail.Remarks, IC_ItemUM.Description AS UOM"
ls_sql = ls_sql & " ,PO_POGRNDetail.expdays,PO_POGRNDetail.expdate,PO_POGRNDetail.FlatAmount FROM PO_POGRNDetail INNER JOIN IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where PO_POGRNDetail.Compcode = '" & Gs_compcode & "' and PO_POGRNDetail.Transcode = '" & txtTransNo & "' order by PO_POGRNDetail.srno"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = IIf(Val(Pr_LoadTrans("Siteid")) = 1, "GODOWN", "SHOWROOM")
                .TextMatrix(.Row, 5) = Val(Pr_LoadTrans("Quantity"))
                .TextMatrix(.Row, 6) = Val(Pr_LoadTrans("BonusQty"))
                .TextMatrix(.Row, 7) = Val(Pr_LoadTrans("Rate"))
                .TextMatrix(.Row, 8) = Val(Pr_LoadTrans("DiscPer"))
                .TextMatrix(.Row, 9) = Val(Pr_LoadTrans("DiscAmount")) - Val(Pr_LoadTrans("FlatAmount"))
                .TextMatrix(.Row, 10) = Val(Pr_LoadTrans("BonusAmount"))
                .TextMatrix(.Row, 11) = Val(Pr_LoadTrans("Amount"))
                .TextMatrix(.Row, 13) = Val(Pr_LoadTrans("GstPer"))
                .TextMatrix(.Row, 14) = Val(Pr_LoadTrans("SedPer"))
                .TextMatrix(.Row, 15) = Val(Pr_LoadTrans("GSTAmount"))
                .TextMatrix(.Row, 16) = Val(Pr_LoadTrans("SedAmount"))
                .TextMatrix(.Row, 18) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 19) = Trim(Pr_LoadTrans("Itemcode") & "")
                .TextMatrix(.Row, 20) = Val(Pr_LoadTrans("SaleCost"))
                .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - (Val(.TextMatrix(.Row, 9)) + Val(.TextMatrix(.Row, 10)))
                .TextMatrix(.Row, 22) = Val(Pr_LoadTrans("ExpDays"))
                .TextMatrix(.Row, 23) = IIf(IsNull(Trim(Pr_LoadTrans("expdate"))), Null, Trim(Pr_LoadTrans("expdate")))
                    
                .TextMatrix(.Row, 17) = Val(.TextMatrix(.Row, 12)) + Val(.TextMatrix(.Row, 15)) + Val(.TextMatrix(.Row, 16))
 
               .CellBackColor = vbWindowBackground
               
                ls_sql = "INSERT into PO_POGRNDetailLog(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,Siteid,SRNo,Computername,emode,expdays,expdate)"
                ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 19)) & "'," & (Val(0 & .TextMatrix(.Row, 5))) & "," & Val(.TextMatrix(.Row, 7)) & "," & Val(.TextMatrix(.Row, 11)) & "," & Val(.TextMatrix(.Row, 13)) & "," & Val(.TextMatrix(.Row, 14)) & "," & Val(.TextMatrix(.Row, 15)) & "," & Val(.TextMatrix(.Row, 16)) & "," & Val(.TextMatrix(.Row, 8)) & "," & Val(.TextMatrix(.Row, 9)) & ",'" & Trim(.TextMatrix(.Row, 17)) & "'," & Val(.TextMatrix(.Row, 6)) & "," & Val(.TextMatrix(.Row, 10)) & "," & ls_siteopt & "," & .Row & " ,'" & Gs_ComputerName & "','" & Mode & "'," & Val(.TextMatrix(.Row, 21)) & ",'" & Format(.TextMatrix(.Row, 22), "YYYY/MM/DD") & "' )"
                gc_dbcon.Execute ls_sql
       
               .TextMatrix(.Row, 17) = Val(.TextMatrix(.Row, 12)) + Val(.TextMatrix(.Row, 15)) + Val(.TextMatrix(.Row, 16))
 
                .Rows = .Rows + 1
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
    Else
        Call SetErr("Transaction not found.!!!", vbCritical)
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub LoadLogTrans()
'On Error GoTo LocalErr
Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = " SELECT PO_POGRNDetail.CustomCode,PO_POGRNDetail.ItemCode,PO_POGRNDetail.siteid, IC_Item.salecost,IC_Item.Description, PO_POGRNDetail.Quantity, PO_POGRNDetail.Rate, PO_POGRNDetail.Amount, PO_POGRNDetail.BonusQty, PO_POGRNDetail.BonusAmount ,PO_POGRNDetail.GSTper,PO_POGRNDetail.Sedper,PO_POGRNDetail.GSTAmount,PO_POGRNDetail.SedAmount,PO_POGRNDetail.Discper,PO_POGRNDetail.Discamount,PO_POGRNDetail.Remarks, IC_ItemUM.Description AS UOM"
ls_sql = ls_sql & " ,PO_POGRNDetail.expdays,PO_POGRNDetail.expdate FROM PO_POGRNDetailLog PO_POGRNDetail INNER JOIN IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where PO_POGRNDetail.Compcode = '" & Gs_compcode & "' and PO_POGRNDetail.computername ='" & Gs_ComputerName & "' order by PO_POGRNDetail.srno"

Pr_LoadTrans.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not Pr_LoadTrans.EOF Then
        With GrdGRN
            Do While Not Pr_LoadTrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(Pr_LoadTrans("CustomCode") & "")
                .TextMatrix(.Row, 2) = Trim(Pr_LoadTrans("Description") & "")
                .TextMatrix(.Row, 3) = Trim(Pr_LoadTrans("UOM") & "")
                .TextMatrix(.Row, 4) = IIf(Val(Pr_LoadTrans("Siteid")) = 1, "GODOWN", "SHOWROOM")
                .TextMatrix(.Row, 5) = Val(Pr_LoadTrans("Quantity"))
                .TextMatrix(.Row, 6) = Val(Pr_LoadTrans("BonusQty"))
                .TextMatrix(.Row, 7) = Val(Pr_LoadTrans("Rate"))
                .TextMatrix(.Row, 8) = Val(Pr_LoadTrans("DiscPer"))
                .TextMatrix(.Row, 9) = Val(Pr_LoadTrans("DiscAmount"))
                .TextMatrix(.Row, 10) = Val(Pr_LoadTrans("BonusAmount"))
                .TextMatrix(.Row, 11) = Val(Pr_LoadTrans("Amount"))
                .TextMatrix(.Row, 13) = Val(Pr_LoadTrans("GstPer"))
                .TextMatrix(.Row, 14) = Val(Pr_LoadTrans("SedPer"))
                .TextMatrix(.Row, 15) = Val(Pr_LoadTrans("GSTAmount"))
                .TextMatrix(.Row, 16) = Val(Pr_LoadTrans("SedAmount"))
                .TextMatrix(.Row, 18) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 19) = Trim(Pr_LoadTrans("Itemcode") & "")
                .TextMatrix(.Row, 20) = Val(Pr_LoadTrans("SaleCost"))
                .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - (Val(.TextMatrix(.Row, 9)) + Val(.TextMatrix(.Row, 10)))
                .TextMatrix(.Row, 22) = Val(Pr_LoadTrans("ExpDays"))
                .TextMatrix(.Row, 23) = IIf(IsNull(Trim(Pr_LoadTrans("expdate") & "")), "", Trim(Pr_LoadTrans("expdate") & ""))
                .Rows = .Rows + 1
                
                    .TextMatrix(.Row, 17) = Val(.TextMatrix(.Row, 12)) + Val(.TextMatrix(.Row, 15)) + Val(.TextMatrix(.Row, 16))
 
                Pr_LoadTrans.MoveNext
                If Pr_LoadTrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
        TotalAmount
        
    End If
    Pr_LoadTrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)

End Sub
Private Sub CheckLogTrans()
Dim pr_dumyLog As New Recordset
Dim res
pr_dumyLog.Open "select * from PO_POGRNDetailLog  where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyLog.EOF Then
    If pr_dumyLog("Emode") = "E" Then
        txtTransNo = pr_dumyLog("Transcode")
        res = MsgBox(txtTransNo & " # you have opened in edit mode not save Do you want to open now", vbYesNo + vbExclamation)
        If res = vbYes Then
        Mode = DentMode(Mode, 2, PR_ICIssue, Me, txtTransNo, txtTransNo, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        If PR_ICIssue.State = 1 Then PR_ICIssue.Close
        PR_ICIssue.Open "select * from PO_POGRN where compcode = '" & Gs_compcode & "' and Transcode = '" & txtTransNo & "' and glstatus = 0 ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not PR_ICIssue.EOF Then
        Call SetVal
        End If
        PR_ICIssue.Close
        LoadLogTrans
        Else
        
        ls_sql = "delete from  PO_POGRNDetailLog where computername = '" & Gs_ComputerName & "' "
        gc_dbcon.Execute ls_sql
           
        End If
    Else
        LoadLogTrans
    End If
End If
pr_dumyLog.Close

End Sub


Private Sub GrdGRN_DblClick()
    GrdGRN.SelectionMode = flexSelectionFree
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
'If GrdGRN.Col = 4 Then
' With GrdGRN
 '   txtsitetype.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
  '  txtsitetype.Visible = True
  '  ClickRow = .Row
  '  txtsitetype.SetFocus
'End With
'End If
 
If KeyCode = 112 And GrdGRN.Col = 1 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    Gs_SQL = "SELECT customCode,Description,Salecost FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    GrdGRN.TextMatrix(GrdGRN.Row, 1) = Text1
    If GrdGRN.TextMatrix(GrdGRN.Row, 1) <> "" Then
        Call GrdGRN_KeyPress(13)
    End If
  ElseIf KeyCode = 113 Then  ' F1 key pressed
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = Text1
    Set PO_DESC = Text2
    
    Gs_SQL = " SELECT GRN.TransCode AS ComputerCode, GRN.GRNCode AS GRNCode,GRN.TransDate AS GRNDate, Vendors.Description AS 'Vendors.Description',"
    Gs_SQL = Gs_SQL & " (((PO_POGRNDetail.Amount - PO_POGRNDetail.DiscAmount) + PO_POGRNDetail.GSTAmount) / (PO_POGRNDetail.Quantity + PO_POGRNDetail.BonusQty)) AS NetRate,PO_POGRNDetail.Quantity , PO_POGRNDetail.BonusQty,  PO_POGRNDetail.Rate  FROM  PO_POGRN GRN INNER JOIN"
    Gs_SQL = Gs_SQL & " IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode INNER JOIN"
    Gs_SQL = Gs_SQL & " PO_POGRNDetail ON GRN.Compcode = PO_POGRNDetail.Compcode AND GRN.TransCode = PO_POGRNDetail.TransCode LEFT OUTER JOIN"
    Gs_SQL = Gs_SQL & " IC_Item ON PO_POGRNDetail.Compcode = IC_Item.Compcode AND PO_POGRNDetail.ItemCode = IC_Item.ItemCode"
    Gs_OtherPara = " where GRN.compcode = '" & Gs_compcode & "' and PO_POGRNDetail.Itemcode = '" & GrdGRN.TextMatrix(GrdGRN.Row, 18) & "'"
    Gs_OrderBy = "ORDER BY GRN.TransCode DESC"
    frmPosearchRecords.Caption = "Vendor Rate Comparison"
    frmPosearchRecords.Show 1
 
 ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             ls_sql = "delete from PO_POGRNDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' and srno = " & .Row & "  "
             gc_dbcon.Execute ls_sql
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
             ls_sql = "delete from PO_POGRNDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' "
             gc_dbcon.Execute ls_sql
            
                InitializeGrid
            End If
            ResetRowSRNO
            TotalAmount
    End With
 ElseIf KeyCode = vbKeyDown And GrdGRN.Row = GrdGRN.Rows - 1 And GrdGRN.TextMatrix(GrdGRN.Row, 1) = "" Then
 txtflatdisc.SetFocus
 ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then 'key down and keyup
    With GrdGRN
    'txtitemname = LoadLastRate(.TextMatrix(.Row, 18))
    txtstock = Val(.TextMatrix(.Row, 21))
    txtbonusamount = .TextMatrix(.Row, 20)
    End With
 End If

    LeftOrRight = "Right" ' So we know if we are going forward or backward in the cells
    If Shift = 1 Then LeftOrRight = "Left" ' Assums we are pressing shift for shift tab

End Sub

Private Sub GrdGRN_KeyPress(KeyAscii As Integer)
'On Error GoTo ErrHandler
With GrdGRN
  If .Col = 4 Then
      If KeyAscii <> 13 Then
      .Text = Chr(KeyAscii)
      If UCase(.Text) = "G" Or UCase(.Text) = "S" Then
      Else
      .Text = ""
      Call MsgBox("Please G or S", vbExclamation)
      End If
      
      
      End If
      
  If UCase(Gc_UserId) <> UCase("Admin") Then
   .TextMatrix(.Row, .Col) = Gs_Siteid
  Else
      If UCase(.Text) = "G" Then
            .TextMatrix(.Row, .Col) = "GODOWN"
            txtstock = CheckBalQTY(.TextMatrix(.Row, 19), 1)
      ElseIf UCase(.Text) = "S" Then
            .TextMatrix(.Row, .Col) = "SHOWROOM"
            txtstock = CheckBalQTY(.TextMatrix(.Row, 19), 2)
      End If
  End If
  End If
  
End With
 Call GetKeysAdd(GrdGRN, KeyAscii)
Exit Sub

'ErrHandler:
'MsgBox ("An Error has Occured In The MSFlexgrid1_KeyPress() Procedure") & vbCr & "Report This Error To Latifjat@hotmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
End Sub
Sub StopTab()
    If FirstPass = True Then Exit Sub
    Dim X As Variant
    ' Dissable the tab stops in each control so the grid tab will work
    For X = 0 To Me.Controls.Count - 1
        Me.Controls(X).TabStop = False
    Next
End Sub

Public Sub GetKeysAdd(argFlexGrid As MSHFlexGrid, KeyAscii As Integer)
'This Procedure is used to display the pressed key into FlexGrid in Addition Mode
'so that when you press Enter Key in the last row then one row will be added.
'When you press the BackSpace Key in an empty Row then a Row will be Removed.
'On Error GoTo ErrHandler

If KeyAscii = 13 Then 'if Enter Key then...
  Opt = ""
  With argFlexGrid
        ' .SelectionMode = flexSelectionByRow
        .Row = .RowSel
    If .Col = 1 Then
        If txtTransNo = "" Then
            Call MsgBox("Enter Invoice No !!!", vbCritical)
            txtTransNo.SetFocus
            Exit Sub
        End If
        
        .CellBackColor = vbWindowBackground
       If .TextMatrix(.Row, 1) <> "" Then
          If PR_IcItem.State = 1 Then PR_IcItem.Close
          PR_IcItem.Open " Select * From Ic_Item Where compcode = '" & Gs_compcode & "' and  CustomCode='" & Trim(.TextMatrix(.Row, 1)) & " ' ", gc_dbcon, adOpenStatic, adLockReadOnly
          
          If PR_IcItem.RecordCount <= 0 Then
              Call MsgBox(Gs_RecNFMsg, vbCritical)
             .TextMatrix(.Row, 1) = ""
             
          Else
             .TextMatrix(.Row, 22) = 5
             .TextMatrix(.Row, 23) = DateAdd("D", Val(.TextMatrix(.Row, 22)), txtvaluedate)
             
             .TextMatrix(.Row, 0) = .Row
             .TextMatrix(.Row, 19) = Trim(PR_IcItem("Itemcode") & "")
             .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(.Row, 7) = Val(PR_IcItem("Purchasecost"))
             .TextMatrix(.Row, 24) = Val(.TextMatrix(.Row, 7))
             .TextMatrix(.Row, 20) = Val(PR_IcItem("Salecost"))
              txtbonusamount = .TextMatrix(.Row, 20)
              txtstock = ""
              txtitemname = .TextMatrix(.Row, 2)
              .TextMatrix(.Row, 8) = LoadLastDisc(.TextMatrix(.Row, 19))
              
              
             .Col = .Col + 3
             ' txtsitetype.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
             ' txtsitetype.Visible = True
             ' ClickRow = .Row
             ' txtsitetype.SetFocus
              'If UCase(Gs_DBDataSource) = "RAHAT" Then
             .TextMatrix(.Row, .Col) = Gs_Siteid
                
                If .TextMatrix(.Row, 4) = "GODOWN" Then
                 ls_siteopt = 1
                Else
                 ls_siteopt = 2
                End If
                
                .TextMatrix(.Row, 21) = CheckBalQTY(.TextMatrix(.Row, 19), ls_siteopt)
                 txtstock = Val(.TextMatrix(.Row, 21))
                 
             
            '  Else
             '   .TextMatrix(.Row, 4) = "GODOWN"
             ' End If
             .CellBackColor = vbHighlight
             
              PR_UOM.Open "Select * From IC_ItemUM Where MCode='" & Trim(PR_IcItem("Mcode") & "") & " '", gc_dbcon, adOpenStatic, adLockReadOnly
              If PR_UOM.RecordCount > 0 Then
                .TextMatrix(.Row, 3) = Trim(PR_UOM("Description") & "")
              End If
            PR_UOM.Close
          
          End If
         PR_IcItem.Close
       Else
           Call GrdGRN_KeyDown(112, vbKeyShift)
       End If
       ElseIf .Col = 2 Then
       ElseIf .Col = 3 Then
       
       ElseIf .Col = 4 Then
           .CellBackColor = vbWindowBackground
           .Col = .Col + 1
       ElseIf .Col = 5 Then
           .CellBackColor = vbWindowBackground
           If .TextMatrix(.Row, 4) = "" Then
             Call MsgBox("Enter Quantity!!!", vbCritical)
             Exit Sub
           End If
            .Col = 6
            .CellBackColor = vbHighlight
       ElseIf .Col = 6 Then
           .CellBackColor = vbWindowBackground
           .Col = .Col + 1
           .CellBackColor = vbHighlight
             
       ElseIf .Col = 7 Then
           .CellBackColor = vbWindowBackground
           If .TextMatrix(.Row, 7) = "" Then
             Call MsgBox("Enter Rate!!!", vbCritical)
             Exit Sub
           End If
            .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 7))
            .Col = .Col + 1
            .CellBackColor = vbHighlight
            .LeftCol = .Col - 3
       ElseIf .Col = 7 Then
            .CellBackColor = vbWindowBackground
            .Col = .Col + 1
            .CellBackColor = vbHighlight
       ElseIf .Col = 8 Then
            .CellBackColor = vbWindowBackground
            If .TextMatrix(.Row, 1) <> "" Then
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
           If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
            End If
            
          .SetFocus
        Else
         Call MsgBox("Enter/Select Item Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
       ElseIf .Col = 13 Then
            .CellBackColor = vbWindowBackground
            .Col = 14
            .CellBackColor = vbHighlight
       ElseIf .Col = 14 Then
            .CellBackColor = vbWindowBackground
            .Col = 17
            .CellBackColor = vbHighlight
            .LeftCol = .Col - 1
       ElseIf .Col = 17 Then
        If .TextMatrix(.Row, 1) <> "" Then
          If .Row = .Rows - 1 Then
           .Rows = .Rows + 1
          End If
          .Col = 1
          .LeftCol = 1
          .Row = .Row + 1
          .SetFocus
        Else
         Call MsgBox("Enter/Select Item Code!!!", vbCritical)
         .Row = .Row
         .Col = 1
        End If
          
        If .RowSel > 10 Then
              .TopRow = .Rows - 1 'To Move the Scrollbar
        End If
            
   End If
   End With
 Exit Sub
End If
      
If KeyAscii = 8 Then  'If BackSpace Key then...
With argFlexGrid
   If .Col = 1 Or .Col = 18 Or .Col = 22 Or .Col = 4 Or .Col = 5 Or .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 13 Or .Col = 14 Or .Col = 15 Or .Col = 23 Or .Col = 25 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
    If Trim(.TextMatrix(.Row, 5)) <> "" Or Trim(.TextMatrix(.Row, 7)) <> "" Then
        .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 7))
        End If
        If Val(.TextMatrix(.Row, 8)) > 0 Then
        .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 8)) * Val(.TextMatrix(.Row, 11)) / 100
        End If

        If Trim(.TextMatrix(.Row, 6)) <> "" Then
        .TextMatrix(.Row, 10) = Val(.TextMatrix(.Row, 6)) * Val(.TextMatrix(.Row, 7))
        End If
       
        .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - Val(.TextMatrix(.Row, 9))

       If Val(.TextMatrix(.Row, 13)) > 0 Then
            If chkdiscpamt.Value = 1 Then
                .TextMatrix(.Row, 15) = Val(.TextMatrix(.Row, 11)) * Val(.TextMatrix(.Row, 13)) / 100
            Else
                .TextMatrix(.Row, 15) = Val(.TextMatrix(.Row, 12)) * Val(.TextMatrix(.Row, 13)) / 100
            End If
        End If
        
'        If Trim(.TextMatrix(.Row, 22)) <> "" Then
'            .TextMatrix(.Row, 23) = DateAdd("D", Val(.TextMatrix(.Row, 22)), txtvaluedate)
'        Else
'            .TextMatrix(.Row, 23) = ""
'        End If
        
        If Trim(.TextMatrix(.Row, 14)) <> "" Then
            If chkdiscpamt.Value = 1 Then
              .TextMatrix(.Row, 16) = Val(.TextMatrix(.Row, 11)) * Val(.TextMatrix(.Row, 14)) / 100
            Else
              .TextMatrix(.Row, 16) = Val(.TextMatrix(.Row, 12)) * Val(.TextMatrix(.Row, 14)) / 100
            
            End If
        End If
        
     .TextMatrix(.Row, 17) = Val(.TextMatrix(.Row, 12)) + Val(.TextMatrix(.Row, 15)) + Val(.TextMatrix(.Row, 16))
              
   
   
    ls_sql = "delete from PO_POGRNDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' and srno = " & .Row & "  "
    gc_dbcon.Execute ls_sql
    
    If .TextMatrix(.Row, 4) = "GODOWN" Then
     ls_siteopt = 1
    Else
     ls_siteopt = 2
    End If
    
    .TextMatrix(.Row, 21) = CheckBalQTY(.TextMatrix(.Row, 19), ls_siteopt)
     txtstock = Val(.TextMatrix(.Row, 21))
     On Error Resume Next
          ls_sql = "INSERT into PO_POGRNDetailLog(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,Siteid,SRNo,Computername,emode,expdays,expdate)"
          ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 19)) & "'," & (Val(0 & .TextMatrix(.Row, 5))) & "," & Val(.TextMatrix(.Row, 7)) & "," & Val(.TextMatrix(.Row, 11)) & "," & Val(.TextMatrix(.Row, 13)) & "," & Val(.TextMatrix(.Row, 14)) & "," & Val(.TextMatrix(.Row, 15)) & "," & Val(.TextMatrix(.Row, 16)) & "," & Val(.TextMatrix(.Row, 8)) & "," & Val(.TextMatrix(.Row, 9)) & ",'" & Trim(.TextMatrix(.Row, 18)) & "'," & Val(.TextMatrix(.Row, 6)) & "," & Val(.TextMatrix(.Row, 10)) & "," & ls_siteopt & "," & .Row & " ,'" & Gs_ComputerName & "' ,'" & Mode & "'," & Val(.TextMatrix(.Row, 22)) & ",'" & Format(.TextMatrix(.Row, 23), "YYYY/MM/DD") & "')"
    gc_dbcon.Execute ls_sql
    TotalAmount
   
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      
      If .Col = 1 Or .Col = 17 Or .Col = 22 Or .Col = 18 Or .Col = 23 Then
        If .CellBackColor = vbHighlight Then
         .Text = "": .CellBackColor = vbWindowBackground
        End If
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
       ElseIf .Col = 23 Then
        dtpexpdate.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
        dtpexpdate.Visible = True
        ClickRow = .Row
        dtpexpdates.SetFocus
     
      ElseIf .Col = 5 Or .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 13 Or .Col = 14 Or .Col = 15 Then
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
        
      If .Col = 22 Then
        .TextMatrix(.Row, 23) = DateAdd("D", Val(.TextMatrix(.Row, 22)), txtvaluedate)
      End If
      
      
        If Val(.TextMatrix(.Row, 7)) > Val(.TextMatrix(.Row, 20)) Then
            Call MsgBox("Enter Purchase price less then sale price")
            .TextMatrix(.Row, 7) = 0
            Exit Sub
        End If
          
        If Trim(.TextMatrix(.Row, 5)) <> "" Or Trim(.TextMatrix(.Row, 7)) <> "" Then
        .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 7))
        End If
        
        If Val(.TextMatrix(.Row, 8)) > 0 Then
        .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 8)) * Val(.TextMatrix(.Row, 11)) / 100
        End If
        
        If Trim(.TextMatrix(.Row, 6)) <> "" Then
        .TextMatrix(.Row, 10) = Val(.TextMatrix(.Row, 6)) * Val(.TextMatrix(.Row, 7))
        End If
        
        .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - Val(.TextMatrix(.Row, 9))

        If Val(.TextMatrix(.Row, 13)) > 0 Then
            If chkdiscpamt.Value = 1 Then
                .TextMatrix(.Row, 15) = Val(.TextMatrix(.Row, 11)) * Val(.TextMatrix(.Row, 13)) / 100
            Else
                .TextMatrix(.Row, 15) = Val(.TextMatrix(.Row, 12)) * Val(.TextMatrix(.Row, 13)) / 100
            End If
        End If
        
        If Trim(.TextMatrix(.Row, 14)) <> "" Then
            If chkdiscpamt.Value = 1 Then
              .TextMatrix(.Row, 16) = Val(.TextMatrix(.Row, 11)) * Val(.TextMatrix(.Row, 14)) / 100
            Else
              .TextMatrix(.Row, 16) = Val(.TextMatrix(.Row, 12)) * Val(.TextMatrix(.Row, 14)) / 100
            
            End If
        End If
        
        .TextMatrix(.Row, 17) = Val(.TextMatrix(.Row, 12)) + Val(.TextMatrix(.Row, 15)) + Val(.TextMatrix(.Row, 16))
        
        ls_sql = "delete from PO_POGRNDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' and srno = " & .Row & "  "
        gc_dbcon.Execute ls_sql
        
        If .TextMatrix(.Row, 4) = "GODOWN" Then
             ls_siteopt = 1
        Else
            ls_siteopt = 2
        End If
        
        On Error Resume Next
               
        ls_sql = "INSERT into PO_POGRNDetailLog(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,Siteid,SRNo,Computername,eMode,expdays,expdate)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txtTransNo) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 19)) & "'," & (Val(0 & .TextMatrix(.Row, 5))) & "," & Val(.TextMatrix(.Row, 7)) & "," & Val(.TextMatrix(.Row, 11)) & "," & Val(.TextMatrix(.Row, 13)) & "," & Val(.TextMatrix(.Row, 14)) & "," & Val(.TextMatrix(.Row, 15)) & "," & Val(.TextMatrix(.Row, 16)) & "," & Val(.TextMatrix(.Row, 8)) & "," & Val(.TextMatrix(.Row, 9)) & ",'" & Trim(.TextMatrix(.Row, 18)) & "'," & Val(.TextMatrix(.Row, 6)) & "," & Val(.TextMatrix(.Row, 10)) & "," & ls_siteopt & "," & .Row & " ,'" & Gs_ComputerName & "','" & Mode & "'," & Val(.TextMatrix(.Row, 22)) & ",'" & Format(.TextMatrix(.Row, 23), "YYYY/MM/DD") & "')"
        gc_dbcon.Execute ls_sql
       
       TotalAmount
 
      
    End With
  End If
End Sub
Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtVendorCode.SetFocus
    End If
End Sub

Public Sub SetFrmEnv(ls_mode As String)
    txtLocCode.Enabled = IIf(ls_mode <> "D", True, False)
    txtpartycode.Enabled = IIf(ls_mode <> "D", True, False)
    TxtRemarks.Enabled = IIf(ls_mode <> "D", True, False)
    Frame2.Enabled = IIf(ls_mode <> "D", True, False)
End Sub


Private Sub txtVendorCode_LostFocus()
 If txtVendorCode <> "" Then Call txtVendorCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
