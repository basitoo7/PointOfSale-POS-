VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPOPurchaseOrder 
   Caption         =   "Purchase Order"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOPurchaseOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   15240
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
      Height          =   1425
      Left            =   45
      TabIndex        =   9
      Top             =   570
      Width           =   15150
      Begin VB.ComboBox txtType 
         Height          =   330
         ItemData        =   "frmPOPurchaseOrder.frx":030A
         Left            =   5160
         List            =   "frmPOPurchaseOrder.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   4050
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   1770
         Picture         =   "frmPOPurchaseOrder.frx":0357
         Style           =   1  'Graphical
         TabIndex        =   19
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
         TabIndex        =   5
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
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   1650
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   7200
         Picture         =   "frmPOPurchaseOrder.frx":04C9
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   6
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
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3510
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
         Picture         =   "frmPOPurchaseOrder.frx":063B
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   -150
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox TxtRemarks 
         Height          =   360
         Left            =   1350
         MaxLength       =   100
         TabIndex        =   4
         Top             =   945
         Width           =   13650
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
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   525
         Width           =   12645
      End
      Begin VB.TextBox TxtBinDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   13815
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1065
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   1995
         Picture         =   "frmPOPurchaseOrder.frx":07AD
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Format          =   62717953
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
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Type #  :"
         Height          =   255
         Left            =   3885
         TabIndex        =   46
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Verified By :"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Approved By :"
         Height          =   255
         Left            =   5700
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "PO #  :"
         Height          =   255
         Left            =   75
         TabIndex        =   23
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   525
         TabIndex        =   22
         Top             =   960
         Width           =   780
      End
      Begin VB.Label label2 
         Caption         =   "PO Date :"
         Height          =   255
         Left            =   12195
         TabIndex        =   21
         ToolTipText     =   "Enter Value Date"
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier :"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   570
         Width           =   945
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
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
               Picture         =   "frmPOPurchaseOrder.frx":091F
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOPurchaseOrder.frx":0D73
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOPurchaseOrder.frx":11C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOPurchaseOrder.frx":161B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOPurchaseOrder.frx":1A6F
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOPurchaseOrder.frx":1EC3
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPOPurchaseOrder.frx":2617
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Save"
         Height          =   270
         Left            =   3180
         TabIndex        =   41
         Top             =   720
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   15
      TabIndex        =   26
      Top             =   1935
      Width           =   15180
      Begin VB.CheckBox chkdiscpamt 
         Caption         =   "Discount/Tax on Purchase Amount"
         Height          =   375
         Left            =   105
         TabIndex        =   47
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
         TabIndex        =   44
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
         TabIndex        =   42
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
         TabIndex        =   39
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Total Issue Value"
         Top             =   5610
         Width           =   1455
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdGRN 
         Height          =   5340
         Left            =   90
         TabIndex        =   7
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
         TabIndex        =   45
         Top             =   6045
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Flat Disc :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3855
         TabIndex        =   43
         Top             =   5640
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Sale Cost :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9645
         TabIndex        =   40
         Top             =   5625
         Width           =   885
      End
      Begin VB.Label Label20 
         Caption         =   " Net Amount :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   12615
         TabIndex        =   38
         Top             =   6045
         Width           =   1395
      End
      Begin VB.Label Label19 
         Caption         =   "Disc Amount :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   9465
         TabIndex        =   37
         Top             =   6045
         Width           =   1395
      End
      Begin VB.Label Label18 
         Caption         =   "SED Amount:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6465
         TabIndex        =   36
         Top             =   6045
         Width           =   1050
      End
      Begin VB.Label Label17 
         Caption         =   "GST Amount:"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6435
         TabIndex        =   35
         Top             =   5625
         Width           =   1485
      End
      Begin VB.Label Label11 
         Caption         =   " Total :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   13095
         TabIndex        =   34
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
Attribute VB_Name = "frmPOPurchaseOrder"
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


Private Function maxtranscode() As String
pr_dumy.Open "select max(transcode) as transcode from PO_POOrder where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
txtCustRef = ClientCoderef("005") + Right(maxtranscode, 6)
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
           If Trim(.TextMatrix(ln_cnt, 13)) <> "" Then
            If chkdiscpamt.Value = 1 Then
                .TextMatrix(ln_cnt, 15) = Val(.TextMatrix(ln_cnt, 11)) * Val(.TextMatrix(ln_cnt, 13)) / 100
            Else
                .TextMatrix(ln_cnt, 15) = Val(.TextMatrix(ln_cnt, 12)) * Val(.TextMatrix(ln_cnt, 13)) / 100
            End If
        End If
        
        If Trim(.TextMatrix(ln_cnt, 14)) <> "" Then
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
Mode = DentMode(Mode, 1, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
Command1.Enabled = False
InitializeGrid
txttransno = maxtranscode
txttransno.Enabled = False
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
Mode = DentMode(Mode, 4, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
End Sub
Private Sub editrecord_Click()
Mode = DentMode(Mode, 2, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
Command1.Enabled = True
txttransno.Enabled = True
txttransno.SetFocus
End Sub
Private Sub Deleterecord_Click()
Mode = DentMode(Mode, 3, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
If txtVendorCode.Enabled Then txtVendorCode.SetFocus
Command1.Enabled = True
txttransno.Enabled = True
End Sub





Private Sub txtacode_LostFocus()
'    If txtacode <> "" Then Call txtACode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtacode1_LostFocus()
  '  If txtacode1 <> "" Then Call txtACode1_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtflatdisc_LostFocus()
If txtflatdisc <> "" Then
    TotalAmount
End If
End Sub
Private Sub txttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvaluedate.SetFocus
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
            If TxtRemarks.Enabled Then TxtRemarks.SetFocus
            
            
            
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
    txtitemname = LoadLastRate(.TextMatrix(.Row, 18))
    If .TextMatrix(.Row, 4) = "GODOWN" Then
        txtstock = Val(.TextMatrix(.Row, 20))
    ElseIf .TextMatrix(.Row, 4) = "SHOWROOM" Then
        txtstock = Val(.TextMatrix(.Row, 20))
    End If
    txtbonusamount = .TextMatrix(.Row, 19)
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
        Set PO_CODE = txttransno
        Set PO_DESC = Text1
        Gs_SQL = "SELECT GRN.TransCode AS ComputerCode, GRN.GRNCode AS GRNCode, Vendors.Description AS 'Vendors.Description', GRN.TransDate AS GRNDate,    GRN.NetAmount AS 'GRN.NetAmount' FROM         PO_POOrder GRN INNER JOIN         IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode"
        Gs_OrderBy = "ORDER BY GRN.TransCode desc"
        Gs_OtherPara = " Where GRN.compcode = '" & Gs_compcode & "' and GRN.glstatus = 0 "
        
        frmPosearchRecords.Caption = "GRN"
        frmPosearchRecords.Show 1
        
        If txttransno <> "" Then Call txtTransNo_KeyDown(vbKeyReturn, vbKeyShift)

End Sub
Private Sub Form_Load()
 Opt = ""
 
  SetToolBar(1) = chkRights("ICISUSTP01")
  SetToolBar(2) = chkRights("ICISUSTP02")
  SetToolBar(3) = chkRights("ICISUSTP03")
  SetToolBar(4) = chkRights("ICISUSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
'  PR_IcItem.Open "Select * from Ic_Item where compcode ='" & Gs_compcode & "' ", gc_dbcon, adOpenDynamic, adLockPessimistic, 1
  'PR_ICIssue.Open "Select * from Ic_TransMaster where compcode ='" & Gs_compcode & "' and transtype in ('D')  order by Transcode", gc_dbcon, adOpenDynamic, adLockOptimistic
  

  txtvaluedate.Value = Date
 
  InitializeGrid


  
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    PR_ICIssue.Close
   ' PR_IcItem.Close
   ' PR_Branch.Close
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
GrdGRN.SetFocus
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
 If KeyCode = vbKeyReturn And Trim(txttransno.Text) <> "" Then
         If PR_ICIssue.State = 1 Then PR_ICIssue.Close
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         PR_ICIssue.Open "select * from PO_POOrder where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and status = 0", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       Select Case Mode
            Case "A"
                If Not PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   If txttransno.Enabled Then txttransno.SetFocus
                Else
                   txtvaluedate.SetFocus
                End If
            Case Else
                If PR_ICIssue.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txttransno.SetFocus
                Else
                   Call SetVal
                   LoadGRNTrans
                   If Mode <> "D" Then
                      txttransno.SetFocus
                   End If
                End If
            End Select
 ElseIf KeyCode = vbKeyReturn And Trim(txttransno.Text) = "" Then
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
       Mode = DentMode(Mode, Button.Index, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
    End If
    If Mode = "A" Then
       Command1.Enabled = False
       txttransno = maxtranscode
       txttransno.Enabled = False
       If txtType.Enabled Then txtType.SetFocus
      CheckLogTrans
    Else
       txttransno.Enabled = True
       Command1.Enabled = True
        
    End If
End Sub


Public Sub SaveValues()
'On Error GoTo RollBack
Dim ln_cnt As Integer
Dim ls_sql As String
Dim ls_transtype As String



     Select Case Mode
           Case "D"
           
           
           
              gc_dbcon.Execute "DELETE FROM PO_POOrder WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
              gc_dbcon.Execute "DELETE FROM PO_POOrderDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"

           
              
           Case Else
                If Mode = "E" Then
                          gc_dbcon.Execute "DELETE FROM PO_POOrder WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                          gc_dbcon.Execute "DELETE FROM PO_POOrderDetail WHERE CompCode = '" & Gs_compcode & "' AND transcode = '" & Trim(txttransno) & "'"
                End If
                If Mode = "A" Then
                    txttransno = maxtranscode
                End If
                      gc_dbcon.BeginTrans
                            ls_sql = "INSERT into PO_POOrder(Compcode,branchcode, TransCode, TransDate, AccountCode,Remarks,userid,adddate,addtime,type,flatdisc)"
                            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & txtVendorCode & "','" & RepApp(TxtRemarks) & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "'," & txtType.ListIndex & " ," & Val(txtflatdisc) & "  )"
                            gc_dbcon.Execute ls_sql
                      gc_dbcon.CommitTrans
                
                With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      If .TextMatrix(ln_cnt, 1) <> "" Then
                        gc_dbcon.BeginTrans
                            ls_sql = "INSERT into PO_POOrderDetail(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,srno)"
                            ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 18)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 5))) & "," & Val(.TextMatrix(ln_cnt, 7)) & "," & Val(.TextMatrix(ln_cnt, 11)) & "," & Val(.TextMatrix(ln_cnt, 13)) & "," & Val(.TextMatrix(ln_cnt, 14)) & "," & Val(.TextMatrix(ln_cnt, 15)) & "," & Val(.TextMatrix(ln_cnt, 16)) & "," & Val(.TextMatrix(ln_cnt, 8)) & "," & Val(.TextMatrix(ln_cnt, 9)) & ",'" & Trim(.TextMatrix(ln_cnt, 17)) & "'," & Val(.TextMatrix(ln_cnt, 6)) & "," & Val(.TextMatrix(ln_cnt, 10)) & "," & ln_cnt & ")"
                            gc_dbcon.Execute ls_sql
                        gc_dbcon.CommitTrans
                     End If
                    Next
                      ls_sql = "delete from PO_POOrderDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'"
                      gc_dbcon.Execute ls_sql
                 End With
                 
                 
     End Select
     


If Mode <> "D" Then
   ls_opt = MsgBox("Print Purchase Order # ?.", vbYesNo)
   If ls_opt = vbYes Then Call PrintPOnote
End If

If Mode = "A" Then
    txttransno = maxtranscode
End If
InitializeGrid

If Mode = "E" Or Mode = "D" Then
txttransno = ""
End If

Exit Sub
RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub PrintPOnote()
On Error GoTo LocalErr

   With rptVoucher
        .WindowTitle = Me.Caption
        .ReportFileName = App.Path & Gs_ICRepoPath & "\" & "POPurchaseOrder.rpt"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(2) = "Reportname = 'Purchase Order'"
        .SelectionFormula = "{PO_POOrderNote.compcode} = '" & Gs_compcode & "'"
        .SelectionFormula = .SelectionFormula & "  and {PO_POOrderNote.transcode} = '" & Trim(txttransno) & "'"
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
     txtType.ListIndex = PR_ICIssue("type")
     txtVendorCode = Trim(PR_ICIssue("AccountCode") & "")
     Call txtVendorCode_KeyDown(vbKeyReturn, vbKeyShift)
     TxtRemarks = Trim(PR_ICIssue("Remarks") & "")
     txtflatdisc = Val(PR_ICIssue("FlatDisc"))
     TotalAmount
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
    If Trim(txttransno) = "" Then
      Call MsgBox("Enter/Select PO # !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtVendorCode) = "" Then
      Call MsgBox("Enter/Select Vendor Code !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(txtType) = "" Then
      Call MsgBox("Enter/Select  Type !!!", vbCritical)
      ChkInputs = False
    ElseIf Trim(TxtRemarks) = "" Then
      Call MsgBox("Enter Remarks !!!", vbCritical)
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
                Call MsgBox("QTY/Rate must be entered !!!", vbCritical)
                lb_opt = False
                Exit For
             Else
                lb_opt = True
             End If
            End If
          Next
        End With
       Call CalTax
       ChkInputs = lb_opt
    End If
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
        .FormatString = "Sr# |<Custom Code|<Item Name|<UOM|<Site|<Qty|<B-Qty|<Rate|<Disc%|<Disc Amount|<B-Amount|<Total|<Net Amount|<GST%|<FED%|<GST Amount|<FED Amount|<Remarks|<ItemCode|<SaleCost|<Stock"
        .ColWidth(1) = 1500
        .ColWidth(2) = 3800
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        
        
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
        .ColWidth(17) = 1500
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 0
        
        .Redraw = True
    End With
    

    
End Sub


Private Sub TotalAmount()
      txttotalamount = ""
      txtgstamount = ""
      txtsedamount = ""
      txtdiscamount = ""
     ' txtbonusamount = ""
      txtNetAmount = ""
    With GrdGRN
        For ln_cnt = 1 To .Rows - 1
            txttotalamount = Round(Val(txttotalamount) + Val(.TextMatrix(ln_cnt, 11)), 2)
            'txtbonusamount = Val(txtbonusamount) + Val(.TextMatrix(ln_cnt, 10))
            txtgstamount = Round(Val(txtgstamount) + Val(.TextMatrix(ln_cnt, 15)), 2)
            txtsedamount = Round(Val(txtsedamount) + Val(.TextMatrix(ln_cnt, 16)), 2)
            txtdiscamount = Round(Val(txtdiscamount) + Val(.TextMatrix(ln_cnt, 9)), 2)
        Next
    End With
    txtNetAmount = Round((Val(txttotalamount) + Val(txtgstamount) + Val(txtsedamount)) - (Val(txtdiscamount) + Val(txtflatdisc)), 0)
    txtNetAmount = Round(Val(txtNetAmount) + Val(txtfreight) + Val(txtloading) + Val(txtmiscCharges) + Val(txtlabour), 0)
End Sub
Private Sub LoadGRNTrans()
'On Error GoTo LocalErr

Dim Pr_LoadTrans As New Recordset
InitializeGrid
Dim ls_sql As String

ls_sql = "delete from PO_POOrderDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'"
gc_dbcon.Execute ls_sql
       

ls_sql = " SELECT PO_POOrderDetail.CustomCode,PO_POOrderDetail.ItemCode, IC_Item.salecost,IC_Item.Description, PO_POOrderDetail.Quantity, PO_POOrderDetail.Rate, PO_POOrderDetail.Amount, PO_POOrderDetail.BonusQty, PO_POOrderDetail.BonusAmount ,PO_POOrderDetail.GSTper,PO_POOrderDetail.Sedper,PO_POOrderDetail.GSTAmount,PO_POOrderDetail.SedAmount,PO_POOrderDetail.Discper,PO_POOrderDetail.Discamount,PO_POOrderDetail.Remarks, IC_ItemUM.Description AS UOM"
ls_sql = ls_sql & " FROM PO_POOrderDetail INNER JOIN IC_Item ON PO_POOrderDetail.Compcode = IC_Item.Compcode AND PO_POOrderDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where PO_POOrderDetail.Compcode = '" & Gs_compcode & "' and PO_POOrderDetail.Transcode = '" & txttransno & "' order by PO_POOrderDetail.srno"

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
                .TextMatrix(.Row, 17) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 18) = Trim(Pr_LoadTrans("Itemcode") & "")
                .TextMatrix(.Row, 19) = Val(Pr_LoadTrans("SaleCost"))
                .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - (Val(.TextMatrix(.Row, 9)) + Val(.TextMatrix(.Row, 10)))
                
               .CellBackColor = vbWindowBackground
                ls_sql = "INSERT into PO_POOrderDetailLog(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,Siteid,SRNo,Computername,emode)"
                ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 18)) & "'," & (Val(0 & .TextMatrix(.Row, 5))) & "," & Val(.TextMatrix(.Row, 7)) & "," & Val(.TextMatrix(.Row, 11)) & "," & Val(.TextMatrix(.Row, 13)) & "," & Val(.TextMatrix(.Row, 14)) & "," & Val(.TextMatrix(.Row, 15)) & "," & Val(.TextMatrix(.Row, 16)) & "," & Val(.TextMatrix(.Row, 8)) & "," & Val(.TextMatrix(.Row, 9)) & ",'" & Trim(.TextMatrix(.Row, 17)) & "'," & Val(.TextMatrix(.Row, 6)) & "," & Val(.TextMatrix(.Row, 10)) & "," & ls_siteopt & "," & .Row & " ,'" & Gs_ComputerName & "','" & Mode & "' )"
                gc_dbcon.Execute ls_sql
       
                
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

ls_sql = " SELECT PO_POOrderDetail.CustomCode,PO_POOrderDetail.ItemCode,PO_POOrderDetail.siteid, IC_Item.salecost,IC_Item.Description, PO_POOrderDetail.Quantity, PO_POOrderDetail.Rate, PO_POOrderDetail.Amount, PO_POOrderDetail.BonusQty, PO_POOrderDetail.BonusAmount ,PO_POOrderDetail.GSTper,PO_POOrderDetail.Sedper,PO_POOrderDetail.GSTAmount,PO_POOrderDetail.SedAmount,PO_POOrderDetail.Discper,PO_POOrderDetail.Discamount,PO_POOrderDetail.Remarks, IC_ItemUM.Description AS UOM"
ls_sql = ls_sql & "  FROM PO_POOrderDetailLog PO_POOrderDetail INNER JOIN IC_Item ON PO_POOrderDetail.Compcode = IC_Item.Compcode AND PO_POOrderDetail.ItemCode = IC_Item.ItemCode INNER JOIN  IC_ItemUM ON IC_Item.MCode = IC_ItemUM.Mcode"
ls_sql = ls_sql & "  where PO_POOrderDetail.Compcode = '" & Gs_compcode & "' and PO_POOrderDetail.computername ='" & Gs_ComputerName & "' order by PO_POOrderDetail.srno"

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
                .TextMatrix(.Row, 17) = Trim(Pr_LoadTrans("Remarks") & "")
                .TextMatrix(.Row, 18) = Trim(Pr_LoadTrans("Itemcode") & "")
                .TextMatrix(.Row, 19) = Val(Pr_LoadTrans("SaleCost"))
                .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - (Val(.TextMatrix(.Row, 9)) + Val(.TextMatrix(.Row, 10)))
                .Rows = .Rows + 1
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
pr_dumyLog.Open "select * from PO_POOrderDetailLog  where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyLog.EOF Then
    If pr_dumyLog("Emode") = "E" Then
        txttransno = pr_dumyLog("Transcode")
        res = MsgBox(txttransno & " # you have opened in edit mode not save Do you want to open now", vbYesNo + vbExclamation)
        If res = vbYes Then
        Mode = DentMode(Mode, 2, PR_ICIssue, Me, txttransno, txttransno, Para_Rs, "IC_ISSuCnt", 10, "txtTransNo", "text1", 1, False, Toolbar1)
        If PR_ICIssue.State = 1 Then PR_ICIssue.Close
        PR_ICIssue.Open "select * from PO_POOrder where compcode = '" & Gs_compcode & "' and Transcode = '" & txttransno & "' and status = 0 ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not PR_ICIssue.EOF Then
        Call SetVal
        End If
        PR_ICIssue.Close
        LoadLogTrans
        Else
        
        ls_sql = "delete from  PO_POOrderDetailLog where computername = '" & Gs_ComputerName & "' "
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
    Gs_SQL = Gs_SQL & " (((PO_POOrderDetail.Amount - PO_POOrderDetail.DiscAmount) + PO_POOrderDetail.GSTAmount) / (PO_POOrderDetail.Quantity + PO_POOrderDetail.BonusQty)) AS NetRate,PO_POOrderDetail.Quantity , PO_POOrderDetail.BonusQty,  PO_POOrderDetail.Rate  FROM  PO_POOrder GRN INNER JOIN"
    Gs_SQL = Gs_SQL & " IC_Supplier Vendors ON GRN.Compcode = Vendors.Compcode AND GRN.AccountCode = Vendors.SupplierCode INNER JOIN"
    Gs_SQL = Gs_SQL & " PO_POOrderDetail ON GRN.Compcode = PO_POOrderDetail.Compcode AND GRN.TransCode = PO_POOrderDetail.TransCode LEFT OUTER JOIN"
    Gs_SQL = Gs_SQL & " IC_Item ON PO_POOrderDetail.Compcode = IC_Item.Compcode AND PO_POOrderDetail.ItemCode = IC_Item.ItemCode"
    Gs_OtherPara = " where GRN.compcode = '" & Gs_compcode & "' and PO_POOrderDetail.Itemcode = '" & GrdGRN.TextMatrix(GrdGRN.Row, 18) & "'"
    Gs_OrderBy = "ORDER BY GRN.TransCode DESC"
    frmPosearchRecords.Caption = "Vendor Rate Comparison"
    frmPosearchRecords.Show 1
 
 ElseIf KeyCode = vbKeyDelete Then 'Delete Key Pressed
    With GrdGRN
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             ls_sql = "delete from PO_POOrderDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' and srno = " & .Row & "  "
             gc_dbcon.Execute ls_sql
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
             ls_sql = "delete from PO_POOrderDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' "
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
    txtstock = Val(.TextMatrix(.Row, 20))
    txtbonusamount = .TextMatrix(.Row, 19)
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
      End If
      If UCase(.Text) = "G" Then
            .TextMatrix(.Row, .Col) = "GODOWN"
            txtstock = CheckBalQTY(.TextMatrix(.Row, 18), 1)
      ElseIf UCase(.Text) = "S" Then
            .TextMatrix(.Row, .Col) = "SHOWROOM"
            txtstock = CheckBalQTY(.TextMatrix(.Row, 18), 2)
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
        If txttransno = "" Then
            Call MsgBox("Enter Invoice No !!!", vbCritical)
            txttransno.SetFocus
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
             
             .TextMatrix(.Row, 0) = .Row
             .TextMatrix(.Row, 18) = Trim(PR_IcItem("Itemcode") & "")
             .TextMatrix(.Row, 2) = Trim(PR_IcItem("Description") & "")
             .TextMatrix(.Row, 7) = Val(PR_IcItem("Purchasecost"))
             .TextMatrix(.Row, 19) = Val(PR_IcItem("Salecost"))
              txtbonusamount = .TextMatrix(.Row, 19)
              txtstock = ""
              txtitemname = .TextMatrix(.Row, 2)
              .TextMatrix(.Row, 8) = LoadLastDisc(.TextMatrix(.Row, 18))
              
              
             .Col = .Col + 4
             ' txtsitetype.Move .Left + .CellLeft - CX, .Top + .CellTop - CY, .CellWidth - 20 ' MSFlexPOS.CellHeight - CZ
             ' txtsitetype.Visible = True
             ' ClickRow = .Row
             ' txtsitetype.SetFocus
              'If UCase(Gs_DBDataSource) = "RAHAT" Then
                .TextMatrix(.Row, 4) = "SHOWROOM"
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
   If .Col = 1 Or .Col = 17 Or .Col = 21 Or .Col = 22 Or .Col = 4 Or .Col = 5 Or .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 13 Or .Col = 14 Then
   If Len(Trim(.Text)) <> 0 Then  'If current cell is not empty then...
      .Text = Left(.Text, (Len(.Text) - 1)) 'Removing a character from the right side of the FlexGrid cell's text
   End If
    If Trim(.TextMatrix(.Row, 5)) <> "" Or Trim(.TextMatrix(.Row, 7)) <> "" Then
        .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 7))
        End If
        If Trim(.TextMatrix(.Row, 8)) <> "" Then
        .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 8)) * Val(.TextMatrix(.Row, 11)) / 100
        End If

        If Trim(.TextMatrix(.Row, 6)) <> "" Then
        .TextMatrix(.Row, 10) = Val(.TextMatrix(.Row, 6)) * Val(.TextMatrix(.Row, 7))
        End If
       
        .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - Val(.TextMatrix(.Row, 9))

       If Trim(.TextMatrix(.Row, 13)) <> "" Then
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
        
   
   
   
    ls_sql = "delete from PO_POOrderDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' and srno = " & .Row & "  "
    gc_dbcon.Execute ls_sql
    
    If .TextMatrix(.Row, 4) = "GODOWN" Then
     ls_siteopt = 1
    Else
     ls_siteopt = 2
    End If
    
    .TextMatrix(.Row, 20) = CheckBalQTY(.TextMatrix(.Row, 18), ls_siteopt)
     txtstock = Val(.TextMatrix(.Row, 20))
               
    ls_sql = "INSERT into PO_POOrderDetailLog(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,SRNo,Computername,emode)"
    ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 18)) & "'," & (Val(0 & .TextMatrix(.Row, 5))) & "," & Val(.TextMatrix(.Row, 7)) & "," & Val(.TextMatrix(.Row, 11)) & "," & Val(.TextMatrix(.Row, 13)) & "," & Val(.TextMatrix(.Row, 14)) & "," & Val(.TextMatrix(.Row, 15)) & "," & Val(.TextMatrix(.Row, 16)) & "," & Val(.TextMatrix(.Row, 8)) & "," & Val(.TextMatrix(.Row, 9)) & ",'" & Trim(.TextMatrix(.Row, 17)) & "'," & Val(.TextMatrix(.Row, 6)) & "," & Val(.TextMatrix(.Row, 10)) & "," & .Row & " ,'" & Gs_ComputerName & "' ,'" & Mode & "')"
    gc_dbcon.Execute ls_sql
    TotalAmount
   
   End If
End With
End If

If KeyAscii <> 27 And KeyAscii <> 8 Then
    With GrdGRN
      
      If .Col = 1 Or .Col = 17 Or .Col = 21 Or .Col = 22 Then
        If .CellBackColor = vbHighlight Then
         .Text = "": .CellBackColor = vbWindowBackground
        End If
        .Text = .Text & Chr(KeyAscii) 'Reset Value in Cell and Append the pressed character to the right.
      
      ElseIf .Col = 5 Or .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 13 Or .Col = 14 Then
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
        
      If .Col = 21 Then
        .TextMatrix(.Row, 22) = DateAdd("D", Val(.TextMatrix(.Row, 21)), txtvaluedate)
      End If
      
      
        If Val(.TextMatrix(.Row, 7)) > Val(.TextMatrix(.Row, 19)) Then
            Call MsgBox("Enter Purchase price less then sale price")
            .TextMatrix(.Row, 7) = 0
            Exit Sub
        End If
          
        If Trim(.TextMatrix(.Row, 5)) <> "" Or Trim(.TextMatrix(.Row, 7)) <> "" Then
        .TextMatrix(.Row, 11) = Val(.TextMatrix(.Row, 5)) * Val(.TextMatrix(.Row, 7))
        End If
        If Trim(.TextMatrix(.Row, 8)) <> "" Then
            
        .TextMatrix(.Row, 9) = Val(.TextMatrix(.Row, 8)) * Val(.TextMatrix(.Row, 11)) / 100
        End If

        If Trim(.TextMatrix(.Row, 6)) <> "" Then
        .TextMatrix(.Row, 10) = Val(.TextMatrix(.Row, 6)) * Val(.TextMatrix(.Row, 7))
        End If
       
        .TextMatrix(.Row, 12) = Val(.TextMatrix(.Row, 11)) - Val(.TextMatrix(.Row, 9))

        If Trim(.TextMatrix(.Row, 13)) <> "" Then
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
        
        ls_sql = "delete from PO_POOrderDetailLog where compcode = '" & Gs_compcode & "' and computername = '" & Gs_ComputerName & "' and srno = " & .Row & "  "
        gc_dbcon.Execute ls_sql
        
      
               
        ls_sql = "INSERT into PO_POOrderDetailLog(Compcode,BranchCode, TransCode, CustomCode,ItemCode, Quantity,Rate,Amount,GSTPer,SEDPER,GSTAmount,SEDAmount,DiscPer,DiscAmount,Remarks,BonusQty,BonusAmount,SRNo,Computername,eMode)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Trim(txttransno) & "','" & Trim(.TextMatrix(.Row, 1)) & "','" & Trim(.TextMatrix(.Row, 18)) & "'," & (Val(0 & .TextMatrix(.Row, 5))) & "," & Val(.TextMatrix(.Row, 7)) & "," & Val(.TextMatrix(.Row, 11)) & "," & Val(.TextMatrix(.Row, 13)) & "," & Val(.TextMatrix(.Row, 14)) & "," & Val(.TextMatrix(.Row, 15)) & "," & Val(.TextMatrix(.Row, 16)) & "," & Val(.TextMatrix(.Row, 8)) & "," & Val(.TextMatrix(.Row, 9)) & ",'" & Trim(.TextMatrix(.Row, 17)) & "'," & Val(.TextMatrix(.Row, 6)) & "," & Val(.TextMatrix(.Row, 10)) & "," & .Row & " ,'" & Gs_ComputerName & "','" & Mode & "')"
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
