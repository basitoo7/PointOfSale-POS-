VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemstp 
   Caption         =   "Item Setup"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmItemStp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6570
      _ExtentX        =   11589
      _ExtentY        =   1058
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Caption         =   "&Listing"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
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
               Picture         =   "FrmItemStp.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemStp.frx":075E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemStp.frx":0BB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemStp.frx":1006
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemStp.frx":145A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemStp.frx":18AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmItemStp.frx":2002
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7440
      Left            =   0
      TabIndex        =   23
      Top             =   600
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   13123
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Basic Info"
      TabPicture(0)   =   "FrmItemStp.frx":2456
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Timer1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TOfferMessage"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtsaleperLstNr"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&Misc Option"
      TabPicture(1)   =   "FrmItemStp.frx":2472
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Dept & Sub Categorey"
         Height          =   3735
         Left            =   -74880
         TabIndex        =   107
         Top             =   3600
         Visible         =   0   'False
         Width           =   6255
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   2055
            Left            =   0
            TabIndex        =   121
            Top             =   1560
            Width           =   6135
            Begin VB.CommandButton CmdSaveInfo 
               Caption         =   "Save Info"
               Height          =   375
               Left            =   2160
               TabIndex        =   136
               Top             =   1560
               Width           =   1695
            End
            Begin VB.CommandButton cmdSearch 
               Caption         =   "Search Info"
               Height          =   375
               Left            =   360
               TabIndex        =   135
               Top             =   1560
               Width           =   1695
            End
            Begin VB.CommandButton CmdReplacement 
               Caption         =   "Replace All Data"
               Height          =   375
               Left            =   3960
               TabIndex        =   134
               Top             =   1560
               Width           =   1815
            End
            Begin VB.TextBox txtClassDesc3 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   315
               Left            =   2160
               TabIndex        =   133
               TabStop         =   0   'False
               Top             =   720
               Width           =   2805
            End
            Begin VB.TextBox txtcatdesc3 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   315
               Left            =   2160
               TabIndex        =   132
               TabStop         =   0   'False
               Top             =   360
               Width           =   2805
            End
            Begin VB.TextBox txtpackdesc3 
               Appearance      =   0  'Flat
               ForeColor       =   &H00800000&
               Height          =   315
               Left            =   2160
               TabIndex        =   131
               TabStop         =   0   'False
               Top             =   1095
               Width           =   2805
            End
            Begin VB.TextBox txtclasscode3 
               Height          =   315
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   127
               Top             =   720
               Width           =   420
            End
            Begin VB.TextBox txtcatcode3 
               Height          =   315
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   126
               Top             =   360
               Width           =   435
            End
            Begin VB.TextBox txtpackcode3 
               Height          =   315
               Left            =   1320
               MaxLength       =   3
               TabIndex        =   125
               Top             =   1095
               Width           =   435
            End
            Begin VB.CommandButton Command17 
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
               Picture         =   "FrmItemStp.frx":248E
               Style           =   1  'Graphical
               TabIndex        =   124
               TabStop         =   0   'False
               Tag             =   "SKIP"
               Top             =   1095
               Width           =   315
            End
            Begin VB.CommandButton Command16 
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
               Picture         =   "FrmItemStp.frx":2600
               Style           =   1  'Graphical
               TabIndex        =   123
               TabStop         =   0   'False
               Tag             =   "SKIP"
               Top             =   720
               Width           =   315
            End
            Begin VB.CommandButton Command15 
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
               Picture         =   "FrmItemStp.frx":2772
               Style           =   1  'Graphical
               TabIndex        =   122
               TabStop         =   0   'False
               Tag             =   "SKIP"
               Top             =   360
               Width           =   315
            End
            Begin VB.Label LON3 
               Caption         =   "LON3"
               Height          =   255
               Left            =   5040
               TabIndex        =   139
               Top             =   1150
               Width           =   735
            End
            Begin VB.Label LON2 
               Caption         =   "LON2"
               Height          =   255
               Left            =   5040
               TabIndex        =   138
               Top             =   740
               Width           =   615
            End
            Begin VB.Label LON1 
               Caption         =   "LON1"
               Height          =   255
               Left            =   5040
               TabIndex        =   137
               Top             =   400
               Width           =   735
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               Caption         =   "Category Code :"
               Height          =   210
               Left            =   120
               TabIndex        =   130
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               Caption         =   "Dept. Code :"
               Height          =   210
               Left            =   405
               TabIndex        =   129
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Sub Cat. Code :"
               Height          =   210
               Left            =   165
               TabIndex        =   128
               Top             =   1095
               Width           =   1185
            End
         End
         Begin VB.TextBox txtClassDesc2 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   720
            Width           =   3645
         End
         Begin VB.TextBox txtcatdesc2 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   360
            Width           =   3645
         End
         Begin VB.TextBox txtpackdesc2 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   1095
            Width           =   3645
         End
         Begin VB.TextBox txtclasscode2 
            Height          =   315
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   114
            Top             =   720
            Width           =   420
         End
         Begin VB.TextBox txtcatcode2 
            Height          =   315
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   113
            Top             =   360
            Width           =   435
         End
         Begin VB.TextBox txtpackcode2 
            Height          =   315
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   112
            Top             =   1095
            Width           =   435
         End
         Begin VB.CommandButton Command14 
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
            Picture         =   "FrmItemStp.frx":28E4
            Style           =   1  'Graphical
            TabIndex        =   111
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1095
            Width           =   315
         End
         Begin VB.CommandButton Command13 
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
            Picture         =   "FrmItemStp.frx":2A56
            Style           =   1  'Graphical
            TabIndex        =   110
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   720
            Width           =   315
         End
         Begin VB.CommandButton Command12 
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
            Picture         =   "FrmItemStp.frx":2BC8
            Style           =   1  'Graphical
            TabIndex        =   109
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   360
            Width           =   315
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Category Code :"
            Height          =   210
            Left            =   120
            TabIndex        =   117
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Dept. Code :"
            Height          =   210
            Left            =   405
            TabIndex        =   116
            Top             =   360
            Width           =   885
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat. Code :"
            Height          =   210
            Left            =   165
            TabIndex        =   115
            Top             =   1095
            Width           =   1185
         End
      End
      Begin VB.TextBox txtsaleperLstNr 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   3850
         Width           =   630
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7065
         Left            =   120
         TabIndex        =   42
         Top             =   345
         Width           =   6330
         Begin VB.TextBox txtProfitPer 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2280
            MaxLength       =   5
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   3120
            Width           =   495
         End
         Begin VB.ComboBox CmbPrintRateNY 
            ForeColor       =   &H00FF0000&
            Height          =   330
            ItemData        =   "FrmItemStp.frx":2D3A
            Left            =   1280
            List            =   "FrmItemStp.frx":2D44
            Style           =   2  'Dropdown List
            TabIndex        =   105
            Top             =   6240
            Width           =   1500
         End
         Begin VB.TextBox txtmodifyDate 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   4400
            Locked          =   -1  'True
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   5880
            Width           =   1800
         End
         Begin VB.TextBox txtsaleperLstNrAfDisc 
            BackColor       =   &H00FFFF00&
            Height          =   315
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   5520
            Width           =   750
         End
         Begin VB.TextBox txtAfterDisc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1250
            Locked          =   -1  'True
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1515
         End
         Begin VB.CheckBox chkanallow 
            Caption         =   "Addition in Qty not Allowed"
            Height          =   210
            Left            =   1215
            TabIndex        =   91
            Top             =   6750
            Width           =   3180
         End
         Begin VB.TextBox txtpackdisc 
            Height          =   315
            Left            =   4410
            MaxLength       =   25
            TabIndex        =   88
            Top             =   5505
            Width           =   1725
         End
         Begin VB.TextBox txtpackqty 
            Height          =   315
            Left            =   4410
            MaxLength       =   25
            TabIndex        =   87
            Top             =   5085
            Width           =   1725
         End
         Begin VB.TextBox txtlmodify 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   5880
            Width           =   1515
         End
         Begin VB.TextBox txtnetprate 
            BackColor       =   &H00FFFF00&
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   83
            Top             =   3525
            Width           =   1500
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Update Stock"
            Height          =   315
            Left            =   4995
            TabIndex        =   82
            Top             =   3870
            Width           =   1185
         End
         Begin VB.TextBox txtstockinG 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   4410
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   4695
            Width           =   1770
         End
         Begin VB.TextBox txtstockinS 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   4425
            Locked          =   -1  'True
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   4275
            Width           =   1755
         End
         Begin VB.TextBox txtsaleper 
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   5415
            Locked          =   -1  'True
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   3135
            Width           =   630
         End
         Begin VB.TextBox txtreorderlabel 
            Height          =   315
            Left            =   4425
            MaxLength       =   25
            TabIndex        =   74
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   3870
            Width           =   570
         End
         Begin VB.TextBox txtdisperamt 
            Height          =   315
            Left            =   1260
            MaxLength       =   25
            TabIndex        =   72
            Top             =   4725
            Width           =   1485
         End
         Begin VB.TextBox txtdiscper 
            Height          =   315
            Left            =   1260
            MaxLength       =   25
            TabIndex        =   70
            Top             =   4335
            Width           =   1485
         End
         Begin VB.TextBox txtavgrate 
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            MaxLength       =   25
            TabIndex        =   68
            Top             =   3930
            Width           =   1500
         End
         Begin VB.CommandButton Command10 
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
            Left            =   5880
            Picture         =   "FrmItemStp.frx":2D51
            Style           =   1  'Graphical
            TabIndex        =   67
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   135
            Width           =   315
         End
         Begin VB.TextBox txtDesc 
            Height          =   315
            Left            =   1290
            MaxLength       =   250
            TabIndex        =   2
            Top             =   525
            Width           =   4920
         End
         Begin VB.TextBox txtcustomCode 
            BackColor       =   &H00FFFF00&
            Height          =   330
            Left            =   3465
            MaxLength       =   256
            TabIndex        =   1
            Tag             =   "SKIP"
            Top             =   150
            Width           =   2415
         End
         Begin VB.TextBox txtSaleCost 
            Height          =   315
            Left            =   4425
            MaxLength       =   25
            TabIndex        =   11
            Top             =   3135
            Width           =   975
         End
         Begin VB.CheckBox Chksalerate 
            Caption         =   "Change Sale Rate"
            Height          =   210
            Left            =   4620
            TabIndex        =   13
            Top             =   6405
            Width           =   1665
         End
         Begin VB.CheckBox chksaleqty 
            Caption         =   "Change Sale Qty"
            Height          =   210
            Left            =   4620
            TabIndex        =   12
            Top             =   6735
            Width           =   1560
         End
         Begin VB.CommandButton Command9 
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
            Left            =   1950
            Picture         =   "FrmItemStp.frx":2EC3
            Style           =   1  'Graphical
            TabIndex        =   54
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1980
            Width           =   315
         End
         Begin VB.TextBox txtmdesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   1980
            Width           =   3915
         End
         Begin VB.TextBox txtmcode 
            Height          =   315
            Left            =   1275
            MaxLength       =   6
            TabIndex        =   6
            Top             =   1980
            Width           =   660
         End
         Begin VB.TextBox txtpackcode 
            Height          =   315
            Left            =   1275
            MaxLength       =   3
            TabIndex        =   5
            Top             =   1620
            Width           =   435
         End
         Begin VB.TextBox txtpackdesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1620
            Width           =   4125
         End
         Begin VB.CommandButton Command8 
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
            Left            =   1725
            Picture         =   "FrmItemStp.frx":3035
            Style           =   1  'Graphical
            TabIndex        =   51
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1605
            Width           =   315
         End
         Begin VB.TextBox txtcatcode 
            Height          =   315
            Left            =   1275
            MaxLength       =   3
            TabIndex        =   4
            Top             =   885
            Width           =   435
         End
         Begin VB.TextBox txtcatdesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   885
            Width           =   4125
         End
         Begin VB.CommandButton Command7 
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
            Left            =   1725
            Picture         =   "FrmItemStp.frx":31A7
            Style           =   1  'Graphical
            TabIndex        =   49
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   870
            Width           =   315
         End
         Begin VB.TextBox txtcode 
            BackColor       =   &H00FFFF00&
            Height          =   315
            Left            =   1290
            MaxLength       =   6
            TabIndex        =   0
            Tag             =   "SKIPN"
            Top             =   180
            Width           =   690
         End
         Begin VB.ComboBox txtvalutionmethod 
            Height          =   330
            ItemData        =   "FrmItemStp.frx":3319
            Left            =   4425
            List            =   "FrmItemStp.frx":332C
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2760
            Width           =   1785
         End
         Begin VB.ComboBox txtitemtype 
            Height          =   330
            ItemData        =   "FrmItemStp.frx":3381
            Left            =   1275
            List            =   "FrmItemStp.frx":338B
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   2730
            Width           =   1515
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
            Left            =   1725
            Picture         =   "FrmItemStp.frx":33AE
            Style           =   1  'Graphical
            TabIndex        =   48
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1245
            Width           =   315
         End
         Begin VB.TextBox txtClassDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2070
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1245
            Width           =   4125
         End
         Begin VB.TextBox txtclasscode 
            Height          =   315
            Left            =   1275
            MaxLength       =   3
            TabIndex        =   3
            Top             =   1245
            Width           =   420
         End
         Begin VB.TextBox txtUcode 
            Height          =   315
            Left            =   1275
            MaxLength       =   3
            TabIndex        =   7
            Top             =   2355
            Width           =   435
         End
         Begin VB.TextBox txtUdesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2085
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   2340
            Width           =   4110
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
            Left            =   1740
            Picture         =   "FrmItemStp.frx":3520
            Style           =   1  'Graphical
            TabIndex        =   45
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   2340
            Width           =   315
         End
         Begin VB.TextBox Textx 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   4320
            MaxLength       =   35
            TabIndex        =   44
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   6240
            Visible         =   0   'False
            Width           =   180
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
            Left            =   2010
            Picture         =   "FrmItemStp.frx":3692
            Style           =   1  'Graphical
            TabIndex        =   43
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   180
            Width           =   315
         End
         Begin VB.TextBox TxtPurchaseCost 
            Height          =   315
            Left            =   1260
            MaxLength       =   25
            TabIndex        =   10
            Top             =   3135
            Width           =   1020
         End
         Begin Crystal.CrystalReport crrpt 
            Left            =   120
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            DiscardSavedData=   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowGroupTree=   -1  'True
            WindowShowCloseBtn=   -1  'True
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Print Rate :"
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   240
            TabIndex        =   100
            Top             =   6330
            Width           =   900
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   " After Disc  Net P-Rate %:"
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   120
            TabIndex        =   99
            Top             =   5520
            Width           =   1875
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   210
            Left            =   6120
            TabIndex        =   98
            Top             =   3550
            Width           =   150
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   " On Last Net Rate :"
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   3600
            TabIndex        =   97
            Top             =   3555
            Width           =   1365
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "After Disc Amt:"
            Height          =   210
            Left            =   120
            TabIndex        =   96
            Top             =   5160
            Width           =   1350
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   " Modify Date :"
            Height          =   285
            Left            =   3360
            TabIndex        =   95
            Top             =   5925
            Width           =   1035
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   210
            Left            =   2880
            TabIndex        =   94
            Top             =   3200
            Width           =   255
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Pack Disc :"
            Height          =   210
            Left            =   3555
            TabIndex        =   90
            Top             =   5520
            Width           =   795
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Qty In Pack :"
            Height          =   210
            Left            =   3465
            TabIndex        =   89
            Top             =   5130
            Width           =   900
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   " Modify By :"
            Height          =   210
            Left            =   300
            TabIndex        =   86
            Top             =   6000
            Width           =   855
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Net Last P-Rate"
            ForeColor       =   &H8000000D&
            Height          =   210
            Left            =   60
            TabIndex        =   84
            Top             =   3555
            Width           =   1170
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Stock In Gowdom :"
            Height          =   420
            Left            =   2820
            TabIndex        =   81
            Top             =   4740
            Width           =   1560
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            Caption         =   "Stock In Showroom :"
            Height          =   420
            Left            =   2820
            TabIndex        =   79
            Top             =   4320
            Width           =   1560
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   210
            Left            =   6060
            TabIndex        =   77
            Top             =   3195
            Width           =   150
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Reorder Level :"
            Height          =   210
            Left            =   3240
            TabIndex        =   75
            Top             =   3900
            Width           =   1110
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Disc Amount :"
            Height          =   210
            Left            =   165
            TabIndex        =   73
            Top             =   4755
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Disc Per :"
            Height          =   210
            Left            =   480
            TabIndex        =   71
            Top             =   4350
            Width           =   690
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "AVG Rate :"
            Height          =   210
            Left            =   345
            TabIndex        =   69
            Top             =   3945
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Sale Rate :"
            Height          =   210
            Left            =   3585
            TabIndex        =   66
            Top             =   3150
            Width           =   780
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Supplier Code :"
            Height          =   210
            Left            =   135
            TabIndex        =   65
            Top             =   2025
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat. Code :"
            Height          =   210
            Left            =   120
            TabIndex        =   64
            Top             =   1650
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dept. Code :"
            Height          =   210
            Left            =   360
            TabIndex        =   63
            Top             =   915
            Width           =   885
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Name :"
            Height          =   210
            Left            =   765
            TabIndex        =   62
            Top             =   555
            Width           =   495
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Valution Method :"
            Height          =   210
            Left            =   3135
            TabIndex        =   61
            Top             =   2775
            Width           =   1245
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Item Type :"
            Height          =   210
            Left            =   465
            TabIndex        =   60
            Top             =   2760
            Width           =   780
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Category Code :"
            Height          =   210
            Left            =   75
            TabIndex        =   59
            Top             =   1275
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Unit :"
            Height          =   210
            Left            =   900
            TabIndex        =   58
            Top             =   2400
            Width           =   360
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Code :"
            Height          =   210
            Left            =   165
            TabIndex        =   57
            Top             =   195
            Width           =   1095
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Custom Code :"
            Height          =   210
            Left            =   2400
            TabIndex        =   56
            Top             =   195
            Width           =   1050
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Last Pur-Rate :"
            Height          =   210
            Left            =   165
            TabIndex        =   55
            Top             =   3150
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H8000000D&
         Height          =   3120
         Left            =   -74925
         TabIndex        =   24
         Top             =   375
         Width           =   6300
         Begin VB.CheckBox ChkNewDept 
            Caption         =   "New & Sub info"
            Height          =   375
            Left            =   4320
            TabIndex        =   108
            Top             =   2520
            Width           =   1575
         End
         Begin VB.ComboBox txtsaletaxoption 
            Height          =   330
            ItemData        =   "FrmItemStp.frx":3804
            Left            =   1695
            List            =   "FrmItemStp.frx":3811
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   210
            Width           =   2010
         End
         Begin VB.ComboBox txtpurchasetaxoption 
            Height          =   330
            ItemData        =   "FrmItemStp.frx":383D
            Left            =   1695
            List            =   "FrmItemStp.frx":384A
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   960
            Width           =   1965
         End
         Begin VB.TextBox txtSaleTaxCode 
            Height          =   315
            Left            =   1695
            MaxLength       =   3
            TabIndex        =   15
            Top             =   585
            Width           =   450
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
            Left            =   2160
            Picture         =   "FrmItemStp.frx":3874
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   585
            Width           =   315
         End
         Begin VB.TextBox txtPurchaseTaxCode 
            Height          =   315
            Left            =   1695
            MaxLength       =   3
            TabIndex        =   17
            Top             =   1350
            Width           =   465
         End
         Begin VB.CommandButton Command4 
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
            Left            =   2175
            Picture         =   "FrmItemStp.frx":39E6
            Style           =   1  'Graphical
            TabIndex        =   31
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox txtsaleschdesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2505
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   30
            Top             =   585
            Width           =   3405
         End
         Begin VB.TextBox txtpurchaseschdesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1350
            Width           =   3405
         End
         Begin VB.TextBox txtSiteID 
            Height          =   315
            Left            =   1695
            MaxLength       =   3
            TabIndex        =   18
            Top             =   1755
            Width           =   435
         End
         Begin VB.TextBox txtBinID 
            Height          =   315
            Left            =   1710
            MaxLength       =   3
            TabIndex        =   19
            Top             =   2145
            Width           =   435
         End
         Begin VB.TextBox txtsed 
            Height          =   315
            Left            =   1710
            MaxLength       =   25
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Text            =   "0"
            Top             =   2505
            Width           =   555
         End
         Begin VB.TextBox txtSiteDesc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1755
            Width           =   3405
         End
         Begin VB.CommandButton Command5 
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
            Left            =   2160
            Picture         =   "FrmItemStp.frx":3B58
            Style           =   1  'Graphical
            TabIndex        =   27
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   1740
            Width           =   315
         End
         Begin VB.TextBox TxtbinDesc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2535
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   26
            Top             =   2145
            Width           =   3390
         End
         Begin VB.CommandButton Command6 
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
            Left            =   2175
            Picture         =   "FrmItemStp.frx":3CCA
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   2145
            Width           =   315
         End
         Begin VB.TextBox txtmaxsaleqty 
            Height          =   315
            Left            =   3510
            MaxLength       =   25
            TabIndex        =   21
            TabStop         =   0   'False
            Tag             =   "SKIPN"
            Text            =   "0"
            Top             =   2505
            Width           =   660
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Sale Tax Option :"
            Height          =   210
            Left            =   435
            TabIndex        =   41
            Top             =   225
            Width           =   1230
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Purchase Tax Option :"
            Height          =   210
            Left            =   60
            TabIndex        =   40
            Top             =   1005
            Width           =   1605
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Tax Schedule ID :"
            Height          =   210
            Left            =   405
            TabIndex        =   39
            Top             =   645
            Width           =   1260
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Tax Schedule ID :"
            Height          =   210
            Left            =   405
            TabIndex        =   38
            Top             =   1380
            Width           =   1260
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Site ID :"
            Height          =   210
            Left            =   1125
            TabIndex        =   37
            Top             =   1770
            Width           =   540
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Bin ID :"
            Height          =   210
            Left            =   1170
            TabIndex        =   36
            Top             =   2175
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "S.E.D :"
            Height          =   210
            Left            =   1170
            TabIndex        =   35
            Top             =   2535
            Width           =   480
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   210
            Left            =   2280
            TabIndex        =   34
            Top             =   2550
            Width           =   150
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Max Sale Qty :"
            Height          =   210
            Left            =   2490
            TabIndex        =   33
            Top             =   2535
            Width           =   1050
         End
      End
      Begin VB.TextBox TOfferMessage 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   92
         Top             =   7080
         Width           =   6315
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   60
         Top             =   6750
      End
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   210
      Left            =   0
      TabIndex        =   93
      Top             =   0
      Width           =   150
   End
   Begin VB.Menu filemenu 
      Caption         =   "File"
      Begin VB.Menu New_Record 
         Caption         =   "New Record"
         Shortcut        =   ^N
      End
      Begin VB.Menu Edit_Record 
         Caption         =   "Edit Record"
         Shortcut        =   ^E
      End
      Begin VB.Menu Delete_Record 
         Caption         =   "Delete Record"
         Shortcut        =   ^D
      End
      Begin VB.Menu Save_Record 
         Caption         =   "Save Record"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmItemstp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_itemsetup As New Recordset
Dim lb_found As Boolean
Dim pr_dumy As New Recordset
Dim pr_itemsetuplog As New Recordset
Dim ln_pcost As Double
Dim ln_ppcost As Double
Dim ln_pDiscPer As Double
Dim ln_pDiscAmt As Double




Private Sub ChkNewDept_Click()

If Gc_UserId = "admin" Then
   If ChkNewDept.Value = 1 Then
      Frame3.Visible = True
   Else
      Frame3.Visible = False
   End If
End If

   

End Sub

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select IC_Item.ItemCode,   IC_Item.Description, IC_ItemCategory.Description as Category,IC_Item.SaleCost,StockS,StockG,isnull(StockS,0)+isnull(StockG,0) as TotalStock from IC_Item left outer join IC_ItemCategory on IC_Item.compcode = IC_ItemCategory.compcode and   IC_Item.catcode = IC_ItemCategory.catcode "
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "Order by IC_Item.Description"
    Gs_OtherPara = " where IC_Item.compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1

    If Len(txtCode) > 0 Then txtCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub CmdReplacement_Click()
   
   If Trim(txtpackcode2) <> "" Then
      ls_sql = "UPDATE IC_Item SET Catcode = '" & txtcatcode3 & "', ClassID= '" & txtclasscode3.Text & "',PackCode = '" & txtpackcode3 & "'  WHERE  CatCode = '" & txtcatcode2 & "' and  compcode = '" & Gs_compcode & "' and ClassID= '" & txtclasscode2.Text & "' and PackCode = '" & txtpackcode2.Text & "'"
      gc_dbcon.Execute ls_sql
   Else
       ls_sql = "UPDATE IC_Item SET Catcode = '" & txtcatcode3 & "', ClassID= '" & txtclasscode3.Text & "' WHERE  CatCode = '" & txtcatcode2 & "' and  compcode = '" & Gs_compcode & "' and ClassID= '" & txtclasscode2.Text & "'"
      gc_dbcon.Execute ls_sql
   End If
   
   
 MsgBox ("All Data Hasbeen Replaced ...."), vbCritical
CmdReplacement.Enabled = False

txtcatcode2.Text = ""
txtclasscode2.Text = ""
txtpackcode2.Text = ""
txtcatdesc2.Text = ""
txtClassDesc2.Text = ""
txtpackdesc2.Text = ""

txtcatcode3.Text = ""
txtclasscode3.Text = ""
txtpackcode3.Text = ""
txtcatdesc3.Text = ""
txtClassDesc3.Text = ""
txtpackdesc3.Text = ""


End Sub

Private Sub CmdSaveInfo_Click()
    '***************************************
       If LON1.Caption = "NEW" Then
           ls_sql = "INSERT into IC_ItemCategory(compcode,CatCode,Description) VALUES ('" & Gs_compcode & "','" & txtcatcode3 & "','" & txtcatdesc3.Text & "')"
           gc_dbcon.Execute ls_sql
         Else
           ls_sql = "UPDATE IC_ItemCategory SET Description= '" & RepApp(txtcatdesc3.Text) & "',profitper =0  WHERE  compcode = '" & Gs_compcode & "' and CatCode= '" & txtcatcode3.Text & "'"
           gc_dbcon.Execute ls_sql
       End If
    '***********************************************
       If LON2.Caption = "NEW" Then
             ls_sql = "INSERT into IC_ItemClass(compcode,deptcode,ClassCode,Description) VALUES ('" & Gs_compcode & "','" & txtcatcode3.Text & "','" & txtclasscode3.Text & "','" & RepApp(txtClassDesc3.Text) & "')"
              gc_dbcon.Execute ls_sql
       Else
             ls_sql = "UPDATE IC_ItemClass SET deptcode = '" & txtcatcode3 & "', Description= '" & RepApp(txtClassDesc3.Text) & "' WHERE  Deptcode = '" & txtcatcode2 & "' and  compcode = '" & Gs_compcode & "' and Classcode= '" & txtclasscode2.Text & "'"
              gc_dbcon.Execute ls_sql
       End If
     '****************************************************
   If Trim(txtpackcode2) <> "" Then
        
          If LON3.Caption = "NEW" Then
        
               ls_sql = "INSERT into IC_ItemPacking(compcode,Deptcode,subcode,PackCode,Description) VALUES ('" & Gs_compcode & "','" & txtcatcode3 & "','" & txtclasscode3 & "','" & txtpackcode3.Text & "','" & RepApp(txtClassDesc3.Text) & "')"
              gc_dbcon.Execute ls_sql
          Else
   
              ls_sql = "UPDATE IC_ItemPacking SET compcode = '" & Gs_compcode & "',deptcode = '" & txtcatcode3 & "', Subcode = '" & txtclasscode3 & "',packcode= '" & txtpackcode3.Text & "', Description= '" & RepApp(txtClassDesc3.Text) & "' WHERE compcode = '" & Gs_compcode & "'and deptcode = '" & txtcatcode2 & "' and subcode = '" & txtclasscode2 & "'  and   packcode= '" & txtpackcode2.Text & "'"
              gc_dbcon.Execute ls_sql
           End If
    Else
    
             If LON3.Caption = "NEW" Then
        
               ls_sql = "INSERT into IC_ItemPacking(compcode,Deptcode,subcode,PackCode,Description) VALUES ('" & Gs_compcode & "','" & txtcatcode3 & "','" & txtclasscode3 & "' ,'" & txtitemcode.Text & "','" & RepApp(txtClassDesc3.Text) & "')"
              gc_dbcon.Execute ls_sql
             Else
   
              ls_sql = "UPDATE IC_ItemPacking SET compcode = '" & Gs_compcode & "',deptcode = '" & txtcatcode3 & "', Subcode = '" & txtclasscode3 & "' WHERE compcode = '" & Gs_compcode & "'and deptcode = '" & txtcatcode2 & "' and subcode = '" & txtclasscode2 & "'"
              gc_dbcon.Execute ls_sql
             End If
    End If
    
    
     '**********************************************
       CmdSaveInfo.Enabled = False
End Sub

Private Sub cmdSearch_Click()
If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_ItemCategory where Upper(Description) Like  '%" & UCase(Trim(txtcatdesc3)) & "%' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        
      
     If pr_dumy.EOF Then
            Call MsgBox("Department code not found !!!", vbCritical)
            txtcatcode3 = ""
           ' txtcatdesc3 = ""
  ' *****************************************************
      If pr_dumy.State = 1 Then pr_dumy.Close
       pr_dumy.Open "select max(CatCode) as transcode from IC_ItemCategory where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If Not pr_dumy.EOF Then
          txtcatcode3 = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
        Else
          txtcatcode3 = DoPad(Trim(str(Int(1))), 3)
       End If
         pr_dumy.Close
         txtcatcode3.SetFocus
        LON1.Caption = "NEW"
        Else
           txtcatcode3 = pr_dumy("CatCode")
            txtcatdesc3 = pr_dumy("Description")
            If txtclasscode3.Enabled Then txtclasscode3.SetFocus
            LON1.Caption = "OLD"
        End If
        'pr_dumy.Close

  '****************************************************
       If pr_dumy.State = 1 Then pr_dumy.Close
       
       pr_dumy.Open "Select * from Ic_ItemClass where Upper(Description) Like  '%" & UCase(Trim(txtClassDesc3)) & "%' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode2 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtclasscode3 = ""
            'txtClassDesc3 = ""
    '****************************************************
            'Dim pr_dumy As New Recordset
            If pr_dumy.State = 1 Then pr_dumy.Close
           pr_dumy.Open "select max(ClassCode) as transcode from IC_ItemClass where compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If Not pr_dumy.EOF Then
             txtclasscode3 = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
           Else
            txtclasscode3 = DoPad(Trim(str(Int(1))), 3)
          End If
           LON2.Caption = "NEW"
       pr_dumy.Close
            
     '**************************************************
            txtclasscode3.SetFocus
        Else
            txtclasscode3 = pr_dumy("ClassCode")
            txtClassDesc3 = pr_dumy("Description")
             If txtpackcode3.Enabled Then txtpackcode3.SetFocus
            LON2.Caption = "OLD"
        End If
       ' pr_dumy.Close
        If pr_dumy.State = 1 Then pr_dumy.Close
       pr_dumy.Open "Select * from IC_ItemPacking  where Upper(Description) Like  '%" & UCase(Trim(txtpackdesc3)) & "%'  and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode2 & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtpackcode3 = ""
            'txtpackdesc3 = ""
         'Dim pr_dumy As New Recordset
         If pr_dumy.State = 1 Then pr_dumy.Close
           pr_dumy.Open "select max(PackCode) as transcode from IC_ItemPacking where compcode = '" & Gs_compcode & "' and subcode = '" & txtclasscode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If Not pr_dumy.EOF Then
             txtpackcode3 = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
            Else
              txtpackcode3 = DoPad(Trim(str(Int(1))), 3)
           End If
            LON3.Caption = "NEW"
         pr_dumy.Close
            
            txtpackcode3.SetFocus
        Else
            txtpackcode3 = pr_dumy("PackCode")
            txtpackdesc3 = pr_dumy("Description")
             LON3.Caption = "Old"
            
        End If
       If pr_dumy.State = 1 Then pr_dumy.Close

End Sub

Private Sub Command10_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcustomCode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select IC_Item.ItemCode,   IC_Item.Description, IC_ItemCategory.Description as Category,IC_Item.SaleCost from IC_Item left outer join IC_ItemCategory on IC_Item.compcode = IC_ItemCategory.compcode and   IC_Item.catcode = IC_ItemCategory.catcode "
    Gs_FindFld = "IC_Item.Description"
    Gs_OrderBy = "Order by IC_Item.Description"
    Gs_OtherPara = " where IC_Item.compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Items"
    
    MyLookupOLDB.Show 1

    If Len(txtcustomCode) > 0 Then txtcustomCode_KeyDown vbKeyReturn, vbKeyShift

End Sub
Private Sub CheckRateDifference()
Dim pr_dumynetrate As New Recordset
If pr_dumynetrate.State = 1 Then pr_dumynetrate.Close
pr_dumynetrate.Open "select Compcode from ic_item where avgrate > salecost", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumynetrate.EOF Then
res = MsgBox("Some Item(s) Sale Rate Less Then Net Rate Do you want to generate report", vbYesNo + vbInformation)
If res = vbYes Then Call PrintRateDiffRpt
End If
pr_dumynetrate.Close
End Sub
Private Sub PrintRateDiffRpt()
    With crrpt
        .ReportFileName = App.Path & Gs_ICRepoPath & "\IC_ItemsRateDiff.RPT"
        .WindowTitle = "Company Items"
        .SelectionFormula = "{Ic_item.Compcode} = '" & Gs_compcode & "'"
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "Reportname = 'Net Rate With Sale Rate Difference Report'"
        .SelectionFormula = ""
        
        .SelectionFormula = "{Ic_Item.compcode} = '" & Gs_compcode & "'"
        
        .SelectionFormula = .SelectionFormula & "  and {Ic_Item.AvgRate} > {Ic_Item.SaleCost}"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
       
    End With

End Sub

Private Sub Command11_Click()
gc_dbcon.Execute "UPDATE IC_Item SET   StockG = 0 ,StockS = 0"
gc_dbcon.Execute "UPDATE IC_Item SET   StockG = StockSummary.Qty FROM   StockSummary INNER JOIN   IC_Item ON StockSummary.ItemCode = IC_Item.ItemCode WHERE (StockSummary.siteid = 1)"
gc_dbcon.Execute "UPDATE IC_Item SET   StockS = StockSummary.Qty FROM   StockSummary INNER JOIN   IC_Item ON StockSummary.ItemCode = IC_Item.ItemCode WHERE (StockSummary.siteid = 2)"
Call MsgBox("Stock Successfully Update !!!", vbInformation)

If txtCode <> "" Then Call SetVal
End Sub







Private Sub Command12_Click()
Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode2
    Set PO_DESC = txtcatdesc2
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtcatcode2 <> "" Then Call txtcatcode2_KeyDown(vbKeyReturn, vbKeyShift)
End Sub



Private Sub Command13_Click()
 Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtclasscode2
    Set PO_DESC = txtClassDesc2
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode2 & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtclasscode2 <> "" Then Call txtclasscode2_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command14_Click()
Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpackcode2
    Set PO_DESC = txtpackdesc2
    Gs_SQL = "Select PackCode,   Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and subcode = '" & txtclasscode2 & "' and deptcode = '" & txtcatcode2 & "'"
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    
    If txtpackcode2 <> "" Then Call txtpackcode2_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command15_Click()
Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode3
    Set PO_DESC = txtcatdesc3
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtcatcode3 <> "" Then Call txtcatcode3_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command16_Click()
Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtclasscode3
    Set PO_DESC = txtClassDesc3
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode3 & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtclasscode3 <> "" Then Call txtclasscode3_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command17_Click()
Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpackcode3
    Set PO_DESC = txtpackdesc3
    Gs_SQL = "Select PackCode,   Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and subcode = '" & txtclasscode3 & "' and deptcode = '" & txtcatcode3 & "'"
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    
    If txtpackcode3 <> "" Then Call txtpackcode3_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtclasscode
    Set PO_DESC = txtClassDesc
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtclasscode <> "" Then Call txtclassCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSaleTaxCode
    Set PO_DESC = txtsaleschdesc
    Gs_SQL = "Select TaxCode, Description from IC_TaxSchedules "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where type =1 "
    MyLookupOLDB.Caption = "Sale Tax Schedule"
    MyLookupOLDB.Show 1
    
    If txtSaleTaxCode <> "" Then Call txtsaletaxcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub Command7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode
    Set PO_DESC = txtcatdesc
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtcatcode <> "" Then Call txtcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command8_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpackcode
    Set PO_DESC = txtpackdesc
    Gs_SQL = "Select PackCode,   Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and subcode = '" & txtclasscode & "' and deptcode = '" & txtcatcode & "'"
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    
    If txtpackcode <> "" Then Call txtpackCode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command9_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtmcode
    Set PO_DESC = txtmdesc
    Gs_SQL = "Select SupplierCode,   Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Suppliers"
    MyLookupOLDB.Show 1
    
    If txtmcode <> "" Then Call txtmcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Delete_record_Click()
Mode = DentMode(Mode, 3, PR_itemsetup, Me, txtcustomCode, txtdesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 1, False, Toolbar1)
       cmdLookup.Enabled = True
       SSTab1.Tab = 0
       txtCode.Enabled = True
       txtCode.SetFocus

End Sub

Private Sub Edit_record_Click()
Mode = DentMode(Mode, 2, PR_itemsetup, Me, txtcustomCode, txtdesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 1, False, Toolbar1)
       cmdLookup.Enabled = True
       SSTab1.Tab = 0
       txtCode.Enabled = True
       txtcustomCode.SetFocus
End Sub

Public Sub New_Record_Click()
    Mode = DentMode(Mode, 1, PR_itemsetup, Me, txtcustomCode, txtdesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 1, False, Toolbar1)
    If Mode = "A" Then
       cmdLookup.Enabled = False
       SSTab1.Tab = 0
       txtCode = maxtranscode
       txtcustomCode.SetFocus
    Else
      cmdLookup.Enabled = True
    End If

End Sub

Private Sub Save_Record_Click()
Mode = DentMode(Mode, 4, PR_itemsetup, Me, txtCode, txtdesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 1, False, Toolbar1)
End Sub





Private Sub txtcatcode2_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode2) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode2 = DoPad(txtcatcode2, 3)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_ItemCategory where Catcode = '" & txtcatcode2 & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Department code not found !!!", vbCritical)
            txtcatcode2 = ""
            txtcatdesc2 = ""
            txtcatcode2.SetFocus
        Else
            txtcatdesc2 = pr_dumy("Description")
            If txtclasscode2.Enabled Then txtclasscode2.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode2) = "" And KeyCode = vbKeyReturn Then
        txtcatcode2 = ""
        txtcatdesc2 = ""
        'Command7_Click
End If
End Sub

Private Sub txtcatcode3_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode3) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode3 = DoPad(txtcatcode3, 3)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_ItemCategory where Catcode = '" & txtcatcode3 & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Department code not found !!!", vbCritical)
            txtcatcode3 = ""
            txtcatdesc3 = ""
            txtcatcode3.SetFocus
        Else
            txtcatdesc3 = pr_dumy("Description")
            If txtclasscode3.Enabled Then txtclasscode3.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode3) = "" And KeyCode = vbKeyReturn Then
        txtcatcode3 = ""
        txtcatdesc3 = ""
        'Command7_Click
End If
End Sub

Private Sub txtclasscode2_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtclasscode2) <> "" And KeyCode = vbKeyReturn Then
        txtclasscode2 = DoPad(txtclasscode2, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
       
       pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode2 & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode2 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       
      '  pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtclasscode2 = ""
            txtClassDesc2 = ""
            txtclasscode2.SetFocus
        Else
            txtClassDesc2 = pr_dumy("Description")
             If txtpackcode2.Enabled Then txtpackcode2.SetFocus
           
        End If
        pr_dumy.Close
        
ElseIf Trim(txtclasscode2) = "" And KeyCode = vbKeyReturn Then
        txtclasscode2 = ""
        txtClassDesc2 = ""
        Command2_Click
End If
End Sub

Private Sub txtclasscode3_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtclasscode3) <> "" And KeyCode = vbKeyReturn Then
        txtclasscode3 = DoPad(txtclasscode3, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
       
       pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode3 & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode3 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtclasscode3 = ""
            txtClassDesc3 = ""
            txtclasscode3.SetFocus
        Else
            txtClassDesc3 = pr_dumy("Description")
             If txtpackcode3.Enabled Then txtpackcode3.SetFocus
           
        End If
        pr_dumy.Close
        
ElseIf Trim(txtclasscode3) = "" And KeyCode = vbKeyReturn Then
        txtclasscode3 = ""
        txtClassDesc3 = ""
        
End If
End Sub

'Private Sub Timer1_Timer()


'If Time() = "9:04:00 PM" Then
     
    'gc_dbcon.Execute ("Offer5PercApply")
    
    'TOfferMessage.Text = "5% Discount has been Applyed from 9-PM to Closing"
    
'    gc_dbcon.Execute ("OfferApply")
'     TOfferMessage.Text = "50% PIZZA OFFER Applyed from 12-AM to Closing"
'End If
 
' If TOfferMessage.Text <> "" Then
'    If TOfferMessage.Text = "PIZZA OFFER has been Applyed from 11AM to 3PM" Then
'    TOfferMessage.Text = " "
'   Else
'    TOfferMessage.Text = "PIZZA OFFER has been Applyed from 11AM to 3PM"
'   End If
'End If

'If Time() = "3:05:00 AM" Then
    
    'gc_dbcon.Execute ("Offer5PercRemove")
    'TOfferMessage.Text = "5% Discount has been Removed from 9-PM to Closing"
   
'   gc_dbcon.Execute ("OfferRemove")
'   TOfferMessage.Text = "50% PIZZA OFFER Removed from 12-AM till Closing"
'End If

'If TOfferMessage.Text <> "" Then
'If TOfferMessage.Text = "PIZZA OFFER has been Removed from 11AM to 3PM" Then
'   TOfferMessage.Text = " "
'Else
'   TOfferMessage.Text = "PIZZA OFFER has been Removed from 11AM to 3PM"
'End If
'End If
'If Val(TOfferMessage.Text) > 10 Then
'   TOfferMessage.Text = 0
'Else
 ' TOfferMessage.Text = Val(TOfferMessage.Text) + 1
'End If
'End Sub



Private Sub txtcustomCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtcustomCode) <> "" Then
     If pr_dumy.State = 1 Then pr_dumy.Close
     pr_dumy.Open "Select * from IC_item where customcode = '" & txtcustomCode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
     If Mode = "A" And Not pr_dumy.EOF Then
            Call MsgBox("Custom code already exist !!!", vbCritical)
            txtcustomCode = ""
            txtcustomCode.SetFocus
     ElseIf Mode <> "A" And Not pr_dumy.EOF Then
            txtCode = pr_dumy("Itemcode")
            pr_dumy.Close
            If txtCode <> "" Then Call txtCode_KeyDown(vbKeyReturn, vbKeyShift)
            If txtdesc.Enabled Then
            txtdesc.SetFocus
            txtdesc.SelStart = Len(txtdesc)
            End If
     ElseIf Mode <> "A" And pr_dumy.EOF Then
            Call MsgBox("Record not found !!!", vbCritical)
            txtcustomCode = ""
            txtcustomCode.SetFocus
            Exit Sub
     Else
            If txtdesc.Enabled Then
                txtdesc.SetFocus
                txtdesc.SelStart = Len(txtdesc)
            End If
     End If
ElseIf KeyCode = vbKeyReturn And Trim(txtcustomCode) = "" Then
    Command10_Click
End If

End Sub

Private Sub txtcustomCode_LostFocus()
If Trim(txtcustomCode) <> "" Then
     If pr_dumy.State = 1 Then pr_dumy.Close
     pr_dumy.Open "Select * from IC_item where customcode = '" & txtcustomCode & "' and compcode = '" & Gs_compcode & "' and itemcode <> '" & txtCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
     If Not pr_dumy.EOF Then
            Call MsgBox("Custom code already exist !!!", vbCritical)
            txtcustomCode = ""
            txtcustomCode.SetFocus
    End If
End If
End Sub

Private Sub txtDesc_LostFocus()
txtdesc = UCase(txtdesc)
txtSaleTaxCode = "001"
'Call txtsaletaxcode_KeyDown(vbKeyReturn, vbKeyShift)
txtPurchaseTaxCode = "002"
'Call txtPurchaseTaxCode_KeyDown(vbKeyReturn, vbKeyShift)
TxtSiteID = "001"
'Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
txtbinID = "001"
'Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
'txtclasscode.SetFocus
End Sub

Private Sub txtdisperamt_Change()

txtAfterDisc = Val(txtSaleCost) - Val(txtdisperamt)
If Val(txtdisperamt) > 0 And Val(txtnetprate) > 0 Then
   txtsaleperLstNrAfDisc = Round((Val(txtAfterDisc) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
Else
 If Val(txtnetprate) > 0 Then
   txtsaleperLstNrAfDisc = Round((Val(txtSaleCost) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
 End If
End If
End Sub

Private Sub txtmaxsaleqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtreorderlabel.SetFocus
End If
End Sub
Private Sub txtmcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtmcode) <> "" And KeyCode = vbKeyReturn Then
        txtmcode.Text = DoPad(txtmcode.Text, 6)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_Supplier where Suppliercode = '" & txtmcode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Supplier code not found !!!", vbCritical)
            txtmcode = ""
            txtmdesc = ""
            txtmcode.SetFocus
        Else
            txtmdesc = pr_dumy("Description")
            If txtUcode.Enabled Then txtUcode.SetFocus
            
        End If
        pr_dumy.Close
ElseIf Trim(txtmcode) = "" And KeyCode = vbKeyReturn Then
        txtmcode = ""
        txtmdesc = ""
        Command9_Click
End If
End Sub
Private Sub txtpackCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtpackcode) <> "" And KeyCode = vbKeyReturn Then
        txtpackcode.Text = DoPad(txtpackcode.Text, 3)
        pr_dumy.Open "Select * from IC_ItemPacking where Packcode = '" & txtpackcode & "' and subcode = '" & txtclasscode & "'  and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtpackcode = ""
            txtpackdesc = ""
            txtpackcode.SetFocus
        Else
            txtpackdesc = pr_dumy("Description")
            If txtmcode.Enabled Then txtmcode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtpackcode) = "" And KeyCode = vbKeyReturn Then
        txtpackcode = ""
        txtpackdesc = ""
        Command8_Click
End If
End Sub



Private Sub txtpackcode2_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtpackcode2) <> "" And KeyCode = vbKeyReturn Then
        txtpackcode2.Text = DoPad(txtpackcode2.Text, 3)
        pr_dumy.Open "Select * from IC_ItemPacking where Packcode = '" & txtpackcode2 & "' and subcode = '" & txtclasscode2 & "'  and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode2 & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtpackcode2 = ""
            txtpackdesc2 = ""
            txtpackcode2.SetFocus
        Else
            txtpackdesc2 = pr_dumy("Description")
            'If txtmcode.Enabled Then txtmcode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtpackcode2) = "" And KeyCode = vbKeyReturn Then
        txtpackcode2 = ""
        txtpackdesc2 = ""
        'Command8_Click
End If

End Sub

Private Sub txtpackcode3_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtpackcode3) <> "" And KeyCode = vbKeyReturn Then
        txtpackcode3.Text = DoPad(txtpackcode3.Text, 3)
        pr_dumy.Open "Select * from IC_ItemPacking where Packcode = '" & txtpackcode3 & "' and subcode = '" & txtclasscode3 & "'  and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode3 & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtpackcode3 = ""
            txtpackdesc3 = ""
            txtpackcode3.SetFocus
        Else
            txtpackdesc3 = pr_dumy("Description")
            'If txtmcode.Enabled Then txtmcode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtpackcode3) = "" And KeyCode = vbKeyReturn Then
        txtpackcode3 = ""
        txtpackdesc3 = ""
        'Command8_Click
End If

End Sub

Private Sub txtpackdisc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then chksaleqty.SetFocus
End Sub

Private Sub txtpackqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpackdisc.SetFocus
End Sub

Private Sub txtProfitPer_Change()
If Val(txtnetprate.Text) > 0 Then
txtSaleCost.Text = Val(TxtPurchaseCost.Text) + Round(Val(txtnetprate.Text) / 100 * Val(txtProfitPer), 2)
End If
End Sub

Private Sub txtProfitPer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
txtsaleper.SetFocus

txtSaleCost.SetFocus
txtSaleCost.SelLength = Len(txtSaleCost.Text)

If Val(txtnetprate.Text) > 0 Then
txtSaleCost.Text = Val(TxtPurchaseCost.Text) + Round(Val(txtnetprate.Text) / 100 * Val(txtProfitPer), 2)
End If

End If
End Sub

Private Sub TxtPurchaseCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

txtProfitPer.SetFocus

'txtSaleCost.SetFocus
'txtSaleCost.SelLength = Len(txtSaleCost.Text)
End If
End Sub

Private Sub txtreorderlabel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtdiscper.SetFocus
End If
End Sub

Private Sub txtSaleCost_Change()

If Val(TxtPurchaseCost) > 0 And Val(txtSaleCost) > 0 Then

'If Val(txtnetprate) > 0 And Val(txtSaleCost) > 0 Then
If Val(txtSaleCost) > 0 Then
   txtsaleper = Round((Val(txtSaleCost) - Val(TxtPurchaseCost)) / (Val(TxtPurchaseCost) / 100), 2)
End If
If Val(txtnetprate) > 0 Then
  txtsaleperLstNr = Round((Val(txtSaleCost) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
  txtsaleperLstNrAfDisc = Round((Val(txtAfterDisc) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
End If
End If
End Sub
Private Sub TxtPurchaseCost_Change()
If Val(txtnetprate) > 0 And Val(txtSaleCost) > 0 Then

'If Val(TxtPurchaseCost) > 0 And Val(txtSaleCost) > 0 Then

If Val(TxtPurchaseCost) > 0 Then
   txtsaleper = Round((Val(txtSaleCost) - Val(TxtPurchaseCost)) / (Val(TxtPurchaseCost) / 100), 2)
End If
If Val(txtnetprate) > 0 Then
  txtsaleperLstNr = Round((Val(txtSaleCost) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
   
    txtsaleperLstNrAfDisc = Round((Val(txtAfterDisc) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
   
End If
End If
End Sub
Private Sub txtsed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtmaxsaleqty.SetFocus
End Sub

Private Sub txtSaleCost_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtreorderlabel.SetFocus
End Sub
Private Sub txtavgrate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdiscper.SetFocus
End Sub
Private Sub txtdiscper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdisperamt.SetFocus
End Sub
Private Sub txtdisperamt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtpackqty.SetFocus
End Sub
Private Sub chksaleqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Chksalerate.SetFocus
End Sub

Private Sub txtsaletaxcode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtSaleTaxCode) <> "" And KeyCode = vbKeyReturn Then
        txtSaleTaxCode = DoPad(txtSaleTaxCode, 3)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_TaxSchedules where type = 1 and taxcode = '" & txtSaleTaxCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sale Tax Code not found !!!", vbCritical)
            txtSaleTaxCode = ""
            Textx = ""
            txtSaleTaxCode.SetFocus
        Else
            txtsaleschdesc = pr_dumy("Description")
            txtpurchasetaxoption.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtSaleTaxCode) = "" And KeyCode = vbKeyReturn Then
        txtSaleTaxCode = ""
        Textx = ""
        Command3_Click
End If

End Sub

Private Sub txtpurchasetaxoption_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtPurchaseTaxCode.SetFocus
End Sub
Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtPurchaseTaxCode
    Set PO_DESC = txtpurchaseschdesc
    Gs_SQL = "Select TaxCode, Description from IC_TaxSchedules "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where type =0 "
    MyLookupOLDB.Caption = "Purchase Tax Schedule"
    MyLookupOLDB.Show 1
    
    If txtPurchaseTaxCode <> "" Then Call txtPurchaseTaxCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtPurchaseTaxCode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtPurchaseTaxCode) <> "" And KeyCode = vbKeyReturn Then
        txtPurchaseTaxCode = DoPad(txtPurchaseTaxCode, 3)
        pr_dumy.Open "Select * from IC_TaxSchedules where type = 0 and taxcode = '" & txtPurchaseTaxCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Purchase Tax Code not found !!!", vbCritical)
            txtPurchaseTaxCode = ""
            Textx = ""
            txtPurchaseTaxCode.SetFocus
        Else
            txtpurchaseschdesc = pr_dumy("Description")
            If TxtSiteID.Enabled Then TxtSiteID.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtPurchaseTaxCode) = "" And KeyCode = vbKeyReturn Then
        txtPurchaseTaxCode = ""
        Textx = ""
        Command4_Click
End If

End Sub
Private Sub Command5_Click()
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
Private Sub txtSiteID_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(TxtSiteID) <> "" And KeyCode = vbKeyReturn Then
        TxtSiteID = DoPad(TxtSiteID, 3)
        pr_dumy.Open "Select * from IC_Sites where Compcode  = '" & Gs_compcode & "' and Sitecode = '" & TxtSiteID & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Site Code not found !!!", vbCritical)
            TxtSiteID = ""
            TxtsiteDesc = ""
            TxtSiteID.SetFocus
        Else
            TxtsiteDesc = pr_dumy("Description")
            If txtbinID.Enabled Then txtbinID.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(TxtSiteID) = "" And KeyCode = vbKeyReturn Then
        TxtSiteID = ""
        TxtsiteDesc = ""
        Command5_Click
End If

End Sub
Private Sub Command6_Click()
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
Private Sub txtBinID_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtbinID) <> "" And KeyCode = vbKeyReturn Then
        txtbinID = DoPad(txtbinID, 3)
        pr_dumy.Open "Select * from IC_SitesBins where  sitecode ='" & TxtSiteID & "' and compcode = '" & Gs_compcode & "'  and bincode = '" & txtbinID & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Bin Code not found !!!", vbCritical)
            txtbinID = ""
            TxtBinDesc = ""
            txtbinID.SetFocus
        Else
            TxtBinDesc = pr_dumy("Description")
            If txtsed.Enabled Then txtsed.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtbinID) = "" And KeyCode = vbKeyReturn Then
        txtbinID = ""
        TxtBinDesc = ""
        Command6_Click
End If

End Sub



Private Sub txtclassCode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtclasscode) <> "" And KeyCode = vbKeyReturn Then
        txtclasscode = DoPad(txtclasscode, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
       
       pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtcatcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
       
      '  pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtclasscode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtclasscode = ""
            txtClassDesc = ""
            txtclasscode.SetFocus
        Else
            txtClassDesc = pr_dumy("Description")
             If txtpackcode.Enabled Then txtpackcode.SetFocus
           
        End If
        pr_dumy.Close
        
ElseIf Trim(txtclasscode) = "" And KeyCode = vbKeyReturn Then
        txtclasscode = ""
        txtClassDesc = ""
        Command2_Click
End If
End Sub
Private Sub txtcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode = DoPad(txtcatcode, 3)
        If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from IC_ItemCategory where Catcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Department code not found !!!", vbCritical)
            txtcatcode = ""
            txtcatdesc = ""
            txtcatcode.SetFocus
        Else
            txtcatdesc = pr_dumy("Description")
            If txtclasscode.Enabled Then txtclasscode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode) = "" And KeyCode = vbKeyReturn Then
        txtcatcode = ""
        txtcatdesc = ""
        Command7_Click
End If
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtUcode
    Set PO_DESC = txtUdesc
    Gs_SQL = "Select MCode, Description from IC_ItemUM "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    MyLookupOLDB.Caption = "Item Units"
    MyLookupOLDB.Show 1
    
    If txtUcode <> "" Then Call txtUcode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Private Sub txtUcode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtUcode) <> "" And KeyCode = vbKeyReturn Then
        txtUcode = DoPad(txtUcode, 3)
        pr_dumy.Open "Select * from IC_ItemUM where Mcode = '" & txtUcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Item Unit not found !!!", vbCritical)
            txtUcode = ""
            txtUdesc = ""
            txtUcode.SetFocus
        Else
            txtUdesc = pr_dumy("Description")
            txtitemtype.SetFocus
            
        End If
        pr_dumy.Close

ElseIf Trim(txtUcode) = "" And KeyCode = vbKeyReturn Then
        txtUcode = ""
        txtUdesc = ""
        Command1_Click
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_itemsetup, Me, txtCode, txtdesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRITE00001")
  SetToolBar(2) = chkRights1("SRITE00002")
  SetToolBar(3) = chkRights1("SRITE00003")
  edit_Record.Enabled = chkRights1("SRITE00002")
  delete_Record.Enabled = chkRights1("SRITE00003")
  SetToolBar(4) = chkRights("SRITE00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)



  'PR_itemsetup.Open "Select * from ic_item where compcode = '" & Gs_compcode & "' Order By itemCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  
 ' PB_BlnkLoca = IIf(PR_itemsetup.EOF, True, False)

txtitemtype.ListIndex = 0
txtvalutionmethod.ListIndex = 2
txtsaletaxoption.ListIndex = 0
txtpurchasetaxoption.ListIndex = 0
CmbPrintRateNY.ListIndex = 0
' Stock upDate

'gc_dbcon.Execute "UPDATE IC_Item SET   StockG = 0 ,StockS = 0"
'gc_dbcon.Execute "UPDATE IC_Item SET   StockG = StockSummary.Qty FROM   StockSummary INNER JOIN   IC_Item ON StockSummary.ItemCode = IC_Item.ItemCode WHERE (StockSummary.siteid = 1)"
'gc_dbcon.Execute "UPDATE IC_Item SET   StockS = StockSummary.Qty FROM   StockSummary INNER JOIN   IC_Item ON StockSummary.ItemCode = IC_Item.ItemCode WHERE (StockSummary.siteid = 2)"


End Sub

Private Sub Form_Unload(Cancel As Integer)
    'PR_itemsetup.Close
End Sub
Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtcatcode.SetFocus
End Sub

Private Sub txtitemtype_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then txtvalutionmethod.SetFocus
End Sub

Private Sub txtsaletaxoption_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtSaleTaxCode.SetFocus
End Sub

Private Sub txtvalutionmethod_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    TxtPurchaseCost.SetFocus
    TxtPurchaseCost.SelLength = Len(TxtPurchaseCost.Text)
  End If
End Sub
Private Function maxtranscode() As String
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open "select max(itemcode) as transcode from ic_item where compcode = '" & Gs_compcode & "' and itemcode<>'88880' and itemcode <>'88881' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 6)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 4)
End If
pr_dumy.Close
End Function
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And txtCode <> "" Then
         
      txtCode.Text = DoPad(txtCode.Text, txtCode.MaxLength)
      PR_itemsetup.Open "Select * from ic_item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtCode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
         
       Select Case Mode
            Case "A"
                If Not PR_itemsetup.EOF Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtCode.SetFocus
                Else
                   txtdesc.SetFocus
                End If
            Case Else
                If PR_itemsetup.EOF Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "A" Then
                     If chkRights1("SRITE00005") Then
                     TxtPurchaseCost.Locked = False
                     Else
                     TxtPurchaseCost.Locked = True
                     End If
                     
                     If chkRights1("SRITE00006") Then
                     txtSaleCost.Locked = False
                     Else
                     txtSaleCost.Locked = True
                     
                     End If
                     End If
                      If txtdesc.Enabled Then txtdesc.SetFocus
                End If

            End Select
        PR_itemsetup.Close
ElseIf KeyCode = vbKeyReturn And Trim(txtCode) = "" Then
        Call cmdLookup_Click
End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
      cmdLookup.Enabled = False
      SSTab1.Tab = 0
    Else
      cmdLookup.Enabled = True
    End If
     Mode = DentMode(Mode, Button.Index, PR_itemsetup, Me, txtcustomCode, txtdesc, Para_Rs, "IC_GrnCnt", 3, "Itemcode", "description", 1, False, Toolbar1)
    
    If Mode = "A" Then
       txtCode = maxtranscode
       txtcustomCode.SetFocus
       CheckRateDifference
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim ls_sql As String



'gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
            txtCode = maxtranscode
            ls_sql = " INSERT into IC_Item(Compcode, ItemCode, Description, ClassID, ItemType, ValuationMethod, MCode, SaleTaxOption, SaleTaxCode"
            ls_sql = ls_sql & " ,PurchaseTaxOption , PurchaseTaxCode, SiteId, BinID, PurchaseCost, SaleCost,Sed,MaxSaleQty,CustomCode,CatCode,PackCode,ManuCode,ReorderQty,PriceDescCStatus,QtyCstatus,avgrate,Discper,discamt,GenericDesc,Packqty,PackDisc,QtyANAllow,AddDateTime,PrintRateNY)"
            ls_sql = ls_sql & " Values('" & Gs_compcode & "', '" & txtCode.Text & "','" & RepApp(txtdesc.Text) & "', '" & txtclasscode & "' , " & txtitemtype.ListIndex & ", " & txtvalutionmethod.ListIndex & ", '" & txtUcode & "', " & txtsaletaxoption.ListIndex & " , '" & txtSaleTaxCode & "',"
            ls_sql = ls_sql & " " & txtpurchasetaxoption.ListIndex & "   ,'" & txtPurchaseTaxCode & "' , '" & TxtSiteID & "' , '" & txtbinID & "', " & Val(TxtPurchaseCost) & "," & Val(txtSaleCost) & "," & txtsed & "," & txtmaxsaleqty & ","
            ls_sql = ls_sql & " '" & txtcustomCode.Text & "'   ,'" & txtcatcode.Text & "' , '" & txtpackcode.Text & "' , '" & txtmcode & "', " & Val(txtreorderlabel) & ", " & Val(Chksalerate.Value) & ", " & Val(chksaleqty.Value) & ", " & Val(txtavgrate) & ", " & Val(txtdiscper) & ", " & Val(txtdisperamt) & ",'" & Gc_UserId & "'," & Val(txtpackqty) & "," & Val(txtpackdisc) & "," & Val(chkanallow.Value) & ",'" & Format(Gd_SysDate, "YYYY/MM/DD HH:MM:SS") & "'," & Val(CmbPrintRateNY.ListIndex) & ")"

            gc_dbcon.Execute ls_sql
              
              
           Case "E"
           
           ' ls_sql = "delete from ic_itemsetuplog where compcode = '" & Gs_compcode & "' and transdate = '" & Format(Gd_SysDate, "YYYY/MM/DD") & "' and itemcode = '" & txtcode & "'"
            'gc_dbcon.Execute ls_sql
             
             ln_pcost = 0
             ln_ppcost = 0
             in_PDiscPer = 0
             in_PDiscPer = 0
            
            If pr_itemsetuplog.State = 1 Then pr_itemsetuplog.Close
            
            pr_itemsetuplog.Open "select Salecost,PurchaseCost,DiscPer,DiscAmt from ic_item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If Not pr_itemsetuplog.EOF Then
            
            ln_pcost = Val(0 & pr_itemsetuplog("SaleCost"))
            ln_ppcost = Val(0 & pr_itemsetuplog("PurchaseCost"))
            in_PDiscPer = Val(0 & pr_itemsetuplog("DiscPer"))
            in_PDiscAmt = Val(0 & pr_itemsetuplog("DiscAmt"))
            
            End If
            pr_itemsetuplog.Close
            
            If Val(TxtPurchaseCost) <> ln_ppcost Or Val(txtSaleCost) <> ln_pcost Or Val(txtdiscper.Text) <> in_PDiscPer Or Val(txtdisperamt.Text) <> in_PDiscAmt Then
               ls_sql = " INSERT into IC_ItemSetuplog(Compcode, ItemCode, transdate,pcost,ccost,Ppcost,CPcost,PDiscPer,CDiscPer,PDiscAmt,CDiscAmt,AddUser) values('" & Gs_compcode & "' ,'" & txtCode & "','" & Format(Gd_SysDate, "YYYY/MM/DD") & "'," & ln_pcost & "  ," & Val(txtSaleCost) & "," & Val(ln_ppcost) & "," & Val(TxtPurchaseCost) & "," & Val(ln_pDiscPer) & "," & Val(txtdiscper.Text) & "," & Val(ln_pDiscAmt) & "," & Val(txtdisperamt.Text) & ",'" & Gc_UserId & "')"
            gc_dbcon.Execute ls_sql
            End If
           
            ls_sql = " UPDATE IC_Item SET description =  '" & RepApp(txtdesc.Text) & "',ClassID = '" & txtclasscode & "' , ItemType= " & txtitemtype.ListIndex & ",ValuationMethod= " & txtvalutionmethod.ListIndex & ", MCode= '" & txtUcode & "', SaleTaxOption = " & txtsaletaxoption.ListIndex & " ,SaleTaxCode= '" & txtSaleTaxCode & "',"
            ls_sql = ls_sql & "  PurchaseTaxOption = " & txtpurchasetaxoption.ListIndex & "   ,PurchaseTaxCode = '" & txtPurchaseTaxCode & "' ,SiteId =  '" & TxtSiteID & "' , BinID ='" & txtbinID & "', PurchaseCost = " & TxtPurchaseCost & ",SaleCost =" & txtSaleCost & ",SED =" & txtsed & ",MaxSaleQty =" & txtmaxsaleqty & ","
            ls_sql = ls_sql & "  CustomCode = '" & txtcustomCode & "'   ,CatCode = '" & txtcatcode & "' ,packcode =  '" & txtpackcode & "' , ManuCode ='" & txtmcode & "', ReorderQty = " & txtreorderlabel & ",PriceDescCStatus = " & Chksalerate.Value & ",QtyCstatus = " & chksaleqty.Value & ",avgrate = " & Val(txtavgrate) & ","
            ls_sql = ls_sql & "  discper = " & Val(txtdiscper) & ",discamt = " & Val(txtdisperamt) & ",GenericDesc = '" & Gc_UserId & "' ,PackQty = " & Val(txtpackqty) & " ,PackDisc = " & Val(txtpackdisc) & ",QtyANAllow = " & Val(chkanallow) & ",EditDateTime='" & Format(Gd_SysDate, "YYYY/MM/DD HH:MM:SS") & "',PrintRateNY = " & Val(CmbPrintRateNY.ListIndex) & ""
            ls_sql = ls_sql & "  WHERE  compcode = '" & Gs_compcode & "' and  Itemcode= '" & txtCode.Text & "'"
            gc_dbcon.Execute ls_sql
            

           Case "D"
           ' gc_dbcon.Execute "DELETE FROM IC_Item WHERE compcode = '" & Gs_compcode & "' and  Itemcode = '" & txtCode.Text & "'"
     
     End Select
     
'gc_dbcon.CommitTrans
 



 If Mode = "A" Then
       txtCode = maxtranscode
       Me.Refresh
 End If
   
 txtcustomCode.SetFocus
Exit Sub
LocalErr:
MsgBox Err.Description
End Sub

Private Sub SetVal()
    On Error Resume Next
     txtdesc = Trim(PR_itemsetup("Description") & "")
     txtlmodify = Trim(PR_itemsetup("GenericDesc") & "")
     txtmodifyDate = Format(PR_itemsetup("editDateTime"), "DD/MM/YYYY HH:MM:SS")
     txtcustomCode = Trim(PR_itemsetup("CustomCode") & "")
     
     txtcatcode = Trim(PR_itemsetup("Catcode") & "")
     If Trim(txtcatcode) <> "" Then Call txtcatcode_KeyDown(vbKeyReturn, vbKeyShift)
     txtclasscode = Trim(PR_itemsetup("ClassID") & "")
     If Trim(txtclasscode) <> "" Then Call txtclassCode_KeyDown(vbKeyReturn, vbKeyShift)
     txtpackcode = Trim(PR_itemsetup("Packcode") & "")
     If Trim(txtpackcode) <> "" Then Call txtpackCode_KeyDown(vbKeyReturn, vbKeyShift)
     
     
         
     
     txtitemtype.ListIndex = Val(0 & PR_itemsetup("ItemType"))

     txtvalutionmethod.ListIndex = Val(0 & PR_itemsetup("ValuationMethod"))

     txtUcode = Trim(PR_itemsetup("Mcode") & "")
     If Trim(txtUcode) <> "" Then Call txtUcode_KeyDown(vbKeyReturn, vbKeyShift)
     
    

     txtmcode = Trim(PR_itemsetup("Manucode") & "")
     If Trim(txtmcode) <> "" Then Call txtmcode_KeyDown(vbKeyReturn, vbKeyShift)
     
     txtsaletaxoption.ListIndex = Val(0 & PR_itemsetup("SaleTaxOption"))
     txtpurchasetaxoption.ListIndex = Val(0 & PR_itemsetup("PurchaseTaxOption"))
     
     txtSaleTaxCode = Trim(PR_itemsetup("Saletaxcode") & "")
     If Trim(txtSaleTaxCode) <> "" Then Call txtsaletaxcode_KeyDown(vbKeyReturn, vbKeyShift)
 
     txtPurchaseTaxCode = Trim(PR_itemsetup("Purchasetaxcode") & "")
     If Trim(txtPurchaseTaxCode) <> "" Then Call txtPurchaseTaxCode_KeyDown(vbKeyReturn, vbKeyShift)
     
     
     TxtSiteID = Trim(PR_itemsetup("SiteID") & "")
     If Trim(TxtSiteID) <> "" Then Call txtSiteID_KeyDown(vbKeyReturn, vbKeyShift)
     
     txtbinID = Trim(PR_itemsetup("BinID") & "")
     If Trim(txtbinID) <> "" Then Call txtBinID_KeyDown(vbKeyReturn, vbKeyShift)
     
     
     TxtPurchaseCost = Val(0 & PR_itemsetup("PurchaseCost"))
     txtSaleCost = Val(0 & PR_itemsetup("SaleCost"))
     txtnetprate = Val(0 & PR_itemsetup("AvgRate1"))
     txtavgrate = Val(0 & PR_itemsetup("AvgRate"))
     txtdiscper = Val(0 & PR_itemsetup("Discper"))
     If PR_itemsetup("DiscAmt") = "NULL" Then
        txtdisperamt = Val(0 & PR_itemsetup("DiscAmt"))
     Else
      txtdisperamt = Val(PR_itemsetup("DiscAmt"))
     End If
      txtAfterDisc = Val(PR_itemsetup("SaleCost")) - Val(PR_itemsetup("DiscAmt"))
    
    
    If Val(txtSaleCost) > 0 And Val(txtnetprate) > 0 Then
       txtsaleperLstNr = Round((Val(txtSaleCost) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
       
        txtsaleper = Round((Val(txtSaleCost) - Val(TxtPurchaseCost)) / (Val(TxtPurchaseCost) / 100), 2)
        
        'txtsaleper = Round((Val(txtSaleCost) - Val(txtnetprate)) / (Val(txtnetprate) / 100), 2)
    
    End If
     
     
     txtsed = Val(0 & PR_itemsetup("Sed"))
     txtmaxsaleqty = Val(0 & PR_itemsetup("MaxSaleQty"))
     txtreorderlabel = Val(0 & PR_itemsetup("ReorderQty"))
     Chksalerate = Val(0 & PR_itemsetup("PriceDescCStatus"))
     chksaleqty = Val(0 & PR_itemsetup("QtyCStatus"))
     txtstockinS = Val(PR_itemsetup("StockS"))
     txtstockinG = Val(PR_itemsetup("StockG"))
     txtpackqty = Val(PR_itemsetup("PackQty"))
     txtpackdisc = Val(PR_itemsetup("PackDisc"))
     chkanallow.Value = Val(PR_itemsetup("QtyANAllow"))
     CmbPrintRateNY.ListIndex = Val(PR_itemsetup("PrintRateNY"))
     SSTab1.Tab = 0
End Sub
Public Function ChkInputs() As Boolean
    If Trim(txtCode.Text) = "" Then
        Call MsgBox("Enter Item Code!!!", vbCritical)
        txtCode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtcustomCode.Text) = "" Then
        Call MsgBox("Enter Item Customer Code!!!", vbCritical)
        txtcustomCode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtdesc.Text) = "" Then
        Call MsgBox("Enter Item Description!!!", vbCritical)
        txtdesc.SetFocus
        ChkInputs = False
    ElseIf Trim(txtclasscode.Text) = "" Then
        Call MsgBox("Enter/Select Item Class!!!", vbCritical)
        txtclasscode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtcatcode) = "" Then
        Call MsgBox("Enter/Select Item Category!!!", vbCritical)
        txtcatcode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtpackcode) = "" Then
        Call MsgBox("Enter/Select Item Packing!!!", vbCritical)
        txtpackcode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtmcode) = "" Then
        Call MsgBox("Enter/Select Item Manufacture !!!", vbCritical)
        txtmcode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtUcode) = "" Then
        Call MsgBox("Enter/Select Item Unit !!!", vbCritical)
        txtUcode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtitemtype) = "" Then
        Call MsgBox("Enter/Select Item Type !!!", vbCritical)
        txtitemtype.SetFocus
        ChkInputs = False
    ElseIf Trim(txtvalutionmethod) = "" Then
        Call MsgBox("Enter/Select Item valution method !!!", vbCritical)
        txtvalutionmethod.SetFocus
        ChkInputs = False
    ElseIf Trim(txtsaletaxoption) = "" Then
        Call MsgBox("Select Sale Tax Option !!!", vbCritical)
        txtsaletaxoption.SetFocus
        ChkInputs = False
    ElseIf Trim(txtSaleTaxCode) = "" Then
        Call MsgBox("Enter/Select Sale Tax Code !!!", vbCritical)
        txtSaleTaxCode.SetFocus
        ChkInputs = False
    ElseIf Trim(txtpurchasetaxoption) = "" Then
        Call MsgBox("Select Purchase Tax Option !!!", vbCritical)
        txtpurchasetaxoption.SetFocus
        ChkInputs = False
    ElseIf Trim(txtPurchaseTaxCode) = "" Then
        Call MsgBox("Enter Select Purchase Tax Code !!!", vbCritical)
        txtPurchaseTaxCode.SetFocus
        ChkInputs = False
    ElseIf Trim(TxtSiteID) = "" Then
        Call MsgBox("Enter Select Site ID !!!", vbCritical)
        TxtSiteID.SetFocus
        ChkInputs = False
    ElseIf Trim(txtbinID) = "" Then
        Call MsgBox("Enter Select Bin ID !!!", vbCritical)
        txtbinID.SetFocus
        ChkInputs = False
    ElseIf Trim(txtSaleCost) = "" Then
        Call MsgBox("Enter Sale Cost !!!", vbCritical)
        txtSaleCost.SetFocus
        ChkInputs = False
    ElseIf Trim(TxtPurchaseCost) = "" Then
        Call MsgBox("Enter Purchase Cost !!!", vbCritical)
        TxtPurchaseCost.SetFocus
        ChkInputs = False
    ElseIf Trim(txtmaxsaleqty) = "" Then
        Call MsgBox("Enter Max QTY !!!", vbCritical)
        txtmaxsaleqty.SetFocus
        ChkInputs = False
    ElseIf Trim(txtsed) = "" Then
        Call MsgBox("Enter SED % !!!", vbCritical)
        txtsed.SetFocus
        ChkInputs = False
    ElseIf Trim(txtreorderlabel) = "" Then
        Call MsgBox("Enter Reorder QTY % !!!", vbCritical)
        txtreorderlabel.SetFocus
        ChkInputs = False
   Else
       ChkInputs = True
    End If
End Function

Public Sub FrmRefresh()
  PR_itemsetup.Requery
End Sub

