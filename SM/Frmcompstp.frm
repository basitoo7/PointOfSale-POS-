VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmcompstp 
   Caption         =   "Company Setup"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmcompstp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   630
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Company"
      TabPicture(0)   =   "Frmcompstp.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtcompcash"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Periods"
      TabPicture(1)   =   "Frmcompstp.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "GL"
      TabPicture(2)   =   "Frmcompstp.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtgrouplevels"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label45"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Other1"
      TabPicture(3)   =   "Frmcompstp.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtstregno"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Label19"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Other2"
      TabPicture(4)   =   "Frmcompstp.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtstregno 
         Height          =   330
         Left            =   -73530
         MaxLength       =   25
         TabIndex        =   0
         Top             =   1545
         Width           =   4515
      End
      Begin VB.TextBox txtcompcash 
         Height          =   315
         Left            =   1530
         MaxLength       =   50
         TabIndex        =   70
         Top             =   2970
         Width           =   4395
      End
      Begin MSMask.MaskEdBox txtgrouplevels 
         Height          =   315
         Left            =   -73710
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   480
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   1
         Format          =   "#"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame4 
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   -74940
         TabIndex        =   36
         Top             =   360
         Width           =   2595
         Begin VB.OptionButton Option3 
            Height          =   315
            Left            =   2040
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   315
            Left            =   2040
            TabIndex        =   65
            Top             =   630
            Width           =   255
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Trading Company :"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   615
            TabIndex        =   66
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Manufacturing Company :"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   120
            TabIndex        =   38
            Top             =   660
            Width           =   1845
         End
      End
      Begin VB.Frame Frame5 
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   -74850
         TabIndex        =   35
         Top             =   360
         Width           =   2595
         Begin VB.OptionButton Option2 
            Height          =   315
            Left            =   2040
            TabIndex        =   34
            Top             =   630
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   315
            Left            =   2040
            TabIndex        =   33
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Purchase  Price  Base :"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   270
            TabIndex        =   64
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Average Price Base :"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   270
            TabIndex        =   63
            Top             =   270
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Account Group Levels"
         ForeColor       =   &H00000080&
         Height          =   2175
         Left            =   -74820
         TabIndex        =   21
         Top             =   900
         Width           =   5775
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "Enter Main account length"
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   23
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub1 account length"
            Top             =   840
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   24
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub2 account length"
            Top             =   1320
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   25
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub3 account length"
            Top             =   1800
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   4
            Left            =   3480
            TabIndex        =   26
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub4 account length"
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   5
            Left            =   3480
            TabIndex        =   27
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub5 account length"
            Top             =   840
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   6
            Left            =   3480
            TabIndex        =   28
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub6 account length"
            Top             =   1320
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   29
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub7 account length"
            Top             =   1800
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   8
            Left            =   5280
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub8 account length"
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtsublen 
            Height          =   255
            Index           =   9
            Left            =   5280
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Enter Sub9 account length"
            Top             =   840
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtdetlLen 
            Height          =   255
            Left            =   5280
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Enter Detailed account length"
            Top             =   1320
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            Caption         =   "Sub 9 Length :"
            Height          =   210
            Left            =   4080
            TabIndex        =   62
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "Sub 8 Length :"
            Height          =   210
            Left            =   4080
            TabIndex        =   61
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "Sub 2 Length :"
            Height          =   210
            Left            =   480
            TabIndex        =   60
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "Sub 5 Length :"
            Height          =   210
            Left            =   2280
            TabIndex        =   59
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Caption         =   "Sub 4 Length :"
            Height          =   210
            Left            =   2280
            TabIndex        =   58
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Sub 6 Length :"
            Height          =   210
            Left            =   2280
            TabIndex        =   57
            Top             =   1320
            Width           =   1050
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Sub 7 Length :"
            Height          =   210
            Left            =   2280
            TabIndex        =   56
            Top             =   1800
            Width           =   1050
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Sub 1 Length :"
            Height          =   210
            Left            =   480
            TabIndex        =   55
            Top             =   840
            Width           =   1050
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Sub 3 Length :"
            Height          =   210
            Left            =   480
            TabIndex        =   54
            Top             =   1800
            Width           =   1050
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Sub 0 Length :"
            Height          =   210
            Left            =   480
            TabIndex        =   53
            Top             =   360
            Width           =   1050
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Detail Length :"
            ForeColor       =   &H00000080&
            Height          =   210
            Left            =   4110
            TabIndex        =   52
            Top             =   1320
            Width           =   1020
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2475
         Left            =   90
         TabIndex        =   3
         Top             =   900
         Width           =   5925
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   5490
            Picture         =   "Frmcompstp.frx":0396
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1320
            Width           =   315
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   5820
            Picture         =   "Frmcompstp.frx":0508
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2040
            Picture         =   "Frmcompstp.frx":067A
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1320
            Width           =   315
         End
         Begin MSMask.MaskEdBox txtcompaddr1 
            Height          =   315
            Left            =   1440
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "Enter Company Address1"
            Top             =   240
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   35
            Format          =   "c"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtcompcity 
            Height          =   315
            Left            =   1440
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Enter City"
            Top             =   960
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   15
            Format          =   "c"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtCompBank 
            Height          =   315
            Left            =   1440
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Enter Default Bank A/c"
            Top             =   1680
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   50
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
         Begin MSMask.MaskEdBox txtcompaddr2 
            Height          =   315
            Left            =   1440
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Enter Company Address2"
            Top             =   600
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   35
            Format          =   "c"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtbasecurrency 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "Default Currency"
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtbranchcode 
            Height          =   315
            Left            =   4890
            TabIndex        =   12
            TabStop         =   0   'False
            ToolTipText     =   "Default Currency"
            Top             =   1320
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   3
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Caption         =   "Branch Code :"
            Height          =   255
            Left            =   3810
            TabIndex        =   68
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Address 1 :"
            Height          =   255
            Left            =   540
            TabIndex        =   51
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "Address 2 :"
            Height          =   255
            Left            =   540
            TabIndex        =   50
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "City :"
            Height          =   255
            Left            =   1020
            TabIndex        =   49
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "NTN # :"
            Height          =   255
            Left            =   45
            TabIndex        =   48
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Name Printing on Reports  : "
            Height          =   390
            Left            =   60
            TabIndex        =   47
            Top             =   2010
            Width           =   1320
         End
         Begin VB.Label Label12 
            Caption         =   "Default Currency :"
            Height          =   255
            Left            =   60
            TabIndex        =   46
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   360
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   3000
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Frame Frame3 
         Height          =   1215
         Left            =   -74850
         TabIndex        =   16
         Top             =   420
         Width           =   5775
         Begin MSComCtl2.DTPicker ttxttodate 
            Height          =   315
            Left            =   3480
            TabIndex        =   20
            Tag             =   "SKIP"
            Top             =   720
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   62914561
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker ftxttodate 
            Height          =   315
            Left            =   3480
            TabIndex        =   18
            Tag             =   "SKIP"
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   62914561
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker ttxtfromdate 
            Height          =   315
            Left            =   1710
            TabIndex        =   19
            Top             =   750
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            _Version        =   393216
            Format          =   62914561
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker ftxtfromdate 
            Height          =   315
            Left            =   1740
            TabIndex        =   17
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   62914561
            CurrentDate     =   37293
         End
         Begin VB.Label Label10 
            Caption         =   "To :"
            Height          =   255
            Left            =   3150
            TabIndex        =   44
            Top             =   720
            Width           =   315
         End
         Begin VB.Label Label11 
            Caption         =   "Tax Year From :"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Financial Year From :"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "To :"
            Height          =   255
            Left            =   3150
            TabIndex        =   41
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   90
         TabIndex        =   2
         Top             =   360
         Width           =   5925
         Begin MSMask.MaskEdBox txtcompname 
            Height          =   315
            Left            =   2340
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Enter Company Name"
            Top             =   180
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   50
            Format          =   "c"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdLookup 
            Height          =   315
            Left            =   1200
            Picture         =   "Frmcompstp.frx":07EC
            Style           =   1  'Graphical
            TabIndex        =   5
            TabStop         =   0   'False
            Tag             =   "SKIP"
            Top             =   180
            Width           =   315
         End
         Begin MSMask.MaskEdBox txtcompcode 
            Height          =   315
            Left            =   720
            TabIndex        =   4
            TabStop         =   0   'False
            Tag             =   "SKIP"
            ToolTipText     =   "Enter Company code"
            Top             =   180
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16776960
            PromptInclude   =   0   'False
            MaxLength       =   3
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
         Begin VB.Label Label2 
            Caption         =   "Name :"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   1740
            TabIndex        =   40
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Code :"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   180
            TabIndex        =   39
            Top             =   180
            Width           =   495
         End
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "S.T.Reg.No.:"
         Height          =   255
         Left            =   -74850
         TabIndex        =   71
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "Group Levels :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   -74790
         TabIndex        =   67
         Top             =   480
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   72
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
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
               Picture         =   "Frmcompstp.frx":095E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmcompstp.frx":0DB2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmcompstp.frx":1206
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmcompstp.frx":165A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmcompstp.frx":1AAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmcompstp.frx":1F02
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frmcompstp.frx":2656
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Frmcompstp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lb_BlnkMast As Boolean
Dim Mode As String

Dim Ls_InvBase As String
Dim Ls_ArBase As String

Dim PR_Crncy  As New Recordset
Dim PR_Sytax As New Recordset
Dim PR_Syfins As New Recordset
Dim PR_SyComp As New Recordset
Dim PR_Branch As New Recordset
Dim PR_GlDetail As New Recordset


Public PO_CODE As Object
Public PO_DESC As Object

Private Sub cmdLookup_Click()
   
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCompCode
    Set PO_DESC = txtcompname
    GoTop PR_SyComp
    MyLookup.Caption = "Companies"
    MyLookup.FillGrid PR_SyComp, "Compcode", "CompName", 5
    MyLookup.Show 1
    
    If Val(txtCompCode) > 0 Then txtcompcode_KeyDown vbKeyReturn, vbKeyShift
    
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCompBank
    Set PO_DESC = Text1
    
    GoTop PR_GlDetail
    PR_GlDetail.Filter = "Acct_Type = 'B'"
    MyLookup.Caption = "Chart of Account"
    MyLookup.FillGrid PR_GlDetail, "AccountNo", "Acct_Desc", IIf(gn_DtlLen = 0, 15, gn_DtlLen)
    MyLookup.Show 1
    
    If Val(txtCompBank) > 0 Then txtCompBank_KeyDown vbKeyReturn, vbKeyShift
    PR_GlDetail.Filter = adFilterNone
End Sub

Private Sub Command2_Click()
'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txtcompcash
'    Set PO_DESC = Text1
'
'    GoTop Pr_Gldetail
'    Pr_Gldetail.Filter = "Acct_Type = 'S'"
'    MyLookup.Caption = "Chart of Account"
'    MyLookup.FillGrid Pr_Gldetail, "AccountNo", "Acct_Desc", IIf(gn_DtlLen = 0, 15, gn_DtlLen)
'    MyLookup.Show 1
'
'    If Val(txtcompcash) > 0 Then txtcompcash_KeyDown vbKeyReturn, vbKeyShift
'    Pr_Gldetail.Filter = adFilterNone

End Sub

Private Sub Command3_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbasecurrency
    Set PO_DESC = Text1
    GoTop PR_Crncy
    MyLookup.Caption = "Currency Types"
    MyLookup.FillGrid PR_Crncy, "Crncy_code", "Crncy_Descrip"
    MyLookup.Show 1
    
    If txtbasecurrency <> "" Then txtbasecurrency_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = Text1
    GoTop PR_Branch
    PR_Branch.Filter = "Compcode = '" & txtCompCode & "'"
    MyLookup.Caption = "Branches"
    MyLookup.FillGrid PR_Branch, "Branchcode", "BranchDesc"
    MyLookup.Show 1
    PR_Branch.Filter = adFilterNone
    If Val(txtbranchcode) > 0 Then txtBranchCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then
      Mode = DentMode(Mode, 4, PR_SyComp, Frmcompstp, txtCompCode, txtcompname, "X", "CompCount", 2, "CompCode", "CompName", 1, False, Toolbar1)
End If
End Sub

Private Sub Form_Load()

  SetToolBar(1) = chkRights("GLCOMPSTP1")
  SetToolBar(2) = chkRights("GLCOMPSTP2")
  SetToolBar(3) = chkRights("GLCOMPSTP3")
  SetToolBar(4) = chkRights("GLCOMPSTP4")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Text1.MaxLength = 50
 
  PR_GlDetail.Open "Select AccountNo,Acct_Type,Acct_Desc from Gl_Detail Where Acct_Type IN ('S','B') And Compcode = '" & Gs_compcode & "' order by AccountNo", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Branch.Open "Select * from SysBranch order by Branchcode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Crncy.Open "Select * from SysCurrency order by Crncy_code", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_SyComp.Open "Select * from Syscomp order by CompCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Syfins.Open "Select * from SysFins Where factiveyear = '" & 1 & "' order by CompCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Sytax.Open "Select * from SysTax Where Tactiveyear = '" & 1 & "' order by CompCode", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

  lb_BlnkMast = IIf(PR_SyComp.EOF, True, False)
  cmdLookup.Enabled = Not lb_BlnkMast
  SSTab1.TabEnabled(2) = Gb_GL
  SSTab1.TabEnabled(3) = True
  SSTab1.TabEnabled(4) = False
       For ln_cnt = 0 To 9
         txtsublen(ln_cnt).Enabled = True
         txtsublen(ln_cnt).Text = ""
       Next
  End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_SyComp.Close
    PR_Syfins.Close
    PR_Sytax.Close
    PR_Branch.Close
    PR_Crncy.Close
    PR_GlDetail.Close
End Sub

Private Sub ftxtFromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       ftxttodate.Enabled = True
       ftxttodate.Value = Format(DateAdd("m", 12, ftxtfromdate.Value) - 1, "dd / mm / yyyy")
       ftxttodate.Enabled = False
       ttxtfromdate.SetFocus
    ElseIf KeyCode = vbKeyPageUp Then
        SSTab1.Tab = 0
        txtcompcash.SetFocus
    End If
End Sub

Private Sub Option1_Click()
  Ls_InvBase = "A"
  SSTab1.Tab = 4
End Sub

Private Sub Option2_Click()
Ls_InvBase = "P"
SSTab1.Tab = 4
End Sub
Private Sub Option3_Click()
  Ls_ArBase = "T"
  SSTab1.Tab = 0
End Sub

Private Sub Option4_Click()
  Ls_ArBase = "M"
  SSTab1.Tab = 0
End Sub

Private Sub ttxtfromdate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       ttxttodate.Enabled = True
       ttxttodate.Value = Format(DateAdd("m", 12, ttxtfromdate.Value) - 1, "dd/mm/yyyy")
       ttxttodate.Enabled = False
       SSTab1.Tab = 0
       txtcompname.SetFocus
       SSTab1.Tab = 2
       txtgrouplevels.SetFocus
 ElseIf KeyCode = vbKeyPageUp Then
    SSTab1.Tab = 0
    txtcompcash.SetFocus
    End If
End Sub


Private Sub txtbasecurrency_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If Lastkey(KeyCode) And txtbasecurrency.Text <> "" Then
         txtbasecurrency = UCase(txtbasecurrency.Text)
         lb_found = MySeek(txtbasecurrency, "Crncy_Code", PR_Crncy)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtbasecurrency.SetFocus
         Else
             txtbranchcode.SetFocus
         End If
 End If

End Sub

Private Sub txtBranchCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) And txtbranchcode.Text <> "" Then
         txtbranchcode = DoPad(txtbranchcode, 3)
         PR_Branch.Filter = "compcode = '" & txtCompCode & "'"
         lb_found = MySeek(txtbranchcode.Text, "BranchCode", PR_Branch)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtbranchcode.SetFocus
         Else
             txtCompBank.SetFocus
         End If
         PR_Branch.Filter = adFilterNone
 End If

End Sub

Private Sub txtcompaddr1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtcompaddr2.SetFocus
 End If
End Sub


Private Sub txtcompaddr2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtcompcity.SetFocus
 End If
End Sub
Private Sub txtCompBank_KeyDown(KeyCode As Integer, Shift As Integer)
 If Lastkey(KeyCode) Then
    txtcompcash.SetFocus
 End If
End Sub

Private Sub txtcompcash_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    SSTab1.Tab = 1
    ftxtfromdate.SetFocus
 End If

End Sub

Private Sub txtcompcity_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
     txtbasecurrency.SetFocus
 End If
End Sub

Private Sub txtcompcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn Then
         txtCompCode = DoPad(txtCompCode, 3)
         PR_SyComp.Requery
         lb_found = MySeek(LTrim(RTrim(txtCompCode.Text)), "CompCode", PR_SyComp)
        
       Select Case Mode
            Case "A"
                If lb_found Then
                   MsgBox "Company already exist.", vbCritical, "E-Counts 2.0"
                   ''Cancel = True
                   txtCompCode.SetFocus
                Else
                   txtcompname.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   MsgBox "Record does not exist.", vbCritical, "E-Counts 2.0"
                   ''Cancel = True
                   txtCompCode.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtCompCode.Enabled = False
                      txtcompname.SetFocus
                   End If
                End If
            End Select
            
       End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
       cmdLookup.Enabled = False
    Else
    cmdLookup.Enabled = True
    End If
    If lb_BlnkMast And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
      ' 'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_SyComp, Frmcompstp, txtCompCode, txtcompname, "X", "CompCount", 2, "CompCode", "CompName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
lb_BlnkMast = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              cntsql.CommandText = "INSERT into syscomp(compcode,BranchCode,compname,compaddr1,compaddr2,compcity,glActlevel,Basecurrency,compbank,compcash,glactsub0,glactsub1,glactsub2,glactsub3,glactsub4,glactsub5,glactsub6,glactsub7,glactsub8,glactsub9,glactdetl,Inv_Base,Ar_Base,Stregno) VALUES ('" & txtCompCode.Text & "','" & txtbranchcode & "','" & txtcompname.Text & "','" & txtcompaddr1.Text & "','" & txtcompaddr2.Text & "','" & txtcompcity.Text & "','" & txtgrouplevels.Text & "','" & txtbasecurrency.Text & "','" & txtCompBank.Text & "','" & txtcompcash.Text & "'," & Val(txtsublen(0).Text) & "," & Val(txtsublen(1).Text) & "," & Val(txtsublen(2).Text) & "," & Val(txtsublen(3).Text) & "," & Val(txtsublen(4).Text) & "," & Val(txtsublen(5).Text) & "," & Val(txtsublen(6).Text) & "," & Val(txtsublen(7).Text) & "," & Val(txtsublen(8).Text) & "," & Val(txtsublen(9).Text) & "," & Val(txtdetlLen.Text) & ",'" & Ls_InvBase & "','" & Ls_ArBase & "','" & txtstregno & "')"
              cntsql.Execute
              cntsql.CommandText = "INSERT into syproc(ProcCode,ProcDesc,ProcStat,ProcType) VALUES ('" & txtCompCode.Text & "','" & txtcompname & "',1,'COM')"
              cntsql.Execute
              cntsql.CommandText = "INSERT into sysFins(compcode,ffromdate,ftodate,factiveyear,fclosed) VALUES ('" & txtCompCode.Text & "','" & Format(ftxtfromdate.Value, "YYYY/MM/DD") & "','" & Format(ftxttodate.Value, "YYYY/MM/DD") & "','" & 1 & "','" & 0 & "')"
              cntsql.Execute
              cntsql.CommandText = "INSERT into sysTax(compcode,Tfromdate,Ttodate,Tactiveyear,Tclosed) VALUES ('" & txtCompCode.Text & "','" & Format(ttxtfromdate.Value, "YYYY/MM/DD") & "','" & Format(ttxttodate.Value, "YYYY/MM/DD") & "','" & 1 & "','" & 0 & "')"
              cntsql.Execute
           Case "E"
              cntsql.CommandText = "UPDATE syscomp SET compcode= '" & txtCompCode.Text & "',BranchCode = '" & txtbranchcode & "', glActlevel = '" & txtgrouplevels.Text & "',compname = '" & txtcompname.Text & "',  compaddr1 = '" & txtcompaddr1 & "',  compaddr2 ='" & txtcompaddr2 & "',  compcity = '" & txtcompcity & "', Basecurrency = '" & txtbasecurrency.Text & "', compbank = '" & txtCompBank & "',  compcash = '" & txtcompcash & "',glactsub0=" & Val(txtsublen(0).Text) & ",glactsub1=" & Val(txtsublen(1).Text) & ",glactsub2=" & Val(txtsublen(2).Text) & ",glactsub3=" & Val(txtsublen(3).Text) & ",glactsub4=" & Val(txtsublen(4).Text) & ",glactsub5=" & Val(txtsublen(5).Text) & ",glactsub6=" & Val(txtsublen(6).Text) & ",glactsub7=" & Val(txtsublen(7).Text) & ",glactsub8=" & Val(txtsublen(8).Text) & ",glactsub9=" & Val(txtsublen(9).Text) & ",glactdetl=" & Val(txtdetlLen.Text) & ",Inv_Base='" & Ls_InvBase & "',Ar_Base='" & Ls_ArBase & "' ,Stregno = '" & txtstregno & "'  WHERE compcode = '" & txtCompCode & "'"
              cntsql.Execute
              
              cntsql.CommandText = "UPDATE sysFins SET FFromDate = '" & Format(ftxtfromdate.Value, "YYYY/MM/DD") & "',FToDate= '" & Format(ftxttodate.Value, "YYYY/MM/DD") & "' Where Compcode = '" & Gs_compcode & "' And fActiveYear = 1"
              cntsql.Execute
              
              cntsql.CommandText = "UPDATE sysTax SET TFromDate = '" & Format(ttxtfromdate.Value, "YYYY/MM/DD") & "',TToDate= '" & Format(ttxttodate.Value, "YYYY/MM/DD") & "' Where Compcode = '" & Gs_compcode & "' And TActiveYear = 1"
              cntsql.Execute
              
           Case "D"
            cntsql.CommandText = "DELETE FROM syscomp WHERE compcode = '" & txtCompCode.Text & "'"
            cntsql.Execute
            
            cntsql.CommandText = "DELETE FROM syproc WHERE Proccode = '" & Trim(txtCompCode.Text) & "'"
            cntsql.Execute
            
           
           
     End Select
     
     PR_SyComp.Requery
     PR_Syfins.Requery
     PR_Sytax.Requery
     gc_dbcon.CommitTrans
     SSTab1.Tab = 0
     
     For ln_cnt = 0 To 9
         txtsublen(ln_cnt).Enabled = True
         txtsublen(ln_cnt).Text = ""
     Next
     
Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
MsgBox Err.Description
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub SetVal()
     On Error Resume Next
     If Not MySeek(txtCompCode, "Compcode", PR_Syfins) Then Call SetErr("Company not found in SysFins table.", vbCritical)
     If Not MySeek(txtCompCode, "Compcode", PR_Sytax) Then Call SetErr("Company not found in SysFins table.", vbCritical)
     txtCompCode = PR_SyComp("compcode")
     txtcompname = PR_SyComp("compname")
     txtcompaddr1 = PR_SyComp("compaddr1")
     txtcompaddr2 = PR_SyComp("compaddr2")
     txtcompcity = PR_SyComp("compcity")
     txtgrouplevels = Val(0 & PR_SyComp("GlActLevel"))
     txtbasecurrency = PR_SyComp("Basecurrency")
     txtCompBank = PR_SyComp("compbank")
     txtcompcash = PR_SyComp("compcash")
     ftxtfromdate = PR_Syfins("ffromdate")
     ftxttodate = PR_Syfins("ftodate")
     ttxtfromdate = PR_Sytax("tfromdate")
     ttxttodate = PR_Sytax("ttodate")
     txtbranchcode = PR_SyComp("BranchCode") & ""
     txtdetlLen = PR_SyComp("glactDetl")
     txtstregno = PR_SyComp("Stregno")
     
     Option1.Value = IIf(PR_SyComp("Inv_Base") = "A", True, False)
     Option2.Value = IIf(PR_SyComp("Inv_Base") = "P", True, False)
     
     Option3.Value = IIf(PR_SyComp("Ar_Base") = "T", True, False)
     Option4.Value = IIf(PR_SyComp("Ar_Base") = "M", True, False)
     
     
     For ln_cnt = 0 To 9
         txtsublen(ln_cnt) = Val(0 & PR_SyComp("glactsub" & ln_cnt))
     Next
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtcompname.Text) > 0 And Len(txtCompCode) = txtCompCode.MaxLength And ftxtfromdate.Value > 0 And ttxtfromdate.Value > 0 Then
       ChkInputs = True
    Else
       Call SetErr("Incomplete Data found", vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtcompname_Change()
txtcompcash = txtcompname
End Sub

Private Sub txtcompname_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtcompaddr1.SetFocus
 End If
End Sub

Private Sub txtdetlLen_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) And SSTab1.Enabled Then SSTab1.Tab = 3 Else SSTab1.Tab = 0
End Sub
Private Sub txtGroupLevels_KeyDown(KeyCode As Integer, Shift As Integer)
xx = txtgrouplevels.Enabled

 If KeyCode = vbKeyReturn Then
    If Val(txtgrouplevels.Text) > 0 Then
       For ln_cnt = 0 To 9
          If ln_cnt <= Val(0 & txtgrouplevels.Text) Then
              txtsublen(ln_cnt).Enabled = True
          Else
              txtsublen(ln_cnt) = ""
              txtsublen(ln_cnt).Enabled = False
          End If
       Next
       txtsublen(0).SetFocus
    End If
 End If
End Sub

Private Sub txtsublen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = txtsublen.Count Then Exit Sub
        If txtsublen(Index + IIf(Index < 9, 1, 0)).Enabled Then
           If Index < 9 Then
             txtsublen(Index + IIf(Index < 9, 1, 0)).SetFocus
           Else
             txtdetlLen.SetFocus
           End If
        Else
           txtdetlLen.SetFocus
        End If
    End If
End Sub

Public Sub FrmRefresh()
  PR_GlDetail.Requery
  PR_Branch.Requery
  PR_Crncy.Requery
  PR_SyComp.Requery
  PR_Syfins.Requery
  PR_Sytax.Requery
End Sub
