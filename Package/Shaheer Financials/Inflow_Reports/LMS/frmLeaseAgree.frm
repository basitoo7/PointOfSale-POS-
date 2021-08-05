VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLeaseAgree 
   Caption         =   "Lease Agreement"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLeaseAgree.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   6465
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5535
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
         Left            =   4860
         Picture         =   "frmLeaseAgree.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   4440
         Width           =   315
      End
      Begin VB.TextBox txtcadre 
         Height          =   315
         Left            =   4260
         MaxLength       =   3
         TabIndex        =   70
         Top             =   4425
         Width           =   585
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
         Left            =   4860
         Picture         =   "frmLeaseAgree.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   4065
         Width           =   315
      End
      Begin VB.TextBox txtcredit 
         Height          =   315
         Left            =   4260
         MaxLength       =   3
         TabIndex        =   67
         Top             =   4072
         Width           =   585
      End
      Begin VB.CheckBox ChkGLAct 
         Alignment       =   1  'Right Justify
         Caption         =   "Open GL Accounts :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   630
         TabIndex        =   66
         Top             =   4470
         Width           =   1785
      End
      Begin VB.TextBox txtleaseno 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1335
         MaxLength       =   3
         TabIndex        =   57
         Tag             =   "SKIP"
         Top             =   540
         Width           =   465
      End
      Begin VB.TextBox txtpaidin 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   56
         Top             =   2340
         Width           =   465
      End
      Begin VB.TextBox txtaccrual 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   55
         Top             =   2700
         Width           =   465
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
         Left            =   4860
         Picture         =   "frmLeaseAgree.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2280
         Width           =   315
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   5130
         Picture         =   "frmLeaseAgree.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin VB.TextBox txtrecoverer 
         Height          =   315
         Left            =   4260
         MaxLength       =   3
         TabIndex        =   18
         Top             =   3720
         Width           =   585
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
         Left            =   4845
         Picture         =   "frmLeaseAgree.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3720
         Width           =   315
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   1965
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
         Left            =   1830
         Picture         =   "frmLeaseAgree.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2700
         Width           =   315
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
         Left            =   1830
         Picture         =   "frmLeaseAgree.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2340
         Width           =   315
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
         Left            =   1830
         Picture         =   "frmLeaseAgree.frx":0D28
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "SKIP"
         Top             =   540
         Width           =   315
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   3750
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.CommandButton cmdLookup 
         Appearance      =   0  'Flat
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
         Picture         =   "frmLeaseAgree.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpdisbr 
         Height          =   285
         Left            =   4260
         TabIndex        =   6
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57409537
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker dtpagrdate 
         Height          =   285
         Left            =   4260
         TabIndex        =   5
         Top             =   570
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57409537
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker DTPschdl 
         Height          =   285
         Left            =   4260
         TabIndex        =   12
         Top             =   1230
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Format          =   57409537
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtdpt_v 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1260
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtrsd_v 
         Height          =   315
         Left            =   2040
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1620
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfef_v 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1980
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtmgtfee 
         Height          =   315
         Left            =   4260
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1560
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtdocfee 
         Height          =   315
         Left            =   4260
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1920
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtinsurance 
         Height          =   315
         Left            =   4260
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2640
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.00000;(##0.00000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtwtax 
         Height          =   315
         Left            =   4260
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3000
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtmisc 
         Height          =   315
         Left            =   4260
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   5205
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   3
         AllowBigSelection=   0   'False
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin MSMask.MaskEdBox txtfrom 
         Height          =   315
         Left            =   1350
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4845
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
         Format          =   "#0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtto 
         Height          =   315
         Left            =   2370
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   4845
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
         Format          =   "#0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txteditrental 
         Height          =   315
         Left            =   3540
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4845
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   4530
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "N"
         ToolTipText     =   "Default Currency"
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTaxRate 
         Height          =   315
         Left            =   2460
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   900
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   6
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
      Begin MSMask.MaskEdBox txtlatefee 
         Height          =   315
         Left            =   4260
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2280
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox txtCustNO 
         Height          =   315
         Left            =   1335
         TabIndex        =   58
         Tag             =   "SKIP"
         Top             =   180
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         MaxLength       =   6
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
      Begin MSMask.MaskEdBox txtdpt_p 
         Height          =   315
         Left            =   1350
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1260
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000;(#,##0.000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtrsd_p 
         Height          =   315
         Left            =   1350
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1620
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000;(#,##0.000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtleaseamount 
         Height          =   315
         Left            =   1350
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   900
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfef_p 
         Height          =   315
         Left            =   1350
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1980
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000;(#,##0.000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtperiod 
         Height          =   315
         Left            =   1350
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3060
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtrental 
         Height          =   315
         Left            =   1350
         TabIndex        =   64
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   3780
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   192
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtirr 
         Height          =   315
         Left            =   1350
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.000000;(#,##0.000000)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Cadre Officer :"
         Height          =   225
         Left            =   2820
         TabIndex        =   72
         Top             =   4470
         Width           =   1425
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Officer :"
         Height          =   225
         Left            =   2805
         TabIndex        =   69
         Top             =   4125
         Width           =   1425
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Company :"
         Height          =   225
         Left            =   2640
         TabIndex        =   53
         Top             =   2310
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "IRR :"
         Height          =   225
         Left            =   930
         TabIndex        =   51
         Top             =   3465
         Width           =   375
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tax Rate"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2400
         TabIndex        =   50
         Top             =   660
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Rental :"
         Height          =   225
         Left            =   750
         TabIndex        =   48
         Top             =   3825
         Width           =   555
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   5490
         Y1              =   4785
         Y2              =   4785
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Rental :"
         Height          =   225
         Left            =   2850
         TabIndex        =   47
         Top             =   4875
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "To :"
         Height          =   225
         Left            =   2010
         TabIndex        =   46
         Top             =   4875
         Width           =   315
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "From :"
         Height          =   225
         Left            =   690
         TabIndex        =   45
         Top             =   4875
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Recovery Officer :"
         Height          =   225
         Left            =   2790
         TabIndex        =   44
         Top             =   3765
         Width           =   1425
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Miscellinous Amt :"
         Height          =   225
         Left            =   2910
         TabIndex        =   43
         Top             =   3390
         Width           =   1305
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Withholding Tax :"
         Height          =   225
         Left            =   2970
         TabIndex        =   42
         Top             =   3030
         Width           =   1245
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Insurance Rate :"
         Height          =   225
         Left            =   3000
         TabIndex        =   41
         Top             =   2670
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Doc. Fee :"
         Height          =   225
         Left            =   3030
         TabIndex        =   40
         Top             =   1950
         Width           =   1185
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Mgt. Fee :"
         Height          =   225
         Left            =   3030
         TabIndex        =   39
         Top             =   1590
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Period :"
         Height          =   225
         Left            =   690
         TabIndex        =   38
         Top             =   3105
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Accrual As :"
         Height          =   225
         Left            =   390
         TabIndex        =   37
         Top             =   2745
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Paid In :"
         Height          =   225
         Left            =   690
         TabIndex        =   36
         Top             =   2385
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "F.E.F % :"
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease Amount :"
         Height          =   225
         Left            =   150
         TabIndex        =   34
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Residual %age :"
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Deposit %age :"
         Height          =   225
         Left            =   150
         TabIndex        =   32
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Schedule :"
         Height          =   210
         Index           =   3
         Left            =   3450
         TabIndex        =   31
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Disb'ment :"
         Height          =   210
         Index           =   2
         Left            =   3435
         TabIndex        =   30
         Top             =   930
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agreement :"
         Height          =   210
         Index           =   1
         Left            =   3330
         TabIndex        =   29
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lease # :"
         Height          =   210
         Index           =   0
         Left            =   630
         TabIndex        =   27
         Top             =   570
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   225
         Width           =   1200
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
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
               Picture         =   "frmLeaseAgree.frx":100C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAgree.frx":1460
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAgree.frx":18B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAgree.frx":1D08
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAgree.frx":215C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAgree.frx":25B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLeaseAgree.frx":2D04
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmLeaseAgree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkSupp As Boolean
Dim PI_CurRow    As Integer
Dim PI_SrNo     As Integer
Dim PS_RowClicked As String
Dim ln_TaxAmt As Double
Dim ln_LastValu As Integer
Dim ln_PrvUpto As Integer
Dim Ln_ReCalc As Integer
Dim lb_Edited As Boolean
Dim ls_LeaseNo As String

Public Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Public PR_AssetType As New Recordset
Public PR_leaseRef As New Recordset

Dim PR_InsurComp As New Recordset
Dim PR_LeaseRSt As New Recordset
Dim PR_RecOfficer As New Recordset
Dim PR_Branch As New Recordset
Dim PR_Accrual As New Recordset
Dim PR_FCMIDs As New Recordset
Dim PR_LMSInfo As New Recordset
Dim PR_AssetInfo As New Recordset
Dim PR_Facility As New Recordset
Dim pr_Customer As New Recordset
Dim Pr_Cib1 As New Recordset


Private Sub Command7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcredit
    Set PO_DESC = Text1
    GoTop PR_RecOfficer
    PR_RecOfficer.Filter = "RecTag = 'CRE'"
    MyLookup.Caption = "Credit Officers"
    MyLookup.FillGrid PR_RecOfficer, "RecCode", "RecName", txtcredit.MaxLength
    MyLookup.Show 1
    PR_RecOfficer.Filter = adFilterNone
    If Len(txtcredit) > 0 Then txtcredit_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command8_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcadre
    Set PO_DESC = Text1
    GoTop PR_RecOfficer
    PR_RecOfficer.Filter = "RecTag = 'CAD'"
    MyLookup.Caption = "Cadre Officers"
    MyLookup.FillGrid PR_RecOfficer, "RecCode", "RecName", txtcadre.MaxLength
    MyLookup.Show 1
    PR_RecOfficer.Filter = adFilterNone
    If Len(txtcadre) > 0 Then txtcadre_KeyDown vbKeyReturn, vbKeyShift
    
End Sub
Private Sub dtpagrdate_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
    frmAssetInfo.Show 1
    If Mode <> "A" Then
       txtmgtfee.SetFocus
    Else
       txtleaseamount.SetFocus
    End If
  End If
End Sub

Private Sub dtpdisbr_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtdpt_p.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtleaseamount.SetFocus
  End If
End Sub

Private Sub DTPschdl_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then
     txtmgtfee.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtperiod.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Dim ln_cnt As Integer
  
  SetToolBar(1) = chkRights("LMLEASEAG1")
  SetToolBar(2) = chkRights("LMLEASEAG2")
  SetToolBar(3) = chkRights("LMLEASEAG3")
  SetToolBar(4) = chkRights("LMLEASEAG4")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
 
'  For ln_cnt = 1 To 8
'     AssetInfo(ln_cnt) = ""
'  Next
  

PR_InsurComp.Open "Select * From LM_InsurComp Order By InsurCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_AssetType.Open "Select * From LM_AssetTypes Order By AssetCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_leaseRef.Open "Select * From LM_SecrtyBy Order By SecrtyCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_RecOfficer.Open "Select * From LM_Recoverer Order By RecCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_Branch.Open "Select * From SysBranch Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_FCMIDs.Open "Select * From FCM_Ids where recid in ('ADP','LMA') Order By IdCode", gc_dbcon, adOpenDynamic, adLockOptimistic, adCmdText
'PR_PaidIn.Open "Select *,RecId+Code As FindFld From Tmp_Table Order By RecId,Code", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

pr_Customer.Open "Select Customer.* from Customer Inner Join Facilities On Customer.CustomerNo = Facilities.CustomerNo Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01' Order By Customer.CustomerNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
Pr_Cib1.Open "Select lm_cib1.* from lm_cib1 left outer Join Facilities On Lm_Cib1.CustomerNo = Facilities.CustomerNo Where LM_Cib1.Compcode = '" & Gs_compcode & "' And Facilities.FacilityNo = '01' Order By LM_Cib1.CustomerNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_Facility.Open "Select *,BranchCode+CustomerNo As FindFld from Facilities where compcode ='" & Gs_compcode & "' And FacilityNo = '01' Order by BranchCode,CustomerNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
PR_LeaseRSt.Open "Select *,BranchCode+CustomerNo+LeaseNo As FindFld From LM_LeaseRSt Where Compcode = '" & Gs_compcode & "' Order By BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
PR_AssetInfo.Open "Select *,BranchCode+CustomerNo+LeaseNo As FindFld from LM_AssetInfo where compcode ='" & Gs_compcode & "' Order by BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
PR_LMSInfo.Open "Select *,BranchCode+CustomerNo+LeaseNo As FindFld from LM_LeaseInfo where compcode ='" & Gs_compcode & "' Order by BranchCode,CustomerNo,LeaseNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1

PB_BlnkSupp = IIf(PR_LMSInfo.EOF, True, False)
txtbranchcode = Gs_BranchCode
InitializeGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload frmAssetInfo
    PR_InsurComp.Close
    pr_Customer.Close
    PR_Branch.Close
    PR_LMSInfo.Close
    Pr_Cib1.Close
    PR_AssetInfo.Close
    PR_Facility.Close
    PR_RecOfficer.Close
    PR_leaseRef.Close
    PR_AssetType.Close
    PR_LeaseRSt.Close
    PR_FCMIDs.Close
    Mode = ""
End Sub
Private Sub Command6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtlatefee
    Set PO_DESC = Text1
    
    GoTop PR_InsurComp
    MyLookup.Caption = "Recovery Officers"
    MyLookup.FillGrid PR_InsurComp, "InsurCode", "InsurDesc", txtlatefee.MaxLength
    MyLookup.Show 1
    If Len(txtlatefee) > 0 Then txtlatefee_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustNO
    Set PO_DESC = Text4
    
    Gs_SQL = "Select Customerno 'Customer No', CustomerName  'Customer Name' from Customer"
    Gs_FindFld = "CustomerName"
    Gs_OrderBy = "Order by CustomerNo,CustomerName"
    Gs_OtherPara = " Where Compcode + BranchCode = '" & Gs_compcode + Gs_BranchCode & "'"
    MyLookupOLDB.Caption = "Customers"
    MyLookupOLDB.Show 1
'
'    Set PO_AnyForm = Nothing
'    Set PO_AnyForm = Me
'    Set PO_CODE = txtCustNO
'    Set PO_DESC = Text4
'    GoTop pr_Customer
'    MyLookup.Caption = "Customer"
'    MyLookup.FillGrid pr_Customer, "CustomerNo", "CustomerName", txtCustNO.MaxLength
'    MyLookup.Show 1

    If Len(txtCustNO) > 0 Then TxtCustNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtrecoverer
    Set PO_DESC = Text1
    GoTop PR_RecOfficer
    PR_RecOfficer.Filter = "RecTag = 'REC'"
    MyLookup.Caption = "Recovery Officers"
    MyLookup.FillGrid PR_RecOfficer, "RecCode", "RecName", txtrecoverer.MaxLength
    MyLookup.Show 1
    PR_RecOfficer.Filter = adFilterNone
    If Len(txtrecoverer) > 0 Then txtrecoverer_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtleaseno
    Set PO_DESC = Text1

    PR_LMSInfo.Filter = "BranchCode = '" & Gs_BranchCode & "' And CustomerNo = '" & txtCustNO & "'"
    GoTop PR_LMSInfo
    MyLookup.Caption = "Lease Agreements"
    MyLookup.FillGrid PR_LMSInfo, "LeaseNo", "LeaseAmount", txtleaseno.MaxLength
    MyLookup.Show 1
    If Len(txtleaseno) > 0 Then txtleaseno_KeyDown vbKeyReturn, vbKeyShift
    PR_LMSInfo.Filter = adFilterNone
End Sub
Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpaidin
    Set PO_DESC = Text1
    PR_FCMIDs.Filter = "RecID = 'ADP'"
    GoTop PR_FCMIDs
    MyLookup.Caption = "Rentals Paid In"
    MyLookup.FillGrid PR_FCMIDs, "IdCode", "IdDescrip", 3
    MyLookup.Show 1
   
    If Len(txtpaidin) > 0 Then txtpaidin_KeyDown vbKeyReturn, vbKeyShift
    PR_FCMIDs.Filter = adFilterNone
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccrual
    Set PO_DESC = Text1
    PR_FCMIDs.Filter = "RecID = 'LMA'"
    GoTop PR_FCMIDs
    MyLookup.Caption = "Rentals Accrued As"
    MyLookup.FillGrid PR_FCMIDs, "IdCode", "IdDescrip", 3
    MyLookup.Show 1
   
    If Len(txtaccrual) > 0 Then txtaccrual_KeyDown vbKeyReturn, vbKeyShift
    PR_FCMIDs.Filter = adFilterNone
End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbranchcode
    Set PO_DESC = Text1
    
    GoTop PR_Branch
    MyLookup.Caption = "Company Branches"
    MyLookup.FillGrid PR_Branch, "BranchCode", "BranchDesc", txtbranchcode.MaxLength
    MyLookup.Show 1

    If Len(txtbranchcode) > 0 Then txtbranchcode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call grid1_DblClick
   If KeyCode = vbKeyDelete Then
       With Grid1
          If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             .RemoveItem .Row
             If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
             End If
       End With
   End If
End Sub

Private Sub txtaccrual_KeyDown(KeyCode As Integer, Shift As Integer)

  If Lastkey(KeyCode) And txtaccrual <> "" Then
     txtaccrual = UCase(txtaccrual)
     PR_FCMIDs.Filter = "Recid = 'LMA'"
     If Not MySeek(txtaccrual, "Idcode", PR_FCMIDs) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtaccrual.SetFocus
     Else
       txtperiod.SetFocus
     End If
     PR_FCMIDs.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtpaidin.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command3_Click
  End If
End Sub

Private Sub txtbranchcode_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
     
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
       dtpagrdate.SetFocus
     End If
ElseIf KeyCode = vbKeyF12 Then
     Command5_Click
  End If
End Sub
Private Sub txtcredit_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
       txtcredit = DoPad(txtcredit, txtcredit.MaxLength)
       PR_RecOfficer.Filter = "RecTag = 'CRE'"
     If Not MySeek(txtcredit, "RecCode", PR_RecOfficer) Then
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtcredit.SetFocus
     Else
       txtcadre.SetFocus
     End If
     PR_RecOfficer.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtmisc.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command7_Click
  End If

End Sub
Private Sub txtcadre_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
       txtcadre = DoPad(txtcadre, txtcadre.MaxLength)
       PR_RecOfficer.Filter = "RecTag = 'CAD'"
     If Not MySeek(txtcadre, "RecCode", PR_RecOfficer) Then
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtcadre.SetFocus
     Else
      If txtfrom.Enabled Then txtfrom.SetFocus
     End If
     PR_RecOfficer.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtmisc.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command8_Click
  End If

End Sub


Private Sub TxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) And Len(txtCustNO.Text) > 0 Then
        txtCustNO.Text = IIf(IsNumeric(txtCustNO), DoPad(txtCustNO, txtCustNO.MaxLength), UCase(txtCustNO))
        lb_found = MySeek(txtCustNO.Text, "CustomerNo", pr_Customer)
       
       Select Case Mode
            Case "A"
                If Not lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   txtCustNO.SetFocus
                Else
                   lb_found = MySeek(txtCustNO.Text, "CustomerNo", Pr_Cib1)
                   'If lb_found Then
                        Text4 = pr_Customer("CustomerName")
                        If MySeek(txtbranchcode + Trim(txtCustNO.Text), "FindFld", PR_Facility) Then
                                txtleaseno = DoPad(PR_Facility("FactyCntr") + 1, txtleaseno.MaxLength)
                                ls_LeaseNo = txtleaseno
                                'txtleaseno.Enabled = False
                                dtpagrdate.SetFocus
                        Else
                                Call SetErr("Customer does't have lease facility.", vbCritical)
                                txtCustNO.SetFocus
                        End If
                  'Else
                       ' Call SetErr("CIB Entry Not Enter For Selected Customer.Enter Cib Entry Before Agreement", vbOKOnly)
                       ' Unload Me
                        'frmlmscib1.Show
                  'End If
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtCustNO.SetFocus
                Else
                   Text4 = pr_Customer("CustomerName")
                   'txtleaseno.Enabled = True
                   txtleaseno.SetFocus
                End If
            End Select
   ElseIf KeyCode = vbKeyF12 Then
    cmdLookup_Click
   End If
  End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Range(Button.Index, 2, 3) Or Button.Index = 7 Then
      Command1.Enabled = True
    ElseIf Button.Index = 1 Then
      Command1.Enabled = False
    End If
    
    If PB_BlnkSupp And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found. ", vbCritical)
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_LMSInfo, Me, txtCustNO, txtbranchcode, "X", "CompCount", 3, " CustomerNo", "CustomerName", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
Dim cntsql As New ADODB.Command
Dim ln_cnt As Integer
Dim ls_CodeID As String
PB_BlnkSupp = False

cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
' Add GL Account Nos here
gc_dbcon.BeginTrans
     
     Select Case Mode
           Case "A"
           ' Creat GL Accounts
            Dim Pr_Temp As New Recordset
            Dim ln_NewAcct As String
            Dim lr_Account As String
            Dim sc_Account As String
            Dim li_Account As String
            
            If ChkGLAct.Value = 1 Then
              Set Pr_Temp = gc_dbcon.Execute("Select MAX(Cast(Acct_Detail as int)) As AcctRecs From Gl_Detail Where Right(Acct_Sub," & Gn_CustLen & ") ='" & Right(txtCustNO, Gn_CustLen) & "' And Left(Acct_Sub,6) = '006001'")
              ln_NewAcct = DoPad(Trim(Str(Val(0 & Pr_Temp("AcctRecs")) + 1)), gn_DtlLen)
              gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,BS_DrLineNo,Bs_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & "006001" + Right(txtCustNO, Gn_CustLen) & "','" & ln_NewAcct & "','" & "006001" + Right(txtCustNO, Gn_CustLen) + ln_NewAcct & "','" & Trim(Text4) + "( Lease Rental ) - " + Trim(Str(txtrental)) & "','G','B','D','PKR','0000003055','0000003055','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
              gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,PF_DrLineNo,PF_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & "200001" + Right(txtCustNO, Gn_CustLen) & "','" & ln_NewAcct & "','" & "200001" + Right(txtCustNO, Gn_CustLen) + ln_NewAcct & "','" & Trim(Text4) + "( Lease Income )" & "','G','P','C','PKR','0000000005','0000000005','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
              If Val(0 & txtdpt_v) > 0 Then gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,BS_DrLineNo,Bs_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & "140001" + Right(txtCustNO, Gn_CustLen) & "','" & ln_NewAcct & "','" & "140001" + Right(txtCustNO, Gn_CustLen) + ln_NewAcct & "','" & Trim(Text4) + "( Lease Amount ) = " + Trim(Str(txtleaseamount)) & "','G','B','D','PKR','0000000130','0000000130','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
              Pr_Temp.Close
              lr_Account = "006001" + Right(txtCustNO, Gn_CustLen) + ln_NewAcct
              sc_Account = "140001" + Right(txtCustNO, Gn_CustLen) + ln_NewAcct
              li_Account = "200001" + Right(txtCustNO, Gn_CustLen) + ln_NewAcct
            End If
            
            cntsql.CommandText = "INSERT into LM_LeaseInfo(Compcode,BranchCode,CustomerNo,LeaseNo,AgreemntDate,DisbrDate,SchdlDate,AssetType,LeaseRef,PaidAs,AccrualType,LeasePeriod,LeaseIRR,LeaseAmount,LeaseDeposit,LeaseResidual,LeaseFEF,MGTFee,DOCFee,OtherFee1,OtherFee2,OtherFee3,InsurCode,TaxRate,RecCode,CreditCode,CadreCode,UserId,TransDate,TransTime,LeaseRental,LR_AccountNo,SC_AccountNo,LI_AccountNo) VALUES  ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNO.Text & "','" & txtleaseno.Text & "','" & Format(dtpagrdate, "YYYY/MM/DD") & "','" & Format(dtpdisbr, "YYYY/MM/DD") & "','" & Format(DTPschdl, "YYYY/MM/DD") & "','" & frmAssetInfo.TxtAssetType & "','" & frmAssetInfo.TxtLeaseRef & "','" & txtpaidin & "',"
            cntsql.CommandText = cntsql.CommandText + "'" & txtaccrual & "'," & Val(0 & txtperiod) & "," & Val(0 & txtirr) & "," & Val(0 & txtleaseamount) & "," & Val(0 & txtdpt_v) & "," & Val(0 & txtrsd_v) & "," & Val(0 & txtfef_v) & "," & Val(0 & txtmgtfee) & "," & Val(0 & txtdocfee) & "," & Val(0 & txtinsurance) & "," & Val(0 & txtwtax) & "," & Val(0 & txtmisc) & ",'" & txtlatefee & "'," & Val(ln_TaxAmt) & ",'" & txtrecoverer & "','" & txtcredit & "','" & txtcadre & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "'," & Val(txtrental) & ",'" & lr_Account & "','" & sc_Account & "','" & li_Account & "')"
            cntsql.Execute
            
            cntsql.CommandText = "INSERT into LM_AssetInfo(Compcode,BranchCode,CustomerNo,LeaseNo,AssetType,AssetDesc,Manufacturer,InstlSite,RegsNo,EnginNo,chasisNo) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNO.Text & "','" & txtleaseno.Text & "','" & frmAssetInfo.TxtAssetType & "','" & frmAssetInfo.txtAssetDecs & "','" & frmAssetInfo.TxtManu & "','" & frmAssetInfo.TxtSite & "','" & frmAssetInfo.TxtReg & "','" & frmAssetInfo.TxtEngine & "','" & frmAssetInfo.Txtchasis & "')"
            cntsql.Execute
            
            If PI_SrNo > 0 Then
            With Grid1
            For ln_cnt = 1 To (Grid1.Rows - 1)
                cntsql.CommandText = "INSERT into LM_LeaseRST(Compcode,BranchCode,CustomerNo,LeaseNo,RangeFrom,RangeTo,EditedRental) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNO.Text & "','" & txtleaseno.Text & "'," & Val(.TextMatrix(ln_cnt, 1)) & "," & Val(.TextMatrix(ln_cnt, 2)) & "," & Val(.TextMatrix(ln_cnt, 3)) & ")"
                cntsql.Execute
            Next
            End With
            End If
            
            Call CreatShdl ' Creat Lease Re-payment Schedule
            If txtleaseno = ls_LeaseNo Then
               PR_Facility("FactyCntr") = PR_Facility("FactyCntr") + 1
               PR_Facility.Update
            End If
           Case "E"
              cntsql.CommandText = "UPDATE LM_LeaseInfo SET AssetType = '" & frmAssetInfo.TxtAssetType & "',AgreemntDate = '" & Format(dtpagrdate, "YYYY/MM/DD") & "',DisbrDate = '" & Format(dtpdisbr, "YYYY/MM/DD") & "',LeaseRef = '" & frmAssetInfo.TxtLeaseRef & "',LeaseFEF=" & Val(0 & txtfef_v) & ",MGTFee=" & Val(0 & txtmgtfee) & ",DOCFee=" & Val(0 & txtdocfee) & ",OtherFee1=" & Val(0 & txtinsurance) & ",OtherFee2 = " & Val(0 & txtwtax) & ",OtherFee3 = " & Val(0 & txtmisc) & ",InsurCode = '" & txtlatefee & "',RecCode= '" & txtrecoverer & "',CreditCode= '" & txtcredit & "',CadreCode= '" & txtcadre & "' WHERE  compcode = '" & Gs_compcode & "' And Branchcode = '" & Gs_BranchCode & "' And CustomerNo='" & txtCustNO.Text & "' And leaseNo = '" & txtleaseno & "'"
              cntsql.Execute
              
              cntsql.CommandText = "UPDATE LM_AssetInfo SET AssetType = '" & frmAssetInfo.TxtAssetType & "',Assetdesc = '" & frmAssetInfo.txtAssetDecs & "',Manufacturer = '" & frmAssetInfo.TxtManu & "',InstlSite = '" & Trim(frmAssetInfo.TxtSite) & "',RegsNo = '" & Trim(frmAssetInfo.TxtReg) & "',EnginNo = '" & Trim(frmAssetInfo.TxtEngine) & "',ChasisNo = '" & Trim(frmAssetInfo.Txtchasis) & "' WHERE  compcode = '" & Gs_compcode & "' And Branchcode = '" & Gs_BranchCode & "' And CustomerNo='" & txtCustNO.Text & "' And leaseNo = '" & txtleaseno & "'"
              cntsql.Execute
           Case "D"
             'Lease Agreement
              cntsql.CommandText = "DELETE FROM LM_LeaseInfo WHERE  compcode = '" & Gs_compcode & "' And Branchcode = '" & Gs_BranchCode & "' And CustomerNo='" & txtCustNO.Text & "' And leaseNo = '" & txtleaseno & "'"
              cntsql.Execute
             ' Lease Asset Information
              cntsql.CommandText = "DELETE FROM LM_AssetInfo WHERE  compcode = '" & Gs_compcode & "' And Branchcode = '" & Gs_BranchCode & "' And CustomerNo='" & txtCustNO.Text & "' And leaseNo = '" & txtleaseno & "'"
              cntsql.Execute
             ' Lease Schedule Re-Structuring
              cntsql.CommandText = "DELETE FROM LM_LeaseRST WHERE  compcode = '" & Gs_compcode & "' And Branchcode = '" & Gs_BranchCode & "' And CustomerNo='" & txtCustNO.Text & "' And leaseNo = '" & txtleaseno & "'"
              cntsql.Execute
             'Lease Schedule
              cntsql.CommandText = "DELETE FROM LM_Schedule WHERE  compcode = '" & Gs_compcode & "' And Branchcode = '" & Gs_BranchCode & "' And CustomerNo='" & txtCustNO.Text & "' And leaseNo = '" & txtleaseno & "'"
              cntsql.Execute
     End Select
           
gc_dbcon.CommitTrans

PR_LMSInfo.Requery
PR_AssetInfo.Requery
PR_LeaseRSt.Requery
SetClear frmAssetInfo

     txtTaxRate.Visible = False
     Label22.Visible = False
     InitializeGrid
     PI_SrNo = 0
     PS_RowClicked = ""
     ls_LeaseNo = ""
End Sub
Private Sub SetVal()
     dtpagrdate = PR_LMSInfo("AgreemntDate")
     dtpdisbr = PR_LMSInfo("DisbrDate")
     DTPschdl = PR_LMSInfo("SchdlDate")
     txtleaseamount = PR_LMSInfo("leaseAmount")
     txtTaxRate = PR_LMSInfo("TaxRate")
     txtdpt_v = PR_LMSInfo("LeaseDeposit")
     txtrsd_v = PR_LMSInfo("LeaseResidual")
     txtfef_v = PR_LMSInfo("LeaseFEF")
     txtdpt_p = Round((PR_LMSInfo("LeaseDeposit") / PR_LMSInfo("leaseAmount")) * 100, 3)
     txtrsd_p = Round((PR_LMSInfo("LeaseResidual") / PR_LMSInfo("leaseAmount")) * 100, 3)
     txtfef_p = Round((PR_LMSInfo("LeaseFef") / PR_LMSInfo("leaseAmount")) * 100, 3)
     txtpaidin = PR_LMSInfo("PaidAS") & ""
     txtaccrual = PR_LMSInfo("AccrualType") & ""
     txtperiod = PR_LMSInfo("LeasePeriod")
     txtirr = PR_LMSInfo("LeaseIRR")
     txtrental = PR_LMSInfo("LeaseRental")
     txtmgtfee = PR_LMSInfo("MgtFee")
     txtdocfee = PR_LMSInfo("DocFee")
     txtinsurance = PR_LMSInfo("OtherFee1")
     txtwtax = PR_LMSInfo("OtherFee2")
     txtmisc = PR_LMSInfo("OtherFee3")
     txtlatefee = PR_LMSInfo("InsurCode") & ""
     txtrecoverer = PR_LMSInfo("RecCode") & ""
     txtcredit = PR_LMSInfo("CreditCode") & ""
     txtcadre = PR_LMSInfo("CadreCode") & ""
     SetClear frmAssetInfo
     frmAssetInfo.TxtAssetType = Trim(PR_LMSInfo("AssetType")) & ""
     frmAssetInfo.TxtLeaseRef = Trim(PR_LMSInfo("LeaseRef")) & ""
     frmAssetInfo.txtAssetDecs = Trim(PR_AssetInfo("AssetDesc")) & ""
     frmAssetInfo.TxtSite = Trim(PR_AssetInfo("InstlSite")) & ""
     frmAssetInfo.TxtManu = Trim(PR_AssetInfo("Manufacturer")) & ""
     frmAssetInfo.TxtReg = Trim(PR_AssetInfo("RegsNo")) & ""
     frmAssetInfo.TxtEngine = Trim(PR_AssetInfo("EnginNo")) & ""
     frmAssetInfo.Txtchasis = Trim(PR_AssetInfo("ChasisNo")) & ""
     LoadGRNTrans
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtCustNO.Text) = txtCustNO.MaxLength And Len(txtleaseno) = txtleaseno.MaxLength And txtleaseamount <> "" And txtirr <> "" And txtaccrual <> "" And txtperiod <> "" And txtpaidin <> "" And txtrental <> "" And frmAssetInfo.TxtAssetType <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub InitializeGrid()
PI_SrNo = 0
    With Grid1
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<From|<To |<Rental  "
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 2000
        .Redraw = True
    End With
End Sub

Private Sub AddGrid()
Dim ln_cnt As Integer

        If PS_RowClicked = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
            With Grid1
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
                
                .TextMatrix(.Row, 1) = Val(0 & txtfrom)
                .TextMatrix(.Row, 2) = Val(0 & txtto)
                .TextMatrix(.Row, 3) = Val(0 & txteditrental)
                
                txtfrom.Text = ""
                txtto.Text = ""
                txteditrental.Text = ""
                PS_RowClicked = ""
                txtfrom.SetFocus
            End With
End Sub
Private Sub LoadGRNTrans()
Dim lb_found As Boolean
InitializeGrid
    
    lb_found = MySeek(txtbranchcode + txtCustNO + txtleaseno, "FindFld", PR_LeaseRSt)
   
    If lb_found Then
        With Grid1
            Do While txtbranchcode + txtCustNO + txtleaseno = PR_LeaseRSt("FindFld")
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = PR_LeaseRSt("RangeFrom")
                .TextMatrix(.Row, 2) = PR_LeaseRSt("RangeTo")
                .TextMatrix(.Row, 3) = PR_LeaseRSt("EditedRental")
                .Rows = .Rows + 1
                PR_LeaseRSt.MoveNext
                If PR_LeaseRSt.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
End Sub
Private Sub grid1_DblClick()
    With Grid1
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
       txtfrom = .TextMatrix(.Row, 1)
       txtto = .TextMatrix(.Row, 2)
       txteditrental = .TextMatrix(.Row, 3)
       PS_RowClicked = "Y"
       txtfrom.SetFocus
    End With
End Sub

Private Sub txtdocfee_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtlatefee.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtmgtfee.SetFocus
  End If
End Sub

Private Sub txtdpt_p_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
        txtdpt_v = Round((Val(0 & txtleaseamount) * Val(0 & txtdpt_p)) / 100, 0)
     If Val(0 & txtdpt_v) > 0 Then
        txtrsd_p.SetFocus
     Else
        txtdpt_v.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     dtpdisbr.SetFocus
  End If
End Sub

Private Sub txtdpt_v_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtdpt_p = Round((Val(0 & txtdpt_v) / Val(0 & txtleaseamount)) * 100, 3)
     txtrsd_p.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtdpt_p.SetFocus
  End If
End Sub


Private Sub txteditrental_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txteditrental <> "" Then
     AddGrid
  ElseIf KeyCode = vbKeyPageUp Then
     txtto.SetFocus
  End If
End Sub

Private Sub txtfef_p_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtfef_v = Round((Val(0 & txtleaseamount) * Val(0 & txtfef_p)) / 100, 0)
     If Val(0 & txtfef_v) > 0 Then
        txtpaidin.SetFocus
     Else
        txtfef_v.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     txtrsd_v.SetFocus
  End If
End Sub

Private Sub txtfef_v_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtfef_p = Round((Val(0 & txtfef_v) / Val(0 & txtleaseamount)) * 100, 3)
     txtpaidin.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtfef_p.SetFocus
  End If
End Sub


Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And Range(Val(0 & txtfrom), ln_PrvUpto + 1, Val(0 & txtperiod)) Then
     txtto.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtrecoverer.SetFocus
  End If
End Sub

Private Sub txtinsurance_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtwtax.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtlatefee.SetFocus
  End If
End Sub

Private Sub txtirr_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtirr <> "" Then
     Dim ln_Rental(1 To 6) As Double
     ln_TaxAmt = 0
     
     If Val(0 & txtleaseamount) > Val(0 & Para_Rs("LMS_VehMax")) Then ' if lease amount > 600,000
        ln_TaxAmt = Round((CalcTaxDepr() / Val(0 & txtleaseamount)) / 12, 6)
        ln_TaxAmt = IIf(ln_TaxAmt < 0 Or ln_TaxAmt = Null, 0, ln_TaxAmt)
     End If
     
     ln_Rental(1) = Val(0 & txtdpt_v) - Val(0 & txtrsd_v)  ' FV
     ln_Rental(1) = IIf(ln_Rental(1) > 0, 0, ln_Rental(1)) 'FV
     ln_Rental(2) = Val(0 & txtleaseamount) - Val(0 & txtdpt_v) ' PV
     ln_Rental(3) = Round((Val(0 & txtirr) / 1200) + ln_TaxAmt, 10) 'IRR + Tax Depr
     ln_Rental(4) = Val(0 & txtperiod) 'Period
     ln_Rental(5) = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
     ln_Rental(6) = IIf(txtpaidin = "A", 1, 0) '0/1 Advance/Arrears
     
     txtrental = Module1.CalcRental(ln_Rental)
     txtirr = Round(ln_Rental(3) * 1200, 6)
     DTPschdl.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
      txtperiod.SetFocus
  End If
End Sub

Private Sub txtlatefee_KeyDown(KeyCode As Integer, Shift As Integer)
   
   If Lastkey(KeyCode) And txtlatefee <> "" Then
       txtlatefee = DoPad(txtlatefee, txtlatefee.MaxLength)
     If Not MySeek(txtlatefee, "InsurCode", PR_InsurComp) Then
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtlatefee.SetFocus
     Else
        txtinsurance.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     txtdocfee.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command6_Click
  End If
End Sub

Private Sub txtleaseamount_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtleaseamount <> "" Then
     If Val(0 & txtleaseamount) > Val(0 & Para_Rs("LMS_VehMax")) Then
        If PR_AssetType("Assetclass") = "V" Then
           txtTaxRate.Visible = True
           Label22.Visible = True
           txtTaxRate.SetFocus
        Else
          dtpdisbr.SetFocus
        End If
     Else
       dtpdisbr.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     frmAssetInfo.Show 1
     dtpagrdate.SetFocus
  End If
End Sub

Private Sub txtleaseno_KeyDown(KeyCode As Integer, Shift As Integer)

If Lastkey(KeyCode) And txtleaseno.Text <> "" Then
   txtleaseno = DoPad(txtleaseno, txtleaseno.MaxLength)
   lb_found = MySeek(txtbranchcode + txtCustNO + txtleaseno, "FindFld", PR_LMSInfo)
   
   If lb_found Then
         If Not MySeek(txtbranchcode + txtCustNO + txtleaseno, "FindFld", PR_AssetInfo) Then
            Call SetErr("Record not found In Asset Info. Table.", vbCritical)
            txtleaseno.SetFocus
            Exit Sub
         End If
        Call SetVal
        If Mode <> "D" Then dtpagrdate.SetFocus
   Else
    If Mode <> "A" Then
      Call SetErr("Record not found", vbCritical)
      SetClear Me
      txtleaseno.SetFocus
    Else
       dtpagrdate.SetFocus
    End If
   End If
ElseIf KeyCode = vbKeyF12 Then
     Command1_Click
End If
End Sub
Private Sub txtmgtfee_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtdocfee.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     If DTPschdl.Enabled Then
        DTPschdl.SetFocus
     Else
        dtpagrdate.SetFocus
     End If
  End If
End Sub

Private Sub txtmisc_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtrecoverer.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtwtax.SetFocus
  End If
End Sub

Private Sub txtpaidin_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtpaidin <> "" Then
     txtpaidin = UCase(txtpaidin)
     PR_FCMIDs.Filter = "Recid = 'ADP'"
     If Not MySeek(txtpaidin, "Idcode", PR_FCMIDs) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtpaidin.SetFocus
     Else
       txtaccrual.SetFocus
     End If
     PR_FCMIDs.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtfef_v.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command2_Click
  End If

End Sub
Private Sub txtPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtperiod <> "" Then
     txtirr.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtaccrual.SetFocus
  End If

End Sub

Private Sub txtrecoverer_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) And txtrecoverer <> "" Then
       txtrecoverer = DoPad(txtrecoverer, txtrecoverer.MaxLength)
     If Not MySeek(txtrecoverer, "RecCode", PR_RecOfficer) Then
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtrecoverer.SetFocus
     Else
        txtcredit.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     txtmisc.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command4_Click
  End If
End Sub

Private Sub txtrsd_p_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtrsd_v = Round((Val(0 & txtleaseamount) * Val(0 & txtrsd_p)) / 100, 0)
     If Val(0 & txtrsd_v) > 0 Then
        txtfef_p.SetFocus
     Else
        txtrsd_v.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     txtdpt_v.SetFocus
  End If
End Sub

Private Sub txtrsd_v_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtrsd_p = Round((Val(0 & txtrsd_v) / Val(0 & txtleaseamount)) * 100, 3)
     txtfef_p.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtrsd_p.SetFocus
  End If
End Sub

Private Sub txtTaxRate_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtTaxRate <> "" Then
     txtdpt_p.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtleaseamount.SetFocus
  End If
End Sub

Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And Range(Val(0 & txtto), Val(0 & txtfrom), Val(0 & txtperiod)) Then
     ln_PrvUpto = Val(0 & txtto)
     txteditrental.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtfrom.SetFocus
  End If

End Sub

Private Sub txtwtax_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtmisc.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtinsurance.SetFocus
  End If

End Sub

Private Function CalcTaxDepr() As Double
'RDC Method of Calculation

Dim ln_LeasAmt As Double
Dim ln_FinLimit As Double
Dim ln_TaxRate As Double
Dim ln_Period As Integer
Dim ln_Allow As Double
Dim ln_NetAllow As Double
Dim ln_LAValue As Double
Dim ln_TxValue As Double
Dim ln_Differ As Double
Dim ln_cnt As Integer

ln_LeaseAmt = Val(0 & txtleaseamount) ' Lease Amount
ln_FinLimit = Val(0 & Para_Rs("LMS_VehMax")) ' Maximum lease facility in case of vehicle
ln_TaxRate = Round(Val(0 & txtTaxRate) / 100, 6) ' Tax Depr Rate
ln_Period = 1 ' Period = 12 months

For ln_cnt = 1 To Val(0 & txtperiod) Step 12
    ln_LAValue = Round((ln_LeaseAmt * Val(0 & Para_Rs("LMS_TaxDepr"))) / 100, 0)
    ln_TxValue = Round((ln_FinLimit * Val(0 & Para_Rs("LMS_TaxDepr"))) / 100, 0)
    ln_Differ = ln_LAValue - ln_TxValue
        
    ln_Allow = Round((ln_Differ * ln_TaxRate) * Round(((1 / (1 + ln_TaxRate)) ^ ln_Period), 6), 0)
    ln_NetAllow = ln_NetAllow + ln_Allow
    ln_LeaseAmt = ln_LeaseAmt - ln_LAValue
    ln_FinLimit = ln_FinLimit - ln_TxValue
    ln_Period = ln_Period + 1
Next
ln_Period = ln_Period - 1
ln_LAValue = Round(Round((ln_LeaseAmt - ln_FinLimit) * ln_TaxRate, 0) * Round(((1 / (1 + ln_TaxRate)) ^ ln_Period), 6), 0)
ln_NetAllow = ln_NetAllow + ln_LAValue
CalcTaxDepr = IIf(ln_NetAllow > 0, ln_NetAllow, 0)
End Function

Private Sub CreatShdl()
Dim ls_InstlCntr As String ' Installment Counter
Dim ld_Accrdate As Date     ' Rental Accrual Date
Dim ln_LeaseAmt As Double
Dim Ln_Profit As Double
Dim ln_IRR As Double
Dim ln_Rental As Double
Dim ln_InsuRental As Double
Dim lb_EndDay As Boolean
Dim ln_cnt, ln_Step, ln_Counter As Integer
Dim ln_Rental2 As Double
Dim Ln_Profit2 As Double
Dim ln_LeaseAmt2 As Double
Dim ln_Cost2 As Double
Dim lb_true As Boolean
Dim ld_AccrDateP As Date

ln_Step = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
ln_LeaseAmt = Val(0 & txtleaseamount) - Val(0 & txtdpt_v) ' Lease Amount - Deposit Paid By Customer
ln_IRR = Round(Val(0 & txtirr) / 1200, 6) * ln_Step ' IRR for the said period
ln_InsuRental = Round(Val(0 & txtleaseamount) * Round((Val(0 & txtinsurance) / 1200) * ln_Step, 6), 0) ' Insurance Rental Calculation per period
ld_Accrdate = DTPschdl.Value
lb_EndDay = IIf(Month(DateAdd("D", 1, DTPschdl.Value)) <> Month(DTPschdl.Value), True, False)
ln_Rental = Val(0 & txtrental) ' Set Default Rental
lb_Edited = False              ' Set Rendom Rentals to False
Ln_ReCalc = 0                  ' Set Re-calculation Tag to false
ln_Cost2 = 0
ln_Counter = 1
ln_Rental2 = ln_Rental      ' Set For Accounting Schedule
ln_LeaseAmt2 = ln_LeaseAmt  ' Set For Accounting Schedule
  
  For ln_cnt = 1 To Val(0 & txtperiod) Step ln_Step
      ' Recovery Schedule Profit
       Ln_Profit = IIf(txtpaidin = "A" And ln_cnt = 1, 0, Round(ln_LeaseAmt * ln_IRR, 0))
      ' Accounting Schedule Profit
       Ln_Profit2 = IIf(txtpaidin = "A" And ln_cnt = 1, 0, Round(ln_LeaseAmt2 * ln_IRR, 0))
       
       ln_InstlCntr = DoPad(LTrim(Str(ln_Counter)), 3)
       If PI_SrNo > 0 Then
         ' Search for Any User Defined Rental For Recovery
          ln_Rental = EditedRental(ln_Counter, ln_Rental)
         ' Search for Any User Defined Rental For Accoaunts
          ln_Rental2 = EditedRental(ln_Counter, ln_Rental2)
          
          If Ln_ReCalc = 1 And lb_Edited = False Then
            ' Recovery Schedule Rental Re-Calculation
             ln_Rental = ReCalcRental(ln_LeaseAmt, Val(0 & txtperiod) - ((ln_Counter - 1) * ln_Step), "R")
            ' Accounting Schedule Rental Re-Calculation
             ln_Rental2 = ReCalcRental(ln_LeaseAmt2, Val(0 & txtperiod) - ((ln_Counter - 1) * ln_Step), "R")
            ' Set Recovery Profit ON/OFF
             'Ln_Profit = IIf(txtpaidin = "A", 0, Ln_Profit)
            ' Set Accounting Profit ON/OFF
             'Ln_Profit2 = IIf(txtpaidin = "A", 0, Ln_Profit2)
             Ln_ReCalc = 0
          End If
       End If
      
      ' Balance Principal For Accounting Schedule
       Ln_Profit2 = IIf((ln_Rental2 - Ln_Profit2) < 0, IIf(ln_Rental2 = 0, 0, ln_Rental2), Ln_Profit2)
       ln_LeaseAmt2 = ln_LeaseAmt2 - (ln_Rental2 - Ln_Profit2)

      ' Balance Principal For Recovery Schedule
       ln_LeaseAmt = ln_LeaseAmt - (ln_Rental - Ln_Profit)
       
       gc_dbcon.Execute ("INSERT into LM_Schedule(Compcode,BranchCode,CustomerNo,LeaseNo,AccrualDate,InstallNo,InsurRental,LeaseRental,CostAmount,CostBalance,ProfitAmount,LeaseRental2,CostAmount2,CostBalance2,ProfitAmount2,RentalStatus,UserID,TransDate,TransTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNO.Text & "','" & txtleaseno.Text & "','" & Format(ld_Accrdate, "YYYY/MM/DD") & "','" & ln_InstlCntr & "'," & ln_InsuRental & "," & ln_Rental & "," & (ln_Rental - Ln_Profit) & "," & ln_LeaseAmt & "," & Ln_Profit & "," & ln_Rental2 & "," & (ln_Rental2 - Ln_Profit2) & "," & ln_LeaseAmt2 & "," & Ln_Profit2 & ",'O','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )")
       ld_Accrdate = DateAdd("M", ln_Step, ld_Accrdate)
       ld_AccrDateP = ld_Accrdate
       If lb_EndDay Then
          ld_Accrdate = DateAdd("M", ln_Step, ld_Accrdate) - Day(ld_Accrdate)
          lb_true = False
          Do While True
             ld_Accrdate = DateAdd("D", 1, ld_Accrdate)
             lb_true = True
             If Month(ld_AccrDateP) <> Month(ld_Accrdate) Then Exit Do
          Loop
          If lb_true Then ld_Accrdate = DateAdd("D", -1, ld_Accrdate)
       ElseIf Day(ld_Accrdate) <> Day(DTPschdl.Value) And Month(ld_Accrdate) <> 2 Then
          lb_true = False
          Do While True
             ld_Accrdate = DateAdd("D", 1, ld_Accrdate)
             lb_true = True
             If Month(ld_AccrDateP) <> Month(ld_Accrdate) Or Day(DTPschdl.Value) < Day(ld_Accrdate) Then Exit Do
          Loop
          If lb_true Then ld_Accrdate = DateAdd("D", -1, ld_Accrdate)
       End If
       ln_Counter = ln_Counter + 1
  Next
End Sub

Private Function EditedRental(ln_InstlNo As Integer, ln_CurRental As Double) As Double
Dim ln_Cnt2 As Integer
EditedRental = ln_CurRental
lb_Edited = False
    With Grid1
    For ln_Cnt2 = 1 To (Grid1.Rows - 1)
        If Range(ln_InstlNo, Val(.TextMatrix(ln_Cnt2, 1)), Val(.TextMatrix(ln_Cnt2, 2))) Then
           lb_Edited = True
           Ln_ReCalc = 1
           EditedRental = Val(.TextMatrix(ln_Cnt2, 3))
           Exit For
        End If
    Next
    End With
End Function

Private Function ReCalcRental(ln_Balance As Double, ln_RemPeriod As Integer, Optional ls_PaidIn As String) As Double
' RunTime Re-Calculations of Rental
Dim ln_Rental(1 To 8) As Double

     ln_Rental(1) = Val(0 & txtdpt_v) - Val(0 & txtrsd_v)  ' FV
     ln_Rental(1) = IIf(ln_Rental(1) > 0, 0, ln_Rental(1)) 'FV
     ln_Rental(2) = ln_Balance ' PV
     ln_Rental(3) = Round((Val(0 & txtirr) / 1200), 6) 'IRR + Tax Depr
     ln_Rental(4) = ln_RemPeriod 'Period
     ln_Rental(5) = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
     ln_Rental(6) = IIf(Trim(ls_PaidIn) = "R", 0, IIf(txtpaidin = "A", 1, 0)) '0/1 Advance/Arrears
     
     ReCalcRental = Module1.CalcRental(ln_Rental)
End Function

Public Sub FrmRefresh()
PR_InsurComp.Requery
PR_AssetType.Requery
PR_leaseRef.Requery
PR_RecOfficer.Requery
PR_Branch.Requery
pr_Customer.Requery
PR_Facility.Requery
PR_LeaseRSt.Requery
PR_AssetInfo.Requery
PR_LMSInfo.Requery
End Sub
