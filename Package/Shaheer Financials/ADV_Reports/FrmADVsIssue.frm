VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmADVsIssue 
   Caption         =   "Loan / Advance Agreement"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmADVsIssue.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
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
      Height          =   5700
      Left            =   -15
      TabIndex        =   21
      Top             =   570
      Width           =   6675
      Begin VB.CheckBox ChkGLAct 
         Alignment       =   1  'Right Justify
         Caption         =   "Open GL Acc :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   5205
         TabIndex        =   65
         Top             =   4830
         Width           =   1410
      End
      Begin VB.CommandButton Command11 
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
         Left            =   6270
         Picture         =   "FrmADVsIssue.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   1575
         Width           =   315
      End
      Begin VB.TextBox TxtPortfolio 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   5430
         MaxLength       =   2
         TabIndex        =   61
         Top             =   1575
         Width           =   840
      End
      Begin VB.CommandButton CmdLookUp3 
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
         Left            =   2205
         Picture         =   "FrmADVsIssue.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   60
         Tag             =   "SKIP"
         Top             =   900
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2550
         MaxLength       =   50
         TabIndex        =   59
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   900
         Width           =   2025
      End
      Begin VB.CheckBox ChkActive 
         Alignment       =   1  'Right Justify
         Caption         =   "Active Agreement :"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4845
         TabIndex        =   57
         Tag             =   "SKIPN"
         Top             =   600
         Width           =   1710
      End
      Begin VB.TextBox txtactno 
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1395
         MaxLength       =   3
         TabIndex        =   1
         Tag             =   "SKIP"
         Top             =   570
         Width           =   810
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   56
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   3240
         Width           =   1740
      End
      Begin VB.CommandButton CmdLookUp5 
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
         Left            =   2220
         Picture         =   "FrmADVsIssue.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3240
         Width           =   315
      End
      Begin VB.CommandButton CmdLookUp1 
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
         Left            =   2205
         Picture         =   "FrmADVsIssue.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   54
         Tag             =   "SKIP"
         Top             =   555
         Width           =   315
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
         Left            =   2190
         Picture         =   "FrmADVsIssue.frx":08D2
         Style           =   1  'Graphical
         TabIndex        =   53
         Tag             =   "SKIP"
         Top             =   225
         Width           =   315
      End
      Begin VB.CommandButton CmdLookUp4 
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
         Left            =   2205
         Picture         =   "FrmADVsIssue.frx":0A44
         Style           =   1  'Graphical
         TabIndex        =   52
         Tag             =   "SKIP"
         Top             =   1230
         Width           =   315
      End
      Begin VB.CommandButton CmdLookUp6 
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
         Left            =   2220
         Picture         =   "FrmADVsIssue.frx":0BB6
         Style           =   1  'Graphical
         TabIndex        =   51
         Tag             =   "SKIP"
         Top             =   2565
         Width           =   315
      End
      Begin VB.CommandButton CmdLookUp2 
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
         Left            =   3345
         Picture         =   "FrmADVsIssue.frx":0D28
         Style           =   1  'Graphical
         TabIndex        =   50
         Tag             =   "SKIP"
         Top             =   555
         Width           =   315
      End
      Begin VB.TextBox txtpaidin 
         Height          =   315
         Left            =   1395
         MaxLength       =   1
         TabIndex        =   8
         Top             =   2895
         Width           =   810
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
         Left            =   2220
         Picture         =   "FrmADVsIssue.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2895
         Width           =   315
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2550
         MaxLength       =   50
         TabIndex        =   48
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1245
         Width           =   2025
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   47
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2910
         Width           =   1740
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2565
         MaxLength       =   50
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2565
         Width           =   1740
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   6615
         MaxLength       =   50
         TabIndex        =   36
         Top             =   210
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton CmdLookUp7 
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
         Left            =   6270
         Picture         =   "FrmADVsIssue.frx":100C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1920
         Width           =   315
      End
      Begin VB.TextBox txtrecoverer 
         Height          =   315
         Left            =   5430
         MaxLength       =   3
         TabIndex        =   10
         Tag             =   "N"
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   2535
         MaxLength       =   50
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   4065
      End
      Begin MSMask.MaskEdBox txtProcessFee 
         Height          =   300
         Left            =   5430
         TabIndex        =   12
         Top             =   2595
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   5
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
      Begin MSMask.MaskEdBox txtDocfee 
         Height          =   300
         Left            =   5430
         TabIndex        =   11
         Top             =   2265
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   5
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
         Height          =   1260
         Left            =   390
         TabIndex        =   18
         Tag             =   "SKIPN"
         Top             =   4365
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2223
         _Version        =   393216
         Cols            =   3
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
      Begin MSMask.MaskEdBox txtfrom 
         Height          =   315
         Left            =   1035
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   4020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Left            =   2295
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   4020
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Left            =   4170
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   4020
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSComCtl2.DTPicker DtpMatuDate 
         Height          =   300
         Left            =   5430
         TabIndex        =   43
         Tag             =   "Skip"
         Top             =   1245
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57016321
         CurrentDate     =   37293
      End
      Begin MSComCtl2.DTPicker DTPInstallDate 
         Height          =   300
         Left            =   5430
         TabIndex        =   13
         Top             =   2925
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
         Format          =   57016321
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   300
         Left            =   1395
         TabIndex        =   6
         Top             =   2235
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   12
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
      Begin MSMask.MaskEdBox txtPeriod 
         Height          =   300
         Left            =   1395
         TabIndex        =   4
         Top             =   1575
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0.000;(##0.000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCustNo 
         Height          =   315
         Left            =   1395
         TabIndex        =   0
         Tag             =   "SKIP"
         Top             =   225
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
      Begin MSMask.MaskEdBox txtCrncyCode 
         Height          =   300
         Left            =   1395
         TabIndex        =   3
         Tag             =   "N"
         Top             =   1245
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
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
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRate 
         Height          =   300
         Left            =   1395
         TabIndex        =   5
         Top             =   1905
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#0.000000;(#0.000000)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCalType 
         Height          =   300
         Left            =   1395
         TabIndex        =   7
         Top             =   2565
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
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
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtrental 
         Height          =   300
         Left            =   5430
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   3255
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox txtaccrual 
         Height          =   300
         Left            =   1395
         TabIndex        =   9
         Top             =   3240
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
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
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTransCode 
         Height          =   315
         Left            =   2535
         TabIndex        =   2
         Tag             =   "SKIP"
         Top             =   555
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSComCtl2.DTPicker dtpValueDate 
         Height          =   300
         Left            =   5430
         TabIndex        =   63
         Top             =   915
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   529
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
         Format          =   57016321
         CurrentDate     =   37293
      End
      Begin MSMask.MaskEdBox txtAgrType 
         Height          =   300
         Left            =   1395
         TabIndex        =   64
         Top             =   915
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
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
         Format          =   "#,##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBalance 
         Height          =   300
         Left            =   5415
         TabIndex        =   66
         Top             =   3615
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         MaxLength       =   12
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Opening Profit :"
         Height          =   210
         Index           =   12
         Left            =   4275
         TabIndex        =   67
         Top             =   3645
         Width           =   1110
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   6600
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Portfolio :"
         Height          =   210
         Index           =   7
         Left            =   4710
         TabIndex        =   58
         Top             =   1609
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Paid In :"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   765
         TabIndex        =   46
         Top             =   2925
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sch'dle Date :"
         ForeColor       =   &H00000080&
         Height          =   210
         Index           =   13
         Left            =   4395
         TabIndex        =   45
         Top             =   2952
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Maturity :"
         Height          =   210
         Index           =   3
         Left            =   4740
         TabIndex        =   44
         Top             =   1277
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Accrual As :"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   465
         TabIndex        =   42
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Installment :"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4395
         TabIndex        =   40
         Top             =   3285
         Width           =   990
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "From :"
         Height          =   225
         Left            =   375
         TabIndex        =   39
         Top             =   4050
         Width           =   615
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "To :"
         Height          =   225
         Left            =   1905
         TabIndex        =   38
         Top             =   4050
         Width           =   315
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Installment :"
         Height          =   225
         Left            =   3135
         TabIndex        =   37
         Top             =   4050
         Width           =   870
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Rec'ry Off. :"
         Height          =   225
         Left            =   4440
         TabIndex        =   35
         Top             =   1941
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Doc Fee :"
         Height          =   210
         Index           =   8
         Left            =   4695
         TabIndex        =   33
         Top             =   2288
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calculation Type :"
         Height          =   210
         Index           =   6
         Left            =   105
         TabIndex        =   32
         Top             =   2580
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Rate :"
         Height          =   210
         Index           =   4
         Left            =   960
         TabIndex        =   31
         Top             =   1935
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Currency :"
         Height          =   210
         Index           =   1
         Left            =   615
         TabIndex        =   30
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agr'ment Type :"
         Height          =   210
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Code :"
         Height          =   210
         Left            =   180
         TabIndex        =   28
         Top             =   255
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Account-Trans # :"
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   27
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Period :"
         Height          =   210
         Index           =   11
         Left            =   840
         TabIndex        =   26
         Top             =   1605
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Process Fee :"
         Height          =   210
         Index           =   10
         Left            =   4380
         TabIndex        =   25
         Top             =   2620
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         Height          =   210
         Index           =   5
         Left            =   735
         TabIndex        =   24
         Top             =   2250
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Agr. Date :"
         Height          =   210
         Index           =   2
         Left            =   4620
         TabIndex        =   23
         Top             =   945
         Width           =   780
      End
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   5820
      MaxLength       =   10
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   3105
      Visible         =   0   'False
      Width           =   270
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
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
            Caption         =   "&Find"
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
               Picture         =   "FrmADVsIssue.frx":117E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmADVsIssue.frx":15D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmADVsIssue.frx":1A26
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmADVsIssue.frx":1E7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmADVsIssue.frx":22CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmADVsIssue.frx":2722
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmADVsIssue.frx":2E76
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmADVsIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PB_Blnk As Boolean
Dim Mode As String
Dim PI_AStatus As Integer
Dim ls_tran As String
Dim PI_CurRow    As Integer
Dim PI_SrNo     As Integer
Dim PS_RowClicked As String
Dim Ln_ReCalc As Integer
Dim lb_Edited As Boolean

Dim ln_PrvUpto As Integer

Public PO_CODE As Object
Public PO_DESC As Object

Dim GL_Detail As New Recordset
Dim Pr_AdvMast As New Recordset
Dim pr_Customer As New Recordset
Dim PR_ADVOpening As New Recordset
Dim PR_FCMIDs As New Recordset
Dim PR_RecOfficer As New Recordset
Dim PR_Instmnt As New Recordset
Dim PR_Crncy As New Recordset
Dim PR_Secrtyby  As New Recordset


Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCustNo
    Set PO_DESC = Text4
    GoTop pr_Customer
    MyLookup.Caption = "Customer"
    MyLookup.FillGrid pr_Customer, "CustomerNo", "CustomerName", txtCustNo.MaxLength
    MyLookup.Show 1
    If Len(txtCustNo) > 0 Then TxtCustNo_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub cmdLookup1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtactno
    Set PO_DESC = Text1
    GoTop PR_ADVOpening
    PR_ADVOpening.Filter = "CustomerNo = '" & txtCustNo & "'"
    MyLookup.Caption = "Customer Accounts"
    MyLookup.FillGrid PR_ADVOpening, "AccountNo", "CustomerNo", txtactno.MaxLength
    MyLookup.Show 1
    If Len(txtactno) > 0 Then txtactno_KeyDown vbKeyReturn, vbKeyShift
    PR_ADVOpening.Filter = adFilterNone
End Sub

Private Sub cmdlookup2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtTransCode
    Set PO_DESC = Text1
    GoTop Pr_AdvMast
    Pr_AdvMast.Filter = "CustomerNo = '" & txtCustNo & "' and Accountno = '" & txtactno & "'"
    MyLookup.Caption = "Trans Code Of Customer"
    MyLookup.FillGrid Pr_AdvMast, "TransCode", "AccountNo", txtTransCode.MaxLength
    MyLookup.Show 1
    Pr_AdvMast.Filter = adFilterNone
    If Len(txtTransCode) > 0 Then txtTransCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub CmdLookUp3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAgrType
    Set PO_DESC = Text2
    GoTop PR_FCMIDs
    PR_FCMIDs.Filter = "Recid = 'ADV'"
    MyLookup.Caption = "Agreement Type"
    MyLookup.FillGrid PR_FCMIDs, "IdCode", "IdDescrip", txtAgrType.MaxLength + 2
    MyLookup.Show 1
    If Len(txtAgrType) > 0 Then txtagrtype_KeyDown vbKeyReturn, vbKeyShift
    PR_FCMIDs.Filter = adFilterNone
End Sub

Private Sub CmdLookup4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCrncyCode
    Set PO_DESC = Text1
    GoTop PR_Crncy
    MyLookup.Caption = "Currency Types"
    MyLookup.FillGrid PR_Crncy, "Crncy_code", "Crncy_Descrip"
    MyLookup.Show 1
    
    If txtCrncyCode <> "" Then txtcrncycode_KeyDown vbKeyReturn, vbKeyShift
End Sub



Private Sub cmdlookup5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtaccrual
    Set PO_DESC = Text6
    GoTop PR_FCMIDs
    PR_FCMIDs.Filter = "Recid = 'LMA'"
    MyLookup.Caption = "Rentals Accrued As"
    MyLookup.FillGrid PR_FCMIDs, "IdCode", "IdDescrip", txtaccrual.MaxLength + 2
    MyLookup.Show 1
    If Len(txtaccrual) > 0 Then txtaccrual_KeyDown vbKeyReturn, vbKeyShift
    PR_FCMIDs.Filter = adFilterNone
End Sub

Private Sub CmdLookUp6_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCalType
    Set PO_DESC = Text5
    GoTop PR_FCMIDs
    PR_FCMIDs.Filter = "Recid = 'ADC'"
    MyLookup.Caption = "Agreement Type"
    MyLookup.FillGrid PR_FCMIDs, "IdCode", "IdDescrip", txtAgrType.MaxLength + 2
    MyLookup.Show 1
    If Len(txtCalType) > 0 Then txtCalType_KeyDown vbKeyReturn, vbKeyShift
    PR_FCMIDs.Filter = adFilterNone
    End Sub

Private Sub CmdLookUp7_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtrecoverer
    Set PO_DESC = Text1
    GoTop PR_RecOfficer
    MyLookup.Caption = "Recovery Officers"
    MyLookup.FillGrid PR_RecOfficer, "RecCode", "RecName", txtrecoverer.MaxLength
    MyLookup.Show 1
    If Len(txtrecoverer) > 0 Then txtrecoverer_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub command11_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtPortfolio
    Set PO_DESC = Text1
    GoTop PR_Secrtyby
    MyLookup.Caption = "Customer Accounts"
    MyLookup.FillGrid PR_Secrtyby, "SecrtyCode", "SecrtyName", TxtPortfolio.MaxLength
    MyLookup.Show 1
    If Len(TxtPortfolio) > 0 Then TxtPortfolio_KeyDown vbKeyReturn, vbKeyShift
End Sub



Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtpaidin
    Set PO_DESC = Text6
    GoTop PR_FCMIDs
    PR_FCMIDs.Filter = "Recid = 'ADP'"
    MyLookup.Caption = "Rentals Accrued As"
    MyLookup.FillGrid PR_FCMIDs, "IdCode", "IdDescrip", txtpaidin.MaxLength + 2
    MyLookup.Show 1
    If Len(txtpaidin) > 0 Then txtpaidin_KeyDown vbKeyReturn, vbKeyShift
    PR_FCMIDs.Filter = adFilterNone
End Sub

Private Sub txtCalType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
Dim ln_Rental(1 To 6) As Double
 If Lastkey(KeyCode) And txtCalType.Text <> "" Then
         txtCalType = UCase(txtCalType.Text)
         PR_FCMIDs.Filter = "Recid = 'ADC'"
         lb_found = MySeek(txtCalType, "IdCode", PR_FCMIDs)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtCalType.SetFocus
         Else
             Text5 = PR_FCMIDs("IdDescrip")
             If txtCalType = "A" Then
                txtpaidin.Enabled = True
                DTPInstallDate.Enabled = True
                txtaccrual.Enabled = True
                Command2.Enabled = True
                CmdLookUp5.Enabled = True
                txtpaidin.SetFocus
             Else
                txtaccrual.Enabled = False
                Command2.Enabled = False
                CmdLookUp5.Enabled = False
                txtpaidin.Enabled = False
                DTPInstallDate.Enabled = False
                
                    If txtCalType = "S" Then
                        If txtrecoverer.Enabled Then txtrecoverer.SetFocus
                        txtaccrual.Enabled = False
                        
                        txtaccrual = ""
                    Else
                        txtaccrual.Enabled = True
                        CmdLookUp5.Enabled = True
                        txtaccrual.SetFocus
                    End If
             End If
         End If
         PR_FCMIDs.Filter = adFilterNone
 ElseIf KeyCode = vbKeyPageUp Then
        txtAmount.SetFocus
 ElseIf KeyCode = vbKeyF12 Then
    CmdLookUp6_Click
 End If
End Sub

Private Sub txtcrncycode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If Lastkey(KeyCode) And txtCrncyCode.Text <> "" Then
         txtCrncyCode = UCase(txtCrncyCode.Text)
         lb_found = MySeek(txtCrncyCode, "Crncy_Code", PR_Crncy)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtCrncyCode.SetFocus
         Else
             Text8 = PR_Crncy("Crncy_Descrip")
             If txtPeriod.Enabled Then txtPeriod.SetFocus
         End If
 ElseIf KeyCode = vbKeyPageUp Then
        txtAgrType.SetFocus
 ElseIf KeyCode = vbKeyF12 Then
    CmdLookup4_Click
 End If
End Sub

Private Sub txtdocfee_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtProcessFee.SetFocus
If KeyCode = vbKeyPageUp Then txtrecoverer.SetFocus
End Sub

Private Sub txtpaidin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
           txtpaidin.Text = UCase(txtpaidin.Text)
           PR_FCMIDs.Filter = "Recid = 'ADP'"
           If Not MySeek(txtpaidin.Text, "IdCode", PR_FCMIDs) Then
                  Call SetErr(Gs_RecNFMsg, vbCritical)
                  If txtpaidin.Enabled Then txtpaidin.SetFocus
           Else
                    Text7 = PR_FCMIDs("IdDescrip")
                    txtaccrual.SetFocus
          End If
        PR_FCMIDs.Filter = adFilterNone
    ElseIf KeyCode = vbKeyPageUp Then
        txtCalType.SetFocus
    ElseIf KeyCode = vbKeyF12 Then
        Command2_Click
    End If
End Sub

Private Sub TxtPortfolio_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim lb_found As Boolean
  If KeyCode = vbKeyReturn And Len(Trim(TxtPortfolio.Text)) > 0 Then
        TxtPortfolio.Text = IIf(IsNumeric(TxtPortfolio), DoPad(TxtPortfolio, TxtPortfolio.MaxLength), UCase(TxtPortfolio))
        lb_found = MySeek(TxtPortfolio.Text, "SecrtyCode", PR_Secrtyby)
              If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   TxtPortfolio.SetFocus
              Else
                   txtactno.SetFocus
              End If
  End If
End Sub

Private Sub txtprocessfee_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
       If DTPInstallDate.Enabled Then
            DTPInstallDate.SetFocus
       Else
            txtBalance.SetFocus
       End If
 End If
If KeyCode = vbKeyPageUp Then txtDocfee.SetFocus
End Sub

Private Sub txtrecoverer_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) And txtrecoverer <> "" Then
       txtrecoverer = DoPad(txtrecoverer, txtrecoverer.MaxLength)
     If Not MySeek(txtrecoverer, "RecCode", PR_RecOfficer) Then
        Call SetErr(Gs_RecNFMsg, vbCritical)
        txtrecoverer.SetFocus
     Else
        txtDocfee.SetFocus
     End If
  ElseIf KeyCode = vbKeyPageUp Then
     If txtCalType = "S" Then
        txtCalType.SetFocus
     Else
        txtaccrual.SetFocus
     End If
  ElseIf KeyCode = vbKeyF12 Then
    CmdLookUp7_Click
  End If
End Sub

Private Sub dtpValueDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtCrncyCode.SetFocus
    If UCase(Trim(txtCalType)) = "S" Or UCase(Trim(txtCalType)) = "I" And Mode = "A" And Trim(Val(txtTransCode)) + 1 <> 1 Then Calc_Balance
End If
If KeyCode = vbKeyPageUp Then txtAgrType.SetFocus
End Sub
Private Sub DTPInstallDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
        If txtCalType = "A" Then
            Dim ln_Days As Double
            ln_Days = Val(0 & txtPeriod) - Int(Val(0 & txtPeriod))
            'ln_Days = ln_Days * Val("1" + String(Len(Trim(Str(ln_Days))) - 1, "0"))
            If ln_Days > 0 Then ln_Days = Val(Mid(txtPeriod, InStr(txtPeriod, ".") + 1, 3))
            DtpMatuDate.Value = Format(DateAdd("m", Int(Val(0 & txtPeriod)), DTPInstallDate.Value), "dd / mm / yyyy")
            If ln_Days > 0 Then DtpMatuDate.Value = Format(DateAdd("D", ln_Days, DtpMatuDate.Value), "dd / mm / yyyy")
            txtfrom.SetFocus
        End If
ElseIf KeyCode = vbKeyPageUp Then
    txtProcessFee.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF11 Then Call DentMode(Mode, 4, Pr_AdvMast, Me, txtCustNo, txtactno, ParaCntr_Rs, "AdvTranCode", 3, "BranchCode", "BranchDesc", 1, False, Toolbar1)
End Sub
Private Sub Form_Load()
  
' Setting up Preveliges
  SetToolBar(1) = chkRights("AGREEMENT1")
  SetToolBar(2) = chkRights("AGREEMENT2")
  SetToolBar(3) = chkRights("AGREEMENT3")
  SetToolBar(4) = chkRights("AGREEMENT4")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  GL_Detail.Open "Select * from Gl_Detail where compcode ='" & Gs_compcode & "' And (Acct_Sub = '00700100001' Or Acct_Sub = '00800100001')  ORder By AccountNo", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  pr_Customer.Open "Select Customer.* from Customer Inner Join Facilities On Customer.CustomerNo = Facilities.CustomerNo Where Customer.Compcode + Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo ='02' Order By Customer.CustomerNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_ADVOpening.Open "Select *,Compcode+BranchCode+Rtrim(Ltrim(CustomerNo))+Rtrim(Ltrim(AccountNo)) As FindFld from ADV_Accounts  where compcode+branchcode ='" & Gs_compcode + Gs_BranchCode & "' Order by CustomerNo,AccountNo", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  Pr_AdvMast.Open "Select *,BranchCode+CustomerNo+AccountNo+TransCode As FindFld,BranchCode+CustomerNo+AccountNo as Findfld1 From ADV_Master where compcode+branchcode ='" & Gs_compcode + Gs_BranchCode & "' and ActiveStatus = 1 Order By BranchCode,CustomerNo,AccountNo", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_FCMIDs.Open "Select * from FCM_IDs where RecId = 'ADV' or RecId = 'ADC'or RecId = 'LMA' or RecId = 'ADP'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_RecOfficer.Open "Select * From LM_Recoverer Order By RecCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_Crncy.Open "Select * From SysCurrency Order By Crncy_Code", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Instmnt.Open "select Adv_Instmnt.*, Branchcode + Customerno+Accountno+Transcode as Findfld from ADV_Instmnt order by Findfld", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  PR_Secrtyby.Open "Select * from LM_SecrtyBy  order by SecrtyCode", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
  PB_Blnk = IIf(Pr_AdvMast.EOF, True, False)
  InitializeGrid
End Sub
Public Sub FrmRefresh()
  GL_Detail.Requery
  pr_Customer.Requery
  PR_ADVOpening.Requery
  Pr_AdvMast.Requery
  PR_FCMIDs.Requery
  PR_RecOfficer.Requery
  PR_Crncy.Requery
  PR_Instmnt.Requery

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Pr_AdvMast.Close
    PR_ADVOpening.Close
    pr_Customer.Close
    PR_FCMIDs.Close
    PR_RecOfficer.Close
    PR_Crncy.Close
    PR_Instmnt.Close
    PR_Secrtyby.Close
    GL_Detail.Close
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 6 Or Button.Index = 7 Then cmdlookup2.Enabled = True
    If PB_Blnk And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
    Else
       Mode = DentMode(Mode, Button.Index, Pr_AdvMast, Me, txtCustNo, txtactno, ParaCntr_Rs, "AdvTranCode", 3, "BranchCode", "BranchDesc", 1, False, Toolbar1)
         If Mode = "A" Then
            ChkActive.Value = 1
            ChkActive.Enabled = False
         Else
            ChkActive.Value = 0
            ChkActive.Enabled = True
         End If
    End If
End Sub

Public Sub SaveValues()
Dim ls_Transcode As String
Dim ls_Accountno As String
Dim ls_AccountDesc As String
Dim ln_cnt As Integer
PB_Blnk = False
Dim cntsql As New ADODB.Command
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
 
If Mode = "D" Then
   cntsql.CommandText = "DELETE FROM ADV_Master WHERE  CompCode+BranchCode+CustomerNo+AccountNo+Transcode = '" & Gs_compcode + Gs_BranchCode + txtCustNo.Text + txtactno + txtTransCode & "'"
   cntsql.Execute
   
   cntsql.CommandText = "DELETE FROM ADV_Trans WHERE  CompCode+BranchCode+CustomerNo+AccountNo+Transcode = '" & Gs_compcode + Gs_BranchCode + txtCustNo.Text + txtactno + txtTransCode & "'"
   cntsql.Execute
   
   cntsql.CommandText = "DELETE FROM ADV_Instmnt WHERE  CompCode+BranchCode+CustomerNo+AccountNo+Transcode = '" & Gs_compcode + Gs_BranchCode + txtCustNo.Text + txtactno + txtTransCode & "'"
   cntsql.Execute
End If
     Select Case Mode
           Case "A"
             If txtTransCode.Enabled And txtCalType <> "A" Then
                    If MySeek(Gs_compcode + Gs_BranchCode + Trim(txtCustNo.Text) + Trim(txtactno), "Findfld", PR_ADVOpening) Then
                        ls_Transcode = DoPad(Val((0 & PR_ADVOpening("AdvTransCode")) + 1), 3)
                        cntsql.CommandText = "Update ADV_Master Set ActiveStatus = 0 WHERE  CompCode+BranchCode+CustomerNo+AccountNo+Transcode = '" & Gs_compcode + Gs_BranchCode + Trim(txtCustNo.Text) + Trim(txtactno) + Trim(txtTransCode) & "'"
                        cntsql.Execute
                        txtTransCode = ls_Transcode
                    Else
                    Call SetErr("Account No Found Please Refresh Form", vbCritical)
                    Exit Sub
                    End If
             End If
              cntsql.CommandText = "INSERT into ADV_Master(CompCode,BranchCode,CustomerNo,AccountNo,TransCode,ValueDate,InstallDate,InstallAmount,AgrType,AgrPeriod,MatuyDate,Agrrate,AgrAmount,CrncyCode,AccrualAs,AgrCalcType,RecCode,DocFee,Processfee,Paidas,PortfolioId,UserId,TransDate,TransTime,ProfitBalance) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo & "','" & txtactno & "','" & txtTransCode & "','" & Format(dtpValueDate, "YYYY/MM/DD") & "','" & Format(DTPInstallDate, "YYYY/MM/DD") & "'," & Val(0 & txtrental) & ",'" & txtAgrType & "'," & txtPeriod & ",'" & Format(DtpMatuDate, "YYYY/MM/DD") & "'," & txtRate & "," & txtAmount & ",'" & txtCrncyCode & "','" & txtaccrual & "','" & txtCalType & "','" & txtrecoverer & "'," & Val(0 & txtDocfee) & "," & Val(0 & txtProcessFee) & ",'" & txtpaidin & "','" & TxtPortfolio & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "'," & Val(0 & txtBalance) & ")"
              cntsql.Execute
              If Mode = "A" Then
                 cntsql.CommandText = "Update ADV_Accounts Set ADVTransCode = " & txtTransCode & " where CompCode+Branchcode+CustomerNo+AccountNo = '" & Gs_compcode + Gs_BranchCode + txtCustNo + txtactno & "'"
                 cntsql.Execute
              End If
              
              If UCase(txtCalType) = "A" Then
                   Call CreatShdl
              ElseIf UCase(txtCalType) = "S" Then
                   Call CreatShdlOnMatrty
              ElseIf UCase(txtCalType) = "I" Then
                    Call CreatShdlOnInstallment
              End If
                
             If Mode = "A" And Trim(txtTransCode) = "001" And ChkGLAct.Value = 1 Then
             ' Open GL Account
              Dim ls_AcctSub As String
              Dim ls_AcctSub2 As String
              
              ls_AcctSub = IIf(Trim(txtAgrType) = "R", "00700100001", "00800100001")
              ls_AcctSub2 = IIf(Trim(txtAgrType) = "R", "00700100002", "00800100002")
              If Not MySeek(ls_AcctSub + Right(txtCustNo, Gn_CustLen), "AccountNo", GL_Detail) Then
                 If Trim(txtAgrType) = "R" Then ' Morabaha Facility
                    gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,BS_DrLineNo,Bs_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_AcctSub & "','" & Right(txtCustNo, Gn_CustLen) & "','" & ls_AcctSub + Right(txtCustNo, Gn_CustLen) & "','" & Trim(Text4) + "( Morabaha Facility )" & "','G','B','D','PKR','0000003020','0000003020','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
                    gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,PF_DrLineNo,PF_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_AcctSub2 & "','" & Right(txtCustNo, Gn_CustLen) & "','" & ls_AcctSub2 + Right(txtCustNo, Gn_CustLen) & "','" & Trim(Text4) + "( Accrual A/c )" & "','G','P','C','PKR','0000000010','0000000010','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
                 ElseIf Trim(txtAgrType) = "M" Or Trim(txtAgrType) = "T" Then ' Musharaka & Operating Lease Facility
                    gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,BS_DrLineNo,Bs_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_AcctSub & "','" & Right(txtCustNo, Gn_CustLen) & "','" & ls_AcctSub + Right(txtCustNo, Gn_CustLen) & "','" & Trim(Text4) + "( Musharaka Facility )" & "','G','B','D','PKR','0000003030','0000003030','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
                    gc_dbcon.Execute ("insert into GL_Detail(Compcode,BranchCode,Acct_Sub,Acct_Detail,AccountNo,Acct_Desc,Acct_Type,Acct_Base,Acct_Status,Crncy_Code,PF_DrLineNo,PF_CrLineNo,UserId,AddDate,AddTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & ls_AcctSub2 & "','" & Right(txtCustNo, Gn_CustLen) & "','" & ls_AcctSub2 + Right(txtCustNo, Gn_CustLen) & "','" & Trim(Text4) + "( Accrual A/c )" & "','G','P','C','PKR','0000000010','0000000010','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')")
                 End If
              Else
                 If Trim(ls_AcctSub) <> "" Then
                  ls_Accountno = GL_Detail("AccountNo") & ""
                  ls_AccountDesc = GL_Detail("Acct_Desc") & ""
                 End If
              End If
              'gc_dbcon.Execute ("Update ADV_Accounts Set PrftAccountNo = '" & ls_Subacct2 & "' + '" & Right(txtCustNo, Gn_CustLen) & "', PrincAccountNo = '" & ls_Accountno & "' Where Compcode+BranchCode+CustomerNo+AccountNo = '" & Gs_compcode + Gs_BranchCode + Trim(txtCustNo) + Trim(txtactno) & "' And len(PrftAccountNo) = 0 And len(PrincAccountNo) = 0 ")
              End If
                
         If PI_SrNo > 0 Then
            With Grid1
            For ln_cnt = 1 To (Grid1.Rows - 1)
                cntsql.CommandText = "INSERT into ADV_Instmnt(Compcode,BranchCode,CustomerNo,AccountNo,TransCode,RangeFrom,RangeTo,EditedRental) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtactno.Text & "','" & txtTransCode.Text & "'," & Val(.TextMatrix(ln_cnt, 1)) & "," & Val(.TextMatrix(ln_cnt, 2)) & "," & Val(.TextMatrix(ln_cnt, 3)) & ")"
                cntsql.Execute
            Next
            End With
         End If
           Case "E"
                    cntsql.CommandText = "Update ADV_Master Set ActiveStatus = " & ChkActive & ",Installamount = " & Val(0 & txtrental) & " WHERE  CompCode+BranchCode+CustomerNo+AccountNo+Transcode = '" & Gs_compcode + Gs_BranchCode + txtCustNo.Text + txtactno + txtTransCode & "'"
                    cntsql.Execute
             
                    
'           Case "D"
'              cntsql.CommandText = "DELETE FROM ADV_Master WHERE  CompCode+BranchCode+CustomerNo+AccountNo+Transcode = '" & Gs_compcode + Gs_BranchCode + txtCustNO.Text + txtactno + txtTransCode & "'"
'              cntsql.Execute
           
     End Select
InitializeGrid
Pr_AdvMast.Requery
PR_ADVOpening.Requery
End Sub
Private Sub SetVal()
dtpValueDate = Pr_AdvMast("Valuedate")
DTPInstallDate = IIf(Pr_AdvMast("InstallDate") <> "", Pr_AdvMast("InstallDate"), Pr_AdvMast("Valuedate"))
txtAgrType = Pr_AdvMast("AgrType") & ""
txtpaidin = Pr_AdvMast("Paidas") & ""
txtPeriod = Val(0 & Pr_AdvMast("AgrPeriod"))
DtpMatuDate = Pr_AdvMast("Matuydate")
txtCrncyCode = Pr_AdvMast("CrncyCode") & ""
txtaccrual = Pr_AdvMast("AccrualAS") & ""
txtRate = Val(0 & Pr_AdvMast("AgrRate"))
txtAmount = Val(0 & Pr_AdvMast("AgrAmount"))
txtCalType = Pr_AdvMast("AgrCalctype") & ""
txtrecoverer = Pr_AdvMast("Reccode") & ""
txtDocfee = Val(0 & Pr_AdvMast("DocFee"))
txtProcessFee = Val(0 & Pr_AdvMast("ProcessFee"))
TxtPortfolio = Pr_AdvMast("PortfolioId") & ""
ChkActive = Val(0 & Pr_AdvMast("ActiveStatus"))
txtrental = Val(0 & Pr_AdvMast("InstallAmount"))
txtBalance = Val(0 & Pr_AdvMast("ProfitBalance"))
If Len(Trim(txtAgrType)) > 0 Then txtagrtype_KeyDown vbKeyReturn, vbKeyShift
If Len(Trim(txtCrncyCode)) > 0 Then txtcrncycode_KeyDown vbKeyReturn, vbKeyShift
If Len(txtCalType) > 0 Then txtCalType_KeyDown vbKeyReturn, vbKeyShift
If Len(Trim(txtpaidin)) > 0 Then txtpaidin_KeyDown vbKeyReturn, vbKeyShift
If Len(Trim(txtaccrual)) > 0 Then txtaccrual_KeyDown vbKeyReturn, vbKeyShift
LoadGRNTrans

End Sub
Private Sub LoadGRNTrans()
Dim lb_found As Boolean
InitializeGrid
    lb_found = MySeek(Gs_BranchCode + Trim(txtCustNo) + Trim(txtactno) + Trim(txtTransCode), "FindFld", PR_Instmnt)
   
    If lb_found Then
        With Grid1
            Do While Gs_BranchCode + Trim(txtCustNo) + Trim(txtactno) + Trim(txtTransCode) = PR_Instmnt("FindFld")
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = PR_Instmnt("RangeFrom")
                .TextMatrix(.Row, 2) = PR_Instmnt("RangeTo")
                .TextMatrix(.Row, 3) = PR_Instmnt("EditedRental")
                .Rows = .Rows + 1
                PR_Instmnt.MoveNext
                If PR_Instmnt.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
            If .Row >= 1 Then .Enabled = True
        End With
    End If
End Sub

Public Function ChkInputs() As Boolean
    If txtCustNo <> "" And txtactno <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function


Private Sub txtactno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If Lastkey(KeyCode) And Len(txtactno.Text) > 0 Then
          txtactno = DoPad(txtactno, txtactno.MaxLength)
            lb_found = MySeek(Gs_compcode + Gs_BranchCode + Trim(txtCustNo.Text) + Trim(txtactno), "Findfld", PR_ADVOpening)
               If Not lb_found Then
                     Call SetErr(Gs_RecNFMsg, vbCritical)
                     txtactno.SetFocus
               Else
                    If Mode = "A" Then
                         ls_tran = DoPad(Val((0 & PR_ADVOpening("AdvTransCode")) + 1), 3)
                         lb_found = MySeek(Gs_BranchCode + Trim(txtCustNo.Text) + Trim(txtactno), "Findfld1", Pr_AdvMast)
                         If ls_tran = "001" Or Not lb_found Then
                            txtTransCode.Enabled = False
                            cmdlookup2.Enabled = False
                            txtTransCode = ls_tran
                            txtAgrType.SetFocus
                         Else
                             txtTransCode.Enabled = True
                             cmdlookup2.Enabled = True
                             txtTransCode.SetFocus
                         End If
                    Else
                        txtTransCode.Enabled = True
                        cmdlookup2.Enabled = True
                        txtTransCode.SetFocus
                        If txtTransCode.Enabled Then txtTransCode.SetFocus
                    End If
                      
               End If
ElseIf KeyCode = vbKeyPageUp Then
     txtCustNo.SetFocus
ElseIf KeyCode = vbKeyF12 Then
    cmdLookup1_Click
End If
End Sub
Private Sub txtagrtype_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
If Lastkey(KeyCode) And Len(txtAgrType.Text) > 0 Then
          txtAgrType = UCase(txtAgrType)
          PR_FCMIDs.Filter = "Recid = 'ADV'"
          lb_found = MySeek(txtAgrType, "IdCode", PR_FCMIDs)
               If Not lb_found Then
                     Call SetErr(Gs_RecNFMsg, vbCritical)
                     txtAgrType.SetFocus
               Else
                    Text2 = PR_FCMIDs("iddescrip")
                   If dtpValueDate.Enabled Then dtpValueDate.SetFocus
               End If
          PR_FCMIDs.Filter = adFilterNone
ElseIf KeyCode = vbKeyPageUp Then
          If txtTransCode.Enabled = True Then
                txtTransCode.SetFocus
          Else
                txtactno.SetFocus
          End If
ElseIf KeyCode = vbKeyF12 Then
    CmdLookUp3_Click
End If
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Val(0 & txtAmount) > 0 Then txtCalType.SetFocus
If KeyCode = vbKeyPageUp Then txtRate.SetFocus
End Sub
Private Sub txtPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtPeriod <> "" Then
        Dim ln_Days As Double
        ln_Days = Val(0 & txtPeriod) - Int(Val(0 & txtPeriod))
        'ln_Days = ln_Days * Val("1" + String(Len(Trim(Str(ln_Days))) - 1, "0"))
        If ln_Days > 0 Then ln_Days = Val(Mid(txtPeriod, InStr(txtPeriod, ".") + 1, 3))
        DtpMatuDate.Value = Format(DateAdd("m", Int(Val(0 & txtPeriod)), dtpValueDate.Value), "dd / mm / yyyy")
        If ln_Days > 0 Then DtpMatuDate.Value = Format(DateAdd("D", ln_Days, DtpMatuDate.Value), "dd / mm / yyyy")
        txtRate.SetFocus
        'DTPMatuDate = DateAdd("M", txtPeriod, dtpValueDate.Value)
End If
If KeyCode = vbKeyPageUp Then txtAgrType.SetFocus
End Sub
Private Sub TxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If Lastkey(KeyCode) And Len(txtCustNo.Text) > 0 Then
        txtCustNo.Text = IIf(IsNumeric(txtCustNo), DoPad(txtCustNo, txtCustNo.MaxLength), UCase(txtCustNo))
        lb_found = MySeek(txtCustNo.Text, "CustomerNo", pr_Customer)
              If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtCustNo.SetFocus
                Else
                   Text4 = pr_Customer("CustomerName")
                   txtactno.SetFocus
                End If
 ElseIf KeyCode = vbKeyF12 Then
    cmdLookup_Click
 End If
End Sub
Private Sub txtaccrual_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ln_Rental(1 To 6) As Double
Dim lb_found As Boolean
    If Lastkey(KeyCode) And Len(txtaccrual.Text) > 0 Then
           txtaccrual.Text = UCase(txtaccrual.Text)
           PR_FCMIDs.Filter = "Recid = 'LMA'"
           lb_found = MySeek(txtaccrual.Text, "IdCode", PR_FCMIDs)
            If Not lb_found Then
                  Call SetErr(Gs_RecNFMsg, vbCritical)
                  txtaccrual.SetFocus
            Else
                    ln_Rental(1) = 0 'Val(0 & txtdpt_v) - Val(0 & txtrsd_v)  ' FV
                    ln_Rental(1) = 0 'IIf(ln_Rental(1) > 0, 0, ln_Rental(1)) 'FV
                    ln_Rental(2) = Val(0 & txtAmount) ' PV
                    ln_Rental(3) = Round((Val(0 & txtRate) / 1200), 10)  'IRR + Tax Depr
                    ln_Rental(4) = Val(0 & txtPeriod) 'Period
                    ln_Rental(5) = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1)))  'Accrual As 3/6/12
                    ln_Rental(6) = IIf(txtpaidin = "A", 1, 0)
                    txtrental = Module1.CalcRental(ln_Rental)
                    txtfrom.Enabled = True
                    txtto.Enabled = True
                    txteditrental.Enabled = True
                    Grid1.Enabled = True
                    Text6 = PR_FCMIDs("IdDescrip")
                    If txtrecoverer.Enabled Then txtrecoverer.SetFocus
            End If
            PR_FCMIDs.Filter = adFilterNone
    ElseIf KeyCode = vbKeyPageUp Then
        If txtpaidin.Enabled Then
            txtpaidin.SetFocus
        Else
            txtCalType.SetFocus
        End If
    ElseIf KeyCode = vbKeyF12 Then
        cmdlookup5_Click
    End If
End Sub

Private Sub txtrate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtRate.Text <> "" Then txtAmount.SetFocus
If KeyCode = vbKeyPageUp Then txtPeriod.SetFocus
End Sub

Private Sub txtTransCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 If KeyCode = vbKeyReturn And Len(txtTransCode.Text) > 0 Then
        txtTransCode = DoPad(txtTransCode, txtTransCode.MaxLength)
         lb_found = MySeek(Gs_BranchCode + txtCustNo + txtactno + txtTransCode, "FindFld", Pr_AdvMast)
              If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   txtTransCode.SetFocus
                Else
                Call SetVal
                If txtAgrType.Enabled Then txtAgrType.SetFocus
              End If
 ElseIf KeyCode = vbKeyPageUp Then
    txtactno.SetFocus
 ElseIf KeyCode = vbKeyF12 Then
    cmdlookup2_Click
 End If
 
End Sub
Private Sub Calc_Balance()

Dim PR_Tmp As New Recordset
Dim ld_ABalance As Double
Dim ld_Balance As Double
Dim ls_Sql As String
    
    ls_Sql = "Select  Sum(ProfitAmount) as ProfitAmount from ADV_Trans "
    ls_Sql = ls_Sql + " Where AccrualDate <= '" & Format(dtpValueDate, "YYYY/MM/DD") & "' And Compcode+BranchCode = '" & Gs_compcode + Gs_BranchCode & "' and Customerno+Accountno+Transcode = '" & Trim(txtCustNo) + Trim(txtactno) + Trim(txtTransCode) & "'"
    ls_Sql = ls_Sql + " Group By  Customerno,Accountno,Transcode"
     
    PR_Tmp.Open ls_Sql, gc_dbcon, adOpenStatic, adLockReadOnly, adCmdText
     
    If Not PR_Tmp.EOF Then
        ld_ABalance = Val(0 & PR_Tmp("ProfitAmount"))
    End If
    PR_Tmp.Close
    
    ls_Sql = "Select  Sum(PaymentAmt) as PaymentAmt from ADV_Payments "
    ls_Sql = ls_Sql + " Where  Compcode+BranchCode = '" & Gs_compcode + Gs_BranchCode & "' and Customerno+Accountno+TransRef = '" & Trim(txtCustNo) + Trim(txtactno) + Trim(txtTransCode) & "' and TransType =0"
    ls_Sql = ls_Sql + " Group By  Customerno,Accountno,TransRef"
    
    PR_Tmp.Open ls_Sql, gc_dbcon, adOpenStatic, adLockReadOnly, adCmdText
     
    If Not PR_Tmp.EOF Then
        ld_Balance = Val(0 & PR_Tmp("PaymentAmt"))
    End If
    PR_Tmp.Close
    If Not Pr_AdvMast.EOF Then
        txtBalance = ld_ABalance + Val(0 & Pr_AdvMast("ProfitBalance")) - ld_Balance
    End If

End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And Range(Val(0 & txtfrom), ln_PrvUpto + 1, Val(0 & txtPeriod)) Then
     txtto.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtrecoverer.SetFocus
  End If
End Sub
Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And Range(Val(0 & txtto), Val(0 & txtfrom), Val(0 & txtPeriod)) Then
     ln_PrvUpto = Val(0 & txtto)
     txteditrental.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtfrom.SetFocus
  End If

End Sub
Private Sub txteditrental_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txteditrental <> "" Then
     AddGrid
  ElseIf KeyCode = vbKeyPageUp Then
     txtto.SetFocus
  End If
End Sub
Public Sub InitializeGrid()
   
    With Grid1
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr# |<From|<To |<Installment  "
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
'Private Sub LoadGRNTrans()
'Dim lb_found As Boolean
'InitializeGrid
'
'    lb_found = MySeek(txtbranchcode + txtCustNO + txtleaseno, "FindFld", PR_LeaseRSt)
'
'    If lb_found Then
'        With Grid1
'            Do While txtbranchcode + txtCustNO + txtleaseno = PR_LeaseRSt("FindFld")
'                .Row = .Rows - 1
'                .TextMatrix(.Row, 0) = .Row
'                 PI_SrNo = Val(.TextMatrix(.Row, 0))
'                .TextMatrix(.Row, 1) = PR_LeaseRSt("RangeFrom")
'                .TextMatrix(.Row, 2) = PR_LeaseRSt("RangeTo")
'                .TextMatrix(.Row, 3) = PR_LeaseRSt("EditedRental")
'                .Rows = .Rows + 1
'                PR_LeaseRSt.MoveNext
'                If PR_LeaseRSt.EOF Then Exit Do
'             Loop
'            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
'        End With
'    End If
'End Sub
Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    With Grid1
        If KeyCode = vbKeyDelete Then
            If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
            .RemoveItem .Row
            If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
            End If
        ElseIf KeyCode = vbKeyReturn Then
            grid1_DblClick
        End If
    End With
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
Private Sub CreatShdlOnMatrty()
Dim ln_InstlCntr As String ' Installment Counter
Dim ls_Sql As String
Dim ln_Count As Double
Dim ld_CDate As Date
Dim PR_Paracount As New Recordset
Dim Pr_Temp As New Recordset
Dim ln_Step As Integer
Dim ln_Days As Double
Dim ln_AccrAmount As Double
Dim ln_AccrTax As Double
Dim ls_Exper2 As String
Dim ln_Choice As Integer
Dim ln_cnt As Integer
Dim ln_outcnt As Integer
Dim ln_UptoCnt As Integer
Dim ln_Counter As Integer
Dim ld_Accrdate As Date
ln_Counter = 1
       
       PR_Paracount.Open "select * from paracount", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
       ld_Accrdate = dtpValueDate
       ln_UptoCnt = 1
       
         ln_InstlCntr = DoPad(LTrim(Str(ln_Counter)), 3)
         ld_CDate = DtpMatuDate
          
          ln_Days = Round(DateDiff("D", ld_Accrdate, ld_CDate) + 0.0001, 4)
            ld_Accrdate = ld_CDate
            If Trim(txtAgrType) = "M" Then
                ls_Sql = PR_Paracount("Mushformula")
            ElseIf Trim(txtAgrType) = "R" Then
                ls_Sql = PR_Paracount("Modarformula")
            ElseIf Trim(txtAgrType) = "T" Then
                ls_Sql = PR_Paracount("Tradformula")
            End If
            
                   If Trim(ls_Sql) <> "" Then
                          ls_Sql = Replace(ls_Sql, "INVESTMENT", Trim(Str(txtAmount)))
                          ls_Sql = Replace(ls_Sql, "YIELD", Trim(Str(Round(Val(txtRate) + 0.00001, 5))))
                          ls_Sql = Replace(ls_Sql, "DAYS", Trim(Str(ln_Days)))
                          If Year(ld_CDate) Mod 4 = 0 Then
                            ls_Sql = Replace(ls_Sql, "CYEAR", 366)
                          Else
                            ls_Sql = Replace(ls_Sql, "CYEAR", 365)
                          End If
                          ls_Sql = "Select (" & ls_Sql & ") As f1"
                          Set Pr_Temp = gc_dbcon.Execute(ls_Sql)
                          ln_AccrAmount = Round(Pr_Temp.Fields("f1").Value, 0)
                          Pr_Temp.Close
                            If ln_AccrAmount > 0 Then
                                gc_dbcon.BeginTrans
                                        gc_dbcon.Execute ("INSERT into ADV_Trans(Compcode,BranchCode,CustomerNo,AccountNo,TransCode,AccrualDate,InstallNo,CostAmount,ProfitAmount,PaidDays,UserID,TransDate,TransTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtactno.Text & "','" & txtTransCode.Text & "','" & Format(ld_Accrdate, "YYYY/MM/DD") & "','" & ln_InstlCntr & "'," & txtAmount & "," & ln_AccrAmount & "," & ln_Days & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )")
                                gc_dbcon.CommitTrans
                                Else
                                  Call SetErr("Calculation Error. On Customer = ", vbCritical)
                            End If
                     Else
                         Call SetErr("Calculation Formula Not Found.", vbCritical)
                     End If
PR_Paracount.Close
End Sub
Private Sub CreatShdlOnInstallment()
Dim ln_InstlCntr As String ' Installment Counter
Dim ls_Sql As String
Dim ln_Count As Double
Dim ld_CDate As Date
Dim ld_Accrdate As Date     ' Rental Accrual Date
Dim PR_Paracount As New Recordset
Dim Pr_Temp As New Recordset
Dim ln_Step As Integer
Dim ln_Days As Double
Dim ln_AccrAmount As Double
Dim ln_AccrTax As Double
Dim ls_Exper2 As String
Dim ln_Choice As Integer
Dim ln_cnt As Integer
Dim ln_outcnt As Integer
Dim ln_Counter As Integer
ln_Counter = 1
       PR_Paracount.Open "select * from paracount", gc_dbcon, adOpenStatic, adLockPessimistic, adCmdText
       ld_Accrdate = dtpValueDate
       ln_Step = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1)))
       
   For ln_outcnt = 1 To 1
       For ln_cnt = 1 To Val(0 & txtPeriod) Step ln_Step
          ln_InstlCntr = DoPad(LTrim(Str(ln_Counter)), 3)
          ld_CDate = DateAdd("M", ln_Step, ld_Accrdate)
          ln_Days = Round(DateDiff("D", ld_Accrdate, ld_CDate) + 0.000001, 6)
            If Trim(txtAgrType) = "M" Then
                ls_Sql = PR_Paracount("Mushformula")
            ElseIf Trim(txtAgrType) = "R" Then
                ls_Sql = PR_Paracount("Modarformula")
            ElseIf Trim(txtAgrType) = "T" Then
                ls_Sql = PR_Paracount("Tradformula")
            End If
                   If Trim(ls_Sql) <> "" Then
                          ls_Sql = Replace(ls_Sql, "INVESTMENT", Trim(Str(txtAmount)))
                          ls_Sql = Replace(ls_Sql, "YIELD", Trim(Str(Round(Val(txtRate) + 0.000001, 6))))
                          ls_Sql = Replace(ls_Sql, "DAYS", Trim(Str(ln_Days)))
                          If Year(ld_CDate) Mod 4 = 0 Then
                            ls_Sql = Replace(ls_Sql, "CYEAR", 366)
                          Else
                            ls_Sql = Replace(ls_Sql, "CYEAR", 365)
                          End If
                          ls_Sql = "Select (" & ls_Sql & ") As f1"
                          Set Pr_Temp = gc_dbcon.Execute(ls_Sql)
                          ln_AccrAmount = Round(Pr_Temp.Fields("f1").Value, 0)
                          Pr_Temp.Close
                          
                            If ln_AccrAmount > 0 Then
                                gc_dbcon.BeginTrans
                                 If ln_outcnt = 1 Then
                                    gc_dbcon.Execute ("INSERT into ADV_Trans(Compcode,BranchCode,CustomerNo,AccountNo,TransCode,AccrualDate,InstallNo,CostAmount,ProfitAmount,PaidDays,UserID,TransDate,TransTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtactno.Text & "','" & txtTransCode.Text & "','" & Format(ld_CDate, "YYYY/MM/DD") & "','" & ln_InstlCntr & "'," & txtAmount & "," & ln_AccrAmount & "," & ln_Days & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )")
                                 ElseIf ln_outcnt = 2 Then
                                    gc_dbcon.Execute ("INSERT into ADV_Accruals(Compcode,BranchCode,CustomerNo,AccountNo,TransCode,ValueDate,AccrualAmt) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtactno.Text & "','" & txtTransCode.Text & "','" & Format(ld_CDate - 1, "YYYY/MM/DD") & "'," & ln_AccrAmount & ")")
                                 End If
                                gc_dbcon.CommitTrans
                                Else
                                  Call SetErr("Calculation Error.", vbCritical)
                            End If
                     Else
                         Call SetErr("Calculation Formula Not Found.", vbCritical)
                     End If
            ld_Accrdate = ld_CDate
            ln_Counter = ln_Counter + 1
       Next
       ld_Accrdate = dtpValueDate
       ln_Step = 1
   Next
PR_Paracount.Close
End Sub
Private Sub CreatShdl()
Dim ln_InstlCntr As String ' Installment Counter
Dim ld_Accrdate As Date     ' Rental Accrual Date
Dim ld_CDate As Date     ' Current Date
Dim ln_Days As Double
Dim ln_LeaseAmt As Double
Dim Ln_Profit As Double
Dim ln_IRR As Double
Dim ln_Rental As Double
Dim ln_InsuRental As Double
Dim ln_cnt, ln_Step, ln_Counter As Integer
Dim ln_Rental2 As Double
Dim Ln_Profit2 As Double
Dim ln_LeaseAmt2 As Double
Dim ln_Cost2 As Double
Dim ln_inncnt As Integer
Dim Ln_AcrAmount As Integer
ld_Accrdate = DTPInstallDate
ln_Step = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
ld_CDate = DateAdd("M", ln_Step, ld_Accrdate)
ln_Days = Round(DateDiff("D", ld_Accrdate, ld_CDate) + 0.0001, 4)
ln_LeaseAmt = Val(0 & txtAmount) ' Lease Amount - Deposit Paid By Customer
ln_IRR = Round(Val(0 & txtRate) / 1200, 8) * ln_Step  ' IRR for the said period
'ln_InsuRental = Round(ln_LeaseAmt * Round((Val(0 & txtinsurance) / 1200) * ln_Step, 6), 0) ' Insurance Rental Calculation per period
ln_Rental = Val(0 & txtrental) ' Set Default Rental
lb_Edited = False              ' Set Rendom Rentals to False
Ln_ReCalc = 0                  ' Set Re-calculation Tag to false
ln_Cost2 = 0
ln_Counter = 1
ln_Rental2 = ln_Rental      ' Set For Accounting Schedule
ln_LeaseAmt2 = ln_LeaseAmt  ' Set For Accounting Schedule
          For ln_cnt = 1 To Val(0 & txtPeriod) Step ln_Step
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
                     ln_Rental = ReCalcRental(ln_LeaseAmt, Val(0 & txtPeriod) - ((ln_Counter - 1) * ln_Step), "R")
                    ' Accounting Schedule Rental Re-Calculation
                     ln_Rental2 = ReCalcRental(ln_LeaseAmt2, Val(0 & txtPeriod) - ((ln_Counter - 1) * ln_Step), "R")
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
               
                gc_dbcon.BeginTrans
                        gc_dbcon.Execute ("INSERT into ADV_Trans(Compcode,BranchCode,CustomerNo,AccountNo,TransCode,AccrualDate,InstallNo,LeaseRental,CostAmount,CostBalance,ProfitAmount,PaidDays,UserID,TransDate,TransTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtactno.Text & "','" & txtTransCode.Text & "','" & Format(ld_Accrdate, "YYYY/MM/DD") & "','" & ln_InstlCntr & "'," & ln_Rental & "," & (ln_Rental - Ln_Profit) & "," & ln_LeaseAmt & "," & Ln_Profit & "," & ln_Days & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )")
                        
                        'Ln_AcrAmount = Ln_Profit2 / ln_Step
                        'ld_CDate = ld_Accrdate
                        'For ln_inncnt = 1 To ln_Step
                         '   gc_dbcon.Execute ("INSERT into ADV_Accruals(Compcode,BranchCode,CustomerNo,AccountNo,TransCode,ValueDate,AccrualAmt) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtactno.Text & "','" & txtTransCode.Text & "','" & Format(ld_CDate - 1, "YYYY/MM/DD") & "'," & Ln_AcrAmount & ")")
                         '   ld_CDate = DateAdd("M", ln_inncnt, ld_CDate)
                        'Next
                gc_dbcon.CommitTrans
                ld_Accrdate = ld_CDate
                ld_CDate = DateAdd("M", ln_Step, ld_Accrdate)
                ln_Days = Round(DateDiff("D", ld_Accrdate, ld_CDate) + 0.0001, 4)
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

     'ln_Rental(1) = Val(0 & txtdpt_v) - Val(0 & txtrsd_v)  ' FV
     ln_Rental(1) = IIf(ln_Rental(1) > 0, 0, ln_Rental(1)) 'FV
     ln_Rental(2) = ln_Balance ' PV
     ln_Rental(3) = Round((Val(0 & txtRate) / 1200), 6)  'IRR + Tax Depr
     ln_Rental(4) = ln_RemPeriod 'Period
     ln_Rental(5) = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
     ln_Rental(6) = IIf(Trim(ls_PaidIn) = "R", 0, IIf(txtpaidin = "A", 1, 0)) '0/1 Advance/Arrears
     
     ReCalcRental = Module1.CalcRental(ln_Rental)
End Function

