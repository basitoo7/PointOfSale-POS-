VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCreditMemo 
   Appearance      =   0  'Flat
   Caption         =   "Credit Memo"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmceditmemo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   6780
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   5865
      Begin MSMask.MaskEdBox txtamtreceipt 
         Height          =   330
         Left            =   1350
         TabIndex        =   45
         Tag             =   "SKIP"
         Top             =   4275
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   5865
         MaxLength       =   10
         TabIndex        =   42
         Top             =   4965
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtltype 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4410
         MaxLength       =   50
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1560
         Width           =   1380
      End
      Begin VB.TextBox dtpofferdate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   4530
         MaxLength       =   50
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox txtofferno 
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   35
         Tag             =   "SKIP"
         Top             =   180
         Width           =   810
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   615
         Left            =   1350
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   33
         Tag             =   "SKIP"
         Top             =   3615
         Width           =   4440
      End
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   31
         Tag             =   "SKIP"
         Top             =   1209
         Width           =   4455
      End
      Begin VB.TextBox txtatype 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   866
         Width           =   4455
      End
      Begin VB.TextBox txtpaidin 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   23
         Tag             =   "SKIP"
         Top             =   2924
         Width           =   1815
      End
      Begin VB.TextBox txtaccrual 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   22
         Tag             =   "SKIP"
         Top             =   3267
         Width           =   1815
      End
      Begin VB.TextBox txtcustname 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1350
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   523
         Width           =   4455
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
         Left            =   2160
         Picture         =   "frmceditmemo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "SKIP"
         Top             =   180
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtdpt_v 
         Height          =   315
         Left            =   2220
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1890
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtrsd_v 
         Height          =   315
         Left            =   2220
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2235
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtfef_v 
         Height          =   315
         Left            =   2220
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2580
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtdocfee 
         Height          =   315
         Left            =   4410
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2910
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtmisc 
         Height          =   315
         Left            =   4410
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   3255
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox txtdpt_p 
         Height          =   315
         Left            =   1350
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1890
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         TabIndex        =   25
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2235
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtfef_p 
         Height          =   315
         Left            =   1350
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2580
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Left            =   4410
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1905
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
         Left            =   4410
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2580
         Width           =   1380
         _ExtentX        =   2434
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
         Left            =   4410
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   2235
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtleaseamount 
         Height          =   330
         Left            =   1350
         TabIndex        =   41
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1545
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Frame Frame1 
         Height          =   1605
         Left            =   45
         TabIndex        =   43
         Top             =   4530
         Width           =   4125
         Begin MSFlexGridLib.MSFlexGrid Grid1 
            Height          =   1380
            Left            =   60
            TabIndex        =   44
            Tag             =   "SKIP"
            Top             =   165
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   2434
            _Version        =   393216
            Cols            =   3
            BackColorBkg    =   -2147483633
            AllowBigSelection=   0   'False
            Enabled         =   0   'False
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   3
            Appearance      =   0
         End
      End
      Begin VB.Frame framsign 
         Enabled         =   0   'False
         Height          =   675
         Left            =   45
         TabIndex        =   47
         Top             =   6060
         Width           =   4125
         Begin VB.TextBox txtsign 
            Height          =   330
            IMEMode         =   3  'DISABLE
            Left            =   1200
            PasswordChar    =   "*"
            TabIndex        =   53
            Tag             =   "SKIP"
            Top             =   225
            Width           =   2835
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Signature :"
            Height          =   210
            Left            =   405
            TabIndex        =   48
            Top             =   270
            Width           =   780
         End
      End
      Begin VB.Frame framdsign 
         Height          =   2565
         Left            =   4185
         TabIndex        =   49
         Tag             =   "SKIP"
         Top             =   4170
         Width           =   1620
         Begin VB.CheckBox ChkTreasury 
            Caption         =   "Approved"
            Height          =   315
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   57
            Tag             =   "SKIP"
            Top             =   2175
            Width           =   1350
         End
         Begin VB.CheckBox chkce 
            Caption         =   "Approved"
            Height          =   315
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   56
            Tag             =   "SKIP"
            Top             =   1620
            Width           =   1350
         End
         Begin VB.CheckBox chkmia 
            Caption         =   "Approved"
            Height          =   315
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   55
            Tag             =   "SKIP"
            Top             =   1065
            Width           =   1350
         End
         Begin VB.CheckBox chkhc 
            Caption         =   "Approved"
            Height          =   315
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   54
            Tag             =   "SKIP"
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label lbltapproved 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Treasury Approval"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   45
            TabIndex        =   58
            Top             =   1950
            Width           =   1530
         End
         Begin VB.Label lblhc 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Head Of Credit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00358852&
            Height          =   210
            Left            =   210
            TabIndex        =   52
            Top             =   150
            Width           =   1200
         End
         Begin VB.Label lblmia 
            Alignment       =   2  'Center
            Caption         =   "Manager Internal Audit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   555
            Left            =   75
            TabIndex        =   51
            Top             =   660
            Width           =   1470
         End
         Begin VB.Label lblce 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Chief Executive"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   210
            TabIndex        =   50
            Top             =   1395
            Width           =   1260
         End
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Amount Receipt :"
         Height          =   210
         Left            =   75
         TabIndex        =   46
         Top             =   4305
         Width           =   1230
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lease Type :"
         Height          =   210
         Left            =   3420
         TabIndex        =   40
         Top             =   1635
         Width           =   945
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "IRR :"
         Height          =   225
         Left            =   3990
         TabIndex        =   37
         Top             =   2295
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Remarks :"
         Height          =   210
         Left            =   585
         TabIndex        =   34
         Top             =   3630
         Width           =   720
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   405
         TabIndex        =   32
         Top             =   1245
         Width           =   900
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Assets Type :"
         Height          =   210
         Left            =   285
         TabIndex        =   29
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Rental :"
         Height          =   210
         Left            =   3825
         TabIndex        =   21
         Top             =   2640
         Width           =   540
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Misc Charges :"
         Height          =   210
         Left            =   3285
         TabIndex        =   20
         Top             =   3300
         Width           =   1080
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Doc. Charges :"
         Height          =   210
         Left            =   3285
         TabIndex        =   19
         Top             =   2955
         Width           =   1080
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Period :"
         Height          =   210
         Left            =   3825
         TabIndex        =   18
         Top             =   1935
         Width           =   540
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Accrual As :"
         Height          =   210
         Left            =   390
         TabIndex        =   17
         Top             =   3300
         Width           =   915
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Paid In :"
         Height          =   210
         Left            =   750
         TabIndex        =   16
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "F.E.F % :"
         Height          =   210
         Left            =   660
         TabIndex        =   15
         Top             =   2625
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lease Amount :"
         Height          =   210
         Left            =   165
         TabIndex        =   14
         Top             =   1575
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Residual %age :"
         Height          =   210
         Left            =   135
         TabIndex        =   13
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Deposit %age :"
         Height          =   210
         Left            =   210
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Offer Date :"
         Height          =   210
         Index           =   1
         Left            =   3645
         TabIndex        =   11
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Offer # :"
         Height          =   210
         Index           =   0
         Left            =   690
         TabIndex        =   10
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   555
         Width           =   1140
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5865
      _ExtentX        =   10345
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
               Picture         =   "frmceditmemo.frx":047C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmceditmemo.frx":08D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmceditmemo.frx":0D24
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmceditmemo.frx":1178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmceditmemo.frx":15CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmceditmemo.frx":1A20
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmceditmemo.frx":2174
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCreditMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_Blnkloffer As Boolean
Dim PI_CurRow    As Integer
Dim PI_SrNo     As Integer
Dim PS_RowClicked As String

Dim ln_TaxAmt As Double
Dim ln_LastValu As Integer
Dim ln_PrvUpto As Integer
Dim Ln_ReCalc As Integer
Dim lb_Edited As Boolean
Dim ls_LeaseNo As String
Dim PStatus As Boolean

Public Mode As String
Public PO_CODE As Object
Public PO_DESC As Object

Dim PR_AssetType As New Recordset
Dim PR_FCMIDs As New Recordset
Dim PR_lmoffer As New Recordset
Dim PR_lmoffer2 As New Recordset
Dim PR_Customer As New Recordset
Dim Pr_Pass As New Recordset

Private Sub chkhc_Click()
       If chkhc.Value = 1 Then
        lblhc.ForeColor = RGB(82, 136, 53)
        chkhc.Caption = "Approved"
        framsign.Enabled = chkRights("SIGNATURED")
        If framsign.Enabled Then txtsign.SetFocus
       Else
        framsign.Enabled = False
        lblhc.ForeColor = RGB(255, 0, 0)
        chkhc.Caption = "Disapproved"
       End If
End Sub
Private Sub chkmia_Click()
       If chkmia.Value = 1 Then
        lblmia.ForeColor = RGB(82, 136, 53)
        chkmia.Caption = "Approved"
        framsign.Enabled = chkRights("SIGNATURED")
        If framsign.Enabled Then txtsign.SetFocus
       Else
       chkmia.Caption = "Disapproved"
       framsign.Enabled = False
       lblmia.ForeColor = RGB(255, 0, 0)
       End If
End Sub
Private Sub chkce_Click()
       If chkce.Value = 1 Then
        lblce.ForeColor = RGB(82, 136, 53)
        chkce.Caption = "Approved"
        framsign.Enabled = True
        framsign.Enabled = chkRights("SIGNATURED")
        If framsign.Enabled Then txtsign.SetFocus
       Else
       framsign.Enabled = False
       lblce.ForeColor = RGB(255, 0, 0)
       chkce.Caption = "Disapproved"
       End If
End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtofferno
    Set PO_DESC = Text1
    GoTop PR_lmoffer
    MyLookup.Caption = "Offer Nos"
    MyLookup.FillGrid PR_lmoffer, "OfferNo", "Customerno", 3
    MyLookup.Show 1
    If Len(txtofferno) > 0 Then txtofferno_KeyDown vbKeyReturn, vbKeyShift
    PR_lmoffer.Filter = adFilterNone
End Sub
Private Sub dtpofferdate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
    txtatype.SetFocus
 ElseIf KeyCode = vbKeyPageUp Then
    txtofferno.SetFocus
 End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_lmoffer, Me, txtCustNO, txtCustNO, "X", "CompCount", 3, " CustomerNo", "CustomerName", 1, False, Toolbar1)
End Sub
Private Sub Form_Load()
  Dim ln_cnt As Integer
  
  SetToolBar(1) = chkRights("LMSCREDIT1")
  SetToolBar(2) = chkRights("LMSCREDIT2")
  SetToolBar(3) = chkRights("LMSCREDIT3")
  SetToolBar(4) = chkRights("LMSCREDIT4")
  framdsign.Enabled = chkRights("DSIGNMAIN1")
  chkhc.Enabled = chkRights("HCAPPROVED")
  chkmia.Enabled = chkRights("MIAAPPROVE")
  chkce.Enabled = chkRights("CEAPPROVED")
  ChkTreasury.Enabled = chkRights("TEAPPROVED")
  txtamtreceipt.Enabled = chkRights("AMTRECEIPT")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
 
Pr_Pass.Open "Select SyUsers.* from syUsers order by 1", gc_dbcon, adOpenStatic, adLockReadOnly
PR_AssetType.Open "Select * From LM_AssetTypes Order By AssetCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
PR_FCMIDs.Open "Select FCM_Ids.*,ltrim(rtrim(FCM_Ids.Idcode))+ltrim(rtrim(FCM_Ids.Recid)) as Findfld From FCM_Ids where recid in ('ADP','LMA') Order By IdCode", gc_dbcon, adOpenDynamic, adLockOptimistic, adCmdText
PR_Customer.Open "Select Customer.* from Customer left Outer Join Facilities On Customer.CustomerNo = Facilities.CustomerNo Where Customer.Compcode+Customer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' And Facilities.FacilityNo = '01' Order By Customer.CustomerNo", gc_dbcon, adOpenStatic, adLockOptimistic, adCmdText
PR_lmoffer.Open "Select lm_offer.*,lm_offer.customerno+lm_offer.offerno as findfld from lm_offer Where lm_offer.Compcode+lm_offer.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' Order By findfld", gc_dbcon, adOpenDynamic, adLockOptimistic, adCmdText
PR_lmoffer2.Open "Select lm_offer2.* from lm_offer2 Where lm_offer2.Compcode+lm_offer2.BranchCode = '" & Gs_compcode + Gs_BranchCode & "' Order By lm_offer2.offerno", gc_dbcon, adOpenDynamic, adLockOptimistic, adCmdText
PB_Blnkloffer = IIf(PR_lmoffer.EOF, True, False)
InitializeGrid
End Sub
Private Sub Form_Unload(Cancel As Integer)
    PR_Customer.Close
    PR_AssetType.Close
    PR_FCMIDs.Close
    PR_lmoffer.Close
    PR_lmoffer2.Close
    Pr_Pass.Close
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
  If KeyCode = vbKeyReturn And txtaccrual <> "" Then
     txtaccrual = UCase(txtaccrual)
     PR_FCMIDs.Filter = " RecID = 'LMA'"
     If Not MySeek(txtaccrual, "Idcode", PR_FCMIDs) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtaccrual.SetFocus
     Else
       txtperiod.SetFocus
     End If
     PR_FCMIDs.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtpaidin.SetFocus
  End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Range(Button.Index, 2, 3) Or Button.Index = 7 Then
       InitializeGrid
    End If
    If PB_Blnkloffer And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found. ", vbCritical)
       Mode = ""
       Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_lmoffer, Me, txtofferno, txtofferno, "X", "CompCount", 3, " CustomerNo", "CustomerName", 1, False, Toolbar1)
    End If

End Sub
Public Sub SaveValues()
Dim SyPassword   As String
Dim cntsql As New ADODB.Command
Dim ln_cnt As Integer
PB_Blnkloffer = False
Dim nextno As Boolean
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText

gc_dbcon.BeginTrans
     Select Case Mode
          Case "A"
          If MySeek(txtofferno, "Offerno", PR_lmoffer) Then
             If txtamtreceipt.Enabled = True Then
                    PR_lmoffer("AmtReceipt") = Val(0 & txtamtreceipt)
             ElseIf framdsign.Enabled Then
                     If PStatus Then
                        If chkhc.Value = 1 And chkhc.Enabled Then
                            PR_lmoffer("HCApproved") = 1
                        End If
                        If chkmia.Value = 1 And chkmia.Enabled Then
                            PR_lmoffer("MiaApproved") = 1
                        End If
                        If chkce.Value = 1 And chkce.Enabled Then
                            PR_lmoffer("ceApproved") = 1
                        End If
                        If ChkTreasury.Value = 1 And ChkTreasury.Enabled Then
                            PR_lmoffer("TreaApproved") = 1
                        End If
                    End If
             End If
          PR_lmoffer.Update
          End If
          Case "E"
              'cntsql.CommandText = "UPDATE LM_offer SET offerDate = '" & Format(dtpofferdate, "YYYY/MM/DD") & "',AssetType ='" & txtatype & "',Assetdescrip ='" & txtdesc & "',leaseamount =" & Val(0 & txtleaseamount) & ",deposit_v =" & Val(0 & txtdpt_v) & ",residual_v =" & Val(0 & txtrsd_v) & ",fef_v =" & Val(0 & txtfef_v) & ","
              'cntsql.CommandText = cntsql.CommandText + " paidas='" & txtpaidin & "', PaymentMode='" & txtaccrual & "',leaseperiod=" & Val(0 & txtperiod) & ", leaserental = " & Val(0 & txtrental) & ",doccharges = " & Val(0 & txtdocfee) & ",misccharges = '" & txtmisc & "',offerremarks= '" & txtremarks & "', "
              'cntsql.CommandText = cntsql.CommandText + " leaseirr = " & Val(0 & txtirr) & " ,leasetype = " & Val(0 & chkleasetype) & " WHERE  compcode+branchcode+customerno+offerno = '" & Gs_compcode + Gs_BranchCode + Trim(txtCustNO) + Trim(txtofferno) & "'"
              'cntsql.Execute
          Case "D"
             ' cntsql.CommandText = "DELETE FROM lm_offer WHERE  compcode+branchcode+customerno+offerno = '" & Gs_compcode + Gs_BranchCode + Trim(txtCustNO) + Trim(txtofferno) & "'"
             ' cntsql.Execute
             ' cntsql.CommandText = "delete from lm_offer2 WHERE  compcode+branchcode+offerno = '" & Gs_compcode + Gs_BranchCode + Trim(txtofferno) & "'"
             ' cntsql.Execute
     End Select
gc_dbcon.CommitTrans
PR_lmoffer.Requery
PR_lmoffer2.Requery
     InitializeGrid
     PI_SrNo = 0
     PS_RowClicked = ""
     ls_LeaseNo = ""
End Sub
Private Sub SetVal()
     dtpofferdate = PR_lmoffer("offerDate")
     txtcustname = IIf(MySeek(Trim(PR_lmoffer("Customerno")), "CustomerNo", PR_Customer), PR_Customer("CustomerName"), "")
     txtltype = IIf(Val(0 & PR_lmoffer("leasetype")) = 1, "Lease And Sale Back", "Direct")
     txtirr = Val(0 & PR_lmoffer("Leaseirr"))
     txtatype = IIf(MySeek(Trim(PR_lmoffer("assettype")), "AssetType", PR_AssetType), PR_AssetType("AssetName"), "")
     txtdesc = PR_lmoffer("assetdescrip")
     txtleaseamount = PR_lmoffer("leaseamount")
     txtdpt_v = PR_lmoffer("Deposit_v")
     txtrsd_v = PR_lmoffer("Residual_v")
     txtfef_v = PR_lmoffer("FEF_v")
     txtdpt_p = Round((PR_lmoffer("Deposit_V") / PR_lmoffer("leaseAmount")) * 100, 3)
     txtrsd_p = Round((PR_lmoffer("Residual_V") / PR_lmoffer("leaseAmount")) * 100, 3)
     txtfef_p = Round((PR_lmoffer("Fef_V") / PR_lmoffer("leaseAmount")) * 100, 3)
     txtpaidin = IIf(MySeek(Trim(PR_lmoffer("PaidAS") + "ADP"), "Findfld", PR_FCMIDs), PR_FCMIDs("Iddescrip"), "")
     txtaccrual = IIf(MySeek(Trim(PR_lmoffer("Paymentmode") + "LMA"), "Findfld", PR_FCMIDs), PR_FCMIDs("Iddescrip"), "")
     txtperiod = PR_lmoffer("LeasePeriod")
     txtrental = PR_lmoffer("LeaseRental")
     txtdocfee = PR_lmoffer("Doccharges")
     txtmisc = PR_lmoffer("Misccharges")
     txtremarks = PR_lmoffer("offerRemarks")
     txtamtreceipt = Val(0 & PR_lmoffer("AmtReceipt"))
     chkhc.Caption = IIf(Val(0 & PR_lmoffer("hcapproved")) = 1, "Approved", "Disapproved")
     chkmia.Caption = IIf(Val(0 & PR_lmoffer("miaapproved")) = 1, "Approved", "Disapproved")
     chkce.Caption = IIf(Val(0 & PR_lmoffer("ceapproved")) = 1, "Approved", "Disapproved")
     ChkTreasury.Caption = IIf(Val(0 & PR_lmoffer("Treaapproved")) = 1, "Approved", "Disapproved")
     chkhc.Value = IIf(Val(0 & PR_lmoffer("hcapproved")) = 1, 1, 0)
     chkmia.Value = IIf(Val(0 & PR_lmoffer("miaapproved")) = 1, 1, 0)
     chkce.Value = IIf(Val(0 & PR_lmoffer("ceapproved")) = 1, 1, 0)
     ChkTreasury.Value = IIf(Val(0 & PR_lmoffer("Treaapproved")) = 1, 1, 0)
     lblhc.ForeColor = IIf(Val(0 & PR_lmoffer("hcapproved")) = 1, RGB(82, 136, 53), RGB(255, 0, 0))
     lblmia.ForeColor = IIf(Val(0 & PR_lmoffer("miaapproved")) = 1, RGB(82, 136, 53), RGB(255, 0, 0))
     lblce.ForeColor = IIf(Val(0 & PR_lmoffer("ceapproved")) = 1, RGB(82, 136, 53), RGB(255, 0, 0))
     lbltapproved.ForeColor = IIf(Val(0 & PR_lmoffer("treaapproved")) = 1, RGB(82, 136, 53), RGB(255, 0, 0))
     LoadGRNTrans
End Sub
Private Sub passstat()
If framdsign.Enabled Then
    If Trim(txtsign) <> "" And (chkhc.Enabled = True Or chkmia.Enabled = True Or chkce.Enabled = True And ChkTreasury.Enabled = True) Then
        If MySeek(UCase(Gc_UserId), "userid", Pr_Pass) Then
            SyPassword = UCase(LTrim(RTrim(DeCode(Pr_Pass.Fields("Password").Value, 50))))
               If UCase(txtsign) = SyPassword Then
                    PStatus = True
                Else
                    PStatus = False
                 End If
        Else
         PStatus = False
        End If
   Else
     PStatus = False
   End If
Else
    If Trim(txtamtreceipt) <> "" Then
     PStatus = True
    Else
     PStatus = False
    End If
End If
End Sub
Public Function ChkInputs() As Boolean
    passstat
    If Len(txtofferno) = txtofferno.MaxLength And PStatus Then
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
        .ColWidth(3) = 1700
        .Redraw = True
    End With
End Sub
Private Sub LoadGRNTrans()
Dim lb_found As Boolean
InitializeGrid

    lb_found = MySeek(txtofferno, "Offerno", PR_lmoffer2)

    If lb_found Then
        With Grid1
            Do While txtofferno = Trim(PR_lmoffer2("offerno"))
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = PR_lmoffer2("FromPeriod")
                .TextMatrix(.Row, 2) = PR_lmoffer2("UpToPeriod")
                .TextMatrix(.Row, 3) = PR_lmoffer2("RentalAmount")
                .Rows = .Rows + 1
                PR_lmoffer2.MoveNext
                If PR_lmoffer2.EOF Then Exit Do
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

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtdesc <> "" Then
   txtleaseamount.SetFocus
ElseIf KeyCode = vbKeyPageUp Then
    txtatype.SetFocus
End If
End Sub

Private Sub txtdocfee_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
     txtmisc.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtdocfee.SetFocus
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
     txtdocfee.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
      txtperiod.SetFocus
  End If
End Sub

Private Sub txtleaseamount_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And txtleaseamount <> "" Then
     'If Val(0 & txtleaseamount) > Val(0 & Para_Rs("LMS_VehMax")) Then
         txtdpt_p.SetFocus
     'End If
  ElseIf KeyCode = vbKeyPageUp Then
        txtdesc.SetFocus
  End If
End Sub
Private Sub txtmisc_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) Then
     txtremarks.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtdocfee.SetFocus
  End If
End Sub

Private Sub txtofferno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtofferno <> "" Then
     txtofferno = DoPad(txtofferno, txtofferno.MaxLength)
     If Not MySeek(txtofferno, "Offerno", PR_lmoffer) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
        txtofferno.SetFocus
     Else
       SetVal
       If txtamtreceipt.Enabled Then txtamtreceipt.SetFocus
     End If
     PR_lmoffer.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtCustNo.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     Command1_Click
  End If
End Sub

Private Sub txtpaidin_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And txtpaidin <> "" Then
     txtpaidin = UCase(txtpaidin)
     PR_FCMIDs.Filter = "Recid = 'ADP'"
     If Not MySeek(txtpaidin, "idcode", PR_FCMIDs) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtpaidin.SetFocus
     Else
       txtaccrual.SetFocus
     End If
     PR_FCMIDs.Filter = adFilterNone
  ElseIf KeyCode = vbKeyPageUp Then
     txtfef_v.SetFocus
  ElseIf KeyCode = vbKeyF12 Then
     'Command2_Click
  End If
End Sub
Private Sub txtPeriod_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtperiod <> "" Then
     txtirr.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtaccrual.SetFocus
  End If
End Sub
Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtfrom.SetFocus
ElseIf KeyCode = vbKeyPageUp Then
    txtmisc.SetFocus
End If
End Sub
Private Sub txtrental_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And txtrental <> "" Then
    txtdocfee.SetFocus
ElseIf KeyCode = vbKeyPageUp Then
    txtperiod.SetFocus
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
Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And Range(Val(0 & txtto), Val(0 & txtfrom), Val(0 & txtperiod)) Then
     ln_PrvUpto = Val(0 & txtto)
     txteditrental.SetFocus
  ElseIf KeyCode = vbKeyPageUp Then
     txtfrom.SetFocus
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
Dim ln_cnt, ln_Step, ln_Counter As Integer
Dim ln_Rental2 As Double
Dim Ln_Profit2 As Double
Dim ln_LeaseAmt2 As Double
Dim ln_Cost2 As Double

ln_Step = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
ln_LeaseAmt = Val(0 & txtleaseamount) - Val(0 & txtdpt_v) ' Lease Amount - Deposit Paid By Customer
ln_IRR = Round(Val(0 & txtirr) / 1200, 6) * ln_Step ' IRR for the said period
ln_InsuRental = Round(ln_LeaseAmt * Round((Val(0 & txtinsurance) / 1200) * ln_Step, 6), 0) ' Insurance Rental Calculation per period
ld_Accrdate = DTPschdl.Value
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
             ln_Rental = ReCalcRental(ln_LeaseAmt, Val(0 & txtperiod) - ((ln_Counter - 1) * ln_Step))
            ' Accounting Schedule Rental Re-Calculation
             ln_Rental2 = ReCalcRental(ln_LeaseAmt2, Val(0 & txtperiod) - ((ln_Counter - 1) * ln_Step))
            ' Set Recovery Profit ON/OFF
             Ln_Profit = IIf(txtpaidin = "A", 0, Ln_Profit)
            ' Set Accounting Profit ON/OFF
             Ln_Profit2 = IIf(txtpaidin = "A", 0, Ln_Profit2)
             Ln_ReCalc = 0
          End If
       End If
      
      ' Balance Principal For Accounting Schedule
       Ln_Profit2 = IIf((ln_Rental2 - Ln_Profit2) < 0, IIf(ln_Rental2 = 0, 0, ln_Rental2), Ln_Profit2)
       ln_LeaseAmt2 = ln_LeaseAmt2 - (ln_Rental2 - Ln_Profit2)

      ' Balance Principal For Recovery Schedule
       ln_LeaseAmt = ln_LeaseAmt - (ln_Rental - Ln_Profit)
       
       gc_dbcon.Execute ("INSERT into LM_Schedule(Compcode,BranchCode,CustomerNo,LeaseNo,AccrualDate,InstallNo,InsurRental,LeaseRental,CostAmount,CostBalance,ProfitAmount,LeaseRental2,CostAmount2,CostBalance2,ProfitAmount2,RentalStatus,UserID,TransDate,TransTime) Values ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtCustNo.Text & "','" & txtleaseno.Text & "','" & Format(ld_Accrdate, "YYYY/MM/DD") & "','" & ln_InstlCntr & "'," & ln_InsuRental & "," & ln_Rental & "," & (ln_Rental - Ln_Profit) & "," & ln_LeaseAmt & "," & Ln_Profit & "," & ln_Rental2 & "," & (ln_Rental2 - Ln_Profit2) & "," & ln_LeaseAmt2 & "," & Ln_Profit2 & ",'O','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "' )")
       ld_Accrdate = DateAdd("M", ln_Step, ld_Accrdate)
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
Private Function ReCalcRental(ln_Balance As Double, ln_RemPeriod As Integer) As Double
' RunTime Re-Calculations of Rental
Dim ln_Rental(1 To 8) As Double

     ln_Rental(1) = Val(0 & txtdpt_v) - Val(0 & txtrsd_v)  ' FV
     ln_Rental(1) = IIf(ln_Rental(1) > 0, 0, ln_Rental(1)) 'FV
     ln_Rental(2) = ln_Balance ' PV
     ln_Rental(3) = Round((Val(0 & txtirr) / 1200), 6) 'IRR + Tax Depr
     ln_Rental(4) = ln_RemPeriod 'Period
     ln_Rental(5) = IIf(txtaccrual = "Q", 3, IIf(txtaccrual = "S", 6, IIf(txtaccrual = "A", 12, 1))) 'Accrual As 3/6/12
     ln_Rental(6) = IIf(txtpaidin = "A", 1, 0) '0/1 Advance/Arrears
     
     ReCalcRental = Module1.CalcRental(ln_Rental)
End Function
Public Sub FrmRefresh()
PR_AssetType.Requery
PR_Customer.Requery
PR_lmoffer.Requery
PR_lmoffer2.Requery
PR_FCMIDs.Requery
End Sub
Private Sub txtsign_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtsign) <> "" Then
        If MySeek(UCase(Gc_UserId), "userid", Pr_Pass) Then
            SyPassword = UCase(LTrim(RTrim(DeCode(Pr_Pass.Fields("Password").Value, 50))))
               If UCase(txtsign) <> SyPassword Then
                    Call SetErr("Invalid Password", vbCritical)
               Else
                    Call SetErr("Password Successfully Verified", vbInformation)
               End If
        End If
End If
End Sub
