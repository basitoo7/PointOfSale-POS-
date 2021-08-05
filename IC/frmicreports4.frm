VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmicreport4 
   Caption         =   "Stock Ledger"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmicreports4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkSum 
      Caption         =   "Summery On Last Net Rate"
      Height          =   495
      Left            =   120
      TabIndex        =   50
      Top             =   4920
      Width           =   2295
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   5505
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Description"
            TextSave        =   "Description"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   105833
            MinWidth        =   105833
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   30
      TabIndex        =   8
      Top             =   -60
      Width           =   6450
      Begin VB.CheckBox ChkNetRate 
         Caption         =   "On Last Net Rate"
         Height          =   495
         Left            =   2520
         TabIndex        =   49
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   4935
         Left            =   90
         TabIndex        =   10
         Top             =   120
         Width           =   6315
         Begin VB.CommandButton Command4 
            Height          =   315
            Left            =   2655
            Picture         =   "frmicreports4.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1350
            Width           =   315
         End
         Begin VB.TextBox txtSupplierdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3000
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1350
            Width           =   3210
         End
         Begin VB.TextBox txtSuppliercode 
            Height          =   315
            Left            =   1680
            TabIndex        =   39
            Top             =   1350
            Width           =   945
         End
         Begin VB.CommandButton Command3 
            Height          =   315
            Left            =   2265
            Picture         =   "frmicreports4.frx":047C
            Style           =   1  'Graphical
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   945
            Width           =   315
         End
         Begin VB.TextBox txtsubcatdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2625
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   945
            Width           =   3585
         End
         Begin VB.TextBox txtsubcatcode 
            Height          =   315
            Left            =   1680
            TabIndex        =   35
            Top             =   945
            Width           =   555
         End
         Begin VB.CommandButton Command2 
            Height          =   315
            Left            =   2265
            Picture         =   "frmicreports4.frx":05EE
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   570
            Width           =   315
         End
         Begin VB.TextBox txtcatdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2625
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   570
            Width           =   3585
         End
         Begin VB.TextBox txtcatcode 
            Height          =   315
            Left            =   1680
            TabIndex        =   31
            Top             =   570
            Width           =   555
         End
         Begin VB.Frame Frame2 
            Height          =   2790
            Left            =   3675
            TabIndex        =   27
            Top             =   2040
            Width           =   2595
            Begin VB.CheckBox Chkprintstocktax 
               Caption         =   "Print Stock Tax Item Only"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   48
               Top             =   2415
               Width           =   2370
            End
            Begin VB.CheckBox chkwamtsum 
               Caption         =   "Print With Amount Summary"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   47
               Top             =   2115
               Width           =   2370
            End
            Begin VB.CheckBox chkwithamount 
               Caption         =   "Print With Amount Item Wise"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   46
               Top             =   1830
               Width           =   2370
            End
            Begin VB.CheckBox chkpwcost 
               Caption         =   "Print With Pur-Sale-Avg Amt"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   45
               Top             =   1560
               Width           =   2475
            End
            Begin VB.CheckBox chkpwcostqty 
               Caption         =   "Print With Cost and Qty"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   44
               Top             =   1275
               Width           =   2370
            End
            Begin VB.CheckBox chkbrl 
               Caption         =   "Base on Reorder Level"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   43
               Top             =   1005
               Width           =   2370
            End
            Begin VB.CheckBox chkItemHistory 
               Caption         =   "Item History"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   30
               Top             =   735
               Width           =   1695
            End
            Begin VB.CheckBox ChkSummary 
               Caption         =   "Summary"
               BeginProperty DataFormat 
                  Type            =   4
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   8
               EndProperty
               Height          =   210
               Left            =   75
               TabIndex        =   29
               Top             =   225
               Value           =   1  'Checked
               Width           =   2160
            End
            Begin VB.CheckBox ChkTransMovementOnly 
               Caption         =   "Transaction Movement Only"
               Height          =   210
               Left            =   75
               TabIndex        =   28
               Top             =   480
               Width           =   2370
            End
         End
         Begin VB.ComboBox txtgrouping 
            Height          =   330
            ItemData        =   "frmicreports4.frx":0760
            Left            =   1680
            List            =   "frmicreports4.frx":0770
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   4425
            Width           =   1860
         End
         Begin VB.ComboBox txtstockbase 
            Height          =   330
            ItemData        =   "frmicreports4.frx":07B5
            Left            =   1680
            List            =   "frmicreports4.frx":07C5
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   3900
            Width           =   1860
         End
         Begin VB.TextBox txtdeptcode 
            Height          =   315
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   0
            Top             =   195
            Width           =   555
         End
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   2625
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   195
            Width           =   3585
         End
         Begin VB.CommandButton Command1 
            Height          =   315
            Left            =   2265
            Picture         =   "frmicreports4.frx":07FA
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   195
            Width           =   315
         End
         Begin VB.TextBox txtitemcodeDirect 
            Height          =   315
            Left            =   3915
            TabIndex        =   18
            Top             =   1350
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.ComboBox txtStoreType 
            Height          =   330
            ItemData        =   "frmicreports4.frx":096C
            Left            =   1680
            List            =   "frmicreports4.frx":0979
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   3375
            Width           =   1860
         End
         Begin VB.TextBox txtitemcode 
            Height          =   315
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   1
            Top             =   1740
            Width           =   960
         End
         Begin VB.TextBox txtitemdesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            ForeColor       =   &H80000002&
            Height          =   315
            Left            =   3000
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1725
            Width           =   3210
         End
         Begin VB.CommandButton Command5 
            Height          =   315
            Left            =   2655
            Picture         =   "frmicreports4.frx":0994
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1725
            Width           =   315
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1680
            TabIndex        =   2
            Top             =   2325
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   109379585
            CurrentDate     =   37293
         End
         Begin MSComCtl2.DTPicker DTPTo 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   2865
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   109379585
            CurrentDate     =   37293
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Supplier Code :"
            Height          =   210
            Left            =   435
            TabIndex        =   42
            Top             =   1380
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Sub Cat Code :"
            Height          =   210
            Left            =   450
            TabIndex        =   38
            Top             =   975
            Width           =   1080
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Category Code :"
            Height          =   210
            Left            =   360
            TabIndex        =   34
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Grouping  :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   780
            TabIndex        =   26
            Top             =   4455
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Stock Base :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   660
            TabIndex        =   24
            Top             =   3930
            Width           =   915
         End
         Begin VB.Label txtselectiveaccount 
            Height          =   300
            Left            =   4590
            TabIndex        =   22
            Top             =   1590
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department Code :"
            Height          =   210
            Left            =   195
            TabIndex        =   21
            Top             =   225
            Width           =   1335
         End
         Begin VB.Label txtselectiveitem 
            Height          =   300
            Left            =   3570
            TabIndex        =   17
            Top             =   1305
            Visible         =   0   'False
            Width           =   2520
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "From Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   735
            TabIndex        =   16
            Top             =   2355
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "To Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   915
            TabIndex        =   15
            Top             =   2895
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Store Type :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   675
            TabIndex        =   14
            Top             =   3420
            Width           =   885
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Item Code :"
            Height          =   210
            Left            =   765
            TabIndex        =   13
            Top             =   1755
            Width           =   795
         End
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   5160
         Top             =   1185
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
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   5280
         TabIndex        =   6
         Top             =   5085
         Width           =   1035
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4230
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Top             =   5085
         Width           =   1035
      End
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   435
         MaxLength       =   50
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmicreport4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pb_BlnkVchr As Boolean
Public PO_CODE As Object
Public PO_DESC As Object
Dim pr_dumy As New Recordset
Public codeid As String
Public Reporttype As String
Dim ls_sql As String
Private Sub Chkprintstocktax_Click()
If Chkprintstocktax.Value = 1 Then
ChkSummary.Value = 0
ChkTransMovementOnly.Value = 0
chkItemHistory.Value = 0
chkpwcostqty.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0
Label8.Caption = "As on Date:"
Else
Label8.Caption = "Date From :"
End If
End Sub
Private Sub chkpwcost_Click()
If chkpwcost.Value = 1 Then
ChkSummary.Value = 0
ChkTransMovementOnly.Value = 0
chkItemHistory.Value = 0
chkpwcostqty.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0
Chkprintstocktax.Value = 0
End If
End Sub
Private Sub chkpwcostqty_Click()
If chkpwcostqty.Value = 1 Then
ChkSummary.Value = 0
ChkTransMovementOnly.Value = 0
chkItemHistory.Value = 0
chkpwcost.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0
chkbrl.Value = 0
Chkprintstocktax.Value = 0

End If

End Sub

Private Sub ChkSummary_Click()
If ChkSummary.Value = 1 Then
 chkbrl.Value = 0
ChkTransMovementOnly.Value = 0
chkItemHistory.Value = 0

chkpwcostqty.Value = 0
chkpwcost.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0

Chkprintstocktax.Value = 0

End If
End Sub
Private Sub chkbrl_Click()
If chkbrl.Value = 1 Then
ChkSummary.Value = 0
ChkTransMovementOnly.Value = 0
chkItemHistory.Value = 0

chkpwcostqty.Value = 0
chkpwcost.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0

Chkprintstocktax.Value = 0

End If
End Sub
Private Sub ChkTransMovementOnly_Click()
If ChkTransMovementOnly.Value = 1 Then
ChkSummary.Value = 0
chkbrl.Value = 0
chkItemHistory.Value = 0

chkpwcostqty.Value = 0
chkpwcost.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0
Chkprintstocktax.Value = 0


End If
End Sub
Private Sub chkItemHistory_Click()
If chkItemHistory.Value = 1 Then
ChkSummary.Value = 0
chkbrl.Value = 0
ChkTransMovementOnly.Value = 0

chkpwcostqty.Value = 0
chkpwcost.Value = 0
chkwithamount.Value = 0
chkwamtsum.Value = 0
Chkprintstocktax.Value = 0


End If
End Sub

Private Sub chkwithamount_Click()
If chkwithamount.Value = 1 Then
    ChkSummary.Value = 0
    chkbrl.Value = 0
    ChkTransMovementOnly.Value = 0
    chkpwcostqty.Value = 0
    chkpwcost.Value = 0
    chkwamtsum.Value = 0
    Chkprintstocktax.Value = 0

End If
End Sub
Private Sub chkwamtsum_Click()
If chkwamtsum.Value = 1 Then
    ChkSummary.Value = 0
    chkbrl.Value = 0
    ChkTransMovementOnly.Value = 0
    chkpwcostqty.Value = 0
    chkpwcost.Value = 0
    chkwithamount.Value = 0
    Chkprintstocktax.Value = 0

   
End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
On Error GoTo localerr



If chkItemHistory.Value = 1 Then
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
With crrpt
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
        
        .ReportFileName = App.Path & Gs_ICRepoPath & "\StockLedgerTransDetail.RPT"
        
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & Me.Caption & "'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
         Call ChkTempTables("Tmp_Stockledger", True)
         ls_sql = "SELECT compcode, TransDate, Transcode, Remarks, ItemCode, ReceiveQty, IssueQty, 0 as OpQty,Rate,InAmount,OutAmount,0 as OpAmount into Tmp_Stockledger From StockLedgerTransDetail"
         ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'"
         ls_sql = ls_sql & " and transdate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
         ls_sql = ls_sql & " and transdate <= '" & Format(DTPTo, "YYYY/MM/DD") & "'"
     
        If txtStoreType.ListIndex <> 2 Then
         ls_sql = ls_sql & " and siteid = " & txtStoreType.ListIndex + 1 & " "
        End If
        
         If txtitemcode <> "" Then
            ls_sql = ls_sql & " and itemcode = '" & txtitemcode & "'"
          
         End If
        
        ls_sql = ls_sql & " Union all"
            
        ls_sql = ls_sql & " SELECT compcode,'" & Format(dtpfrom, "YYYY/MM/DD") & "' as TransDate,'0000000000' as TransCode,'Opening Balance' as Remarks, ItemCode, 0 as ReceiveQty,0 as IssueQty,SUM(ReceiveQty) - SUM(IssueQty) as OpQty,sum(InAmount)-Sum(OutAmount)/case when SUM(ReceiveQty) - SUM(IssueQty)>0 then SUM(ReceiveQty) - SUM(IssueQty)else 1 end as Rate,0 as InAmount,0 as OutAmount,sum(InAmount)-Sum(OutAmount) as OpAmount  From  StockLedgerTransDetail"
        ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "' "
        ls_sql = ls_sql & " and transdate < '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
        
        If txtStoreType.ListIndex <> 2 Then
        ls_sql = ls_sql & " and siteid = " & txtStoreType.ListIndex + 1 & " "
        End If
          
        If txtitemcode <> "" Then
                 ls_sql = ls_sql & " and itemcode = '" & txtitemcode & "'"
        End If
         
        ls_sql = ls_sql & " group by compcode,itemcode"
        
        gc_dbcon.Execute ls_sql
        .RetrieveSQLQuery
         .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
End With
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub
End If
If Me.Caption = "Stock Ledger Detail" Then
    Call ChkTempTables("Tmp_Stockledger", True)
    ls_sql = " SELECT compcode,Branchcode, TransDate, printtranscode, Remarks, ItemCode, ReceiveQty, IssueQty,0 as OpeningQty into Tmp_Stockledger"
    ls_sql = ls_sql & " FROM StockLedger where Compcode = '" & Gs_compcode & "' and transdate >='" & Format(dtpfrom, "YYYY/MM/DD") & "'  and transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "' "
    If txtdeptcode <> "" Then
    ls_sql = ls_sql & " and itemcode = '" & txtdeptcode & "'"
    End If
    gc_dbcon.Execute ls_sql
    'Opening balance
    ls_sql = " insert into Tmp_Stockledger (compcode,Branchcode, TransDate, printtranscode, Remarks, ItemCode, ReceiveQty, IssueQty,OpeningQty)"
    ls_sql = ls_sql & " SELECT compcode,Branchcode, '" & Format(dtpfrom, "YYYY/MM/DD") & "' as TransDate, 'OP' as printtranscode, 'Opening Balance' as Remarks, ItemCode, 0 as ReceiveQty,0 as  IssueQty ,SUM(ReceiveQty) - SUM(IssueQty) as OpeningQty "
    ls_sql = ls_sql & " FROM StockLedger where Compcode = '" & Gs_compcode & "' and transdate <'" & Format(dtpfrom, "YYYY/MM/DD") & "' "
    If txtdeptcode <> "" Then
    ls_sql = ls_sql & " and itemcode = '" & txtdeptcode & "'"
    End If
    ls_sql = ls_sql & " group by compcode, itemcode"
    gc_dbcon.Execute ls_sql
    With crrpt
        
        .ReportFileName = App.Path & Gs_ICRepoPath & "\Stockledgerdetail.RPT"
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = 'Stock Ledger'"
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
Else
MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
    With crrpt
        .Formulas(0) = ""
        .Formulas(1) = ""
        .Formulas(2) = ""
        .Formulas(3) = ""
       
        If ChkSum.Value = 1 Then
          .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodicAtNetRateSumPurDept.rpt"
          '.ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodicatNetRateSum.RPT"
        ElseIf ChkNetRate.Value = 1 Then
           .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodicatNetRate.RPT"
        ElseIf ChkSummary.Value = 1 Or chkbrl.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodic.RPT"
        ElseIf chkpwcostqty.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodic1.RPT"
        ElseIf chkpwcost.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodic2.RPT"
        ElseIf chkwithamount.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodic3.RPT"
        ElseIf chkwamtsum.Value = 1 Then
            .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodic4.RPT"
       
        Else
        .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerSuppliesWise.rpt"
        '.ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodicDetail.RPT"
        End If
        
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & Me.Caption + " (" + txtstockbase.Text & "  " & txtdesc.Text & ")" & "'"
        If txtstockbase.ListIndex = 3 Then
        .Formulas(2) = "Period = '" & "As on Date : " & DTPTo & "'"
        Else
        .Formulas(2) = "Period = '" & "From " & dtpfrom & " to " & DTPTo & "'"
        
        End If
        
      
        
           
        .Formulas(3) = "groupstatus = " & txtgrouping.ListIndex & ""
        
        
            Call ChkTempTables("Tmp_Stockledger1", True)
            Call ChkTempTables("Tmp_Stockledger", True)
         
            ls_sql = "SELECT siteid,compcode, ItemCode, ReceiveQty, IssueQty,0 as OPqty ,Rate,InAmount,OutAmount,0 as OpAmount into Tmp_Stockledger1 From StockLedgerTransDetail"
            ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'"
        
            ls_sql = ls_sql & " and transdate >= '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
            ls_sql = ls_sql & " and transdate <= '" & Format(DTPTo, "YYYY/MM/DD") & "'"
            
            If txtStoreType.ListIndex <> 2 Then
            ls_sql = ls_sql & " and siteid = " & txtStoreType.ListIndex + 1 & " "
            End If
            
            If txtitemcode <> "" Then
                    ls_sql = ls_sql & " and itemcode = '" & txtitemcode & "'"
            End If
            
            ls_sql = ls_sql & " Union all"
            
            ls_sql = ls_sql & " SELECT siteid,compcode, ItemCode, 0 as ReceiveQty,0 as IssueQty,SUM(ReceiveQty) - SUM(IssueQty) as OpQty,sum(InAmount)-Sum(OutAmount)/case when SUM(ReceiveQty) - SUM(IssueQty)>0 then SUM(ReceiveQty) - SUM(IssueQty)else 1 end as Rate,0 as InAmount,0 as OutAmount,sum(InAmount)-Sum(OutAmount) as OpAmount   From StockLedgerTransDetail"
            ls_sql = ls_sql & " where compcode = '" & Gs_compcode & "'"
         
            ls_sql = ls_sql & " and transdate < '" & Format(dtpfrom, "YYYY/MM/DD") & "'"
            
            If txtStoreType.ListIndex <> 2 Then
            ls_sql = ls_sql & " and siteid = " & txtStoreType.ListIndex + 1 & " "
            End If
            
            If txtitemcode <> "" Then
                 ls_sql = ls_sql & " and itemcode = '" & txtitemcode & "'"
            End If
                    
            ls_sql = ls_sql & " group by siteid,compcode,itemcode"
            gc_dbcon.Execute ls_sql
            
            
            
            
         
        
          
          ls_sql = "SELECT siteid,compcode, ItemCode, SUM(ReceiveQty) AS ReceiveQty, SUM(IssueQty) AS IssueQty, SUM(OPqty) AS OPqty,sum(InAmount) as InAmount ,sum(OutAmount) as OutAmount,  sum(OpAmount) as OpAmount Into Tmp_Stockledger From Tmp_Stockledger1 GROUP BY siteid,compcode, ItemCode"
          gc_dbcon.Execute ls_sql
          
    If ChkTransMovementOnly.Value = 1 Then
        ls_sql = "delete From Tmp_Stockledger WHERE((compcode+ ItemCode) not IN  (SELECT     compcode+ Itemcode  From Tmp_Stockledger  WHERE      Receiveqty > 0 OR   issueqty > 0 ))"
         gc_dbcon.Execute ls_sql
    End If
         
       
         .SQLQuery = ""
           
           
           
            .SQLQuery = "SELECT Tmp_Stockledger. siteid , Tmp_Stockledger. ItemCode , Tmp_Stockledger. ReceiveQty , Tmp_Stockledger. IssueQty , Tmp_Stockledger. OPqty ,"
            .SQLQuery = .SQLQuery & " IC_Item. Description , IC_Item. ClassID , IC_Item. CatCode , IC_Item. Manucode , IC_Item. AvgRate ,"
            .SQLQuery = .SQLQuery & " IC_ItemCategory. Description ,   IC_ItemUM. Description ,    IC_Manufacturer. Description ,"
            .SQLQuery = .SQLQuery & " IC_ItemClass.Description FROM  dbo . Tmp_Stockledger  Tmp_Stockledger LEFT OUTER JOIN  dbo . IC_Item  IC_Item ON"
            .SQLQuery = .SQLQuery & " Tmp_Stockledger.Compcode = IC_Item.Compcode And Tmp_Stockledger.ItemCode = IC_Item.ItemCode"
            .SQLQuery = .SQLQuery & " LEFT OUTER JOIN  dbo . IC_ItemUM  IC_ItemUM ON         IC_Item. MCode  = IC_ItemUM. Mcode"
            .SQLQuery = .SQLQuery & " LEFT OUTER JOIN  dbo . IC_Manufacturer  IC_Manufacturer ON         IC_Item. Compcode  = IC_Manufacturer. Compcode  AND"
            .SQLQuery = .SQLQuery & " IC_Item. Manucode  = IC_Manufacturer. MCode      LEFT OUTER JOIN  dbo . IC_ItemClass  IC_ItemClass ON"
            .SQLQuery = .SQLQuery & " IC_Item. Compcode  = IC_ItemClass. Compcode  AND     IC_Item. CatCode  = IC_ItemClass. Deptcode  AND"
            .SQLQuery = .SQLQuery & " IC_Item. ClassID  = IC_ItemClass. ClassCode      LEFT OUTER JOIN  dbo . IC_ItemCategory  IC_ItemCategory ON"
            .SQLQuery = .SQLQuery & " IC_Item.Compcode = IC_ItemCategory.Compcode And IC_Item.CatCode = IC_ItemCategory.CatCode"
         
         
'         .SQLQuery = "SELECT  Tmp_Stockledger.siteid, Tmp_Stockledger.PQty, Tmp_Stockledger.RPQty, Tmp_Stockledger.SQty, Tmp_Stockledger.SRQty, Tmp_Stockledger.IQTY,"
'
         .SQLQuery = .SQLQuery & " where Tmp_Stockledger.compcode = '" & Gs_compcode & "'"
      
        If ChkSum.Value = 1 Then
           .SQLQuery = .SQLQuery & "  and IC_Item.CatCode in ('012','013','014','015','016','017','018','019','020','022','023','024','025','026','027','028','029','034','046')"
        End If
        
         If txtdeptcode <> "" Then
         .SQLQuery = .SQLQuery & "  and IC_Item.CatCode = '" & txtdeptcode & "'"
         End If
            
         If txtcatcode <> "" Then
         .SQLQuery = .SQLQuery & "  and IC_Item.ClassID = '" & txtcatcode & "'"
         End If
            
         If txtsubcatcode <> "" Then
         .SQLQuery = .SQLQuery & "  and IC_Item.PackCode = '" & txtsubcatcode & "'"
         End If
            
         If txtSuppliercode <> "" Then
         .SQLQuery = .SQLQuery & "  and IC_Item.Manucode = '" & txtSuppliercode & "'"
         End If
         
         
         If chkbrl.Value = 1 Then
           .SQLQuery = .SQLQuery & "  and  ((Tmp_Stockledger.OPqty + Tmp_Stockledger.PQty + Tmp_Stockledger.SRQty) - (Tmp_Stockledger.IQTY + Tmp_Stockledger.SQty + Tmp_Stockledger.RPQty)) < IC_Item.ReorderQty "
         End If
         
         If txtstockbase.ListIndex = 1 Then
            ls_sql = "delete from Tmp_Stockledger where  ((OPqty + PQty + SRQty) - (IQTY + SQty + RPQty)) <> 0"
            gc_dbcon.Execute ls_sql
         End If
        
         If txtstockbase.ListIndex = 2 Then
             ls_sql = "delete from Tmp_Stockledger where  ((OPqty + PQty + SRQty) - (IQTY + SQty + RPQty))>=  0"
             gc_dbcon.Execute ls_sql
         End If
        
         If txtstockbase.ListIndex = 3 Then
            ls_sql = "delete from Tmp_Stockledger where itemcode in ( SELECT SO_Trans.ItemCode FROM  SO_TransMaster INNER JOIN  SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode"
            ls_sql = ls_sql & " where SO_TransMaster.compcode = '" & Gs_compcode & "'   and SO_TransMaster.transdate <='" & Format(DTPTo, "YYYY/MM/DD") & "')"
            gc_dbcon.Execute ls_sql
         End If
        
        If Chkprintstocktax.Value = 1 Then
        .ReportFileName = App.Path & Gs_ICRepoPath & "\StockledgerPeriodic.RPT"
        ls_sql = "delete from Tmp_Stockledger where itemcode not in (SELECT PO_POGRNDetail.ItemCode FROM  PO_POGRN  INNER JOIN"
        ls_sql = ls_sql & " PO_POGRNDetail ON PO_POGRN.Compcode = PO_POGRNDetail.Compcode AND PO_POGRN.TransCode = PO_POGRNDetail.TransCode"
        ls_sql = ls_sql & " Where (PO_POGRNDetail.GSTAmount > 0 and PO_POGRN.compcode = '" & Gs_compcode & "'   and PO_POGRN.transdate <='" & Format(dtpfrom, "YYYY/MM/DD") & "'))"
        gc_dbcon.Execute ls_sql
          
        End If
        
        
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
End If

Exit Sub

localerr:
Call SetErr(Err.Description, vbCritical)
On Error GoTo 0
End Sub
Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtdeptcode
    Set PO_DESC = txtdesc
    Gs_SQL = "Select CatCode,   Description from IC_ItemCategory "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Departments"
    MyLookupOLDB.Show 1
    
    If txtdeptcode <> "" Then Call txtdeptcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command2_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcatcode
    Set PO_DESC = txtcatdesc
    Gs_SQL = "Select ClassCode,   Description from IC_ItemClass "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "'"
    MyLookupOLDB.Caption = "Categories"
    MyLookupOLDB.Show 1
    
    If txtcatcode <> "" Then Call txtcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsubcatcode
    Set PO_DESC = txtsubcatdesc
    Gs_SQL = "Select PackCode,   Description from IC_ItemPacking "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' and subcode = '" & txtcatcode & "' and deptcode = '" & txtdeptcode & "' "
    MyLookupOLDB.Caption = "Sub Categories"
    MyLookupOLDB.Show 1
    
    If txtsubcatcode <> "" Then Call txtsubcatcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub Command4_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtSuppliercode
    Set PO_DESC = txtSupplierdesc
    Gs_SQL = "Select SupplierCode,   Description from IC_Supplier "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Supplier"
    MyLookupOLDB.Show 1

    
    If txtSuppliercode <> "" Then Call txtSuppliercode_KeyDown(vbKeyReturn, vbKeyShift)


End Sub

Private Sub txtcatcode_Change()
If txtcatcode = "" Then
txtcatdesc = ""
End If
End Sub

Private Sub txtcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtcatcode = DoPad(txtcatcode, 3)
       If pr_dumy.State = 1 Then pr_dumy.Close
        pr_dumy.Open "Select * from Ic_ItemClass where Classcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Category Code not found !!!", vbCritical)
            txtcatcode = ""
            txtcatdesc = ""
            txtcatcode.SetFocus
        Else
            txtcatdesc = pr_dumy("Description")
             If txtsubcatcode.Enabled Then txtsubcatcode.SetFocus
           
        End If
        pr_dumy.Close
        
ElseIf Trim(txtcatcode) = "" And KeyCode = vbKeyReturn Then
        txtcatcode = ""
        txtcatdesc = ""
        Command2_Click
End If
End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtitemcode
    Set PO_DESC = txtitemdesc
    
    Gs_SQL = "SELECT itemCode, Description  FROM IC_Item "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where compcode = '" & Gs_compcode & "' "
    
    MyLookupOLDB.Caption = "Items"
    MyLookupOLDB.Show 1
    
    If txtitemcode <> "" Then Call txtItemcode_KeyDown(vbKeyReturn, vbKeyShift)

End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    DTPTo.SetFocus
End If
End Sub
Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtStoreType.SetFocus
End If
End Sub

Private Sub Form_Load()
  
  dtpfrom = Date
  DTPTo = Date
  txtStoreType = "SHOWROOM"
  txtgrouping.ListIndex = 0
End Sub

Private Sub txtdeptcode_Change()
If txtdeptcode = "" Then
txtdesc = ""
End If
End Sub

Private Sub txtdeptcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtdeptcode <> "" And KeyCode = vbKeyReturn Then
    txtdeptcode = DoPad(txtdeptcode, txtdeptcode.MaxLength)
    ls_sql = "Select Catcode,Description from IC_ItemCategory where compcode = '" & Gs_compcode & "' and Catcode = '" & txtdeptcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Department Code not found", vbCritical)
            Else
                txtdesc = pr_dumy("description")
               txtcatcode.SetFocus
            End If
         pr_dumy.Close
ElseIf txtdeptcode = "" And KeyCode = vbKeyReturn Then
Command1_Click
End If
End Sub


Private Sub txtgrouping_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    cmdGenerate.SetFocus
End If
End Sub

Private Sub txtItemcode_Change()
If txtitemcode = "" Then
txtitemdesc = ""
End If

End Sub

Private Sub txtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
If txtitemcode <> "" And KeyCode = vbKeyReturn Then
    txtitemcode = DoPad(txtitemcode, txtitemcode.MaxLength)
    ls_sql = "Select itemcode,Description from IC_Item where compcode = '" & Gs_compcode & "' and itemcode = '" & txtitemcode & "' "
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Item Code not found", vbCritical)
            Else
                txtitemdesc = pr_dumy("description")
                 dtpfrom.SetFocus
            End If
         pr_dumy.Close
ElseIf txtitemcode = "" And KeyCode = vbKeyReturn Then
    Command5_Click
End If
End Sub

Private Sub txtselectedcode_Change()

End Sub

Private Sub txtstockbase_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtgrouping.SetFocus
End If
End Sub

Private Sub txtStoreType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtstockbase.SetFocus
End If
End Sub

Private Sub txtsubcatcode_Change()
If txtsubcatcode = "" Then
txtsubcatdesc = ""
End If
End Sub

Private Sub txtsubcatcode_KeyDown(KeyCode As Integer, Shift As Integer)
If Trim(txtsubcatcode) <> "" And KeyCode = vbKeyReturn Then
        txtsubcatcode.Text = DoPad(txtsubcatcode.Text, 3)
        pr_dumy.Open "Select * from IC_ItemPacking where Packcode = '" & txtsubcatcode & "'  and subcode = '" & txtcatcode & "' and compcode = '" & Gs_compcode & "' and deptcode = '" & txtdeptcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Sub Category code not found !!!", vbCritical)
            txtsubcatcode = ""
            txtsubcatdesc = ""
            txtsubcatcode.SetFocus
        Else
            txtsubcatdesc = pr_dumy("Description")
            If txtSuppliercode.Enabled Then txtSuppliercode.SetFocus
            
        End If
        pr_dumy.Close
        
ElseIf Trim(txtsubcatcode) = "" And KeyCode = vbKeyReturn Then
        txtsubcatcode = ""
        txtsubcatdesc = ""
        Command3_Click
End If

End Sub

Private Sub txtSuppliercode_Change()
If txtSuppliercode = "" Then
txtSupplierdesc = ""
End If
End Sub

Private Sub txtSuppliercode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtSuppliercode) <> "" And KeyCode = vbKeyReturn Then
        txtSuppliercode.Text = DoPad(txtSuppliercode.Text, 6)
        pr_dumy.Open "Select * from IC_Supplier where Suppliercode = '" & txtSuppliercode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Supplier code not found !!!", vbCritical)
            txtSuppliercode = ""
            txtSupplierdesc = ""
            txtSuppliercode.SetFocus
        Else
            txtSupplierdesc = pr_dumy("Description")
            If txtitemcode.Enabled Then txtitemcode.SetFocus
            
        End If
        pr_dumy.Close
ElseIf Trim(txtSuppliercode) = "" And KeyCode = vbKeyReturn Then
        txtSuppliercode = ""
        txtSupplierdesc = ""
        Command4_Click
End If

End Sub
