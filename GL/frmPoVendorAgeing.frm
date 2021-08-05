VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPoVendorAging 
   Caption         =   "Account Payable Ageing"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPoVendorAgeing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3495
      Width           =   6015
      _ExtentX        =   10610
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
      Height          =   3555
      Left            =   30
      TabIndex        =   13
      Top             =   -60
      Width           =   5940
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   4650
         TabIndex        =   12
         Top             =   3135
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
         Left            =   3570
         MaskColor       =   &H00000000&
         TabIndex        =   11
         Top             =   3135
         Width           =   1035
      End
      Begin Crystal.CrystalReport crrpt 
         Left            =   90
         Top             =   1770
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
      Begin VB.TextBox txtVchrDesc 
         Height          =   315
         Left            =   2850
         MaxLength       =   50
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   855
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame Frame3 
         ForeColor       =   &H00000080&
         Height          =   2955
         Left            =   75
         TabIndex        =   14
         Top             =   135
         Width           =   5655
         Begin VB.TextBox Txtday41 
            Height          =   345
            Left            =   3045
            MaxLength       =   4
            TabIndex        =   10
            Text            =   "180"
            Top             =   2385
            Width           =   705
         End
         Begin VB.TextBox txtday4 
            Height          =   345
            Left            =   1545
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "91"
            Top             =   2385
            Width           =   705
         End
         Begin VB.TextBox Txtday21 
            Height          =   345
            Left            =   3045
            MaxLength       =   4
            TabIndex        =   5
            Text            =   "60"
            Top             =   1545
            Width           =   705
         End
         Begin VB.TextBox Txtday31 
            Height          =   345
            Left            =   3045
            MaxLength       =   4
            TabIndex        =   8
            Text            =   "90"
            Top             =   1965
            Width           =   705
         End
         Begin VB.TextBox txtday2 
            Height          =   345
            Left            =   1545
            MaxLength       =   4
            TabIndex        =   6
            Text            =   "31"
            Top             =   1545
            Width           =   705
         End
         Begin VB.TextBox txtday3 
            Height          =   345
            Left            =   1545
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "61"
            Top             =   1965
            Width           =   705
         End
         Begin VB.TextBox txtday11 
            Height          =   345
            Left            =   3045
            MaxLength       =   4
            TabIndex        =   4
            Text            =   "30"
            Top             =   1140
            Width           =   705
         End
         Begin VB.TextBox txtday1 
            Height          =   345
            Left            =   1545
            MaxLength       =   4
            TabIndex        =   3
            Text            =   "1"
            Top             =   1140
            Width           =   705
         End
         Begin MSComCtl2.DTPicker dtpfrom 
            Height          =   315
            Left            =   1545
            TabIndex        =   2
            Top             =   780
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
            Format          =   63242241
            CurrentDate     =   37293
         End
         Begin VB.Frame Frame5 
            Height          =   585
            Left            =   30
            TabIndex        =   17
            Top             =   135
            Width           =   5580
            Begin VB.TextBox txtselectedcode 
               Height          =   315
               Left            =   1200
               TabIndex        =   1
               Top             =   180
               Width           =   1545
            End
            Begin VB.TextBox txtdesc 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               ForeColor       =   &H80000002&
               Height          =   315
               Left            =   3075
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   180
               Width           =   2445
            End
            Begin VB.CommandButton Command5 
               Height          =   315
               Left            =   2730
               Picture         =   "frmPoVendorAgeing.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   180
               Width           =   315
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Account Code :"
               Height          =   210
               Left            =   60
               TabIndex        =   20
               Top             =   225
               Width           =   1125
            End
         End
         Begin VB.Label Label10 
            Caption         =   "From Day :"
            Height          =   225
            Left            =   720
            TabIndex        =   28
            Top             =   2430
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "To Day :"
            Height          =   225
            Left            =   2430
            TabIndex        =   27
            Top             =   2430
            Width           =   630
         End
         Begin VB.Label Label4 
            Caption         =   "From Day :"
            Height          =   225
            Left            =   720
            TabIndex        =   26
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            Caption         =   "To Day :"
            Height          =   225
            Left            =   2430
            TabIndex        =   25
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "From Day :"
            Height          =   225
            Left            =   720
            TabIndex        =   24
            Top             =   1590
            Width           =   825
         End
         Begin VB.Label Label6 
            Caption         =   "To Day :"
            Height          =   225
            Left            =   2430
            TabIndex        =   23
            Top             =   1590
            Width           =   645
         End
         Begin VB.Label Label1 
            Caption         =   "From Day :"
            Height          =   225
            Left            =   720
            TabIndex        =   22
            Top             =   2010
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "To Day :"
            Height          =   225
            Left            =   2430
            TabIndex        =   21
            Top             =   2010
            Width           =   630
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "As on Date :"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   600
            TabIndex        =   15
            Top             =   810
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "frmPoVendorAging"
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




Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
'On Error GoTo LocalErr

MDIForm1.StatusBar1.Panels(7).Text = "Processing Data Please Wait..."
    With crrpt
        .WindowTitle = "" & Me.Caption & ""
        .Formulas(0) = "CompanyName = '" & Gs_CompName & "'"
        .Formulas(1) = "ReportName = '" & Me.Caption & "'"
        .ReportFileName = App.Path & Gs_GlRepoPath & "\VendorAgeingReportSum.RPT"
        .Formulas(2) = "Period = '" & " As on date " & dtpfrom & "'"
        '.Formulas(3) = "Currentdate = cdate(" & Format(dtpfrom, "yyyy,MM,dd") & ")"
        
        .Formulas(4) = "Cday1 = '" & txtday1 + "-" + txtday11 + " Days " & "'"
        .Formulas(5) = "Cday2 = '" & txtday2 + "-" + Txtday21 + " Days " & "'"
        .Formulas(6) = "Cday3 = '" & txtday3 + "-" + Txtday31 + " Days " & "'"
        .Formulas(7) = "Cday4 = '" & txtday4 + "-" + Txtday41 + " Days " & "'"
        .Formulas(8) = "Cday5 = '" & "Above " + Txtday41 + " Days " & "'"
        
'        .Formulas(9) = "day1 = " & Val(txtday1) & ""
'        .Formulas(10) = "day11 = " & Val(txtday11) & ""
'        .Formulas(11) = "day2 = " & Val(txtday2) & ""
'        .Formulas(12) = "day21 = " & Val(Txtday21) & ""
'        .Formulas(13) = "day3 = " & Val(txtday3) & ""
'        .Formulas(14) = "day31 = " & Val(Txtday31) & ""
'        .Formulas(15) = "day4 = " & Val(txtday4) & ""
'        .Formulas(16) = "day41 = " & Val(Txtday41) & ""
'        .Formulas(17) = "day5 = " & Val(Txtday41) & ""
          Call ChkTempTables("Tmp_Aging1", True)
          Call ChkTempTables("Tmp_Aging", True)
          ls_sql = "SELECT Gl_Trans.accountno, round(SUM(Gl_Trans.Cr_Amount),0) AS Amount1, 0 AS Amount2, 0 AS Amount3, 0 AS Amount4, 0 AS Amount5,0 as DrAmount into Tmp_Aging1"
          ls_sql = ls_sql & " FROM Gl_Trans INNER JOIN  Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND Gl_Trans.accountno = Gl_Detail.AccountNo"
          ls_sql = ls_sql & " WHERE     (Gl_Detail.Acct_Type = 'C') "
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') >=" & Val(txtday1) & ""
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') <=" & Val(txtday11) & ""
          
          
          
          ls_sql = ls_sql & " GROUP BY Gl_Trans.accountno"
          ls_sql = ls_sql & " Union All"
            
          
          
          ls_sql = ls_sql & " SELECT Gl_Trans.accountno, 0 AS Amount1, round(SUM(Gl_Trans.Cr_Amount),0) AS Amount2, 0 AS Amount3, 0 AS Amount4, 0 AS Amount5,0 as DrAmount"
          ls_sql = ls_sql & " FROM Gl_Trans INNER JOIN  Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND Gl_Trans.accountno = Gl_Detail.AccountNo"
          ls_sql = ls_sql & " WHERE     (Gl_Detail.Acct_Type = 'C') "
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') >=" & Val(txtday2) & ""
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') <=" & Val(Txtday21) & ""
          
          
          ls_sql = ls_sql & " GROUP BY Gl_Trans.accountno"
          ls_sql = ls_sql & " Union All"
          
            
          ls_sql = ls_sql & " SELECT Gl_Trans.accountno, 0 AS Amount1, 0 AS Amount2,  round(SUM(Gl_Trans.Cr_Amount),0) AS Amount3, 0 AS Amount4, 0 AS Amount5,0 as DrAmount"
          ls_sql = ls_sql & " FROM Gl_Trans INNER JOIN  Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND Gl_Trans.accountno = Gl_Detail.AccountNo"
          ls_sql = ls_sql & " WHERE     (Gl_Detail.Acct_Type = 'C') "
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') >=" & Val(txtday3) & ""
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') <=" & Val(Txtday31) & ""
          
          
          ls_sql = ls_sql & " GROUP BY Gl_Trans.accountno"
          ls_sql = ls_sql & " Union All"
        
          
          ls_sql = ls_sql & " SELECT Gl_Trans.accountno, 0 AS Amount1, 0 AS Amount2, 0 AS Amount3,  round(SUM(Gl_Trans.Cr_Amount),0) AS Amount4, 0 AS Amount5,0 as DrAmount"
          ls_sql = ls_sql & " FROM Gl_Trans INNER JOIN  Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND Gl_Trans.accountno = Gl_Detail.AccountNo"
          ls_sql = ls_sql & " WHERE     (Gl_Detail.Acct_Type = 'C') "
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') >=" & Val(txtday4) & ""
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') <=" & Val(Txtday41) & ""
          
          ls_sql = ls_sql & " GROUP BY Gl_Trans.accountno"
          
          ls_sql = ls_sql & " Union All"
        
          
          ls_sql = ls_sql & " SELECT Gl_Trans.accountno, 0 AS Amount1, 0 AS Amount2, 0 AS Amount3, 0 as Amount4, round(SUM(Gl_Trans.Cr_Amount),0) AS Amount5,0 as DrAmount"
          ls_sql = ls_sql & " FROM Gl_Trans INNER JOIN  Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND Gl_Trans.accountno = Gl_Detail.AccountNo"
          ls_sql = ls_sql & " WHERE     (Gl_Detail.Acct_Type = 'C') "
          ls_sql = ls_sql & " AND  DATEDIFF(day, Value_Date, '" & Format(dtpfrom, "YYYY/MM/DD") & "') >" & Val(Txtday41) & ""
          
          ls_sql = ls_sql & " GROUP BY Gl_Trans.accountno"
          
          ls_sql = ls_sql & " Union All"
        
          
          ls_sql = ls_sql & " SELECT Gl_Trans.accountno, 0 AS Amount1, 0 AS Amount2, 0 AS Amount3, 0 as Amount4,0 AS Amount5 , round(SUM(Gl_Trans.Dr_Amount),0) as DrAmount"
          ls_sql = ls_sql & " FROM Gl_Trans INNER JOIN  Gl_Detail ON Gl_Trans.compcode = Gl_Detail.compcode AND Gl_Trans.accountno = Gl_Detail.AccountNo"
          ls_sql = ls_sql & " WHERE     (Gl_Detail.Acct_Type = 'C') "
          
          
          ls_sql = ls_sql & " GROUP BY Gl_Trans.accountno"
          
          
          gc_dbcon.Execute ls_sql
          ls_sql = "SELECT accountno, SUM(Amount1) AS Amount1, SUM(Amount2) AS Amount2, SUM(Amount3) AS Amount3, SUM(Amount4) AS Amount4, SUM(Amount5)AS Amount5, SUM(DrAmount) AS DrAmount into Tmp_Aging  FROM Tmp_Aging1 GROUP BY accountno"
          gc_dbcon.Execute ls_sql
       
          gc_dbcon.Execute "delete from tmp_aging where amount1+amount2+amount3+amount4+amount5+dramount =0"
          
          gc_dbcon.Execute "delete from tmp_aging where amount1+amount2+amount3+amount4+amount5=dramount "
          
          If txtselectedcode <> "" Then
            gc_dbcon.Execute "delete from tmp_aging where Accountno <> '" & txtselectedcode & "' "
          End If
         
        .Connect = "DNS=Censoft;UID=Sa"
        .Action = 1
    End With
MDIForm1.StatusBar1.Panels(7).Text = ""
Exit Sub

LocalErr:
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Sub Command5_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtselectedcode
    Set PO_DESC = txtdesc
    Gs_SQL = "SELECT Accountno,Acct_desc  from Gl_detail"
    Gs_FindFld = "Acct_desc"
    Gs_OrderBy = "Order by Acct_desc"
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and acct_type = 'C'"
    MyLookupOLDB.Caption = "Suppliers"
    MyLookupOLDB.Show 1
    SendKeys "{Tab}"
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub Form_Load()
  
  dtpfrom = Date
  
  
End Sub
Private Sub txtselectedcode_Validate(Cancel As Boolean)
If txtselectedcode <> "" Then
    txtselectedcode = DoPad(txtselectedcode, txtselectedcode.MaxLength)
    'If Me.Caption = "Vendor Account History" Then
        ls_sql = "Select accountno,acct_desc from gl_detail where compcode = '" & Gs_compcode & "' and Accountno = '" & txtselectedcode & "' "
    'Else
      '  ls_sql = "Select clientcode,Description from Ic_clients where compcode = '" & Gs_compcode & "' and clientcode = '" & txtselectedcode & "' "
    'End If
          pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If pr_dumy.EOF Then
                Call MsgBox("Supplier Code Not Found", vbCritical)
                'Cancel = True
            Else
                txtdesc = pr_dumy("Acct_desc")
            End If
         pr_dumy.Close

ElseIf txtselectedcode = "" Then
txtdesc = ""
End If
End Sub

