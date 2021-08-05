VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSOPostSale 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Post Sale Voucher"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frmSOPostSale.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdisccash 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3555
      MaxLength       =   64
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   915
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   4905
      Begin VB.TextBox txtsalereturn 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1635
      End
      Begin VB.TextBox txtdisccredit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3540
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1305
         Width           =   1290
      End
      Begin VB.TextBox txtadjustmentN 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   1245
         MaxLength       =   64
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2490
         Width           =   1620
      End
      Begin VB.TextBox txtadjustmentP 
         BackColor       =   &H80000004&
         Height          =   375
         Left            =   1245
         MaxLength       =   64
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1620
      End
      Begin VB.TextBox txtNetAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1620
      End
      Begin VB.TextBox txtCreditAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1305
         Width           =   1620
      End
      Begin VB.TextBox txtCashAmount 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1245
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   930
         Width           =   1635
      End
      Begin VB.ComboBox txtcasher1 
         Height          =   315
         ItemData        =   "frmSOPostSale.frx":030A
         Left            =   1230
         List            =   "frmSOPostSale.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   570
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Post Sale Voucher"
         Height          =   315
         Left            =   2685
         TabIndex        =   2
         Top             =   3855
         Width           =   2130
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   30
         TabIndex        =   1
         Top             =   3285
         Width           =   4830
         Begin VB.Label lblStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   75
            TabIndex        =   3
            Top             =   180
            Width           =   4185
         End
      End
      Begin MSComCtl2.DTPicker dtpfrom 
         Height          =   315
         Left            =   1245
         TabIndex        =   4
         Top             =   165
         Width           =   2085
         _ExtentX        =   3678
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
         CustomFormat    =   "dd-MM-yyyy HH:mm:ss"
         Format          =   54394881
         CurrentDate     =   37293
      End
      Begin VB.ComboBox txtcasher 
         Height          =   315
         ItemData        =   "frmSOPostSale.frx":030E
         Left            =   1245
         List            =   "frmSOPostSale.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   570
         Width           =   2505
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sale Return :"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   24
         Top             =   2970
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Disc :"
         Height          =   195
         Index           =   6
         Left            =   3060
         TabIndex        =   22
         Top             =   1365
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Disc :"
         Height          =   195
         Index           =   5
         Left            =   3075
         TabIndex        =   21
         Top             =   945
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Adjustment[-] :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   2550
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Adjustment[+] :"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   2100
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Net Amount :"
         Height          =   210
         Index           =   1
         Left            =   255
         TabIndex        =   13
         Top             =   1695
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Credit Amount :"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   1335
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cash Amount :"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   9
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Date :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   690
         TabIndex        =   6
         Top             =   180
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Casher :"
         Height          =   210
         Left            =   570
         TabIndex        =   5
         Top             =   585
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmSOPostSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ls_transcode As String
Dim ls_transcodePS As String
Dim ls_sql  As String
Dim PR_Dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim ln_cnt

Private Function maxtranscode() As String
pr_dumy1.Open "select max(transcode) as transcode from IC_TransMaster where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy1("transcode")) + 1)), 10)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy1.Close
End Function
Private Function maxtranscodePS(ls_transtype As String) As String
pr_dumy1.Open "select max(InvoiceNo) as transcode from IC_TransMaster where compcode = '" & Gs_compcode & "' and transtype = '" & ls_transtype & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
    maxtranscodePS = DoPad(Trim(str(Int(0 & pr_dumy1("transcode")) + 1)), 10)
Else
    maxtranscodePS = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy1.Close
End Function


Private Sub Command1_Click()
If Trim(txtcasher.Text) = "" Then
 Call MsgBox("Select Casher !!!", vbCritical)
 txtcasher1.SetFocus
 Exit Sub
End If

If Val(txtCashAmount) <> 0 Then
lblStatus.Caption = "Cash Voucher Posting ... "
Me.Refresh
 If PostCashSaleVoucher(Gs_compcode, Val(txtCashAmount), Val(txtdisccash), dtpfrom, txtcasher1.Text, txtcasher.Text) Then
    Call MsgBox("Cash Sale Voucher Successfully Posted !!!", vbInformation)
 Else
    Call MsgBox("Voucher not Posted !!!", vbCritical)
 End If
 
End If

If Val(txtadjustmentN) <> 0 Or Val(txtadjustmentP) <> 0 Then
lblStatus.Caption = "Cash Adjustment Voucher Posting ... "
Me.Refresh
 
 If PostCashAdjVoucher(Gs_compcode, Val(txtadjustmentN), Val(txtadjustmentP), dtpfrom, txtcasher1.Text, txtcasher.Text) Then
    Call MsgBox("Cash Adjustment Voucher Successfully Posted !!!", vbInformation)
 Else
    Call MsgBox("Voucher not Posted !!!", vbCritical)
 End If
 
End If


If Val(txtCreditAmount) <> 0 Then
lblStatus.Caption = "Credit Sale Voucher Posting ... "
Me.Refresh
 
 
 If PostCreditSaleVoucher(Gs_compcode, Val(txtdisccredit), dtpfrom, txtcasher1.Text, txtcasher.Text) Then
 Call MsgBox("Credit Sale Voucher Successfully Posted !!!", vbInformation)
 Else
 Call MsgBox("Voucher not Posted !!!", vbCritical)
 End If
End If

If Val(txtCreditAmount) <> 0 Or Val(txtCashAmount) <> 0 Then
lblStatus.Caption = "Consumption Voucher Posting ... "
Me.Refresh
 
 If PostSaleConsumptionVoucher(Gs_compcode, dtpfrom, txtcasher1.Text, txtcasher.Text) Then
 Call MsgBox("Sale Consumption Voucher Successfully Posted !!!", vbInformation)
 Else
 Call MsgBox("Voucher not Posted !!!", vbCritical)
 End If
End If

If Val(txtsalereturn) <> 0 Then
    lblStatus.Caption = "Sale Return Voucher Posting ... "
    Me.Refresh

 If PostCreditSaleVoucher(Gs_compcode, Val(txtdisccredit), dtpfrom, txtcasher1.Text, txtcasher.Text) Then
 Call MsgBox("Credit Sale Voucher Successfully Posted !!!", vbInformation)
 Else
 Call MsgBox("Voucher not Posted !!!", vbCritical)
 End If
End If

If Val(txtCreditAmount) = 0 And Val(txtCashAmount) = 0 Then
 Call MsgBox("No entry for posting !!!", vbInformation)
End If

lblStatus.Caption = ""
Me.Refresh

LoadCasherData

End Sub

Private Sub Form_Load()
  dtpfrom = Date
  LoadCasher
End Sub
Private Sub LoadCasher()
Dim pr_loadcasher As New Recordset
pr_loadcasher.Open "SELECT ltrim(rtrim(UserCode)) as UserCode, ltrim(rtrim(UserName)) as UserName  from SyUsers where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_loadcasher.EOF Then
Do While Not pr_loadcasher.EOF
   txtcasher1.AddItem pr_loadcasher("Username")
   txtcasher.AddItem pr_loadcasher("UserCode")
pr_loadcasher.MoveNext
Loop
End If
pr_loadcasher.Close
End Sub

Private Sub LoadCasherData()
Dim ln_cashamount
Dim ln_Creditamount
Dim ln_cashdisc
Dim ln_Creditdisc

ls_sql = " SELECT SUM(SO_Trans.Amount) AS TotalAmount FROM SO_TransMaster INNER JOIN  SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode"
ls_sql = ls_sql & " Where (SO_TransMaster.UserCode = " & txtcasher.Text & " ) And (SO_TransMaster.GLStatus = 0)  And (SO_TransMaster.SaleStatus = 0) and convert(varchar,SO_TransMaster.transdate,111) = '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

If PR_Dumy.State = 1 Then PR_Dumy.Close
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
If Not PR_Dumy.EOF Then
    ln_cashamount = Val(0 & PR_Dumy("Totalamount"))
End If
PR_Dumy.Close

ls_sql = " SELECT SUM(SO_Trans.Amount) AS TotalAmount FROM SO_TransMaster INNER JOIN  SO_Trans ON SO_TransMaster.Compcode = SO_Trans.Compcode AND SO_TransMaster.TransCode = SO_Trans.TransCode"
ls_sql = ls_sql & " Where (SO_TransMaster.UserCode = " & txtcasher.Text & " ) And (SO_TransMaster.GLStatus = 0)  And (SO_TransMaster.SaleStatus = 1) and convert(varchar,SO_TransMaster.transdate,111) = '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

If PR_Dumy.State = 1 Then PR_Dumy.Close
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
If Not PR_Dumy.EOF Then
    ln_Creditamount = Val(0 & PR_Dumy("Totalamount"))
End If
PR_Dumy.Close

'discount

ls_sql = " SELECT SUM(discAmount) AS DiscAmount FROM SO_TransMaster "
ls_sql = ls_sql & " Where (UserCode = " & txtcasher.Text & " ) And (GLStatus = 0)  And (SaleStatus = 0) and convert(varchar,transdate,111) = '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

If PR_Dumy.State = 1 Then PR_Dumy.Close
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
If Not PR_Dumy.EOF Then
    ln_cashdisc = Val(0 & PR_Dumy("Discamount"))
End If
PR_Dumy.Close

ls_sql = " SELECT SUM(discAmount) AS DiscAmount FROM SO_TransMaster "
ls_sql = ls_sql & " Where (UserCode = " & txtcasher.Text & " ) And (GLStatus = 0)  And (SaleStatus = 1) and convert(varchar,transdate,111) = '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

If PR_Dumy.State = 1 Then PR_Dumy.Close
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
If Not PR_Dumy.EOF Then
    ln_Creditdisc = Val(0 & PR_Dumy("Discamount"))
End If
PR_Dumy.Close

ls_sql = " SELECT SUM(SO_TransReturn.Amount) AS TotalAmount FROM SO_TransReturnMaster INNER JOIN  SO_TransReturn ON SO_TransReturnMaster.Compcode = SO_TransReturn.Compcode AND SO_TransReturnMaster.TransCode = SO_TransReturn.TransCode"
ls_sql = ls_sql & " Where (SO_TransReturnMaster.UserCode = " & txtcasher.Text & " ) And (SO_TransReturnMaster.GLStatus = 0)  and convert(varchar,SO_TransReturnMaster.transdate,111) = '" & Format(dtpfrom, "YYYY/MM/DD") & "' "

If PR_Dumy.State = 1 Then PR_Dumy.Close
PR_Dumy.Open ls_sql, gc_dbcon, adOpenStatic, 1
If Not PR_Dumy.EOF Then
    txtsalereturn = Val(0 & PR_Dumy("Totalamount"))
End If
PR_Dumy.Close


txtdisccash = ln_cashdisc
txtdisccredit = ln_Creditdisc

txtCashAmount = ln_cashamount - ln_cashdisc
txtCreditAmount = ln_Creditamount - ln_Creditdisc

txtnetamount = Val(txtCashAmount) + Val(txtCreditAmount)
txtadjustmentN = ""
txtadjustmentP = ""
End Sub

Private Sub txtadjustmentP_Change()
If txtadjustmentP <> "" Then
txtadjustmentN = ""
End If
End Sub
Private Sub txtadjustmentn_Change()
If txtadjustmentN <> "" Then
txtadjustmentP = ""
End If
End Sub

Private Sub txtadjustmentP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtadjustmentN.SetFocus
End Sub

Private Sub txtcasher1_Click()
txtcasher.ListIndex = txtcasher1.ListIndex
LoadCasherData
txtadjustmentP.SetFocus
End Sub
