VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adjustment"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK "
      Height          =   345
      Left            =   2700
      TabIndex        =   4
      Top             =   1560
      Width           =   1665
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1470
      TabIndex        =   3
      Top             =   495
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   609
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   41061
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1470
      TabIndex        =   2
      Text            =   "BPP"
      Top             =   135
      Width           =   780
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   1470
      TabIndex        =   5
      Top             =   930
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   609
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   41061
   End
   Begin VB.Label Label3 
      Caption         =   "To Date :"
      Height          =   255
      Left            =   675
      TabIndex        =   6
      Top             =   990
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "From Date :"
      Height          =   255
      Left            =   495
      TabIndex        =   1
      Top             =   555
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Voucher Type :"
      Height          =   495
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ls_sql As String
Dim ls_CrAccount As String
Dim pr_dumy As New Recordset
Dim pr_dumy1 As New Recordset

ls_sql = "select * from gl_ref where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(DTPicker1, "YYYY/MM/DD") & "' and value_date <= '" & Format(DTPicker2, "YYYY/MM/DD") & "' and vchrtype = '" & Text1 & "'  order by voucher_no"
If pr_dumy.State = 1 Then pr_dumy.Close
pr_dumy.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
Do While Not pr_dumy.EOF

ls_sql = "select * from gl_trans where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(DTPicker1, "YYYY/MM/DD") & "' and value_date <= '" & Format(DTPicker2, "YYYY/MM/DD") & "' and vchrtype = '" & Text1 & "'  and voucher_no = '" & pr_dumy("Voucher_no") & "' and cr_Amount > 0  "
pr_dumy1.Open ls_sql, gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy1.EOF Then
ls_CrAccount = pr_dumy1("Accountno")

ls_sql = "delete  from gl_trans where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(DTPicker1, "YYYY/MM/DD") & "' and value_date <= '" & Format(DTPicker2, "YYYY/MM/DD") & "' and vchrtype = '" & Text1 & "'  and voucher_no = '" & pr_dumy("Voucher_no") & "' and cr_Amount > 0  and accountno = '" & ls_CrAccount & "' "
gc_dbcon.Execute ls_sql

ls_sql = "Insert into gl_trans (Compcode, BranchCode, accountno, SerialNo, AcctName, Acct_Nirration, Value_Date, Voucher_No, VchrType, ExchgRate, Dr_Amount, Cr_Amount, "
ls_sql = ls_sql & " AddDate, Addtime, userid, pflag, instrumentno, newVoucherno) "
ls_sql = ls_sql & " select Compcode, BranchCode, '" & ls_CrAccount & "' as accountno, SerialNo, AcctName, Acct_Nirration, Value_Date, Voucher_No, VchrType, ExchgRate, 0 as Dr_Amount, Dr_Amount as CR_Amount ,AddDate, Addtime, userid, pflag, instrumentno, newVoucherno  from gl_trans where compcode = '" & Gs_compcode & "' and value_date >= '" & Format(DTPicker1, "YYYY/MM/DD") & "' and value_date <= '" & Format(DTPicker2, "YYYY/MM/DD") & "' and vchrtype = '" & Text1 & "'  and voucher_no = '" & pr_dumy("Voucher_no") & "' and Dr_Amount > 0  "
gc_dbcon.Execute ls_sql

End If
pr_dumy1.Close
pr_dumy.MoveNext
Loop
End If
pr_dumy.Close
Call MsgBox("Successfully Updated", vbInformation)
End Sub

Private Sub Form_Load()
DTPicker1 = Date
DTPicker2 = Date
End Sub
