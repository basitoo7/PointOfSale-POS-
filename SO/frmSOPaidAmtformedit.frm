VERSION 5.00
Begin VB.Form frmSOPaidAmtformEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receive Amount"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSOPaidAmtformedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6090
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   6000
      Begin VB.ComboBox txtdiscBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1785
         Width           =   3030
      End
      Begin VB.ComboBox txtdiscBy1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2790
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1785
         Visible         =   0   'False
         Width           =   3030
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   19
         Top             =   5550
         Width           =   1530
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4275
         TabIndex        =   18
         Top             =   5550
         Width           =   1530
      End
      Begin VB.TextBox txtdisamount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2790
         TabIndex        =   0
         Top             =   1005
         Width           =   2970
      End
      Begin VB.TextBox txtnetamount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2250
         Width           =   2970
      End
      Begin VB.TextBox txtCreditCode 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   5
         Top             =   4605
         Width           =   780
      End
      Begin VB.TextBox txtCreditDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   30
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   5100
         Width           =   5760
      End
      Begin VB.CommandButton Command5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3555
         Picture         =   "frmSOPaidAmtformedit.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4590
         Width           =   360
      End
      Begin VB.CheckBox ChkCreditSale 
         Caption         =   "Credit Sale"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3960
         TabIndex        =   9
         Top             =   4620
         Width           =   1710
      End
      Begin VB.TextBox txtBalAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2760
         TabIndex        =   4
         Top             =   3825
         Width           =   2970
      End
      Begin VB.TextBox txtRecAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   2760
         TabIndex        =   3
         Top             =   3030
         Width           =   2970
      End
      Begin VB.TextBox txttotalamount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2970
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount By :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   60
         TabIndex        =   17
         Top             =   1740
         Width           =   2670
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Discount Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   75
         TabIndex        =   16
         Top             =   1110
         Width           =   2670
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Net Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   15
         Top             =   2370
         Width           =   2235
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit ID :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   810
         TabIndex        =   14
         Top             =   4545
         Width           =   1830
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance  Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   135
         TabIndex        =   13
         Top             =   3960
         Width           =   2580
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Receive Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   12
         Top             =   3105
         Width           =   2580
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   2580
      End
   End
End
Attribute VB_Name = "frmSOPaidAmtformEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_Dumy As New Recordset
Public PO_DESC As Object
Public PO_CODE As Object
Dim ls_sql As String
Dim ls_res
Private Sub ChkCreditSale_Click()
If ChkCreditSale.Value = 1 Then
    txtCreditCode.Enabled = True
    txtCreditDesc.Enabled = True
    Command5.Enabled = True
    txtCreditCode.SetFocus
Else
    txtCreditCode = ""
    txtCreditDesc = ""
    txtCreditCode.Enabled = False
    txtCreditDesc.Enabled = False
    Command5.Enabled = False
End If
End Sub

Private Sub ChkDisc_Click()
End Sub

Private Sub Command1_Click()
Dim ls_clientcode As String
With frmSO_PosformEdit
.txtstatus = "OK"
If ChkCreditSale.Value = 1 Then
    ls_clientcode = txtCreditCode
Else
    ls_clientcode = "000019"
End If

gc_dbcon.Execute "delete from SO_TransMaster where compcode = '" & Gs_compcode & "' and transcode = '" & .TXTINVNUMBER & "' "
gc_dbcon.Execute "delete from SO_Trans where compcode = '" & Gs_compcode & "' and transcode = '" & .TXTINVNUMBER & "' "

ls_sql = "Insert into SO_TransMaster(Compcode, TransCode, TransDate, AccountCode, Remarks, TotalAmount, DiscPer, DiscAmount, NetAmount, RecAmount, BalAmount,DiscBy,SaleStatus,usercode)"
ls_sql = ls_sql & " Values ('" & Gs_compcode & "' , '" & .TXTINVNUMBER & "', '" & Format(.TXTINVDATE, "YYYY/MM/DD HH:MM:SS") & "', '" & ls_clientcode & "' , 'Sale' , " & Val(txttotalamount) & ", " & Val(txtdiscper) & ", " & Val(txtdisamount) & ", " & Val(txtnetamount) & ", " & Val(txtRecAmount) & ", " & Val(txtBalAmount) & ",'" & txtdiscBy1.Text & "'," & ChkCreditSale.Value & "," & Gn_UserCode & ")"
gc_dbcon.Execute ls_sql

 With frmSO_PosformEdit.MSFlexPOS
       For ln_cnt = 1 To .Rows - 1
       If .TextMatrix(ln_cnt, 1) <> "" Then
        ls_sql = "INSERT into SO_Trans(Compcode, TransCode,customcode, ItemCode, Quantity,Itemrate,Amount)"
        ls_sql = ls_sql & " VALUES ('" & Gs_compcode & "','" & Trim(frmSO_PosformEdit.TXTINVNUMBER) & "','" & Trim(.TextMatrix(ln_cnt, 1)) & "','" & Trim(.TextMatrix(ln_cnt, 10)) & "'," & (Val(0 & .TextMatrix(ln_cnt, 3))) & "," & Val(.TextMatrix(ln_cnt, 4)) & "," & Val(.TextMatrix(ln_cnt, 5)) & ")"
        gc_dbcon.Execute ls_sql
      End If
      Next
  End With
Unload Me

End With
End Sub

Private Sub Command2_Click()
frmSO_Posform.txtstatus = "Cancel"
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
frmSO_Posform.txtstatus = "Cancel"
Unload Me
End If
End Sub

Private Sub Form_Load()
LoadDiscBy
End Sub
Private Sub LoadDiscBy()
Dim pr_dumyload As New Recordset
txtdiscBy.Clear

pr_dumyload.Open "select * from SO_AuthorityPerson order by aname", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumyload.EOF Then
 Do While Not pr_dumyload.EOF

    txtdiscBy.AddItem pr_dumyload("Aname")
    txtdiscBy1.AddItem pr_dumyload("Acode")
    pr_dumyload.MoveNext
 Loop
End If
pr_dumyload.Close
End Sub

Private Sub txtdisamount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtdiscBy.SetFocus
End Sub

Private Sub txtdisamount_LostFocus()
If txtdisamount <> "" Then
    If Val(txtdisamount) > Val(frmSO_Posform.txtdiscamount) Then
    ls_res = MsgBox("Discount amount entered not allow are you sure ?", vbYesNo + vbCritical)
    If ls_res = vbNo Then
     txtdisamount = 0
     txtdisamount.SetFocus
    End If
  End If
Else
 txtdisamount = 0
End If
txtnetamount = Val(txttotalamount) - Val(txtdisamount)
End Sub

Private Sub txtdiscBy_Click()
txtdiscBy1.ListIndex = txtdiscBy.ListIndex
End Sub
Private Sub txtdiscBy_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtRecAmount.SetFocus
End Sub

Private Sub txtRecAmount_Change()
If txtRecAmount <> "" Then
    If Not IsNumeric(txtRecAmount) Then
    Call MsgBox("Numeric entery only !!!", vbCritical)
    txtRecAmount = ""
    txtRecAmount.SetFocus
    Exit Sub
End If
txtBalAmount = Val(txtRecAmount) - Val(txtnetamount)
End If
End Sub
Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtCreditCode
    Set PO_DESC = txtCreditDesc
    Gs_SQL = "Select ClientCode, Description from IC_clients "
    Gs_FindFld = "Description"
    Gs_OrderBy = "Order by Description"
    Gs_OtherPara = " where Compcode ='001'"
    MyLookupOLDB.Caption = "Credit Clients"
    MyLookupOLDB.Show 1
    
    If txtCreditCode <> "" Then Call txtCreditCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub txtCreditCode_KeyDown(KeyCode As Integer, Shift As Integer)

If Trim(txtCreditCode) <> "" And KeyCode = vbKeyReturn Then
        txtCreditCode = DoPad(txtCreditCode, 6)
        PR_Dumy.Open "Select * from IC_clients where Compcode  = '001' and Clientcode = '" & txtCreditCode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If PR_Dumy.EOF Then
            Call MsgBox("Client Code not found !!!", vbCritical)
            txtCreditCode = ""
            txtCreditCode = ""
            txtCreditCode.SetFocus
        Else
            txtCreditDesc = PR_Dumy("Description")
            Command1.SetFocus
        End If
        PR_Dumy.Close

ElseIf Trim(txtCreditCode) = "" And KeyCode = vbKeyReturn Then
        txtCreditCode = ""
        txtCreditDesc = ""
End If

End Sub

Private Sub txtRecAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then ChkCreditSale.SetFocus
End Sub
Private Sub ChkCreditSale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Command1.SetFocus
End Sub

