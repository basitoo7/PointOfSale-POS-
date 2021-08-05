VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmArRelRcpts 
   Caption         =   "Release to GL."
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmArRelRcpts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame3 
      Height          =   570
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   4815
      Begin VB.TextBox TxtVchrNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   2
         Top             =   180
         Width           =   1215
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
         Left            =   1980
         Picture         =   "FrmArRelRcpts.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   315
      End
      Begin MSMask.MaskEdBox txtVchrType 
         Height          =   315
         Left            =   1380
         TabIndex        =   0
         Top             =   180
         Width           =   615
         _ExtentX        =   1085
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
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher # :"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher Type :"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   270
      TabIndex        =   5
      Top             =   525
      Width           =   4335
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Release to GL."
         Height          =   375
         Left            =   300
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmArRelRcpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_GlDetail As New Recordset
Dim PR_VchType As New Recordset
Dim PR_VchCntr As New Recordset

Dim lb_found As Boolean
Dim Ls_Date As Date

Public PO_DESC As Object
Public PO_CODE As Object
Dim ls_VchRem As String


Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVchrType
    Set PO_DESC = Text1
    
    PR_VchType.Filter = adFilterNone
    PR_VchType.Filter = "VchrBase = '" & frmICPayments.ls_base & "' and Branchcode = '" & Gs_BranchCode & "'"
    GoTop PR_VchType
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    
    If Len(txtVchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
On Error GoTo RollbackRrr
Dim ln_Counter As Integer
ln_Counter = 1
gc_dbcon.BeginTrans
  ' Save Debit Details of Voucher
           If MySeek(frmICPayments.TxtGLAccountNO, "AccountNo", PR_GlDetail) Then
             ls_VchRem = Trim(PR_GlDetail("Acct_desc"))
             gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & frmICPayments.TxtGLAccountNO & "'," & ln_Counter & ",'" & Format(frmICPayments.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "'," & Val(frmICPayments.txtAmount) & ",0,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
           Else
            Call SetErr("GL account not found voucher not posted", vbCritical)
            GoTo RollbackRrr
           End If
           ln_Counter = ln_Counter + 1
           
           If MySeek(frmICPayments.ls_CustomerCode, "AccountNo", PR_GlDetail) Then
             ls_VchRem = Trim(PR_GlDetail("Acct_desc"))
             gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & frmICPayments.ls_CustomerCode & "'," & ln_Counter & ",'" & Format(frmICPayments.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "',0," & Val(frmICPayments.txtAmount) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
           Else
                Call SetErr("GL account not found voucher not posted", vbCritical)
                GoTo RollbackRrr
           End If
  If ln_Counter > 1 Then
 ' Save References of Voucher
        ls_VchRem = "Amount Received from " & frmICPayments.txtJobDesc & IIf(frmICPayments.txtpaymentmode = "Bank", " Vide Cheque # " & frmICPayments.txtinstrument & " of " & frmICPayments.txtbankdesc, "")
        gc_dbcon.Execute "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks, userid,adddate,addtime,exchgrate,InstrumentNo) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(frmICPayments.txtvaluedate, "YYYY/MM/DD") & "','" & Format(frmICPayments.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo.Text & "','" & txtVchrType.Text & "','" & ls_VchRem & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "',0,'" & frmICPayments.txtinstrument & "')"
       
       If MySeek(txtVchrType.Text, "VchrType", PR_VchType) Then
        PR_VchCntr.Filter = adFilterNone
        PR_VchCntr.Filter = "Compcode = '" & Gs_compcode & "' and  VchrType = '" & txtVchrType & "'"
        If PR_VchType.Fields("VchrFrequency") = "1" Then
            PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmICPayments.txtvaluedate.Value)))) = PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmICPayments.txtvaluedate.Value)))) + 1
        Else
            PR_VchCntr.Fields("VchrCount") = PR_VchCntr.Fields("VchrCount") + 1
        End If
        PR_VchCntr.Update
        PR_VchCntr.Requery
       End If
    End If
gc_dbcon.CommitTrans
frmICPayments.ls_VchDesc = Text1
frmICPayments.ls_VchType = txtVchrType
frmICPayments.ls_VchNo = TxtVchrNo
Unload Me
Exit Sub
RollbackRrr:
gc_dbcon.RollbackTrans
Unload Me
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  PR_GlDetail.Open "Select Gl_Detail.AccountNo,Gl_Detail.Acct_Desc,Gl_Detail.Acct_Type from Gl_Detail where CompCode ='" & Gs_compcode & "' Order By AccountNo", gc_dbcon, adOpenStatic, adLockReadOnly
  PR_VchType.Open "Select * from GlVchrType where glVchrType.CompCode ='" & Gs_compcode & "' and VchrType <> '0OB' order by GlVchrType.VchrType ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_VchCntr.Open "SELECT * FROM Gl_VchrCntrs WHERE CompCode = '" & Gs_compcode & "' And VchrYear = " & Year(Gs_Fnperiod) & " ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_VchType.Close
    PR_GlDetail.Close
    PR_VchCntr.Close
End Sub

Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
lb_found = False
If KeyCode = vbKeyReturn Then
    If txtFrom.Text <> "" Then
        lb_found = MySeek(txtFrom.Text, "AccountNo", PR_GlDetail)
        If lb_found Then
            Command1.SetFocus
        Else
            Call SetErr("Record not found", vbCritical)
        End If
    End If
End If
End Sub

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) Then
    txtVchrType = UCase(txtVchrType)
    lb_found = MySeek(txtVchrType.Text, "VchrType", PR_VchType)
    If Not lb_found Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtVchrType.SetFocus
    Else
       Text1 = PR_VchType.Fields("VchrDescrip")
       
       PR_VchCntr.Filter = "Compcode = '" & Gs_compcode & "' and  VchrType = '" & txtVchrType & "'"
       If PR_VchType.Fields("VchrFrequency") = "1" Then
           TxtVchrNo = DoPad((LTrim(Str(0 + PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmICPayments.txtvaluedate.Value))))) + 1)), 10)
       Else
           TxtVchrNo = DoPad((LTrim(Str(0 + PR_VchCntr.Fields("VchrCount")) + 1)), 10)
       End If
       Command1.SetFocus
    End If
  End If
  End Sub

