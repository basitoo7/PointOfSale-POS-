VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpettycash 
   Caption         =   "Payments"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPettyCash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   0
      TabIndex        =   7
      Top             =   570
      Width           =   6000
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   5745
         MaxLength       =   15
         TabIndex        =   21
         Top             =   540
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox txtTransNo 
         BackColor       =   &H00FFFF00&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   165
         Width           =   1095
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   2535
         Picture         =   "frmPettyCash.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   165
         Width           =   315
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1995
         Width           =   1035
      End
      Begin VB.TextBox txtremarks 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1425
         MaxLength       =   200
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1635
         Width           =   4500
      End
      Begin VB.ComboBox txttype 
         Height          =   330
         ItemData        =   "frmPettyCash.frx":047C
         Left            =   4845
         List            =   "frmPettyCash.frx":0486
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker txtvaluedate 
         Height          =   315
         Left            =   1425
         TabIndex        =   1
         Tag             =   "SKIPN"
         Top             =   540
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Format          =   24641537
         CurrentDate     =   37580
      End
      Begin VB.TextBox txtJobDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2445
         MaxLength       =   64
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1260
         Width           =   3495
      End
      Begin VB.TextBox txtJobCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1425
         MaxLength       =   6
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1275
         Width           =   660
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   2115
         Picture         =   "frmPettyCash.frx":049A
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   1260
         Width           =   315
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   2115
         Picture         =   "frmPettyCash.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   315
      End
      Begin VB.TextBox TxtPartyCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "XXX"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1425
         MaxLength       =   3
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   915
         Width           =   660
      End
      Begin VB.TextBox TxtPartyDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2445
         MaxLength       =   64
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   900
         Width           =   3480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference #  :"
         Height          =   255
         Left            =   90
         TabIndex        =   20
         Top             =   165
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount :"
         Height          =   255
         Left            =   45
         TabIndex        =   17
         Top             =   2010
         Width           =   1305
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks :"
         Height          =   255
         Left            =   45
         TabIndex        =   16
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Employee  Code :"
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Payment Type :"
         Height          =   210
         Index           =   8
         Left            =   3690
         TabIndex        =   14
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Fund Code :"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   945
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Value Date :"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   10
         ToolTipText     =   "Enter Value Date"
         Top             =   555
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
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
            Caption         =   "&Slip"
            Description     =   "Print Listing."
            Object.ToolTipText     =   "Print listing."
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
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
               Picture         =   "frmPettyCash.frx":077E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPettyCash.frx":0BD2
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPettyCash.frx":1026
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPettyCash.frx":147A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPettyCash.frx":18CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPettyCash.frx":1D22
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPettyCash.frx":2476
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmpettycash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGRN As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim ls_CodeID As String
Dim Pr_ICParty As New Recordset
Dim PR_Expense As New Recordset
Dim PR_Payments As New Recordset


Private Sub cmdLookup_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txttransno
    Set PO_DESC = Text1
    
    GoTop PR_Payments
    MyLookup.Caption = "Transactions"
    MyLookup.FillGrid PR_Payments, "TransCode", "ValueDate", 6
    MyLookup.Show 1
    If Len(txttransno) > 0 Then txtTransNo_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command3_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = TxtPartyCode
    Set PO_DESC = txtpartydesc
    
    GoTop PR_Expense
    MyLookup.Caption = "Expense Types"
    MyLookup.FillGrid PR_Expense, "ExpCode", "ExpName", 6
    MyLookup.Show 1
    If Len(TxtPartyCode) > 0 Then txtPartyCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command5_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtJobCode
    Set PO_DESC = txtJobDesc
    GoTop Pr_ICParty
    MyLookup.Caption = "Employees"
    MyLookup.FillGrid Pr_ICParty, "SupplierCode", "Description", 6
    MyLookup.Show 1
    If Len(txtJobCode) > 0 Then txtJobCode_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then Mode = DentMode(Mode, 4, PR_Payments, Me, txttransno, txttype, Para_Rs, "IC_PettyCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
End Sub

Private Sub Form_Load()
 
  SetToolBar(1) = chkRights("ICISUSTP01")
  SetToolBar(2) = chkRights("ICISUSTP02")
  SetToolBar(3) = chkRights("ICISUSTP03")
  SetToolBar(4) = chkRights("ICISUSTP04")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Pr_ICParty.Open "Select * from Ic_Supplier where Codeid = 'D' order by SupplierCode", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_Expense.Open "Select * from Ic_Expense order by expcode ", gc_dbcon, adOpenDynamic, adLockOptimistic
  PR_Payments.Open "Select *  from PettyCash order by Transcode ", gc_dbcon, adOpenDynamic, adLockOptimistic
  txtvaluedate.Value = Date
  txttype = "Issue"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Pr_ICParty.Close
  PR_Expense.Close
  PR_Payments.Close
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyPageUp Then TxtPartyCode.SetFocus
End Sub
Private Sub txtJobCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And txtJobCode.Text <> "" Then
        If txttype = "Issue" Then
            txtremarks = "Paid for " + Trim(LCase(txtpartydesc)) + " of " + Trim(LCase(txtJobDesc))
        Else
            txtremarks = "Received from  " + Trim(LCase(txtJobDesc)) + " for Petty Cash"
        End If
         txtJobCode.Text = DoPad(txtJobCode.Text, txtJobCode.MaxLength)
             If Not MySeek(txtJobCode.Text, "SupplierCode", Pr_ICParty) Then
                    Call SetErr(Gs_RecNFMsg, vbCritical)
                    txtJobCode.SetFocus
                    txtJobDesc.Text = ""
                Else
                    txtJobDesc.Text = Pr_ICParty("Description")
                    txtremarks.Enabled = True
                    txtremarks.SetFocus
             End If
 ElseIf KeyCode = vbKeyF12 Then
    Command5_Click
 End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
       Mode = DentMode(Mode, Button.Index, PR_Payments, Me, txttransno, txttype, Para_Rs, "IC_PettyCnt", 10, "txtTransNo", "text1", 0, False, Toolbar1)
End Sub

Public Sub SaveValues()
'On Error GoTo RollBack
Dim ls_sql As String
gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              ls_sql = "Insert into PettyCash(TransCode,ValueDate,PaymentType,ExpCode,Employeecode,Narration,Amount) Values ('" & txttransno & "','" & Format(txtvaluedate, "YYYY/MM/DD") & "','" & Left(txttype, 1) & "','" & TxtPartyCode & "','" & txtJobCode & "','" & txtremarks & "' ," & Val(0 & txtAmount) & ")"
              gc_dbcon.Execute ls_sql
           Case "E"
              
               ls_sql = "Update  PettyCash set PaymentType = '" & Left(txttype, 1) & "',ValueDate = '" & Format(txtvaluedate, "YYYY/MM/DD") & "',ExpCode = '" & TxtPartyCode & "',employeecode = '" & txtJobCode & "',Amount = " & Val(txtAmount) & ",Narration = '" & txtremarks & "' where TransCode = '" & txttransno & "'"
               gc_dbcon.Execute ls_sql
           Case Else
               ls_sql = "Delete from  PettyCash where TransCode = '" & txttransno & "'"
                gc_dbcon.Execute ls_sql
           End Select
gc_dbcon.CommitTrans
PR_Payments.Requery
Exit Sub

RollBack:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)

End Sub
Public Sub ClearVal()
     '
End Sub
Private Sub SetVal()
     TxtPartyCode = Trim(PR_Payments("ExpCode") & "")
     txtvaluedate = PR_Payments("ValueDate")
     If TxtPartyCode <> "" Then Call txtPartyCode_KeyDown(vbKeyReturn, vbKeyShift)
     txtAmount = Val(0 & PR_Payments("Amount"))
     txttype = IIf(Trim(PR_Payments("PaymentType") & "") = "I", "Issue", "Receipt")
     txtremarks = Trim(PR_Payments("Narration") & "")
     txtJobCode = Trim(PR_Payments("EmployeeCode") & "")
     If txtJobCode <> "" Then Call txtJobCode_KeyDown(vbKeyReturn, vbKeyShift)
End Sub
Public Function ChkInputs() As Boolean
    If Len(txttransno.Text) = txttransno.MaxLength And Len(txtJobCode) = txtJobCode.MaxLength And Len(TxtPartyCode) = TxtPartyCode.MaxLength And Val(txtAmount) > 0 And txttype <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function
Private Sub txtPartyCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean
 
 If KeyCode = vbKeyReturn And Len(TxtPartyCode.Text) > 0 Then
         TxtPartyCode.Text = IIf(IsNumeric(LTrim(RTrim(TxtPartyCode.Text))), DoPad(TxtPartyCode.Text, 3), UCase(TxtPartyCode.Text))
         lb_found = MySeek(TxtPartyCode.Text, "ExpCode", PR_Expense)
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             TxtPartyCode.SetFocus
             txtpartydesc.Text = ""
         Else
             txtpartydesc.Text = PR_Expense("expname")
             txtJobCode.SetFocus
         End If
 ElseIf KeyCode = vbKeyF12 Then
    Command3_Click
 End If
End Sub
Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtAmount.SetFocus
End Sub
Private Sub txtTransNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Len(txttransno.Text) > 0 Then
         txttransno.Text = DoPad(UCase(txttransno.Text), 10)
         If Not MySeek(txttransno.Text, "TransCode", PR_Payments) Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
                txttransno.SetFocus
         Else
               Call SetVal
               txttype.SetFocus
         End If
ElseIf KeyCode = vbKeyF12 Then
    Call cmdLookup_Click
End If
End Sub
Private Sub txttype_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtvaluedate.SetFocus
End Sub
Private Sub txtvaluedate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then TxtPartyCode.SetFocus
 If KeyCode = vbKeyPageUp Then txttype.SetFocus
End Sub
