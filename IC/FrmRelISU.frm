VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRelISU 
   Caption         =   "Release to GL."
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRelISU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   4770
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2700
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame3 
      Height          =   1365
      Left            =   0
      TabIndex        =   13
      Top             =   720
      Width           =   4875
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   1980
         Picture         =   "FrmRelISU.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2325
         TabIndex        =   18
         Top             =   210
         Width           =   2355
      End
      Begin VB.TextBox TxtVchrNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   7
         Top             =   585
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
         Left            =   1995
         Picture         =   "FrmRelISU.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   570
         Width           =   315
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
         Left            =   4380
         Picture         =   "FrmRelISU.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txtFrom 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Account No"
         Top             =   960
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtVchrType 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   585
         Width           =   600
         _ExtentX        =   1058
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
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   1350
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Default Currency"
         Top             =   210
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   3
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Branch Code :"
         Height          =   210
         Left            =   270
         TabIndex        =   21
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher # :"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher Type :"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   585
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Dr. Account # :"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   960
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   15
      TabIndex        =   12
      Top             =   2010
      Width           =   4860
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2610
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   630
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   225
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transaction  Type"
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "By JVS A/c"
         Height          =   330
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "By Cash A/c"
         Height          =   330
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "By Bank A/c"
         Height          =   330
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "None"
         Height          =   330
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmRelISU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PR_GlDetail As New Recordset
Dim PR_VchType As New Recordset
Dim PR_VchCntr As New Recordset
Dim PR_Branch As New Recordset
Dim lb_found As Boolean
Public PO_DESC As Object
Public PO_CODE As Object


Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtVchrType
    Set PO_DESC = Text1
    
    GoTop PR_VchType
    PR_VchType.Filter = "BranchCode = '" & txtbranchcode & "'"
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    
    If Len(txtVchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
End Sub
Private Sub Command1_Click()
Frame3.Enabled = True
Dim ln_Counter As Integer
Dim ls_VchRem  As String
ln_Counter = 1

Dim ln_RAmount As Double
Dim ln_DAmount As Double
Dim ln_TAmount As Double
Dim ln_CustAmount As Double

   'Count The Amount
 With frmIssue.GrdGRN
     For Ln_Cnt = 1 To .Rows - 1
         If Val(.TextMatrix(Ln_Cnt, 12)) > 0 Then ln_RAmount = ln_RAmount + Val(.TextMatrix(Ln_Cnt, 12))
         If Val(.TextMatrix(Ln_Cnt, 13)) > 0 Then ln_DAmount = ln_DAmount + Val(.TextMatrix(Ln_Cnt, 13))
         If Val(.TextMatrix(Ln_Cnt, 14)) > 0 Then ln_TAmount = ln_TAmount + Val(.TextMatrix(Ln_Cnt, 14))
         If Val(.TextMatrix(Ln_Cnt, 15)) > 0 Then ln_CustAmount = ln_CustAmount + Val(.TextMatrix(Ln_Cnt, 15))
     Next
     
End With
'Save voucher
If txtVchrType <> "" Then
    If ln_RAmount + ln_TAmount = ln_DAmount + ln_CustAmount Then
   ' Save References of Voucher
       ls_VchRem = "Sale against Invoice # " & frmIssue.txtTransNo & " to Customer " & frmIssue.txtJobDesc
       gc_dbcon.Execute "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks, userid,adddate,addtime,exchgrate,CrncyCode) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(frmIssue.txtvaluedate, "YYYY/MM/DD") & "','" & Format(frmIssue.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "','" & ls_VchRem & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "',0,'" & Gs_BaseCrncy & "')"
    
    ' Save Detail of Voucher
    If ln_RAmount + ln_TAmount > 0 And ln_DAmount + ln_CustAmount > 0 Then
             If ln_CustAmount > 0 Then
                ls_VchRem = Trim(IIf(MySeek(frmIssue.ls_CustAccount, "AccountNo", PR_GlDetail), PR_GlDetail.Fields("Acct_Desc"), "Not Found."))
                ln_Counter = 1
                gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & frmIssue.ls_CustAccount & "'," & ln_Counter & ",'" & Format(frmIssue.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "'," & ln_CustAmount & ",0,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
             End If
             If ln_DAmount > 0 Then
                ls_VchRem = Trim(IIf(MySeek(frmIssue.ls_DAccount, "AccountNo", PR_GlDetail), PR_GlDetail.Fields("Acct_Desc"), "Not Found."))
                ln_Counter = ln_Counter + 1
                gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & frmIssue.ls_DAccount & "'," & ln_Counter & ",'" & Format(frmIssue.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "'," & ln_DAmount & ",0,'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
             End If
             If ln_RAmount > 0 Then
                ls_VchRem = Trim(IIf(MySeek(frmIssue.ls_RAccount, "AccountNo", PR_GlDetail), PR_GlDetail.Fields("Acct_Desc"), "Not Found."))
                ln_Counter = ln_Counter + 1
                gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & frmIssue.ls_RAccount & "'," & ln_Counter & ",'" & Format(frmIssue.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "',0," & ln_RAmount & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
             End If
             If ln_TAmount > 0 Then
                ls_VchRem = Trim(IIf(MySeek(frmIssue.ls_TAccount, "AccountNo", PR_GlDetail), PR_GlDetail.Fields("Acct_Desc"), "Not Found."))
                ln_Counter = ln_Counter + 1
                gc_dbcon.Execute "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & frmIssue.ls_TAccount & "'," & ln_Counter & ",'" & Format(frmIssue.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "',0," & ln_TAmount & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
             End If
             
            'Increasing Counter
            If MySeek(txtVchrType.Text, "VchrType", PR_VchType) Then
             PR_VchCntr.Filter = adFilterNone
             PR_VchCntr.Filter = "BranchCode = '" & txtbranchcode & "' and  VchrType = '" & txtVchrType & "'"
             If PR_VchType.Fields("VchrFrequency") = "1" Then
                 PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmIssue.txtvaluedate.Value)))) = PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmIssue.txtvaluedate.Value)))) + 1
             Else
                 PR_VchCntr.Fields("VchrCount") = PR_VchCntr.Fields("VchrCount") + 1
             End If
           End If
             PR_VchCntr.Update
             PR_VchCntr.Requery
    Else
        Call MsgBox("Voucher Not Enter With Zero Amount")
    End If
    Else
        Call MsgBox("Voucher Amount Not Tally")
    End If
    frmIssue.ls_VchNo = TxtVchrNo
    frmIssue.ls_VchDesc = Text1
    frmIssue.ls_VchType = txtVchrType
    Unload Me
Else
       Call MsgBox("Please Select Voucher Type")
       txtVchrType.SetFocus
End If
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  PR_GlDetail.Open "Select * from Gl_Detail where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
  PR_VchType.Open "Select * from GlVchrType where glVchrType.CompCode ='" & Gs_compcode & "' and VchrType <> '0OB' order by GlVchrType.VchrType ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_VchCntr.Open "SELECT Gl_VchrCntrs.*,BranchCode+VchrType as FindFld FROM Gl_VchrCntrs WHERE CompCode = '" & Gs_compcode & "' And VchrYear = " & Year(Gs_Fnperiod) & " ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_Branch.Open "Select * From SysBranch Where compcode = '" & Gs_compcode & "' Order By BranchCode", gc_dbcon, adOpenStatic, adLockOptimistic, 1
  txtbranchcode = Gs_BranchCode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_VchType.Close
    PR_GlDetail.Close
    PR_VchCntr.Close
    PR_Branch.Close
End Sub

Private Sub Command3_Click()
Dim ls_Sting As String

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtFrom
    Set PO_DESC = Text1
    
    ls_String = IIf(Option1.Value = True, "B", IIf(Option2.Value = True, "S", ""))
    If ls_String <> "" Then
       PR_GlDetail.Filter = "Acct_Type = '" & ls_String & "'"
    Else
       PR_GlDetail.Filter = "Acct_Type = 'G' or Acct_Type = 'D' or Acct_Type = 'C'"
    End If
    
    GoTop PR_GlDetail
    MyLookup.Caption = "Account Nos."
    MyLookup.FillGrid PR_GlDetail, "AccountNo", "Acct_Desc", Len(PR_GlDetail.Fields("AccountNo"))
    MyLookup.Show 1
    
    If Len(txtFrom) > 0 Then txtFrom_KeyDown vbKeyReturn, vbKeyShift
    PR_GlDetail.Filter = adFilterNone
End Sub

Private Sub Option4_Click()
   Form_Unload False
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
Private Sub txtbranchcode_KeyDown(KeyCode As Integer, Shift As Integer)
  If Lastkey(KeyCode) And txtbranchcode <> "" Then
     txtbranchcode = DoPad(txtbranchcode, txtbranchcode.MaxLength)
     
     If Not MySeek(txtbranchcode, "Branchcode", PR_Branch) Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtbranchcode.SetFocus
     Else
        Text2 = PR_Branch("BranchDesc")
        txtVchrType.SetFocus
     End If
  ElseIf KeyCode = vbKeyF12 Then
     Command5_Click
 End If
End Sub

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
lb_found = False

 If Lastkey(KeyCode) Then
    txtVchrType = UCase(txtVchrType)
    lb_found = MySeek(txtVchrType.Text, "VchrType", PR_VchType)
    If Not lb_found Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtVchrType.SetFocus
    Else
       Text1 = PR_VchType.Fields("VchrDescrip")
       PR_VchCntr.Filter = "BranchCode = '" & Gs_BranchCode & "' and  VchrType = '" & txtVchrType & "'"
       If PR_VchType.Fields("VchrFrequency") = "1" Then
           TxtVchrNo = DoPad((LTrim(Str(0 + PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmIssue.txtvaluedate.Value))))) + 1)), 10)
       Else
           TxtVchrNo = DoPad((LTrim(Str(0 + PR_VchCntr.Fields("VchrCount")) + 1)), 10)
       End If
       Command1.SetFocus
    End If
  End If
  End Sub

