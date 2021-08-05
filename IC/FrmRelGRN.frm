VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmRelGRN 
   Caption         =   "Release to GL."
   ClientHeight    =   2835
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
   Icon            =   "FrmRelGRN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   4950
      MaxLength       =   50
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1980
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame3 
      Height          =   1350
      Left            =   60
      TabIndex        =   13
      Top             =   720
      Width           =   4815
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   2355
         TabIndex        =   21
         Top             =   195
         Width           =   2340
      End
      Begin VB.CommandButton Command5 
         Height          =   315
         Left            =   2010
         Picture         =   "FrmRelGRN.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   195
         Width           =   315
      End
      Begin VB.TextBox TxtVchrNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   7
         Top             =   570
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
         Picture         =   "FrmRelGRN.frx":047C
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
         Left            =   4395
         Picture         =   "FrmRelGRN.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   930
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
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Account No"
         Top             =   930
         Width           =   3015
      End
      Begin MSMask.MaskEdBox txtVchrType 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Top             =   570
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
      Begin MSMask.MaskEdBox txtbranchcode 
         Height          =   315
         Left            =   1380
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Default Currency"
         Top             =   195
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
         Left            =   300
         TabIndex        =   20
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Voucher # :"
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Voucher Type :"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cr. Account # :"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   930
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   75
      TabIndex        =   12
      Top             =   2025
      Width           =   4830
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   375
         Left            =   315
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
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
Attribute VB_Name = "FrmRelGRN"
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
    PR_VchType.Filter = "Branchcode = '" & txtbranchcode & "'"
    MyLookup.Caption = "Voucher Types"
    MyLookup.FillGrid PR_VchType, "VchrType", "VchrDescrip", 5
    MyLookup.Show 1
    PR_VchCntr.Filter = adFilterNone
    If Len(txtVchrType) > 0 Then txtVchrType_KeyDown vbKeyReturn, vbKeyShift
End Sub

Private Sub Command1_Click()
Frame3.Enabled = True
Dim Ln_Cnt As Integer
Dim ln_Counter As Integer
Dim ls_VchRem As String
Dim cntsql As New ADODB.Command
cntsql.ActiveConnection = gc_dbcon
cntsql.CommandType = adCmdText
ln_Counter = 1
                 
  ' Save Debit Details of Voucher
   With FrmGRN.GrdGRN
     For Ln_Cnt = 1 To .Rows - 1
         If .TextMatrix(Ln_Cnt, 9) <> "" Then
           If MySeek(.TextMatrix(Ln_Cnt, 9), "AccountNo", PR_GlDetail) Then
             ls_VchRem = PR_GlDetail.Fields("Acct_Desc")
           Else
               Call SetErr("GL account not found Voucher not posted", vbCritical)
               Unload Me
               Exit Sub
           End If
           cntsql.CommandText = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & .TextMatrix(Ln_Cnt, 9) & "'," & Ln_Cnt & ",'" & Format(FrmGRN.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "'," & Val(.TextMatrix(.Row, 5)) & "," & 0 & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
           cntsql.Execute
           ln_Counter = ln_Counter + 1
         End If
     Next
   End With
   
  If ln_Counter > 1 Then
' Save Credit Details of Voucher
   ls_VchRem = Trim(IIf(MySeek(txtFrom.Text, "AccountNo", PR_GlDetail), PR_GlDetail.Fields("Acct_Desc"), "Not Found."))
   cntsql.CommandText = "INSERT into Gl_Trans(compcode,BranchCode,AccountNo, SerialNo, Value_Date, Voucher_No, VchrType, Dr_Amount, Cr_Amount, userid,adddate,addtime,Acct_Nirration,AcctName) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & txtFrom.Text & "'," & ln_Counter & ",'" & Format(FrmGRN.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "'," & 0 & "," & Val(FrmGRN.TxtGrnTotal.Text) & ",'" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "','" & ls_VchRem & "','" & ls_VchRem & "')"
   cntsql.Execute
 
 ' Save References of Voucher
   cntsql.CommandText = "INSERT into Gl_Ref(compcode,BranchCode,Value_Date,Trans_Date, Voucher_No, VchrType, Vchr_Remarks, userid,adddate,addtime,Exchgrate,CrncyCode) VALUES ('" & Gs_compcode & "','" & Gs_BranchCode & "','" & Format(FrmGRN.txtvaluedate, "YYYY/MM/DD") & "','" & Format(FrmGRN.txtvaluedate, "YYYY/MM/DD") & "','" & TxtVchrNo & "','" & txtVchrType & "','Goods Received.','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "',0,'" & Gs_BaseCrncy & "')"
   cntsql.Execute
       If MySeek(txtVchrType.Text, "VchrType", PR_VchType) Then
        PR_VchCntr.Filter = "Branchcode = '" & txtbranchcode & "' and  VchrType = '" & txtVchrType & "'"
        If PR_VchType.Fields("VchrFrequency") = "1" Then
            PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(FrmGRN.txtvaluedate.Value)))) = PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(FrmGRN.txtvaluedate.Value)))) + 1
        Else
            PR_VchCntr.Fields("VchrCount") = PR_VchCntr.Fields("VchrCount") + 1
        End If
        PR_VchCntr.Update
        PR_VchCntr.Requery
       End If
    End If
FrmGRN.ls_VchDesc = Text1
FrmGRN.ls_VchNo = TxtVchrNo
FrmGRN.ls_VchType = txtVchrType
Unload Me
End Sub

Private Sub Command2_Click()
   Unload Me
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



Private Sub Form_Load()
  PR_GlDetail.Open "Select * from Gl_Detail where CompCode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
  PR_VchType.Open "Select * from GlVchrType where glVchrType.CompCode ='" & Gs_compcode & "' and VchrType <> '0OB' order by GlVchrType.VchrType ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_VchCntr.Open "SELECT * FROM Gl_VchrCntrs WHERE CompCode = '" & Gs_compcode & "' And VchrYear = " & Year(Gs_Fnperiod) & " ", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
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
    If Not PR_GlDetail.EOF Then
        MyLookup.Caption = "Account Nos."
        MyLookup.FillGrid PR_GlDetail, "AccountNo", "Acct_Desc", Len(PR_GlDetail.Fields("AccountNo"))
        MyLookup.Show 1
        
        If Len(txtFrom) > 0 Then txtFrom_KeyDown vbKeyReturn, vbKeyShift
        PR_GlDetail.Filter = adFilterNone
    Else
        Call SetErr("Account Not Found", vbCritical)
    End If
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

Private Sub txtVchrType_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If Lastkey(KeyCode) Then
    txtVchrType = UCase(txtVchrType)
    lb_found = MySeek(Trim(txtbranchcode.Text) + Trim(txtVchrType.Text), "FindFld", PR_VchType)
    If Not lb_found Then
       Call SetErr(Gs_RecNFMsg, vbCritical)
       txtVchrType.SetFocus
    Else
       FrmGRN.ls_VchDesc = PR_VchType.Fields("VchrDescrip")
       PR_VchCntr.Filter = "BranchCode = '" & Gs_BranchCode & "' and  VchrType = '" & txtVchrType & "'"
       If PR_VchType.Fields("VchrFrequency") = "1" Then
           TxtVchrNo = DoPad((LTrim(Str(0 + PR_VchCntr.Fields("VchrMonth" & LTrim(Str(Month(frmIssue.txtvaluedate.Value))))) + 1)), 10)
       Else
           TxtVchrNo = DoPad((LTrim(Str(0 + PR_VchCntr.Fields("VchrCount")) + 1)), 10)
       End If
       txtFrom.SetFocus
    
    End If
  End If
  End Sub

