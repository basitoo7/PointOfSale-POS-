VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCashflowsetup 
   Caption         =   "Cash Flow Setup"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashflowSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   30
      TabIndex        =   1
      Top             =   570
      Width           =   6015
      Begin VB.TextBox txtaccountdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2655
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   11
         Tag             =   "SKIP"
         Top             =   960
         Width           =   3285
      End
      Begin VB.TextBox TxtAccountNo 
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
         Left            =   975
         MaxLength       =   13
         TabIndex        =   10
         ToolTipText     =   "Account No"
         Top             =   975
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2325
         Picture         =   "frmCashflowSetup.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtCodedesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   210
         Width           =   4065
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFF00&
         Height          =   330
         Left            =   975
         MaxLength       =   3
         TabIndex        =   6
         Tag             =   "SKIPN"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   975
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   600
         Width           =   2025
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
         Left            =   1545
         Picture         =   "frmCashflowSetup.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   225
         Width           =   315
      End
      Begin MSFlexGridLib.MSFlexGrid GrdGRN 
         Height          =   1725
         Left            =   45
         TabIndex        =   7
         Top             =   1350
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   3043
         _Version        =   393216
         Rows            =   1
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Account no :"
         Height          =   195
         Left            =   45
         TabIndex        =   12
         Top             =   990
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   210
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   210
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   465
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6075
      _ExtentX        =   10716
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
            Caption         =   "&Listing"
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
               Picture         =   "frmCashflowSetup.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashflowSetup.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashflowSetup.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashflowSetup.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashflowSetup.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashflowSetup.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCashflowSetup.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCashflowsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkLoca As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_CashFlow As Recordset
Dim pr_dumy As New Recordset
Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String


Private Sub cmdLookup_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtcode
    Set PO_DESC = txtCodedesc
    Gs_SQL = "Select DCode as  Code, DDescription  as Description  from GL_DailyRptCode "
    Gs_FindFld = "DDescription"
    Gs_OrderBy = "Order By DDescription"
    Gs_OtherPara = " Where compcode = '" & Gs_compcode & "' group by DCode, DDescription "
    MyLookupOLDB.Caption = "Cash Flow Codes"
    MyLookupOLDB.Show 1
    If Len(txtcode) > 0 Then txtCode_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub Command1_Click()
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtAccountNo
    Set PO_DESC = txtaccountdesc
    Gs_SQL = "Select AccountNo, acct_desc  Description from gl_detail"
    Gs_FindFld = "acct_desc"
    Gs_Subon = True
    Gs_OrderBy = "Order by acct_desc,accountno"
    Gs_OtherPara = " Where compcode = '" & Gs_compcode & "'"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    If txtAccountNo <> "" Then Call TxtAccountNo_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub TxtAccountNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And Trim(txtAccountNo) <> "" Then
        Dim lb_found As Boolean
        pr_dumy.Open "Select * from gl_detail where accountno = '" & txtAccountNo & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Account No not found", vbCritical)
            txtAccountNo = ""
            txtaccountdesc = ""
        Else
            txtaccountdesc = pr_dumy("acct_desc")
            Call AddToGrid
        End If
        pr_dumy.Close
Else
   txtAccountNo = ""
   txtaccountdesc = ""
End If

End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("SRSIB00001")
  SetToolBar(2) = chkRights("SRSIB00002")
  SetToolBar(3) = chkRights("SRSIB00003")
  SetToolBar(4) = chkRights("SRSIB00004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  Set PR_CashFlow = New Recordset
   
  PR_CashFlow.Open "Select * from GL_DailyRptCode where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkLoca = IIf(PR_CashFlow.EOF, True, False)
  
  InitializeGrid

End Sub


Private Sub Form_Unload(Cancel As Integer)
    PR_CashFlow.Close
End Sub

Private Sub GrdGRN_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call GrdGRN_DblClick
   If KeyCode = vbKeyDelete Then
       With GrdGRN
          If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             .RemoveItem .Row
             If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
             End If
       End With
   End If
End Sub


Private Sub GrdGRN_DblClick()
  With GrdGRN
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
       txtAccountNo = .TextMatrix(.Row, 1)
       txtaccountdesc = .TextMatrix(.Row, 2)

       PS_RowClicked = "Y"
       txtdesc.SetFocus
    End With
End Sub


Private Sub txtDesc_Change()
If txtdesc <> "" Then
txtCodedesc = txtdesc
End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtdesc) <> "" Then
    txtAccountNo.SetFocus
ElseIf KeyCode = vbKeyReturn And Trim(txtdesc) = "" Then
    Call MsgBox("Enter Description !!!", vbCritical)
    txtdesc.SetFocus
End If
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn Then
         
         txtcode.Text = IIf(IsNumeric(txtcode.Text), DoPad(UCase(txtcode.Text), txtcode.MaxLength), UCase(txtcode.Text))
         lb_found = MySeek(txtcode.Text, "DCode", PR_CashFlow)
         Select Case Mode
         
         Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                   txtcode.SetFocus
                Else
                   txtCodedesc = PR_CashFlow("dDescription")
                   txtdesc.SetFocus
                End If
         Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                   txtcode.SetFocus
                Else
                   txtCodedesc = PR_CashFlow("dDescription")
                   txtdesc = PR_CashFlow("dDescription")
                   txtdesc.SetFocus
                   SetVal
                End If
         End Select
            
         
 End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkLoca And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_CashFlow, Me, txtcode, txtdesc, "X", "CompCount", 3, "SiteCode", "Description", 1, False, Toolbar1)
    End If
    If Button.Index = 7 Then InitializeGrid

    If Button.Index = 1 Then
    InitializeGrid
    txtcode = maxtranscode
    cmdLookup.Enabled = False
    txtdesc.SetFocus
    Else
    cmdLookup.Enabled = True
    End If
End Sub
Private Sub AddToGrid()
Dim ln_cnt As Integer
            If Trim(txtdesc) <> "" Then
                    If PS_RowClicked = "" Then
                        If PI_SrNo = 0 Then
                            PI_SrNo = 1
                        Else
                            PI_SrNo = PI_SrNo + 1
                         End If
                     End If
        
                    If Trim(txtdesc) <> "" Then
                        With GrdGRN
                            If PS_RowClicked = "" Then
                                    If Not PI_SrNo = 1 Then .Rows = .Rows + 1
                                        .Row = .Rows - 1
                                    Else
                                        .Row = PI_CurRow
                                    End If
                                    If PS_RowClicked = "" Then
                                        .TextMatrix(.Row, 0) = PI_SrNo
                                    Else
                                        .TextMatrix(.Row, 0) = PI_CurRow
                                    End If
                                   
                                   .TextMatrix(.Row, 1) = Trim(txtAccountNo)
                                   .TextMatrix(.Row, 2) = Trim(txtaccountdesc)
                                    
                                    txtaccountdesc = ""
                                    txtAccountNo = ""
                                    txtAccountNo.SetFocus
                                    
 
        
                                
                                PS_RowClicked = ""
                        End With
                     End If
                        
                    
                   
        Else
            Call SetErr("Enter bin Description", vbCritical)
            txtdesc.SetFocus
       End If
End Sub

Private Sub InitializeGrid()
    With GrdGRN
        .Redraw = False
        .Clear
        .Rows = 2
        .FormatString = "Sr#|<Account No|<Account Desc"
        .ColWidth(1) = 1500
        .ColWidth(2) = 3700
        .Redraw = True
    End With
    PI_SrNo = 0
    PI_CurRow = 0
    PS_RowClicked = ""
End Sub

Public Sub SaveValues()
'On Error GoTo LocalErr
Dim ls_sql As String
PB_BlnkLoca = False
gc_dbcon.BeginTrans
     Select Case Mode
           Case "D"
              ls_sql = "DELETE FROM GL_DailyRptCode WHERE dCode = '" & txtcode.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
            
           Case Else
              ls_sql = "DELETE FROM GL_DailyRptCode WHERE dCode = '" & txtcode.Text & "' and compcode = '" & Gs_compcode & "'"
              gc_dbcon.Execute ls_sql
            
              With GrdGRN
                    For ln_cnt = 1 To .Rows - 1
                      ls_sql = "INSERT into GL_DailyRptCode (compcode,DCode,DDescription,AccountNo) VALUES ('" & Gs_compcode & "','" & txtcode.Text & "','" & txtdesc.Text & "','" & .TextMatrix(ln_cnt, 1) & "' )"
                      gc_dbcon.Execute ls_sql
                    Next
              End With
              
              
     End Select

gc_dbcon.CommitTrans
PR_CashFlow.Requery
InitializeGrid
txtcode.Text = ""
txtCodedesc = ""
If Mode = "A" Then
    txtcode = maxtranscode
    txtdesc.SetFocus
End If

Exit Sub
LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub
Private Function maxtranscode() As String
pr_dumy.Open "select max(Dcode) as transcode from GL_DailyRptCode where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
If Not pr_dumy.EOF Then
    maxtranscode = DoPad(Trim(str(Int(0 & pr_dumy("transcode")) + 1)), 3)
Else
    maxtranscode = DoPad(Trim(str(Int(1))), 10)
End If
pr_dumy.Close
End Function

Public Sub ClearVal()
     txtcode = ""
     txtdesc = ""
End Sub

Private Sub SetVal()
On Error GoTo LocalErr

Dim pr_dumyloadtrans As New Recordset

InitializeGrid
    
    pr_dumyloadtrans.Open "select * from GL_DailyRptCode where compcode = '" & Gs_compcode & "' and DCode = '" & txtcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly
   
    If Not pr_dumyloadtrans.EOF Then
        With GrdGRN
            Do While Not pr_dumyloadtrans.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = Trim(pr_dumyloadtrans("AccountNo") & "")
                 pr_dumy.Open "Select * from gl_detail where accountno = '" & .TextMatrix(.Row, 1) & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                 If Not pr_dumy.EOF Then
                 .TextMatrix(.Row, 2) = pr_dumy("acct_desc")
                 End If
                 pr_dumy.Close

                .Rows = .Rows + 1
                pr_dumyloadtrans.MoveNext
                If pr_dumyloadtrans.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
    pr_dumyloadtrans.Close
Exit Sub
LocalErr:
Call MsgBox(Err.Description)
     
End Sub
Public Function ChkInputs() As Boolean
    If Len(txtcode.Text) = txtcode.MaxLength And PI_SrNo > 0 Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub SetFrmEnv(ls_mode As String)
    txtdesc.Enabled = IIf(ls_mode <> "D", True, False)
End Sub
