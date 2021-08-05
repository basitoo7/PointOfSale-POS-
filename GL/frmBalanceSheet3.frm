VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbalancesheet3 
   Caption         =   "Balance Sheet SubNotes Setup"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "frmBalanceSheet3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   30
      TabIndex        =   7
      Top             =   615
      Width           =   5205
      Begin VB.CommandButton Command1 
         Height          =   315
         Left            =   2715
         Picture         =   "frmBalanceSheet3.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1650
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup 
         Height          =   315
         Left            =   1995
         Picture         =   "frmBalanceSheet3.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   885
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup1 
         Height          =   315
         Left            =   1980
         Picture         =   "frmBalanceSheet3.frx":05EE
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   510
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   1890
         Picture         =   "frmBalanceSheet3.frx":0760
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   315
      End
      Begin VB.TextBox txtbnsdesc 
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
         TabIndex        =   3
         ToolTipText     =   "Account No"
         Top             =   1275
         Width           =   3690
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Add To Grid"
         Height          =   360
         Left            =   3885
         TabIndex        =   5
         Top             =   1650
         Width           =   1125
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1125
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox txtbnscode 
         Height          =   300
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   2
         Tag             =   "SKIPN"
         Top             =   900
         Width           =   600
      End
      Begin VB.TextBox txtbndesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2325
         MaxLength       =   64
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   510
         Width           =   2730
      End
      Begin VB.TextBox txtbncode 
         Height          =   300
         Left            =   1380
         MaxLength       =   4
         TabIndex        =   1
         Tag             =   "SKIPN"
         Top             =   525
         Width           =   600
      End
      Begin VB.TextBox txtbcode 
         Height          =   300
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   0
         Tag             =   "SKIPN"
         Top             =   150
         Width           =   510
      End
      Begin VB.TextBox txtbdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2235
         MaxLength       =   64
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "SKIPN"
         Top             =   135
         Width           =   2820
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
         Left            =   1365
         MaxLength       =   13
         TabIndex        =   4
         ToolTipText     =   "Account No"
         Top             =   1665
         Width           =   1335
      End
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
         Left            =   1365
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   9
         Tag             =   "SKIP"
         Top             =   2055
         Width           =   3660
      End
      Begin MSFlexGridLib.MSFlexGrid grdVoucher 
         Height          =   1815
         Left            =   75
         TabIndex        =   8
         Top             =   2445
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   1
         BackColorFixed  =   -2147483637
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   1335
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "BS Item Code:"
         Height          =   195
         Left            =   330
         TabIndex        =   15
         Top             =   930
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "BS Note Code:"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   540
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BS Code:"
         Height          =   195
         Left            =   690
         TabIndex        =   11
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Account no :"
         Height          =   195
         Left            =   405
         TabIndex        =   10
         Top             =   1710
         Width           =   915
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   1005
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
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
            Caption         =   "Refresh"
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
               Picture         =   "frmBalanceSheet3.frx":08D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet3.frx":0D26
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet3.frx":117A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet3.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet3.frx":1A22
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet3.frx":1E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmBalanceSheet3.frx":25CA
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmbalancesheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls0 As Boolean
Dim Mode As String
Public PO_CODE As Object
Public PO_DESC As Object
Dim PR_GlBS01 As Recordset
Dim PR_GLBS02 As Recordset
Dim PR_GLBS03 As Recordset
Dim pr_dumy As New Recordset
Dim pr_dumy1 As New Recordset
Dim PI_CurRow    As Integer
Dim PI_SrNo      As Integer
Dim PS_RowClicked As String

Private Sub Command2_Click()

End Sub

Private Sub grdVoucher_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call grdVoucher_DblClick
   If KeyCode = vbKeyDelete Then
       With grdVoucher
          If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             .RemoveItem .Row
             If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
             End If
       End With
   End If
End Sub



Private Sub cmdLookup_Click()
Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbnscode
    Set PO_DESC = Text1
    GoTop PR_GLBS03
    PR_GLBS03.Filter = "Bcode = '" & txtbcode & "' and bncode = '" & txtbncode & "'"
    
    MyLookup.Caption = "Balance Sheet Sub Notes"
    MyLookup.FillGrid PR_GLBS03, "BniCODE", "BniDESC", 5
    MyLookup.Show 1
    PR_GLBS03.Filter = adFilterNone
    If Len(txtbcode) > 0 Then
        txtbnscode_Validate False
        SendKeys vbTab
    End If
End Sub

Private Sub cmdLookup0_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbcode
    Set PO_DESC = txtbdesc
    GoTop PR_GlBS01
    MyLookup.Caption = "Balance Sheet Main Head"
    MyLookup.FillGrid PR_GlBS01, "BCODE", "BDESC", 5
    MyLookup.Show 1
    
    If Len(txtbcode) > 0 Then
        txtbcode_Validate False
        SendKeys vbTab
    End If

End Sub

Private Sub cmdLookup1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtbncode
    Set PO_DESC = txtbndesc
    PR_GLBS02.Filter = "Bcode = '" & txtbcode & "'"
    GoTop PR_GLBS02
    MyLookup.Caption = "Balance Sheet Notes"
    MyLookup.FillGrid PR_GLBS02, "BnCODE", "BnDESC", 5
    MyLookup.Show 1
    PR_GLBS02.Filter = adFilterNone
    If Len(txtbncode) > 0 Then
        txtbncode_Validate False
        SendKeys vbTab
    End If

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
    
    If Len(txtAccountNo) > 0 Then
        TxtAccountNo_Validate False
        SendKeys vbTab
    End If


    
End Sub

Private Sub Command9_Click()

If txtAccountNo <> "" Then
    Call AddGrid
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub
Private Sub AddGrid()
Dim ln_cnt As Integer

        If PS_RowClicked1 = "" Then
            If PI_SrNo = 0 Then
                PI_SrNo = 1
            Else
                PI_SrNo = PI_SrNo + 1
            End If
        End If
            With grdVoucher
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
                .TextMatrix(.Row, 1) = txtAccountNo
                .TextMatrix(.Row, 2) = txtaccountdesc
                txtAccountNo = ""
                txtaccountdesc = ""
                PS_RowClicked1 = ""
                txtAccountNo.SetFocus
            End With
         
End Sub

Private Sub Form_Load()
  
  SetToolBar(1) = chkRights("GLFRM30001")
  SetToolBar(2) = chkRights("GLFRM30002")
  SetToolBar(3) = chkRights("GLFRM30003")
  SetToolBar(4) = chkRights("GLFRM30004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  
  Set PR_GlBS01 = New Recordset
  Set PR_GLBS02 = New Recordset
  Set PR_GLBS03 = New Recordset

  PR_GlBS01.Open "Select Gl_BSheet1.* from Gl_BSheet1 where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GLBS02.Open "Select Gl_BSheet2.* from Gl_BSheet2 where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
  PR_GLBS03.Open "Select Gl_BSheet3.* from Gl_BSheet3 where compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkGls0 = IIf(PR_GLBS02.EOF, True, False)
  
  Call InitializeGrid
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GLBS03.Close
    PR_GLBS02.Close
    PR_GlBS01.Close
End Sub







Public Sub InitializeGrid()
    With grdVoucher
        .Redraw = False
        .Clear
        .Rows = 2
        .Cols = 2
        .FormatString = "Sr# |<Account No|<Account Narration"
        .ColWidth(1) = 1500 + 450
        .ColWidth(2) = 2500
        .Redraw = True
    End With
    PI_SrNo = 0
    
End Sub
Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then Call grdVoucher_DblClick
   If KeyCode = vbKeyDelete Then
       With Grid1
          If .Row = 1 And Not .Rows > 2 Then .Rows = .Rows + 1
             .RemoveItem .Row
             If .Rows = 2 And .TextMatrix(.Row, 1) = "" Then
                .TextMatrix(.Row, 0) = ""
                PI_SrNo = 0
             End If
       End With
   End If
End Sub


Private Sub grdVoucher_DblClick()
  With grdVoucher
        If .Row > 0 Then
            PI_CurRow = .Row
        End If
       txtAccountNo = .TextMatrix(.Row, 1)
       txtaccountdesc = .TextMatrix(.Row, 2)
       PS_RowClicked = "Y"
       txtAccountNo.SetFocus
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If PB_BlnkGls0 And Range(Button.Index, 2, 3) Then
       MsgBox "Data not found :", vbCritical
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GLBS03, Me, txtbcode, txtbdesc, "X", "CompCount", 3, "Acct_sub0", "Acct_Desc", 1, False, Toolbar1)
        If Mode = "A" Then
            cmdLookup.Enabled = False
            txtbnscode.Enabled = False
        Else
            cmdLookup.Enabled = True
            txtbnscode.Enabled = True
        End If
    End If
End Sub

Public Function ChkInputs() As Boolean
    If Len(txtbcode.Text) = txtbcode.MaxLength And Len(RTrim(txtbnsdesc)) > 0 And Len(txtbncode.Text) = txtbncode.MaxLength Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Public Sub FrmRefresh()
   PR_GlBS01.Requery
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim ls_accNature As String

PB_BlnkGls0 = False
gc_dbcon.BeginTrans
    
    If Mode = "A" Then
                pr_dumy.Open "select max(bnicode) as bncode from Gl_BSheet3 where BCODE = '" & txtbcode.Text & "' and bncode = '" & txtbncode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If Not pr_dumy.EOF Then
                    txtbnscode = DoPad(Trim(str(Val(0 & pr_dumy("bncode")) + 1)), txtbnscode.MaxLength)
                Else
                    txtbnscode = DoPad(Trim(str(1)), txtbnscode.MaxLength)
                End If
                pr_dumy.Close
    End If
     
     
     Select Case Mode
           Case "D"
                 gc_dbcon.Execute "DELETE FROM Gl_BSheet3 WHERE BCODE = '" & txtbcode & "' AND BNCODE = '" & txtbncode & "' AND BNICODE = '" & txtbnscode & "' and compcode = '" & Gs_compcode & "' "
                 gc_dbcon.Execute "DELETE FROM Gl_BSheet3DETAIL WHERE BCODE = '" & txtbcode & "' AND BNCODE = '" & txtbncode & "' AND BNICODE = '" & txtbnscode & "' and compcode = '" & Gs_compcode & "'"
              
           Case Else
                If Mode = "E" Then
                    gc_dbcon.Execute "DELETE FROM Gl_BSheet3 WHERE BCODE = '" & txtbcode & "' AND BNCODE = '" & txtbncode & "' AND BNICODE = '" & txtbnscode & "' and compcode = '" & Gs_compcode & "'"
                    gc_dbcon.Execute "DELETE FROM Gl_BSheet3DETAIL WHERE BCODE = '" & txtbcode & "' AND BNCODE = '" & txtbncode & "' AND BNICODE = '" & txtbnscode & "' and compcode = '" & Gs_compcode & "'"
                End If
                
                gc_dbcon.Execute "INSERT INTO  Gl_BSheet3 (Compcode,BCODE, BNCODE, BNICODE,BNIDESC) VALUES('" & Gs_compcode & "', '" & txtbcode & "','" & txtbncode & "','" & txtbnscode & "','" & txtbnsdesc & "')"
                
                
                With grdVoucher
                       For ln_cnt = 1 To .Rows - 1
                          If Len(Trim(.TextMatrix(ln_cnt, 1))) > 0 Then
                           gc_dbcon.Execute "INSERT INTO  Gl_BSheet3DETAIL (CompCode,BCODE, BNCODE, BNICODE,accountno) VALUES('" & Gs_compcode & "','" & txtbcode & "','" & txtbncode & "','" & txtbnscode & "','" & .TextMatrix(ln_cnt, 1) & "')"
                          End If
                       Next
                End With

            
            
     End Select
gc_dbcon.CommitTrans
PR_GLBS03.Requery

Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub


Sub SetVal()
    txtdesc = Trim(PR_GLBS03("BNIDESC"))
    Call LoadTrans
End Sub
Private Sub LoadTrans()
Call InitializeGrid
 pr_dumy.Open "Select * from  Gl_BSheet3DETAIL where bcode = '" & txtbcode & "' and  bncode = '" & txtbncode & "' and  bnicode = '" & txtbnscode & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
    If Not pr_dumy.EOF Then
        With grdVoucher
            Do While Not pr_dumy.EOF
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                 PI_SrNo = Val(.TextMatrix(.Row, 0))
                .TextMatrix(.Row, 1) = pr_dumy("accountno")
                 pr_dumy1.Open "Select * from gl_detail where accountno = '" & .TextMatrix(.Row, 1) & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                 If Not pr_dumy1.EOF Then
                  .TextMatrix(.Row, 2) = pr_dumy1("acct_desc")
                 End If
                pr_dumy1.Close
                
                
                .Rows = .Rows + 1
                
                pr_dumy.MoveNext
                If pr_dumy.EOF Then Exit Do
             Loop
            If .TextMatrix(.Rows - 1, 1) = "" Then .RemoveItem .Rows - 1
        End With
    End If
 pr_dumy.Close
End Sub
Private Sub TxtAccountNo_Validate(Cancel As Boolean)
Dim lb_found As Boolean

    If Trim(txtAccountNo) <> "" Then
        pr_dumy.Open "Select * from gl_detail where accountno = '" & txtAccountNo & "' and compcode = '" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
        If pr_dumy.EOF Then
            Call MsgBox("Account No not found", vbCritical)
            txtAccountNo = ""
            txtaccountdesc = ""
        Else
            txtaccountdesc = pr_dumy("acct_desc")
            Call AddGrid
        End If
        pr_dumy.Close
    Else
        txtAccountNo = ""
        txtaccountdesc = ""
    End If

End Sub


Private Sub txtbnscode_Validate(Cancel As Boolean)
Dim lb_found As Boolean

    If Trim(txtbnscode) <> "" And Mode <> "A" Then
        txtbnscode = DoPad(txtbnscode, txtbnscode.MaxLength)
        PR_GLBS03.Filter = "Bcode = '" & txtbcode & "' and bncode = '" & txtbncode & "'"
        lb_found = MySeek(txtbnscode, "bniCODE", PR_GLBS03)
       
        If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            'Cancel = True
            txtbnscode = ""
            txtbnsdesc = ""
            
        Else
            txtbnsdesc = PR_GLBS03("bniDESC")
            Call SetVal
            txtbncode.SetFocus
            PR_GLBS03.Filter = adFilterNone
        End If
    End If

End Sub

Private Sub txtbcode_Validate(Cancel As Boolean)
Dim lb_found As Boolean

    If Trim(txtbcode) <> "" Then
        txtbcode = DoPad(txtbcode, txtbcode.MaxLength)
        lb_found = MySeek(txtbcode.Text, "bCODE", PR_GlBS01)
       
        If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            'Cancel = True
        Else
            txtbdesc = PR_GlBS01("bDESC")
            'txtbncode.SetFocus
           
        End If
    Else
        txtbcode = ""
        txtbdesc = ""
    End If

End Sub
Private Sub txtbncode_Validate(Cancel As Boolean)
Dim lb_found As Boolean

    If Trim(txtbncode) <> "" Then
        txtbncode = DoPad(txtbncode, txtbncode.MaxLength)
     PR_GLBS02.Filter = "Bcode = '" & txtbcode & "'"
        lb_found = MySeek(txtbncode.Text, "bnCODE", PR_GLBS02)
       
        If Not lb_found Then
            Call SetErr(Gs_RecNFMsg, vbCritical)
            'Cancel = True
        Else
            If Mode = "A" Then
                pr_dumy.Open "select max(bnicode) as bncode from Gl_BSheet3 where BCODE = '" & txtbcode.Text & "' and bncode = '" & txtbncode & "' and compcode = '" & Gs_compcode & "' ", gc_dbcon, adOpenStatic, adLockReadOnly, 1
                If Not pr_dumy.EOF Then
                    txtbnscode = DoPad(Trim(str(Val(0 & pr_dumy("bncode")) + 1)), txtbnscode.MaxLength)
                Else
                    txtbnscode = DoPad(Trim(str(1)), txtbnscode.MaxLength)
                End If
                pr_dumy.Close
                
                cmdLookup.Enabled = False
                txtbnscode.Enabled = False
                
            End If
            txtbndesc = PR_GLBS02("bndesc")
        End If
        PR_GLBS02.Filter = adFilterNone
    Else
        txtbncode = ""
        txtbndesc = ""
    End If

End Sub

