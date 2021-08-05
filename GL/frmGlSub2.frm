VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGlSub2 
   Caption         =   "Sub Ledger Accounts"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmGlSub2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsub1 
      Height          =   330
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   10
      Tag             =   "SKIPN"
      Top             =   1155
      Width           =   435
   End
   Begin VB.TextBox txtsub0 
      Height          =   330
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   9
      Tag             =   "SKIPN"
      Top             =   780
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   4935
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2415
         MaxLength       =   64
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   2475
      End
      Begin VB.CommandButton cmdLookup1 
         Height          =   315
         Left            =   1770
         Picture         =   "frmGlSub2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   585
         Width           =   315
      End
      Begin VB.CommandButton cmdLookup0 
         Height          =   315
         Left            =   2085
         Picture         =   "frmGlSub2.frx":047C
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "SKIP"
         Top             =   210
         Width           =   315
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   975
         Width           =   3540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Code :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Detail Code :"
         Height          =   195
         Left            =   330
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Description :"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   990
         Width           =   1005
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4950
      _ExtentX        =   8731
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
               Picture         =   "frmGlSub2.frx":05EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub2.frx":0A42
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub2.frx":0E96
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub2.frx":12EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub2.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub2.frx":1B92
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmGlSub2.frx":22E6
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmGlSub2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PB_BlnkGls1 As Boolean
Dim ls_PFields As String
Dim ls_TFields As String
Dim ls_CurrAlia As String
Dim ls_PrvAlia As String
Public Mode As String

Public PO_CODE As Object
Public PO_DESC As Object

Dim PR_GlSub0 As Recordset
Dim PR_GLSUB1 As Recordset
Dim pr_dumy As New Recordset



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF11 Then
    Mode = DentMode(Mode, 4, PR_GLSUB1, frmGlSub1, txtsub1, txtsub1, "X", "CompCount", 3, "Acct_sub1", "Acct_Desc", 1, False, Toolbar1)
End If
End Sub

Private Sub Form_Load()
' Setting up Preveliges
  
  SetToolBar(1) = chkRights("GLSUB10001")
  SetToolBar(2) = chkRights("GLSUB10002")
  SetToolBar(3) = chkRights("GLSUB10003")
  SetToolBar(4) = chkRights("GLSUB10004")
  
  Toolbar1.Buttons(1).Enabled = SetToolBar(1)
  Toolbar1.Buttons(2).Enabled = SetToolBar(2)
  Toolbar1.Buttons(3).Enabled = SetToolBar(3)
  Toolbar1.Buttons(5).Enabled = SetToolBar(4)
  
  Set PR_GlSub0 = New Recordset
  Set PR_GLSUB1 = New Recordset
  
  PR_GlSub0.Open "Select *,acct_sub0+acct_sub1 As MastAcctNo from gl_sub1 where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
  PR_GLSUB1.Open "Select *, acct_sub2 AcctNo from gl_sub2 where compcode ='" & Gs_compcode & "'", gc_dbcon, adOpenDynamic, adLockOptimistic, 1
   
  PB_BlnkGls1 = IIf(PR_GLSUB1.EOF, True, False)


End Sub

Private Sub Form_Unload(Cancel As Integer)
    PR_GlSub0.Close
    PR_GLSUB1.Close
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
   If Lastkey(KeyCode) Then txtdesc = StrConv(txtdesc, vbProperCase)
End Sub


Private Sub txtsub0_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Len(txtsub0.Text) > 0 And Val(txtsub0.Text) <> 0 Then
         txtsub0 = DoPad(txtsub0, txtsub0.MaxLength)
         lb_found = MySeek(txtsub0.Text, "MastAcctNo", PR_GlSub0)
        
         If Not lb_found Then
             Call SetErr(Gs_RecNFMsg, vbCritical)
             txtsub0.SetFocus
             txtsub0.Text = ""
         Else
             
             
             
             Text1.Text = PR_GlSub0("Acct_Desc")
            'next counter
             If Mode = "A" Then
                txtsub1 = maxacccountno
             End If
             If Mode = "A" Then txtdesc.SetFocus Else txtsub1.SetFocus
         End If
  ElseIf KeyCode = vbKeyF12 Then
       cmdLookup0_Click
 End If
End Sub
Function maxacccountno() As String
pr_dumy.Open "select max(acct_sub2) as acctno from gl_sub2 where compcode = '" & Gs_compcode & "' and acct_sub1 = '" & txtsub0 & "'", gc_dbcon, adOpenStatic, adLockReadOnly, 1
            If Not pr_dumy.EOF Then
                maxacccountno = DoPad(Val(0 & pr_dumy("acctno")) + 1, txtsub1.MaxLength)
            Else
                maxacccountno = DoPad("1", txtsub1.MaxLength)
            End If
pr_dumy.Close
End Function
Private Sub txtsub1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lb_found As Boolean

 If KeyCode = vbKeyReturn And Val(txtsub1.Text) > 0 Then
          
         If Mode = "A" Then
            txtsub1.Text = DoPad(UCase(txtsub1.Text), txtsub1.MaxLength)
         End If
         PR_GLSUB1.Filter = "acct_sub1 = '" & txtsub0 & "'"
         lb_found = MySeek(txtsub1.Text, "AcctNo", PR_GLSUB1)
         
       Select Case Mode
            Case "A"
                If lb_found Then
                   Call SetErr(Gs_RecFdMsg, vbCritical)
                   'Cancel = True
                   SetClear Me
                   txtsub0.SetFocus
                Else
                 If txtdesc.Enabled Then txtdesc.SetFocus
                End If
            Case Else
                If Not lb_found Then
                   Call SetErr(Gs_RecNFMsg, vbCritical)
                   'Cancel = True
                    SetClear Me
                   txtsub1.Enabled = True
                   txtsub1.SetFocus
                Else
                   Call SetVal
                   If Mode <> "D" Then
                      txtsub1.Enabled = True
                      txtdesc.SetFocus
                   End If
                End If
            End Select
            PR_GLSUB1.Filter = adFilterNone
    ElseIf KeyCode = vbKeyF12 And cmdLookup1.Enabled Then
       cmdLookup1_Click
    End If
  End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Index = 1 Then
         cmdLookup1.Enabled = False
         cmdLookup0.Enabled = True
         txtsub0.Enabled = True
         SetClear (Me)
    ElseIf Range(Button.Index, 2, 3) Then
         cmdLookup1.Enabled = True
         cmdLookup0.Enabled = True
         txtsub0.Enabled = True
         SetClear (Me)
         txtsub1 = ""
         txtsub0.SetFocus
         
    End If
    
    If PB_BlnkGls1 And Range(Button.Index, 2, 3) Then
       Call SetErr("Data not found.", vbCritical)
       Mode = ""
       'Cancel = True
    Else
       Mode = DentMode(Mode, Button.Index, PR_GLSUB1, Me, txtsub0, txtsub1, "X", "X", 3, "X", "X", 1, False, Toolbar1)
    End If
    
End Sub

Public Sub SaveValues()
On Error GoTo LocalErr
Dim cntsql As New ADODB.Command
Dim Appfields As String
Dim ReplFields As String
PB_BlnkGls1 = False



gc_dbcon.BeginTrans
     Select Case Mode
           Case "A"
              gc_dbcon.Execute "INSERT into gl_sub2 (compcode, Acct_Sub1, Acct_Sub2, Acct_Desc, UserId, AddDate, AddTime) VALUES ('" & Gs_compcode & "', '" & txtsub0 & "','" & txtsub1 & "' ,'" & txtdesc.Text & "','" & Gc_UserId & "','" & Format(Date, "YYYY/MM/DD") & "','" & Time & "')"
              
           Case "E"
              gc_dbcon.Execute "UPDATE gl_sub2 SET Acct_Desc= '" & txtdesc.Text & "' WHERE  compcode = '" & Gs_compcode & "' and acct_sub1= '" & txtsub0.Text & "' and acct_sub2= '" & txtsub1.Text & "'"
              
           Case "D"
              gc_dbcon.Execute "delete from gl_sub2  WHERE  compcode = '" & Gs_compcode & "' and acct_sub1= '" & txtsub0.Text & "' and acct_sub2= '" & txtsub1.Text & "'"
              
     End Select
gc_dbcon.CommitTrans
     PR_GLSUB1.Requery
     If Mode = "A" Then
       txtsub1 = maxacccountno
     End If
     If Mode = "A" Then
        txtdesc.SetFocus
     Else
        txtsub1.SetFocus
        txtsub1 = ""
     End If
     
     Exit Sub

LocalErr:
gc_dbcon.RollbackTrans
Call SetErr("Processing Error, Operation Terminated. Call System Administrator,", vbCritical)
On Error GoTo 0
End Sub

Private Sub SetVal()
     txtdesc = Trim(PR_GLSUB1("Acct_Desc"))
End Sub
Public Function ChkInputs() As Boolean
    If txtsub0.Text <> "" And txtsub1.Text <> "" And RTrim(txtdesc.Text) <> "" Then
       ChkInputs = True
    Else
       Call SetErr(Gs_InvldMsg, vbCritical)
       ChkInputs = False
    End If
End Function

Private Sub txtsub1_LostFocus()
   Call txtsub1_KeyDown(vbKeyReturn, vbKeyShift)
End Sub

Private Sub cmdLookup0_Click()

    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub0
    Set PO_DESC = Text1
    Gs_SQL = "Select acct_sub0+acct_sub1 'Account No', Acct_Desc  'Description' from gl_sub1"
    Gs_FindFld = "Acct_Desc"
    Gs_Subon = False
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "'"
    Gs_OrderBy = "Order by Acct_Desc,acct_sub0+acct_sub1"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    
    If Len(txtsub0) > 0 Then txtsub0_KeyDown vbKeyReturn, vbKeyShift

End Sub

Private Sub cmdLookup1_Click()
    
    Set PO_AnyForm = Nothing
    Set PO_AnyForm = Me
    Set PO_CODE = txtsub1
    Set PO_DESC = txtdesc
    Gs_SQL = "Select acct_sub2 'Account No', Acct_Desc  'Description' from gl_sub2"
    Gs_FindFld = "Acct_Desc"
    Gs_Subon = False
    Gs_OtherPara = " Where Compcode = '" & Gs_compcode & "' and acct_sub1 = '" & txtsub0 & "' "
    Gs_OrderBy = "Order by Acct_Desc,acct_sub2"
    MyLookupOLDB.Caption = "Account Nos."
    MyLookupOLDB.Show 1
    
    If Len(txtsub1) > 0 Then txtsub1_KeyDown vbKeyReturn, vbKeyShift
End Sub

Public Sub FrmRefresh()
  PR_GlSub0.Requery
  PR_GLSUB1.Requery
End Sub
